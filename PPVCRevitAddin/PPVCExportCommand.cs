using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitView = Autodesk.Revit.DB.View;

namespace PPVCRevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class PPVCExportCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1) Ask user where to save
                string outDir = PickOutputFolder();
                if (string.IsNullOrEmpty(outDir))
                {
                    TaskDialog.Show("PPVC Export", "Export cancelled.");
                    return Result.Cancelled;
                }

                // 2) Collect nodes (via utils) - STRICT: only nodes in whitelist
                // Look for whitelist in the project directory, not the DLL directory
                string whitelistPath = Path.Combine(
                    Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                    "..",
                    "..",
                    "annotation_whitelist.csv");
                
                // Also try the bin\Debug folder directly
                if (!File.Exists(whitelistPath))
                {
                    whitelistPath = Path.Combine(
                        Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                        "annotation_whitelist.csv");
                }
                
                var whitelist = LoadAnnotationWhitelist(whitelistPath, doc.Title);
                
                // Determine mode and collect nodes accordingly
                HashSet<long> whitelistNodeIds = null;
                if (whitelist.Count > 0)
                {
                    // WHITELIST mode: we have annotation whitelists, so we need model element IDs separately
                    // For now, collect all nodes (don't filter by annotation whitelist)
                    whitelistNodeIds = null;
                }
                // If whitelist is empty, we'll use SMART mode and collect all nodes
                
                var nodes = PPVCExportUtils.CollectNodes(doc, whitelistNodeIds);

                // 3) Collect edges (via utils)
                var edges = PPVCExportUtils.CollectEdges(doc, nodes);

                // DIAGNOSTIC: Log what we found
                string diagnosticPath = Path.Combine(outDir, "whitelist_diagnostic.txt");
                string ppvcNameNoExt = Path.GetFileNameWithoutExtension(doc.Title).Trim();
                File.WriteAllText(diagnosticPath, $"PPVC Name: {doc.Title}\n" +
                    $"PPVC Name (no ext): {ppvcNameNoExt}\n" +
                    $"Whitelist Path: {whitelistPath}\n" +
                    $"Whitelist Exists: {File.Exists(whitelistPath)}\n" +
                    $"Whitelist IDs loaded: {whitelist.Count}\n" +
                    $"First 10 IDs: {string.Join(", ", whitelist.Take(10))}\n\n");

                // 5) Collect annotations
                // If whitelist is empty, use SMART mode (export all Dimensions + TextNotes)
                // If whitelist has entries, use WHITELIST mode (export only whitelisted)
                List<AnnotationRecord> anns;
                string modeStr;
                if (whitelist.Count == 0)
                {
                    // Smart mode: export ALL Dimensions and TextNotes
                    anns = CollectAnnotationsSmartMode(doc, outDir);
                    modeStr = "SMART (all Dims+TextNotes)";
                }
                else
                {
                    // Whitelist mode: export only whitelisted IDs
                    anns = CollectAnnotations(doc, whitelist);
                    modeStr = "WHITELIST";
                }

                // 6) Write CSV
                string nodesPath = Path.Combine(outDir, "nodes.csv");
                string edgesPath = Path.Combine(outDir, "edges.csv");
                string annPath = Path.Combine(outDir, "annotation.csv");

                PPVCExportUtils.WriteNodesCsv(nodesPath, nodes);
                PPVCExportUtils.WriteEdgesCsv(edgesPath, edges);
                WriteAnnotationsCsv(annPath, anns);

                TaskDialog.Show("PPVC Export",
                    $"Export done! [Mode: {modeStr}]\n\n" +
                    $"Nodes: {nodes.Count}\n" +
                    $"Edges: {edges.Count}\n" +
                    $"Annotations: {anns.Count}\n" +
                    $"Whitelist IDs: {whitelist.Count}\n\n" +
                    $"Check whitelist_diagnostic.txt in output folder\n\n" +
                    $"Saved to:\n{outDir}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("PPVC Export ERROR", ex.ToString());
                return Result.Failed;
            }
        }

        // Load whitelist for specific PPVC
        private HashSet<long> LoadAnnotationWhitelist(string whitelistPath, string ppvcName)
        {
            var whitelistedIds = new HashSet<long>();
            
            if (!File.Exists(whitelistPath))
                return whitelistedIds;

            try
            {
                // Strip file extension from ppvcName (e.g., "PPVC 01, 20_Typ.rvt" -> "PPVC 01, 20_Typ")
                string ppvcNameNoExt = Path.GetFileNameWithoutExtension(ppvcName).Trim();
                
                var lines = File.ReadAllLines(whitelistPath);
                string diagnosticLog = $"WHITELIST LOADING DEBUG:\n";
                diagnosticLog += $"Document Title: '{ppvcName}'\n";
                diagnosticLog += $"Looking for PPVC: '{ppvcNameNoExt}'\n";
                diagnosticLog += $"Total lines in whitelist: {lines.Length}\n\n";
                
                // First pass: get all unique PPVC names for debugging
                var uniquePpvcNames = new HashSet<string>();
                foreach (var line in lines.Skip(1))
                {
                    var parts = ParseCsvLine(line);
                    if (parts.Length >= 1)
                    {
                        uniquePpvcNames.Add(parts[0].Trim());
                    }
                }
                
                diagnosticLog += $"UNIQUE PPVC NAMES IN WHITELIST ({uniquePpvcNames.Count}):\n";
                foreach (var name in uniquePpvcNames.OrderBy(x => x))
                {
                    bool isMatch = name.Equals(ppvcNameNoExt, StringComparison.OrdinalIgnoreCase);
                    diagnosticLog += $"  - '{name}'{(isMatch ? " ? MATCH!" : "")}\n";
                }
                diagnosticLog += $"\n";
                
                int matchCount = 0;
                foreach (var line in lines.Skip(1)) // Skip header
                {
                    var parts = ParseCsvLine(line);
                    if (parts.Length >= 3)
                    {
                        string ppvc = parts[0].Trim();
                        string viewName = parts.Length > 1 ? parts[1].Trim() : "";
                        string elementIdStr = parts[2].Trim();
                        
                        // Case-insensitive comparison
                        if (ppvc.Equals(ppvcNameNoExt, StringComparison.OrdinalIgnoreCase))
                        {
                            if (long.TryParse(elementIdStr, out long id))
                            {
                                whitelistedIds.Add(id);
                                matchCount++;
                            }
                            else
                            {
                                diagnosticLog += $"? FAILED PARSE: element_id='{elementIdStr}'\n";
                            }
                        }
                    }
                }
                
                diagnosticLog += $"\n========================================\n";
                diagnosticLog += $"Total whitelisted IDs loaded: {whitelistedIds.Count}\n";
                diagnosticLog += $"Match count: {matchCount}\n";
                
                // Write diagnostic to temp directory
                string diagnosticPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "PPVCRevitAddin",
                    "whitelist_loading_debug.txt");
                
                Directory.CreateDirectory(Path.GetDirectoryName(diagnosticPath));
                File.WriteAllText(diagnosticPath, diagnosticLog);
            }
            catch (Exception ex)
            {
                // Still return what we could load
            }

            return whitelistedIds;
        }

        // Robust CSV parser that handles commas inside field values
        // Expects format: ppvc_name,view_name,element_id,annotation_type
        // where ppvc_name may contain commas but view_name and element_id/annotation_type do not
        private string[] ParseCsvLine(string line)
        {
            // Strategy: find the last 3 commas and split on those
            // Format: ppvc_name, view_name, element_id, annotation_type
            // The ppvc_name is everything before the LAST 3 commas
            
            var parts = line.Split(',');
            
            if (parts.Length < 4)
            {
                // Not enough parts, return as-is
                return parts;
            }
            
            // We expect: ppvc_name, view_name, element_id, annotation_type
            // If we have more than 4 parts, it means ppvc_name contains commas
            
            // Get the last 3 parts (they are definitely view_name, element_id, annotation_type)
            string annotationType = parts[parts.Length - 1].Trim();
            string elementId = parts[parts.Length - 2].Trim();
            string viewName = parts[parts.Length - 3].Trim();
            
            // Everything else is the ppvc_name
            string ppvcName = string.Join(",", parts.Take(parts.Length - 3)).Trim();
            
            return new[] { ppvcName, viewName, elementId, annotationType };
        }

        // SMART MODE: Export ALL Dimensions and TextNotes (for new PPVCs)
        private List<AnnotationRecord> CollectAnnotationsSmartMode(Document doc, string outDir)
        {
            var anns = new List<AnnotationRecord>();
            int annId = 0;

            // Find ALL Elevation views
            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitView))
                .Cast<RevitView>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.Elevation)
                .ToList();

            var elevIds = new HashSet<ElementId>(allViews.Select(v => v.Id));
            if (elevIds.Count == 0)
                return anns;

            // --- ALL DIMENSIONS ---
            var dims = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .Where(d => elevIds.Contains(d.OwnerViewId))
                .ToList();

            // DIAGNOSTIC: Log all dimension IDs found
            string dimLog = $"DIMENSION IDs FOUND ({dims.Count}):\n";
            foreach (var d in dims)
            {
                dimLog += $"  ID: {d.Id.Value}\n";
            }
            File.WriteAllText(Path.Combine(outDir, "dimension_ids_found.txt"), dimLog);

            foreach (var d in dims)
            {
                annId++;
                XYZ mid = SafeCurveMidpoint(d.Curve, d.get_BoundingBox(null));
                string valueStr = d.ValueString ?? "";
                
                var ids = new List<string>();
                if (d.References != null && d.References.Size > 0)
                {
                    foreach (Reference r in d.References)
                    {
                        try
                        {
                            Element e = doc.GetElement(r);
                            if (e != null && !e.ViewSpecific)
                            {
                                ids.Add(e.Id.Value.ToString());
                            }
                        }
                        catch { }
                    }
                }
                string targets = string.Join("|", ids.Distinct());

                var ownerView = doc.GetElement(d.OwnerViewId) as RevitView;

                anns.Add(new AnnotationRecord
                {
                    AnnotationId = annId,
                    ElementId = d.Id.Value,
                    ViewName = ownerView?.Name ?? "",
                    ViewType = ownerView?.ViewType.ToString() ?? "",
                    Category = "Dimensions",
                    AnnotationType = d.DimensionType?.Name ?? "dimension",
                    Value = valueStr,
                    X = mid.X,
                    Y = mid.Y,
                    Extra = d.Name ?? "",
                    TargetElementIds = targets
                });
            }

            // --- ALL TEXT NOTES ---
            var texts = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .Where(t => elevIds.Contains(t.OwnerViewId))
                .ToList();

            // DIAGNOSTIC: Log all text note IDs found
            string textLog = $"TEXTNOTE IDs FOUND ({texts.Count}):\n";
            foreach (var t in texts)
            {
                textLog += $"  ID: {t.Id.Value}\n";
            }
            File.WriteAllText(Path.Combine(outDir, "textnote_ids_found.txt"), textLog);

            foreach (var t in texts)
            {
                annId++;
                XYZ p = t.Coord;
                var ownerView = doc.GetElement(t.OwnerViewId) as RevitView;

                // Extract target element IDs for text notes (elements near the text location)
                string targetIds = ExtractTargetElementsForTextNote(doc, t);

                anns.Add(new AnnotationRecord
                {
                    AnnotationId = annId,
                    ElementId = t.Id.Value,
                    ViewName = ownerView?.Name ?? "",
                    ViewType = ownerView?.ViewType.ToString() ?? "",
                    Category = "TextNotes",
                    AnnotationType = "text_note",
                    Value = t.Text ?? "",
                    X = p.X,
                    Y = p.Y,
                    Extra = t.TextNoteType?.Name ?? "",
                    TargetElementIds = targetIds
                });
            }

            return anns;
        }

        // =========================================================
        // ==================== ANNOTATIONS ========================
        // =========================================================

        private List<AnnotationRecord> CollectAnnotations(Document doc, HashSet<long> whitelist)
        {
            var anns = new List<AnnotationRecord>();
            int annId = 0;

            // --- Find ALL Elevation views ---
            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitView))
                .Cast<RevitView>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.Elevation)
                .ToList();

            var elevIds = new HashSet<ElementId>(allViews.Select(v => v.Id));
            if (elevIds.Count == 0)
                return anns;

            string extractionLog = $"=== STRICT WHITELIST EXTRACTION LOG ===\n";
            extractionLog += $"Whitelist contains: {whitelist.Count} IDs\n";
            extractionLog += $"Looking in {allViews.Count} Elevation views\n\n";

            int foundCount = 0;
            int skippedCount = 0;

            // --- DIMENSIONS (STRICT: only whitelisted) ---
            var dims = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .Where(d => elevIds.Contains(d.OwnerViewId))
                .ToList();

            extractionLog += $"DIMENSIONS FOUND IN DOCUMENT: {dims.Count}\n";

            foreach (var d in dims)
            {
                // STRICT: ONLY add if in whitelist
                if (!whitelist.Contains(d.Id.Value))
                {
                    skippedCount++;
                    extractionLog += $"  ? SKIPPED: ID {d.Id.Value} (NOT in whitelist)\n";
                    continue;
                }

                foundCount++;
                annId++;
                XYZ mid = SafeCurveMidpoint(d.Curve, d.get_BoundingBox(null));
                string valueStr = d.ValueString ?? "";
                
                var ids = new List<string>();
                if (d.References != null && d.References.Size > 0)
                {
                    foreach (Reference r in d.References)
                    {
                        try
                        {
                            Element e = doc.GetElement(r);
                            if (e != null && !e.ViewSpecific)
                            {
                                ids.Add(e.Id.Value.ToString());
                            }
                        }
                        catch { }
                    }
                }
                string targets = string.Join("|", ids.Distinct());

                var ownerView = doc.GetElement(d.OwnerViewId) as RevitView;

                extractionLog += $"  ? EXTRACTED: ID {d.Id.Value} from {ownerView?.Name}\n";

                anns.Add(new AnnotationRecord
                {
                    AnnotationId = annId,
                    ElementId = d.Id.Value,
                    ViewName = ownerView?.Name ?? "",
                    ViewType = ownerView?.ViewType.ToString() ?? "",
                    Category = "Dimensions",
                    AnnotationType = d.DimensionType?.Name ?? "dimension",
                    Value = valueStr,
                    X = mid.X,
                    Y = mid.Y,
                    Extra = d.Name ?? "",
                    TargetElementIds = targets
                });
            }

            extractionLog += $"\n--- DIMENSIONS: Found {foundCount}, Skipped {skippedCount} ---\n\n";

            foundCount = 0;
            skippedCount = 0;

            // --- TEXT NOTES (STRICT: only whitelisted) ---
            var texts = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .Where(t => elevIds.Contains(t.OwnerViewId))
                .ToList();

            extractionLog += $"TEXT NOTES FOUND IN DOCUMENT: {texts.Count}\n";

            foreach (var t in texts)
            {
                // STRICT: ONLY add if in whitelist
                if (!whitelist.Contains(t.Id.Value))
                {
                    skippedCount++;
                    extractionLog += $"  ? SKIPPED: ID {t.Id.Value} (NOT in whitelist)\n";
                    continue;
                }

                foundCount++;
                annId++;
                XYZ p = t.Coord;
                var ownerView = doc.GetElement(t.OwnerViewId) as RevitView;

                // Extract target element IDs for text notes (elements near the text location)
                string targetIds = ExtractTargetElementsForTextNote(doc, t);

                extractionLog += $"  ? EXTRACTED: ID {t.Id.Value} from {ownerView?.Name}\n";

                anns.Add(new AnnotationRecord
                {
                    AnnotationId = annId,
                    ElementId = t.Id.Value,
                    ViewName = ownerView?.Name ?? "",
                    ViewType = ownerView?.ViewType.ToString() ?? "",
                    Category = "TextNotes",
                    AnnotationType = "text_note",
                    Value = t.Text ?? "",
                    X = p.X,
                    Y = p.Y,
                    Extra = t.TextNoteType?.Name ?? "",
                    TargetElementIds = targetIds
                });
            }

            extractionLog += $"\n--- TEXT NOTES: Found {foundCount}, Skipped {skippedCount} ---\n";
            extractionLog += $"\n? TOTAL EXTRACTED: {anns.Count} annotations (must match whitelist count)\n";

            // Write extraction log for verification
            File.AppendAllText(
                Path.Combine(
                    Path.GetDirectoryName(Path.Combine(
                        Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                        "..",
                        "..",
                        "annotation_whitelist.csv")) ?? "",
                    "extraction_verification.txt"), 
                extractionLog);

            return anns;
        }

        // =========================================================
        // ================ Folder picker ===========================
        // =========================================================

        private string PickOutputFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder to save PPVC export CSVs";
                var result = dialog.ShowDialog();
                if (result == DialogResult.OK)
                    return dialog.SelectedPath;
            }
            return null;
        }

        // Extract target elements for TextNote by finding nearby model elements
        private string ExtractTargetElementsForTextNote(Document doc, TextNote textNote)
        {
            try
            {
                var targetIds = new List<string>();
                XYZ textPos = textNote.Coord;

                // Collect all model elements (Walls, Floors, Columns, etc.)
                BuiltInCategory[] cats =
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_Rooms
                };

                var filter = new ElementMulticategoryFilter(
                    cats.Select(c => new ElementId(c)).ToList());

                var elems = new FilteredElementCollector(doc)
                    .WherePasses(filter)
                    .WhereElementIsNotElementType()
                    .ToElements();

                // Use a reasonable search radius (10 feet)
                double searchRadius = 10.0;

                foreach (var elem in elems)
                {
                    if (elem.ViewSpecific)
                        continue;

                    try
                    {
                        // Get bounding box in world coordinates
                        BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                        if (bb == null)
                            continue;

                        // Check if text note position is within or near the element's bounding box
                        XYZ elemCenter = (bb.Min + bb.Max) * 0.5;
                        double distance = textPos.DistanceTo(elemCenter);

                        if (distance <= searchRadius)
                        {
                            targetIds.Add(elem.Id.Value.ToString());
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                return string.Join("|", targetIds.Distinct());
            }
            catch
            {
                return "";
            }
        }

        // SAFE mid-point for possibly unbound curves
        private XYZ SafeCurveMidpoint(Curve c, BoundingBoxXYZ bbFallback)
        {
            try
            {
                if (c != null)
                {
                    if (c.IsBound)
                    {
                        // Normalize OK
                        return c.Evaluate(0.5, true);
                    }
                    else
                    {
                        // Unbound: use non-normalized parameter safely
                        double p0 = c.GetEndParameter(0);
                        double p1 = c.GetEndParameter(1);
                        double pm = (p0 + p1) * 0.5;
                        return c.Evaluate(pm, false);
                    }
                }
            }
            catch
            {
                // fall through to bbox
            }

            if (bbFallback != null)
                return (bbFallback.Min + bbFallback.Max) * 0.5;

            return XYZ.Zero;
        }

        private XYZ GetElementPoint(Element e)
        {
            if (e.Location is LocationPoint lp)
                return lp.Point;

            if (e.Location is LocationCurve lc && lc.Curve != null)
            {
                Curve c = lc.Curve;
                if (c.IsBound)
                    return c.Evaluate(0.5, true);

                double p0 = c.GetEndParameter(0);
                double p1 = c.GetEndParameter(1);
                return c.Evaluate((p0 + p1) * 0.5, false);
            }

            BoundingBoxXYZ bb = e.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) * 0.5;

            return XYZ.Zero;
        }

        private string GetElementLabel(Element e)
        {
            Parameter p = e.LookupParameter("Text");
            if (p != null && p.StorageType == StorageType.String)
                return p.AsString() ?? "";

            return e.Name ?? "";
        }

        // =========================================================
        // ======================= CSV IO ==========================
        // =========================================================

        private void WriteAnnotationsCsv(string path, List<AnnotationRecord> anns)
        {
            var sb = new StringBuilder();
            sb.AppendLine(AnnotationRecord.Header());
            foreach (var a in anns) sb.AppendLine(a.ToCsv());
            File.WriteAllText(path, sb.ToString());
        }
    }

    // =========================================================
    // ====================== RECORDS ==========================
    // =========================================================

    public class NodeRecord
    {
        public long Id;
        public string Category;
        public string Family;
        public string TypeName;
        public long LevelId;
        public long HostId;
        public double MinX, MinY, MinZ;
        public double MaxX, MaxY, MaxZ;
        public double Cx, Cy, Cz;
        public double Length, Height, Thickness, Area, Width, Depth;
        public double DirX, DirY, DirZ;
        public string RoomName;
        public string RoomNumber;
        public double LevelElevation;

        public static string Header()
        {
            return "id,category,family,type,level_id,host_id," +
                   "min_x,min_y,min_z,max_x,max_y,max_z," +
                   "cx,cy,cz,length,height,thickness,area,width,depth," +
                   "dir_x,dir_y,dir_z,room_name,room_number,level_elevation";
        }

        public string ToCsv()
        {
            string S(string s) => (s ?? "").Replace(",", "_");

            return string.Join(",",
                Id,
                S(Category),
                S(Family),
                S(TypeName),
                LevelId,
                HostId,
                MinX, MinY, MinZ,
                MaxX, MaxY, MaxZ,
                Cx, Cy, Cz,
                Length, Height, Thickness, Area, Width, Depth,
                DirX, DirY, DirZ,
                S(RoomName),
                S(RoomNumber),
                LevelElevation);
        }
    }

    public class EdgeRecord
    {
        public long Src;
        public long Dst;
        public string Type;
    }

    public class AnnotationRecord
    {
        public int AnnotationId;
        public long ElementId;
        public string ViewName;
        public string ViewType;
        public string Category;
        public string AnnotationType;
        public string Value;
        public double X;
        public double Y;
        public string Extra;
        public string TargetElementIds;

        public static string Header()
        {
            return "annotation_id,element_id,view_name,view_type,category," +
                   "annotation_type,value,x,y,extra,target_element_ids";
        }

        public string ToCsv()
        {
            string S(string s)
            {
                if (s == null) s = "";
                s = s.Replace("\"", "\"\"");
                if (s.Contains(",") || s.Contains("\n"))
                    return $"\"{s}\"";
                return s;
            }

            string F(double d) => d.ToString("G", CultureInfo.InvariantCulture);

            return string.Join(",",
                AnnotationId,
                ElementId,
                S(ViewName),
                S(ViewType),
                S(Category),
                S(AnnotationType),
                S(Value),
                F(X),
                F(Y),
                S(Extra),
                S(TargetElementIds));
        }
    }
}
