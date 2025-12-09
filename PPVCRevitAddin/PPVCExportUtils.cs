using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PPVCRevitAddin
{
    /// <summary>
    /// Represents a whitelist entry with compound key
    /// </summary>
    public class WhitelistEntry
    {
        public string PpvcName { get; set; }
        public string ViewName { get; set; }
        public long ElementId { get; set; }
        public string AnnotationType { get; set; }
    }

    /// <summary>
    /// Shared helpers for exporting nodes/edges from the Revit model.
    /// </summary>
    public static class PPVCExportUtils
    {
        /// <summary>
        /// Load annotation whitelist from embedded resource or file.
        /// Returns a dictionary keyed by (ppvc_name, view_name, element_id) tuple.
        /// </summary>
        public static Dictionary<(string ppvc, string view, long elementId), WhitelistEntry> LoadAnnotationWhitelistFull()
        {
            var whitelist = new Dictionary<(string ppvc, string view, long elementId), WhitelistEntry>();
            
            try
            {
                // Try to load from embedded resource first
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                string resourceName = "PPVCRevitAddin.annotation_whitelist.csv";
                
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (var reader = new System.IO.StreamReader(stream))
                        {
                            string content = reader.ReadToEnd();
                            ParseWhitelistContentFull(content, whitelist);
                        }
                    }
                }
                
                // If no embedded resource, try file in same directory as assembly
                if (whitelist.Count == 0)
                {
                    string assemblyPath = assembly.Location;
                    string directory = System.IO.Path.GetDirectoryName(assemblyPath);
                    string filePath = System.IO.Path.Combine(directory, "annotation_whitelist.csv");
                    
                    if (System.IO.File.Exists(filePath))
                    {
                        string content = System.IO.File.ReadAllText(filePath);
                        ParseWhitelistContentFull(content, whitelist);
                    }
                }
            }
            catch
            {
                // If loading fails, return empty dictionary
            }
            
            return whitelist;
        }

        /// <summary>
        /// Load annotation whitelist as a simple set of element IDs (for backward compatibility)
        /// </summary>
        public static HashSet<long> LoadAnnotationWhitelist()
        {
            var whitelistIds = new HashSet<long>();
            
            try
            {
                // Try to load from embedded resource first
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                string resourceName = "PPVCRevitAddin.annotation_whitelist.csv";
                
                System.Diagnostics.Debug.WriteLine($"[LoadAnnotationWhitelist] Trying embedded resource: {resourceName}");
                
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (var reader = new System.IO.StreamReader(stream))
                        {
                            string content = reader.ReadToEnd();
                            ParseWhitelistContent(content, whitelistIds);
                            System.Diagnostics.Debug.WriteLine($"[LoadAnnotationWhitelist] Loaded {whitelistIds.Count} IDs from embedded resource");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[LoadAnnotationWhitelist] Embedded resource not found");
                    }
                }
                
                // If no embedded resource, try file in same directory as assembly
                if (whitelistIds.Count == 0)
                {
                    string assemblyPath = assembly.Location;
                    string directory = System.IO.Path.GetDirectoryName(assemblyPath);
                    string filePath = System.IO.Path.Combine(directory, "annotation_whitelist.csv");
                    
                    System.Diagnostics.Debug.WriteLine($"[LoadAnnotationWhitelist] Trying file: {filePath}");
                    
                    if (System.IO.File.Exists(filePath))
                    {
                        string content = System.IO.File.ReadAllText(filePath);
                        ParseWhitelistContent(content, whitelistIds);
                        System.Diagnostics.Debug.WriteLine($"[LoadAnnotationWhitelist] Loaded {whitelistIds.Count} IDs from file");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadAnnotationWhitelist] File not found: {filePath}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadAnnotationWhitelist] ERROR: {ex.Message}");
            }
            
            return whitelistIds;
        }

        /// <summary>
        /// Get all whitelisted element IDs for a specific PPVC and view
        /// </summary>
        public static HashSet<long> GetWhitelistedElementsForView(
            Dictionary<(string ppvc, string view, long elementId), WhitelistEntry> whitelist,
            string ppvcName,
            string viewName)
        {
            var result = new HashSet<long>();
            
            foreach (var kvp in whitelist)
            {
                // Match PPVC name (case-insensitive, partial match)
                bool ppvcMatch = kvp.Key.ppvc.Equals(ppvcName, StringComparison.OrdinalIgnoreCase) ||
                                 ppvcName.Contains(kvp.Key.ppvc) ||
                                 kvp.Key.ppvc.Contains(ppvcName);
                
                // Match view name (case-insensitive)
                bool viewMatch = kvp.Key.view.Equals(viewName, StringComparison.OrdinalIgnoreCase);
                
                if (ppvcMatch && viewMatch)
                {
                    result.Add(kvp.Key.elementId);
                }
            }
            
            return result;
        }

        private static void ParseWhitelistContentFull(string content, Dictionary<(string ppvc, string view, long elementId), WhitelistEntry> whitelist)
        {
            var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            // Skip header: ppvc_name,view_name,element_id,annotation_type
            for (int i = 1; i < lines.Length; i++)
            {
                var parts = lines[i].Split(',');
                if (parts.Length >= 4)
                {
                    string ppvcName = parts[0].Trim();
                    string viewName = parts[1].Trim();
                    string annotationType = parts[3].Trim();
                    
                    if (long.TryParse(parts[2].Trim(), out long elementId))
                    {
                        var key = (ppvcName, viewName, elementId);
                        if (!whitelist.ContainsKey(key))
                        {
                            whitelist[key] = new WhitelistEntry
                            {
                                PpvcName = ppvcName,
                                ViewName = viewName,
                                ElementId = elementId,
                                AnnotationType = annotationType
                            };
                        }
                    }
                }
            }
        }

        private static void ParseWhitelistContent(string content, HashSet<long> whitelistIds)
        {
            var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            // Skip header: ppvc_name,view_name,element_id,annotation_type
            for (int i = 1; i < lines.Length; i++)
            {
                var parts = lines[i].Split(',');
                if (parts.Length >= 3)
                {
                    // element_id is the 3rd column (index 2)
                    if (long.TryParse(parts[2].Trim(), out long elementId))
                    {
                        whitelistIds.Add(elementId);
                    }
                }
            }
        }

        // STRICT MODE: Only export nodes that are referenced in the whitelist
        public static List<NodeRecord> CollectNodes(Document doc, HashSet<long> whitelistNodeIds = null)
        {
            var nodes = new List<NodeRecord>();

            BuiltInCategory[] cats =
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_Rooms,
                BuiltInCategory.OST_Levels
            };

            var filter = new ElementMulticategoryFilter(
                cats.Select(c => new ElementId(c)).ToList());

            var elements = new FilteredElementCollector(doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (var e in elements)
            {
                if (e.ViewSpecific) continue; // only model elements

                long elementId = e.Id.Value;

                // STRICT: If whitelist is provided, only include whitelisted nodes
                if (whitelistNodeIds != null && !whitelistNodeIds.Contains(elementId))
                    continue;

                var nr = BuildNodeRecord(doc, e);
                if (nr != null) nodes.Add(nr);
            }

            return nodes;
        }

        public static List<EdgeRecord> CollectEdges(Document doc, List<NodeRecord> nodes)
        {
            var edges = new List<EdgeRecord>();
            var idSet = new HashSet<long>(nodes.Select(n => n.Id));

            // host edges (STRICT: only if both nodes exist)
            foreach (var n in nodes)
            {
                if (n.HostId != -1 && idSet.Contains(n.HostId))
                {
                    edges.Add(new EdgeRecord
                    {
                        Src = n.HostId,
                        Dst = n.Id,
                        Type = "host"
                    });
                }
            }

            // level edges (STRICT: only if both nodes exist)
            foreach (var n in nodes)
            {
                if (n.LevelId != -1 && idSet.Contains(n.LevelId))
                {
                    edges.Add(new EdgeRecord
                    {
                        Src = n.LevelId,
                        Dst = n.Id,
                        Type = "level"
                    });
                }
            }

            // adjacency edges (bounding boxes touching) - STRICT: only between whitelisted nodes
            for (int i = 0; i < nodes.Count; i++)
            {
                for (int j = i + 1; j < nodes.Count; j++)
                {
                    if (BoxesTouch(nodes[i], nodes[j]))
                    {
                        edges.Add(new EdgeRecord
                        {
                            Src = nodes[i].Id,
                            Dst = nodes[j].Id,
                            Type = "adjacent"
                        });
                    }
                }
            }

            return edges;
        }

        public static void WriteNodesCsv(string path, List<NodeRecord> nodes)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(NodeRecord.Header());
            foreach (var n in nodes) sb.AppendLine(n.ToCsv());
            System.IO.File.WriteAllText(path, sb.ToString());
        }

        public static void WriteEdgesCsv(string path, List<EdgeRecord> edges)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("src,dst,type");
            foreach (var e in edges) sb.AppendLine($"{e.Src},{e.Dst},{e.Type}");
            System.IO.File.WriteAllText(path, sb.ToString());
        }

        /// <summary>
        /// Write annotation CSV - ONLY exports annotations that exist in the whitelist.
        /// Exports from ALL elevation views, not just the current view.
        /// This is hardcoded for proof of concept.
        /// </summary>
        public static void WriteAnnotationCsv(string path, Document doc, string currentPpvcName = null, string currentViewName = null)
        {
            // HARDCODED: Load ALL element IDs from whitelist (ignore PPVC/view matching for now)
            var whitelistElementIds = LoadAnnotationWhitelist();
            
            System.Diagnostics.Debug.WriteLine($"[WriteAnnotationCsv] Loaded {whitelistElementIds.Count} element IDs from whitelist");
            
            // If whitelist is empty, don't export anything (fail-safe)
            if (whitelistElementIds.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[WriteAnnotationCsv] WARNING: Whitelist is empty! No annotations will be exported.");
            }
            
            var sb = new System.Text.StringBuilder();
            
            // Exact PPVC 13 format header - 11 columns
            sb.AppendLine("annotation_id,element_id,view_name,view_type,category,annotation_type,value,x,y,extra,target_element_ids");

            int annotationId = 1;
            int skippedCount = 0;

            // Get ALL elevation views in the document
            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && 
                           (v.ViewType == ViewType.Elevation || 
                            v.ViewType == ViewType.Section ||
                            v.ViewType == ViewType.FloorPlan ||
                            v.ViewType == ViewType.Detail))
                .ToList();

            System.Diagnostics.Debug.WriteLine($"[WriteAnnotationCsv] Found {allViews.Count} views to search for annotations");

            // Collect TextNotes from ALL views - ONLY those in whitelist
            var textNoteCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNote));

            foreach (TextNote tn in textNoteCollector)
            {
                try
                {
                    long elementId = tn.Id.Value;
                    
                    // STRICT: Only include if element ID is in whitelist
                    if (!whitelistElementIds.Contains(elementId))
                    {
                        skippedCount++;
                        continue;
                    }

                    string text = tn.Text;
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    // Escape quotes in text for CSV
                    text = text.Replace("\"", "\"\"");
                    
                    // Get location of text note
                    XYZ pos = tn.Coord;
                    
                    // Get the view that contains this text note
                    string viewName = "";
                    string viewType = "";
                    try
                    {
                        var ownerView = doc.GetElement(tn.OwnerViewId) as View;
                        if (ownerView != null)
                        {
                            viewName = ownerView.Name;
                            viewType = ownerView.ViewType.ToString();
                        }
                    }
                    catch { }

                    // Category - use PLURAL form to match training data
                    string category = "TextNotes";

                    // Annotation type - use "text" for TextNotes
                    string annotationType = "text";

                    // Value is the text itself
                    string value = text;

                    // X, Y coordinates
                    string x = pos.X.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string y = pos.Y.ToString(System.Globalization.CultureInfo.InvariantCulture);

                    // Extra field (empty)
                    string extra = "";

                    // Target element IDs (empty - will be filled by Python)
                    string targetElementIds = "";

                    // Build CSV line with exact column order
                    sb.AppendLine($"{annotationId},{elementId},\"{viewName}\",\"{viewType}\",\"{category}\",\"{annotationType}\",\"{value}\",{x},{y},\"{extra}\",\"{targetElementIds}\"");
                    
                    annotationId++;
                }
                catch { }
            }

            // Collect Dimensions from ALL views - ONLY those in whitelist
            var dimensionCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension));

            foreach (Dimension dim in dimensionCollector)
            {
                try
                {
                    long elementId = dim.Id.Value;
                    
                    // STRICT: Only include if element ID is in whitelist
                    if (!whitelistElementIds.Contains(elementId))
                    {
                        skippedCount++;
                        continue;
                    }

                    // Get dimension value
                    string valueStr = "";
                    try
                    {
                        if (dim.Value.HasValue)
                        {
                            // Convert from feet to mm
                            double valueMm = dim.Value.Value * 304.8;
                            valueStr = valueMm.ToString("F0", System.Globalization.CultureInfo.InvariantCulture) + " mm";
                        }
                        else if (!string.IsNullOrEmpty(dim.ValueString))
                        {
                            valueStr = dim.ValueString;
                        }
                    }
                    catch { }

                    if (string.IsNullOrWhiteSpace(valueStr))
                        continue;

                    // Escape quotes in value for CSV
                    valueStr = valueStr.Replace("\"", "\"\"");
                    
                    // Get location (use origin of dimension curve)
                    double posX = 0, posY = 0;
                    try
                    {
                        if (dim.Curve != null)
                        {
                            XYZ midPoint = dim.Curve.Evaluate(0.5, true);
                            posX = midPoint.X;
                            posY = midPoint.Y;
                        }
                    }
                    catch { }
                    
                    // Get the view that contains this dimension
                    string viewName = "";
                    string viewType = "";
                    try
                    {
                        var ownerView = doc.GetElement(dim.OwnerViewId) as View;
                        if (ownerView != null)
                        {
                            viewName = ownerView.Name;
                            viewType = ownerView.ViewType.ToString();
                        }
                    }
                    catch { }

                    // Category - use PLURAL form to match training data
                    string category = "Dimensions";

                    // Annotation type
                    string annotationType = "dimension";

                    // X, Y coordinates
                    string x = posX.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string y = posY.ToString(System.Globalization.CultureInfo.InvariantCulture);

                    // Extra field (empty)
                    string extra = "";

                    // Target element IDs - try to get referenced elements
                    string targetElementIds = "";
                    try
                    {
                        var refArray = dim.References;
                        if (refArray != null && refArray.Size > 0)
                        {
                            var targetIds = new List<string>();
                            foreach (Reference r in refArray)
                            {
                                if (r.ElementId != ElementId.InvalidElementId)
                                {
                                    targetIds.Add(r.ElementId.Value.ToString());
                                }
                            }
                            targetElementIds = string.Join(";", targetIds);
                        }
                    }
                    catch { }

                    // Build CSV line with exact column order
                    sb.AppendLine($"{annotationId},{elementId},\"{viewName}\",\"{viewType}\",\"{category}\",\"{annotationType}\",\"{valueStr}\",{x},{y},\"{extra}\",\"{targetElementIds}\"");
                    
                    annotationId++;
                }
                catch { }
            }

            System.Diagnostics.Debug.WriteLine($"[WriteAnnotationCsv] Exported {annotationId - 1} annotations, skipped {skippedCount} (not in whitelist)");
            
            System.IO.File.WriteAllText(path, sb.ToString());
        }

        // Overload for backward compatibility (no PPVC/view filtering)
        public static void WriteAnnotationCsv(string path, Document doc)
        {
            WriteAnnotationCsv(path, doc, null, null);
        }

        /// <summary>
        /// Determine annotation type based on text content
        /// </summary>
        private static string DetermineAnnotationType(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "text";

            text = text.ToLower();

            // Check for dimension-like content (contains numbers and measurements)
            if (text.Contains("mm") || text.Contains("m²") || text.Contains("m") && char.IsDigit(text[0]))
                return "dimension";

            // Check for area notation
            if (text.Contains("area") || text.Contains("m²"))
                return "dimension";

            // Check for height notation
            if (text.Contains("height") || text.Contains("h="))
                return "dimension";

            // Check for thickness notation
            if (text.Contains("thick") || text.Contains("thk") || text.Contains("t="))
                return "dimension";

            // Check for length notation
            if (text.Contains("length") || text.Contains("l="))
                return "dimension";

            // Default to text
            return "text";
        }

        // ----------------- Private helpers copied from your original class -----------------

        private static NodeRecord BuildNodeRecord(Document doc, Element e)
        {
            var nr = new NodeRecord
            {
                Id = e.Id.Value,
                Category = e.Category?.Name ?? "Unknown",
                Family = GetFamilyName(e),
                TypeName = doc.GetElement(e.GetTypeId())?.Name ?? "UnknownType",
                LevelId = GetLevelId(e),
                HostId = GetHostId(e)
            };

            BoundingBoxXYZ bb = e.get_BoundingBox(null);
            if (bb != null)
            {
                nr.MinX = bb.Min.X;
                nr.MinY = bb.Min.Y;
                nr.MinZ = bb.Min.Z;
                nr.MaxX = bb.Max.X;
                nr.MaxY = bb.Max.Y;
                nr.MaxZ = bb.Max.Z;
                nr.Cx = (nr.MinX + nr.MaxX) / 2.0;
                nr.Cy = (nr.MinY + nr.MaxY) / 2.0;
                nr.Cz = (nr.MinZ + nr.MaxZ) / 2.0;
            }

            if (e is Wall wall)
            {
                nr.Length = wall.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                nr.Height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0;
                nr.Thickness = wall.WallType?.Width ?? 0;

                if (wall.Location is LocationCurve lc && lc.Curve is Line line)
                {
                    XYZ d = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
                    nr.DirX = d.X;
                    nr.DirY = d.Y;
                    nr.DirZ = d.Z;
                }
            }
            else if (e is Floor floor)
            {
                nr.Thickness = floor.FloorType?
                                    .get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?
                                    .AsDouble() ?? 0;
                nr.Area = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0;
            }
            else if (e.Category != null &&
                     e.Category.Id.Value == (long)BuiltInCategory.OST_StructuralFraming)
            {
                nr.Length = e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                nr.Width = e.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH)?.AsDouble() ?? 0;
                nr.Depth = e.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_HEIGHT)?.AsDouble() ?? 0;
            }
            else if (e is Room rm)
            {
                nr.RoomName = rm.Name;
                nr.RoomNumber = rm.Number;
                nr.Area = rm.Area;
            }
            else if (e is Level lvl)
            {
                nr.LevelElevation = lvl.Elevation;
            }

            return nr;
        }

        private static string GetFamilyName(Element e)
        {
            if (e is FamilyInstance fi)
                return fi.Symbol?.FamilyName ?? "UnknownFamily";

            var t = e.Document.GetElement(e.GetTypeId()) as ElementType;
            return t?.FamilyName ?? t?.Name ?? "UnknownFamily";
        }

        private static long GetLevelId(Element e)
        {
            Parameter p = e.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (p != null && p.StorageType == StorageType.ElementId)
                return p.AsElementId().Value;

            if (e is Room rm) return rm.LevelId.Value;
            return -1;
        }

        private static long GetHostId(Element e)
        {
            if (e is FamilyInstance fi && fi.Host != null)
                return fi.Host.Id.Value;

            return -1;
        }

        private static bool BoxesTouch(NodeRecord a, NodeRecord b, double tol = 0.05)
        {
            bool x = a.MinX <= b.MaxX + tol && a.MaxX >= b.MinX - tol;
            bool y = a.MinY <= b.MaxY + tol && a.MaxY >= b.MinY - tol;
            bool z = a.MinZ <= b.MaxZ + tol && a.MaxZ >= b.MinZ - tol;
            return x && y && z;
        }
    }
}
