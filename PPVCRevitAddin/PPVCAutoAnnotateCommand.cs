using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using RevitView = Autodesk.Revit.DB.View;

namespace PPVCRevitAddin
{
    [Transaction(TransactionMode.Manual)]
    public class PPVCAutoAnnotateCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Create a log file path on Desktop
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "PPVCAutoAnnotate_Diagnostic.txt");

            void LogToFile(string msg)
            {
                Debug.WriteLine(msg);
                File.AppendAllText(logPath, msg + "\n");
            }

            try
            {
                // Clear previous log and write header
                File.WriteAllText(logPath, $"=== PPVC Auto Annotate Diagnostic Log ===\n{DateTime.Now:yyyy-MM-dd HH:mm:ss}\n\n");

                // 1) Ask user for working folder
                string outDir = PickOutputFolder();
                if (string.IsNullOrEmpty(outDir))
                {
                    TaskDialog.Show("PPVC Auto Annotate", "Operation cancelled.");
                    return Result.Cancelled;
                }

                string nodesPath = Path.Combine(outDir, "nodes.csv");
                string edgesPath = Path.Combine(outDir, "edges.csv");
                string annotationPath = Path.Combine(outDir, "annotation.csv");
                string predPath = Path.Combine(outDir, "predictions.csv");

                // Get current view info for whitelist filtering
                RevitView activeView = uidoc.ActiveView as RevitView;
                string currentViewName = activeView?.Name ?? "";
                
                // Extract PPVC name from document title or view name
                // Document titles are typically like "PPVC 13_Typ" or contain "PPVC XX"
                string currentPpvcName = ExtractPpvcName(doc.Title, currentViewName);
                
                LogToFile($"Document Title: {doc.Title}");
                LogToFile($"Extracted PPVC Name: {currentPpvcName}");
                LogToFile($"Current View: {currentViewName}");

                // 2) Export nodes + edges using the same logic as export command
                var nodes = PPVCExportUtils.CollectNodes(doc);
                var nodeDict = nodes.ToDictionary(n => n.Id, n => n);
                var edges = PPVCExportUtils.CollectEdges(doc, nodes);

                PPVCExportUtils.WriteNodesCsv(nodesPath, nodes);
                PPVCExportUtils.WriteEdgesCsv(edgesPath, edges);

                // 3) Export annotations with whitelist filtering by PPVC and view
                PPVCExportUtils.WriteAnnotationCsv(annotationPath, doc, currentPpvcName, currentViewName);

                // 4) Call Python script with annotation CSV parameter
                if (!RunPythonInference(nodesPath, edgesPath, annotationPath, predPath))
                {
                    return Result.Failed;
                }

                if (!File.Exists(predPath))
                {
                    TaskDialog.Show("PPVC Auto Annotate",
                        $"predictions.csv not found at:\n{predPath}");
                    return Result.Failed;
                }

                // 5) Load predictions
                var predictions = LoadPredictions(predPath);
                if (predictions.Count == 0)
                {
                    TaskDialog.Show("PPVC Auto Annotate",
                        "No predictions found in predictions.csv.");
                    return Result.Succeeded;
                }

                LogToFile($"✓ Loaded {predictions.Count} predictions from CSV");

                // 6) Use the already defined activeView for annotation
                // (activeView was already declared above for whitelist filtering)
                
                if (activeView == null)
                {
                    TaskDialog.Show("PPVC Auto Annotate",
                        "Please open a view (elevation, floor plan, or section) to annotate.");
                    return Result.Failed;
                }

                // Allow elevation, floor plan, section, and detail views
                ViewType[] supportedViewTypes = 
                {
                    ViewType.Elevation,
                    ViewType.FloorPlan,
                    ViewType.Section,
                    ViewType.Detail
                };

                if (!supportedViewTypes.Contains(activeView.ViewType))
                {
                    TaskDialog.Show("PPVC Auto Annotate",
                        $"Active view type '{activeView.ViewType}' is not supported.\n" +
                        $"Please use: Elevation, Floor Plan, Section, or Detail view.");
                    return Result.Failed;
                }

                LogToFile($"✓ Active view: {activeView.Name} ({activeView.ViewType})\n");
                LogToFile($"\n=== DIMENSION PLACEMENT LOG ===");

                int textCount = 0;
                int dimCount = 0;

                using (Transaction tx = new Transaction(doc, "PPVC Auto Annotate"))
                {
                    tx.Start();

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

                    var elems = new FilteredElementCollector(doc)
                        .WherePasses(filter)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    int elementsCandidateForDim = 0;
                    int elementsPredictedDim = 0;
                    int elementsAttemptedDim = 0;
                    int elementsSucceededDim = 0;

                    foreach (var e in elems)
                    {
                        long id = e.Id.Value;
                        if (!predictions.ContainsKey(id))
                            continue;

                        var pred = predictions[id];
                        nodeDict.TryGetValue(id, out NodeRecord nodeInfo);

                        elementsCandidateForDim++;

                        if (pred.NeedText)
                        {
                            if (CreateTextNoteForElement(doc, activeView, e, nodeInfo, pred))
                            {
                                textCount++;
                            }
                        }

                        if (pred.NeedDimension)
                        {
                            elementsPredictedDim++;
                            
                            // Check if this is a type we can dimension
                            bool isWall = e is Wall;
                            bool isFloor = e is Floor;
                            bool isLevel = e is Level;
                            bool isRoom = e is Room;
                            bool isStructuralFraming = e.Category != null && 
                                e.Category.Id.Value == (long)BuiltInCategory.OST_StructuralFraming;
                            bool isGenericModel = e.Category != null && 
                                e.Category.Id.Value == (long)BuiltInCategory.OST_GenericModel;
                            
                            if (isWall || isFloor || isLevel || isRoom || isStructuralFraming || isGenericModel)
                            {
                                elementsAttemptedDim++;
                                
                                if (CreateDimensionsForElement(doc, activeView, e, nodeInfo, pred, LogToFile))
                                {
                                    elementsSucceededDim++;
                                    dimCount++;
                                }
                            }
                            else
                            {
                                LogToFile($"  [SKIP] {e.Category?.Name} ID={id}: Unsupported category for dimensions");
                            }
                        }
                    }

                    tx.Commit();

                    // Log diagnostic info
                    LogToFile($"\n=== DIMENSION DIAGNOSTIC ===");
                    LogToFile($"Elements with predictions: {elementsCandidateForDim}");
                    LogToFile($"Predicted by model for dimensions: {elementsPredictedDim}");
                    LogToFile($"Attempted to create (Walls/Floors): {elementsAttemptedDim}");
                    LogToFile($"Successfully created: {elementsSucceededDim}");
                    LogToFile($"Failed to create: {elementsAttemptedDim - elementsSucceededDim}");
                    LogToFile($"=========================\n");
                }

                LogToFile($"\nFINAL RESULTS:");
                LogToFile($"Text notes created: {textCount}");
                LogToFile($"Dimensions created: {dimCount}");
                LogToFile($"\nLog file saved to: {logPath}");

                TaskDialog.Show("PPVC Auto Annotate",
                    $"Auto-annotation complete on: {activeView.Name}\n\n" +
                    $"Text notes created: {textCount}\n" +
                    $"Dimensions created: {dimCount}\n\n" +
                    $"Diagnostic log saved to Desktop:\n" +
                    $"PPVCAutoAnnotate_Diagnostic.txt");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                string errorMsg = $"ERROR: {ex.ToString()}";
                LogToFile(errorMsg);
                TaskDialog.Show("PPVC Auto Annotate ERROR", ex.ToString());
                return Result.Failed;
            }
        }

        // ---------------- Extract PPVC name from document or view -----------------

        /// <summary>
        /// Extract PPVC name from document title or view name.
        /// Matches patterns like "PPVC 13_Typ", "PPVC 02_Typ 1", etc.
        /// </summary>
        private string ExtractPpvcName(string docTitle, string viewName)
        {
            // Try to match PPVC pattern in document title first
            // Patterns: "PPVC 13_Typ", "PPVC 02_Typ 1", "PPVC 01, 20_Typ", etc.
            
            string[] sources = { docTitle, viewName };
            
            foreach (string source in sources)
            {
                if (string.IsNullOrEmpty(source))
                    continue;
                    
                // Look for "PPVC" followed by number and optional suffix
                var match = System.Text.RegularExpressions.Regex.Match(
                    source, 
                    @"PPVC\s*[\d,\s]+_?[^\\/]*", 
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    
                if (match.Success)
                {
                    return match.Value.Trim();
                }
                
                // Simpler pattern: just "PPVC XX"
                match = System.Text.RegularExpressions.Regex.Match(
                    source, 
                    @"PPVC\s*\d+", 
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    
                if (match.Success)
                {
                    return match.Value.Trim();
                }
            }
            
            // If no pattern found, return document title as-is (for partial matching)
            return docTitle ?? "";
        }

        // ---------------- Folder picker -----------------

        private string PickOutputFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder to run PPVC GNN annotation";
                var result = dialog.ShowDialog();
                if (result == DialogResult.OK)
                    return dialog.SelectedPath;
            }
            return null;
        }

        // ---------------- Call batch file with all CSV paths -----------------

        private bool RunPythonInference(string nodesPath, string edgesPath, string annotationPath, string predPath)
        {
            const string pythonScriptPath = @"C:\Users\charm\Spatial GNN\model\automated_inference_workflow.py";
            const string condaPath = @"C:\Users\charm\anaconda3\condabin\conda.bat";
            const string condaEnv = "cling";
            const string modelFile = @"C:\Users\charm\Spatial GNN\model\greattrained_model.pth";

            // Verify all required files exist
            if (!File.Exists(pythonScriptPath))
            {
                TaskDialog.Show("PPVC Auto Annotate", "Python script not found:\n" + pythonScriptPath);
                return false;
            }

            if (!File.Exists(condaPath))
            {
                TaskDialog.Show("PPVC Auto Annotate", "Conda not found at:\n" + condaPath + "\n\nRun in PowerShell: where conda");
                return false;
            }

            if (!File.Exists(modelFile))
            {
                TaskDialog.Show("PPVC Auto Annotate", "Model file not found:\n" + modelFile);
                return false;
            }

            try
            {
                // Expected output files from Python with --save_cleaned flag
                string outputDir = Path.GetDirectoryName(predPath);
                string annotationWithTargetsPath = Path.Combine(outputDir, "annotation_with_targets.csv");
                string graphmlPath = Path.Combine(outputDir, "graph.graphml");

                // Create a temporary batch file that activates conda and runs Python
                string tempBatchPath = Path.Combine(Path.GetTempPath(), "ppvc_inference_" + Guid.NewGuid().ToString("N") + ".bat");

                // Build batch content using string concatenation to avoid escaping issues
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("@echo off");
                sb.AppendLine("setlocal enabledelayedexpansion");
                sb.AppendLine("chcp 65001 > nul");
                sb.AppendLine();
                sb.AppendLine("REM === PPVC Inference via Conda ===");
                sb.AppendLine();
                sb.AppendLine("set \"CONDA_PATH=" + condaPath + "\"");
                sb.AppendLine("set \"CONDA_ENV=" + condaEnv + "\"");
                sb.AppendLine("set \"PYTHON_SCRIPT=" + pythonScriptPath + "\"");
                sb.AppendLine("set \"MODEL_FILE=" + modelFile + "\"");
                sb.AppendLine();
                sb.AppendLine("set \"NODES_CSV=" + nodesPath + "\"");
                sb.AppendLine("set \"EDGES_CSV=" + edgesPath + "\"");
                sb.AppendLine("set \"PRED_CSV=" + predPath + "\"");
                sb.AppendLine("set \"ANN_CSV=" + annotationPath + "\"");
                sb.AppendLine();
                sb.AppendLine("REM Verify Conda exists");
                sb.AppendLine("if not exist \"%CONDA_PATH%\" (");
                sb.AppendLine("    echo ERROR: Conda not found at: %CONDA_PATH%");
                sb.AppendLine("    exit /b 2");
                sb.AppendLine(")");
                sb.AppendLine();
                sb.AppendLine("REM Activate conda environment");
                sb.AppendLine("call \"%CONDA_PATH%\" activate %CONDA_ENV%");
                sb.AppendLine("if errorlevel 1 (");
                sb.AppendLine("    echo ERROR: Failed to activate conda environment %CONDA_ENV%");
                sb.AppendLine("    exit /b 3");
                sb.AppendLine(")");
                sb.AppendLine();
                sb.AppendLine("REM Run Python script with --save_cleaned flag");
                sb.AppendLine("python \"%PYTHON_SCRIPT%\" ^");
                sb.AppendLine("  --nodes_csv \"%NODES_CSV%\" ^");
                sb.AppendLine("  --edges_csv \"%EDGES_CSV%\" ^");
                sb.AppendLine("  --out_csv \"%PRED_CSV%\" ^");
                sb.AppendLine("  --annotation_csv \"%ANN_CSV%\" ^");
                sb.AppendLine("  --model \"%MODEL_FILE%\" ^");
                sb.AppendLine("  --save_cleaned");
                sb.AppendLine();
                sb.AppendLine("set \"EXIT_CODE=%ERRORLEVEL%\"");
                sb.AppendLine("if %EXIT_CODE% neq 0 (");
                sb.AppendLine("    echo ERROR: Python script failed with exit code %EXIT_CODE%");
                sb.AppendLine("    exit /b %EXIT_CODE%");
                sb.AppendLine(")");
                sb.AppendLine();
                sb.AppendLine("echo SUCCESS: Inference complete");
                sb.AppendLine("exit /b 0");

                string batchContent = sb.ToString();

                // Write the temporary batch file
                File.WriteAllText(tempBatchPath, batchContent);

                // Execute the batch file
                var psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c \"" + tempBatchPath + "\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = System.Text.Encoding.UTF8,
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };

                using (var proc = Process.Start(psi))
                {
                    string stdout = proc.StandardOutput.ReadToEnd();
                    string stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();

                    // Clean up temporary batch file
                    try { File.Delete(tempBatchPath); }
                    catch { }

                    if (proc.ExitCode != 0)
                    {
                        string errorMsg = "Python inference failed with exit code " + proc.ExitCode + " (0x" + proc.ExitCode.ToString("X8") + ")\n\n";
                        
                        if (!string.IsNullOrWhiteSpace(stdout))
                            errorMsg += "Output:\n" + stdout + "\n\n";
                        
                        if (!string.IsNullOrWhiteSpace(stderr))
                            errorMsg += "Error:\n" + stderr + "\n\n";

                        errorMsg += "Input Files:\n";
                        errorMsg += "  Script: " + (File.Exists(pythonScriptPath) ? "OK" : "MISSING") + " " + pythonScriptPath + "\n";
                        errorMsg += "  Nodes CSV: " + (File.Exists(nodesPath) ? "OK" : "MISSING") + " " + nodesPath + "\n";
                        errorMsg += "  Edges CSV: " + (File.Exists(edgesPath) ? "OK" : "MISSING") + " " + edgesPath + "\n";
                        errorMsg += "  Annotation CSV: " + (File.Exists(annotationPath) ? "OK" : "MISSING") + " " + annotationPath + "\n";
                        errorMsg += "  Conda: " + (File.Exists(condaPath) ? "OK" : "MISSING") + " " + condaPath + "\n";
                        errorMsg += "  Model: " + (File.Exists(modelFile) ? "OK" : "MISSING") + " " + modelFile;

                        TaskDialog.Show("PPVC Auto Annotate", errorMsg);
                        return false;
                    }

                    // Verify all 3 output files were created
                    var missingFiles = new List<string>();
                    
                    if (!File.Exists(predPath))
                        missingFiles.Add("predictions.csv: " + predPath);
                    
                    if (!File.Exists(annotationWithTargetsPath))
                        missingFiles.Add("annotation_with_targets.csv: " + annotationWithTargetsPath);
                    
                    if (!File.Exists(graphmlPath))
                        missingFiles.Add("graph.graphml: " + graphmlPath);

                    if (missingFiles.Count > 0)
                    {
                        string warningMsg = "Python completed but some output files were not created:\n\n";
                        warningMsg += string.Join("\n", missingFiles);
                        warningMsg += "\n\nMake sure --save_cleaned flag is supported by Python script.\n\n";
                        warningMsg += "Python Output:\n" + stdout;
                        
                        // Show warning but continue if predictions.csv exists
                        if (!File.Exists(predPath))
                        {
                            TaskDialog.Show("PPVC Auto Annotate - Error", warningMsg);
                            return false;
                        }
                        else
                        {
                            // Just log warning, continue with predictions
                            Debug.WriteLine(warningMsg);
                        }
                    }

                    // Log success with file info
                    Debug.WriteLine("Python inference completed successfully:");
                    Debug.WriteLine("  predictions.csv: " + (File.Exists(predPath) ? "OK" : "MISSING"));
                    Debug.WriteLine("  annotation_with_targets.csv: " + (File.Exists(annotationWithTargetsPath) ? "OK" : "MISSING"));
                    Debug.WriteLine("  graph.graphml: " + (File.Exists(graphmlPath) ? "OK" : "MISSING"));

                    return true;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("PPVC Auto Annotate", "Error running Python inference:\n" + ex.Message);
                return false;
            }
        }

        // ---------------- Read predictions.csv -----------------

        private Dictionary<long, PredictionRecord> LoadPredictions(string csvPath)
        {
            var dict = new Dictionary<long, PredictionRecord>();
            var lines = File.ReadAllLines(csvPath);
            if (lines.Length <= 1)
                return dict;

            // header: node_id,predicted_class,confidence,annotation_type
            // where predicted_class is 0 (no_annotation), 1 (dimension), 2 (text), or 3 (both)
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split(',');
                if (parts.Length < 3)
                    continue;

                if (!long.TryParse(parts[0], out long nodeId))
                    continue;

                if (!int.TryParse(parts[1], out int predictedClass))
                    continue;

                double confidence = 0.0;
                double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out confidence);

                string annType = parts.Length > 3 ? parts[3] : "";

                dict[nodeId] = new PredictionRecord
                {
                    NodeId = nodeId,
                    PredictedClass = predictedClass,
                    Confidence = confidence,
                    AnnotationType = annType
                };
            }

            return dict;
        }

        // ---------------- Geometry helper for point -----------------

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

        // ---------------- RULES: what text to show -----------------

        private string BuildTextContent(Element e, NodeRecord nodeInfo)
        {
            double FtToMm(double ft) => ft * 304.8;

            if (e is Room rm)
            {
                string areaStr = $"{FtToMm(rm.Area):0} mm²";
                return $"{rm.Name} ({rm.Number})\nArea: {areaStr}";
            }

            string typeName = e.Document.GetElement(e.GetTypeId())?.Name ?? "";
            string category = e.Category?.Name ?? "";
            string familyName = "";

            if (e is FamilyInstance fi)
                familyName = fi.Symbol?.FamilyName ?? "";

            // Walls: show type + length + height + thickness
            if (e is Wall && nodeInfo != null)
            {
                string len = $"{FtToMm(nodeInfo.Length):0}mm";
                string h = $"{FtToMm(nodeInfo.Height):0}mm";
                string t = $"{FtToMm(nodeInfo.Thickness):0}mm";
                return $"{typeName}\nL={len}, H={h}, t={t}";
            }

            // Floors: thickness + area
            if (e is Floor && nodeInfo != null)
            {
                string t = $"{FtToMm(nodeInfo.Thickness):0}mm";
                string a = $"{FtToMm(nodeInfo.Area):0}mm²";
                return $"{typeName}\nThk={t}, A={a}";
            }

            // Skip text for Generic Models - they don't have meaningful labels
            if (e.Category != null && e.Category.Id.Value == (long)BuiltInCategory.OST_GenericModel)
            {
                return null;
            }

            // Generic fallback for other types
            return $"{category} - {typeName} {familyName}".Trim();
        }

        private bool CreateTextNoteForElement(
            Document doc,
            RevitView view,
            Element e,
            NodeRecord nodeInfo,
            PredictionRecord pred)
        {
            try
            {
                string content = BuildTextContent(e, nodeInfo);
                // Skip if no content (e.g., Generic Models)
                if (string.IsNullOrWhiteSpace(content))
                    return false;

                // Get element's bounding box for better positioning
                BoundingBoxXYZ bb = e.get_BoundingBox(null);
                if (bb == null)
                    return false;

                // Use the top-right corner of the bounding box as the base point
                // This naturally spreads annotations across different elements
                XYZ basePoint = new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z);

                // Get view-local directions for consistent offset
                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;
                XYZ viewOut = view.ViewDirection;

                // Calculate intelligent offset based on element size
                double elementSize = (bb.Max - bb.Min).GetLength();
                double offsetRight = Math.Max(2.0, elementSize * 0.15);
                double offsetUp = Math.Max(3.0, elementSize * 0.2);

                // Place text to the upper-right of the element's bounding box
                // This naturally separates annotations for different elements
                XYZ textPoint = basePoint + viewRight * offsetRight + viewUp * offsetUp + viewOut * 0.5;

                ElementId textTypeId =
                    doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
                if (textTypeId == ElementId.InvalidElementId)
                    return false;

                TextNote tn = TextNote.Create(doc, view.Id, textPoint, content, textTypeId);
                return tn != null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CreateTextNoteForElement failed: {ex.Message}");
                return false;
            }
        }

        // ---------------- RULES: where to place dimensions -----------------

        private bool CreateDimensionsForElement(
            Document doc,
            RevitView view,
            Element e,
            NodeRecord nodeInfo,
            PredictionRecord pred,
            Action<string> logToFile)
        {
            string elementInfo = $"{e.Category?.Name} ID={e.Id.Value} Name='{e.Name}'";
            
            // Rule 1: Walls → height dimension (bottom to top)
            if (e is Wall wall)
            {
                bool result = CreateWallHeightDimension(doc, view, wall, out string reason);
                logToFile($"  [WALL] {elementInfo}: {(result ? "✓ PLACED" : $"✗ FAILED - {reason}")}");
                return result;
            }

            // Rule 2: Floors → thickness dimension
            if (e is Floor floor)
            {
                bool result = CreateFloorThicknessDimension(doc, view, floor, out string reason);
                logToFile($"  [FLOOR] {elementInfo}: {(result ? "✓ PLACED" : $"✗ FAILED - {reason}")}");
                return result;
            }

            // Rule 3: Levels → elevation dimension (creates text, not dimension)
            if (e is Level level)
            {
                bool result = CreateLevelElevationDimension(doc, view, level, out string reason);
                logToFile($"  [LEVEL] {elementInfo}: {(result ? "✓ PLACED (text)" : $"✗ FAILED - {reason}")}");
                return result;
            }

            // Rule 4: Rooms → area or height dimension (creates text, not dimension)
            if (e is Room room)
            {
                bool result = CreateRoomDimension(doc, view, room, nodeInfo, out string reason);
                logToFile($"  [ROOM] {elementInfo}: {(result ? "✓ PLACED (text)" : $"✗ FAILED - {reason}")}");
                return result;
            }

            // Rule 5: Structural Framing → dimension
            if (e.Category != null && e.Category.Id.Value == (long)BuiltInCategory.OST_StructuralFraming)
            {
                bool result = CreateStructuralFramingDimension(doc, view, e, out string reason);
                logToFile($"  [FRAMING] {elementInfo}: {(result ? "✓ PLACED" : $"✗ FAILED - {reason}")}");
                return result;
            }

            // Rule 6: Generic Models → dimension
            if (e.Category != null && e.Category.Id.Value == (long)BuiltInCategory.OST_GenericModel)
            {
                bool result = CreateGenericModelDimension(doc, view, e, out string reason);
                logToFile($"  [GENERIC] {elementInfo}: {(result ? "✓ PLACED" : $"✗ FAILED - {reason}")}");
                return result;
            }

            logToFile($"  [UNKNOWN] {elementInfo}: ✗ SKIPPED - Unsupported element type");
            return false;
        }

        // Wall: dimension from bottom edge to top edge
        private bool CreateWallHeightDimension(Document doc, RevitView view, Wall wall, out string reason)
        {
            reason = "";
            try
            {
                var opt = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine
                };

                Solid solid = null;
                foreach (GeometryObject go in wall.get_Geometry(opt))
                {
                    if (go is Solid s && s.Faces.Size > 0 && s.Volume > 0)
                    {
                        solid = s;
                        break;
                    }
                }

                if (solid == null)
                {
                    reason = "No solid geometry found";
                    return false;
                }

                // Find top and bottom horizontal edges
                Edge topEdge = null;
                Edge bottomEdge = null;
                double maxZ = double.MinValue;
                double minZ = double.MaxValue;

                foreach (Edge e in solid.Edges)
                {
                    Curve c = e.AsCurve();
                    if (c is Line line)
                    {
                        XYZ ep0 = line.GetEndPoint(0);
                        XYZ ep1 = line.GetEndPoint(1);
                        
                        if (Math.Abs(ep0.Z - ep1.Z) < 0.01)
                        {
                            if (ep0.Z > maxZ)
                            {
                                maxZ = ep0.Z;
                                topEdge = e;
                            }
                            if (ep0.Z < minZ)
                            {
                                minZ = ep0.Z;
                                bottomEdge = e;
                            }
                        }
                    }
                }

                if (topEdge == null || bottomEdge == null)
                {
                    reason = "Could not find top/bottom edges";
                    return false;
                }

                var refs = new ReferenceArray();
                refs.Append(topEdge.Reference);
                refs.Append(bottomEdge.Reference);

                // Get wall centerline point
                XYZ wallPoint = XYZ.Zero;
                if (wall.Location is LocationCurve lc && lc.Curve is Line wallLine)
                {
                    wallPoint = lc.Curve.Evaluate(0.5, true);
                }

                // Calculate wall width for intelligent offset
                BoundingBoxXYZ wallBB = wall.get_BoundingBox(null);
                double wallWidth = (wallBB.Max - wallBB.Min).GetLength();
                double offsetDistance = Math.Max(5.0, wallWidth * 0.5);

                // Create dimension line in view-local coordinates
                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;
                
                XYZ midHeight = wallPoint + (topEdge.Evaluate(0.5) - bottomEdge.Evaluate(0.5)) * 0.5;
                XYZ p0 = midHeight + viewRight * offsetDistance - viewUp * (maxZ - minZ) * 0.6;
                XYZ p1 = midHeight + viewRight * offsetDistance + viewUp * (maxZ - minZ) * 0.6;

                Line dimLine = Line.CreateBound(p0, p1);

                Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
                if (dim == null)
                {
                    reason = "NewDimension returned null";
                    return false;
                }
                
                reason = $"Height={maxZ - minZ:F2}ft at ({p0.X:F1},{p0.Y:F1})";
                return true;
            }
            catch (Exception ex)
            {
                reason = $"Exception: {ex.Message}";
                return false;
            }
        }

        // Floor: dimension from bottom edge to top edge (thickness)
        private bool CreateFloorThicknessDimension(Document doc, RevitView view, Floor floor, out string reason)
        {
            reason = "";
            try
            {
                var opt = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine
                };

                Solid solid = null;
                foreach (GeometryObject go in floor.get_Geometry(opt))
                {
                    if (go is Solid s && s.Faces.Size > 0 && s.Volume > 0)
                    {
                        solid = s;
                        break;
                    }
                }

                if (solid == null)
                {
                    reason = "No solid geometry found";
                    return false;
                }

                // Find top and bottom horizontal edges
                Edge topEdge = null;
                Edge bottomEdge = null;
                double maxZ = double.MinValue;
                double minZ = double.MaxValue;

                foreach (Edge e in solid.Edges)
                {
                    Curve c = e.AsCurve();
                    if (c is Line line)
                    {
                        XYZ ep0 = line.GetEndPoint(0);
                        XYZ ep1 = line.GetEndPoint(1);
                        
                        if (Math.Abs(ep0.Z - ep1.Z) < 0.01)
                        {
                            if (ep0.Z > maxZ)
                            {
                                maxZ = ep0.Z;
                                topEdge = e;
                            }
                            if (ep0.Z < minZ)
                            {
                                minZ = ep0.Z;
                                bottomEdge = e;
                            }
                        }
                    }
                }

                if (topEdge == null || bottomEdge == null)
                {
                    reason = "Could not find top/bottom edges";
                    return false;
                }

                var refs = new ReferenceArray();
                refs.Append(topEdge.Reference);
                refs.Append(bottomEdge.Reference);

                BoundingBoxXYZ floorBB = floor.get_BoundingBox(null);
                XYZ floorCenter = (floorBB.Min + floorBB.Max) * 0.5;

                double floorSize = (floorBB.Max - floorBB.Min).GetLength();
                double offsetDistance = Math.Max(5.0, floorSize * 0.3);

                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;
                
                double thickness = maxZ - minZ;
                XYZ p0 = floorCenter + viewRight * offsetDistance - viewUp * thickness * 1.5;
                XYZ p1 = floorCenter + viewRight * offsetDistance + viewUp * thickness * 1.5;

                Line dimLine = Line.CreateBound(p0, p1);

                Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
                if (dim == null)
                {
                    reason = "NewDimension returned null";
                    return false;
                }
                
                reason = $"Thickness={thickness:F2}ft at ({p0.X:F1},{p0.Y:F1})";
                return true;
            }
            catch (Exception ex)
            {
                reason = $"Exception: {ex.Message}";
                return false;
            }
        }

        // Level: dimension with arrows
        private bool CreateLevelElevationDimension(Document doc, RevitView view, Level level, out string reason)
        {
            reason = "";
            try
            {
                double elevation = level.Elevation;
                
                XYZ levelPoint = new XYZ(0, 0, elevation);
                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;
                
                double FtToMm(double ft) => ft * 304.8;
                string elevText = $"{FtToMm(elevation):F0}";
                
                XYZ annotationPoint = levelPoint + viewRight * 3.0 + viewUp * 1.0;
                XYZ leaderEndPoint = levelPoint + viewRight * 0.5;
                
                bool result = CreateAnnotationWithArrows(doc, view, elevText, annotationPoint, leaderEndPoint);
                reason = result ? $"Elevation={elevation:F2}ft" : "TextNote creation failed";
                return result;
            }
            catch (Exception ex)
            {
                reason = $"Exception: {ex.Message}";
                return false;
            }
        }

        // Room: dimension with arrows
        private bool CreateRoomDimension(Document doc, RevitView view, Room room, NodeRecord nodeInfo, out string reason)
        {
            reason = "";
            try
            {
                IList<IList<BoundarySegment>> segs = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
                
                if (segs.Count == 0)
                {
                    reason = "No boundary segments";
                    return false;
                }

                if (segs[0].Count == 0)
                {
                    reason = "Empty boundary segment list";
                    return false;
                }

                BoundarySegment firstSeg = segs[0][0];
                Curve curve = firstSeg.GetCurve();
                
                if (curve == null)
                {
                    reason = "Boundary curve is null";
                    return false;
                }

                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;
                
                double FtToMm(double ft) => ft * 304.8;
                string roomText = $"Area: {FtToMm(room.Area):F0}mm²";
                
                if (nodeInfo != null && nodeInfo.Height > 0)
                {
                    roomText += $"\nHeight: {FtToMm(nodeInfo.Height):F0}mm";
                }
                
                XYZ midPoint = (curve.GetEndPoint(0) + curve.GetEndPoint(1)) * 0.5;
                XYZ annotationPoint = midPoint + viewUp * 2.5 + viewRight * 2.0;
                XYZ leaderEndPoint = midPoint;
                
                bool result = CreateAnnotationWithArrows(doc, view, roomText, annotationPoint, leaderEndPoint);
                reason = result ? $"Area={room.Area:F2}sqft" : "TextNote creation failed";
                return result;
            }
            catch (Exception ex)
            {
                reason = $"Exception: {ex.Message}";
                return false;
            }
        }

        // Helper to create a text annotation (used for Levels and Rooms)
        private bool CreateAnnotationWithArrows(
            Document doc,
            RevitView view,
            string text,
            XYZ annotationPoint,
            XYZ leaderEndPoint)
        {
            try
            {
                ElementId textTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);
                if (textTypeId == ElementId.InvalidElementId)
                    return false;

                TextNote tn = TextNote.Create(doc, view.Id, annotationPoint, text, textTypeId);
                return tn != null;
            }
            catch
            {
                return false;
            }
        }

        // Structural Framing: create dimension
        private bool CreateStructuralFramingDimension(Document doc, RevitView view, Element framing, out string reason)
        {
            reason = "";
            try
            {
                if (!(framing.Location is LocationCurve lc) || lc.Curve == null)
                {
                    reason = "No location curve";
                    return false;
                }

                Curve curve = lc.Curve;
                XYZ p0 = curve.GetEndPoint(0);
                XYZ p1 = curve.GetEndPoint(1);
                
                var refs = new ReferenceArray();
                
                var opt = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
                GeometryElement geom = framing.get_Geometry(opt);
                
                int edgeCount = 0;
                foreach (GeometryObject obj in geom)
                {
                    if (obj is Solid solid)
                    {
                        foreach (Edge edge in solid.Edges)
                        {
                            refs.Append(edge.Reference);
                            edgeCount++;
                            if (edgeCount >= 2) break;
                        }
                        if (edgeCount >= 2) break;
                    }
                }
                
                if (refs.Size < 2)
                {
                    reason = $"Only found {refs.Size} edge references (need 2)";
                    return false;
                }

                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;
                
                XYZ dimP0 = p0 + viewRight * 3.0 - viewUp * 2.0;
                XYZ dimP1 = p1 + viewRight * 3.0 + viewUp * 2.0;
                
                Line dimLine = Line.CreateBound(dimP0, dimP1);
                
                Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
                if (dim == null)
                {
                    reason = "NewDimension returned null";
                    return false;
                }
                
                reason = $"Length={curve.Length:F2}ft";
                return true;
            }
            catch (Exception ex)
            {
                reason = $"Exception: {ex.Message}";
                return false;
            }
        }

        // Generic Model: create actual dimension arrows
        private bool CreateGenericModelDimension(Document doc, RevitView view, Element genericModel, out string reason)
        {
            reason = "";
            try
            {
                var opt = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
                GeometryElement geom = genericModel.get_Geometry(opt);
                
                if (geom == null)
                {
                    reason = "No geometry found";
                    return false;
                }

                var refs = new ReferenceArray();
                BoundingBoxXYZ bb = null;
                
                foreach (GeometryObject obj in geom)
                {
                    if (obj is Solid solid && solid.Faces.Size > 0)
                    {
                        bb = solid.GetBoundingBox();
                        
                        int edgeCount = 0;
                        foreach (Edge edge in solid.Edges)
                        {
                            refs.Append(edge.Reference);
                            edgeCount++;
                            if (edgeCount >= 2) break;
                        }
                        if (refs.Size >= 2) break;
                    }
                }
                
                if (bb == null)
                {
                    reason = "No bounding box from solid";
                    return false;
                }
                
                if (refs.Size < 2)
                {
                    reason = $"Only found {refs.Size} edge references (need 2)";
                    return false;
                }

                XYZ center = (bb.Min + bb.Max) * 0.5;
                XYZ viewUp = view.UpDirection;
                XYZ viewRight = view.RightDirection;
                
                XYZ dimP0 = center - viewUp * 3.0 - viewRight * 2.0;
                XYZ dimP1 = center - viewUp * 3.0 + viewRight * 2.0;
                
                Line dimLine = Line.CreateBound(dimP0, dimP1);
                Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
                
                if (dim == null)
                {
                    reason = "NewDimension returned null";
                    return false;
                }
                
                XYZ size = bb.Max - bb.Min;
                reason = $"Size=({size.X:F2},{size.Y:F2},{size.Z:F2})ft";
                return true;
            }
            catch (Exception ex)
            {
                reason = $"Exception: {ex.Message}";
                return false;
            }
        }
    }
}
