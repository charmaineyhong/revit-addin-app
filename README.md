# Revit Add-in for Spatial Design Studio Auto Annotator

This repository contains a Revit add-in application for exporting model data in PPVC (Point-Process Virtual Construction) format. The add-in collects nodes, edges, and annotations from Autodesk Revit models and exports them to CSV files for use in downstream graph-based machine learning tasks.

## Features

- **Node Export**: Extracts model elements with properties like category, family, bounding box, and geometric attributes.
- **Edge Export**: Identifies relationships between elements, such as adjacency or connectivity.
- **Annotation Export**: Collects dimensions and text notes, with options for whitelist-based filtering or smart mode.
- **CSV Output**: Generates `nodes.csv`, `edges.csv`, and `annotation.csv` files.
- **Auto Annotation**: Command to apply predictions back to the model as annotations.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/charmaineyhong/revit-addin-app.git
   cd revit-addin-app
   ```

2. Build the project using Visual Studio or your preferred C# IDE.
3. Copy the compiled DLL to your Revit add-ins folder (e.g., `C:\Users\charmaineyhong\AppData\Roaming\Autodesk\Revit\Addins\2025` - adjust the version number for your Revit installation).
4. Ensure `annotation_whitelist.csv` is placed in the appropriate directory if using whitelist mode.

## User-Specific Configurations

- **Add-in Folder Path**: The default Revit add-ins folder is `C:\Users\<username>\AppData\Roaming\Autodesk\Revit\Addins\<version>`. Replace `<username>` with your Windows username (e.g., `charmaineyhong`) and `<version>` with your Revit version (e.g., `2025`).
- **Whitelist Path**: The add-in looks for `annotation_whitelist.csv` in the assembly directory or project folder. Place your whitelist CSV in the same directory as the DLL or update the code in `PPVCExportCommand.cs` to point to a custom path like `C:\Users\charmaineyhong\Documents\whitelist.csv`.
- **Output Directory**: When exporting, select a user-specific folder. For automation, hardcode a path in the code if needed, e.g., change `PickOutputFolder()` to default to `C:\Users\charmaineyhong\Exports\`.
- **Diagnostic Logs**: Logs are written to `C:\Users\charmaineyhong\AppData\Roaming\PPVCRevitAddin\whitelist_loading_debug.txt`. Ensure the folder exists or adjust the path in the code.
- **Model Matching**: For specific PPVC names, ensure your Revit project title matches the entries in the whitelist CSV. Edit the CSV to include your project names.

## Usage

1. Open a Revit project.
2. Run the "PPVC Export" command from the add-in ribbon.
3. Select an output folder.
4. The add-in will export nodes, edges, and annotations to CSV files.

For auto-annotation after predictions:
- Place `predictions.csv` in the project directory.
- Run the "PPVC Auto Annotate" command.

## Integrated Workflow with Python Model

This add-in works seamlessly with the [spatial_gnn_automation](https://github.com/charmaineyhong/spatial_gnn_automation) Python repository for automated spatial GNN predictions:

1. **Export Data**: In Revit, use the "PPVC Export" command to export `nodes.csv`, `edges.csv`, and `annotation.csv` from your model. These files represent the graph structure and annotations of the BIM data.

2. **Run GNN Inference**: Transfer the exported CSVs to your Python environment. Use scripts from spatial_gnn_automation to run inference with a trained GAT model, e.g.:
   ```bash
   python automated_inference_workflow.py --nodes_csv nodes.csv --edges_csv edges.csv --annotation_csv annotation.csv --out_csv predictions.csv
   ```
   Or use the batch file `run_ppvc_inference.bat nodes.csv edges.csv predictions.csv` after updating the paths in the batch file as described above.

3. **Import Predictions**: Copy the resulting `predictions.csv` back to your Revit project directory. Run the "PPVC Auto Annotate" command in Revit to automatically apply these predictions as model annotations, completing the labeling workflow.

This integration enables BIM model enrichment through machine learning, automating the annotation of spatial elements.

## Requirements

- Autodesk Revit (tested on recent versions)
- .NET Framework
