# Direct Python Execution Guide - CORRECTED FOR YOUR SYSTEM

## Fixed Issue

The error "Conda not found at: C:\Users\charm\anaconda3\Scripts\conda.bat" has been fixed.

### ? Correct Conda Path
```
C:\Users\charm\anaconda3\condabin\conda.bat
```

NOT: `C:\Users\charm\anaconda3\Scripts\conda.bat`

This path has been updated in your C# code.

---

## What Changed in Your Code

### Current Settings (Updated)

```csharp
const string pythonScriptPath = @"C:\Users\charm\Spatial GNN\model\automated_inference_workflow.py";
const string condaPath = @"C:\Users\charm\anaconda3\condabin\conda.bat";  // ? FIXED
const string condaEnv = "cling";  // ?? IMPORTANT: Verify this is correct
```

---

## ?? IMPORTANT: Verify Your Conda Environment

I found these environments on your system:
```
base                     C:\Users\charm\anaconda3
cling                    C:\Users\charm\anaconda3\envs\cling
```

**The code currently uses `"cling"`** - but you need to verify this is the environment where you installed PyTorch, numpy, and your ML dependencies.

### To check which environment has your packages:

```powershell
# In PowerShell, activate the environment and check:
conda activate cling
pip list

# Look for: torch, pytorch-geometric, numpy, pandas, networkx, etc.
```

### If cling is NOT the correct environment:

You can either:
1. **Install packages in cling**: `conda activate cling` then `pip install torch pytorch-geometric ...`
2. **OR** Change the C# code to use base:
   ```csharp
   const string condaEnv = "base";
   ```

---

## How It Works Now

The updated method creates a temporary batch file that:

1. ? Uses the **correct conda path** (`condabin\conda.bat`)
2. ? Changes directory to your model folder
3. ? Activates your conda environment
4. ? Runs Python with your script and arguments
5. ? Cleans up the temporary batch file afterward

```batch
@echo off
chcp 65001 > nul
cd /d "C:\Users\charm\Spatial GNN\model"
call "C:\Users\charm\anaconda3\condabin\conda.bat" activate cling
python "C:\Users\charm\Spatial GNN\model\automated_inference_workflow.py" --nodes_csv "path\to\nodes.csv" --edges_csv "path\to\edges.csv" --out_csv "path\to\predictions.csv" --model greattrained_model.pth
exit /b %ERRORLEVEL%
```

---

## Next Steps

1. ? **Verify your conda environment has all packages**:
   ```powershell
   conda activate cling
   pip list | findstr /i "torch numpy"
   ```

2. **If packages are missing**, install them:
   ```powershell
   conda activate cling
   pip install torch pytorch-geometric numpy pandas networkx
   ```

3. **Test the Python script directly** (before running from Revit):
   ```powershell
   conda activate cling
   cd "C:\Users\charm\Spatial GNN\model"
   python automated_inference_workflow.py --nodes_csv "test.csv" --edges_csv "test.csv" --out_csv "out.csv" --model greattrained_model.pth
   ```

4. **If it works in PowerShell**, it will work from Revit!

---

## Summary of Fixes

| Item | Was | Now | Status |
|------|-----|-----|--------|
| Conda Path | `anaconda3\Scripts\conda.bat` | `anaconda3\condabin\conda.bat` | ? Fixed |
| Conda Env | `spatial_gnn` | `cling` | ?? Verify |
| Build | N/A | Successful | ? OK |

---

## If You Get "Conda Not Found" Again

Run this in PowerShell to verify the conda path:
```powershell
Test-Path "C:\Users\charm\anaconda3\condabin\conda.bat"
```

Should return: `True`

If it returns `False`, conda is installed elsewhere. Run:
```powershell
where conda
```

And update the path in your C# code accordingly.
