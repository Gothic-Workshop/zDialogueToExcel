# zDialogueToExcel
**zDialogueToExcel** is a simple utility tool (written using ChatGPT in Python) for extracting, organizing, and formatting dialogue data from Gothic engine `.d` script files. It converts in-game conversations into structured, multi-sheet Excel spreadsheets, making them easy to review, annotate, and edit.

You can find the example of the program being used [here, on my Google Drive](https://drive.google.com/drive/folders/1r7uP4U5n9Nm6krZx19DohA_cF-lkkouU?usp=drive_link).

## ✅ Features
### 🗂️ Input Processing
- Supports multiple input directories (`Input_Gothic`, `Input_Gothic2`, `Input_NotR`)
- Automatically groups dialogue lines by character
- Handles multiple dialogue files for the same character
- Merges same-named `.d` files across folders (e.g. `DIA_Diego.d` from Gothic2 and NotR)

### 📑 Spreadsheet Generation
- One `.xlsx` file per character, with each input folder as a separate sheet
- Filename is based on the dialogue file name (e.g. `DIA_BDT_1081_Wache.d` → `BDT_1081_Wache.xlsx`)
- Custom mappings supported via `config.yaml` (optional)

### 📄 Excel Structure
Each spreadsheet contains the following columns:
| Column             | Description                            |
|--------------------|----------------------------------------|
| `Output Name`      | Tag string used in AI_Output           |
| `What character says` | The spoken dialogue from the script |
| `Who says it`      | The speaker in the line (e.g. `self`)  |
| `Chapter`          | (empty, for manual use)                |
| `Condition`        | (empty, for manual use)                |

### 🎨 Formatting
- Column widths match a manual reference layout
- Header row:
  - Bold text
  - Light gray background
- Cell formatting:
  - Column A: bold
  - Column B: italic
  - Column C: rows highlighted in **green** if spoken by `self`, **blue** if by `hero` or `other`
- All cells:
  - Vertically centered
  - Wrapped text
  - Dotted borders for all inner cells
  - Thick border around the outer edge of the table
- Frozen header row for easy navigation

# 🚀 Usage and 📌 Requirements
- Python 3.7+
- pandas
- openpyxl
- pyyaml

Make sure you have the required libraries:
```bash
pip install pandas openpyxl pyyaml
```
Then run the script:
```bash
python extract_dialogues_formatted.py
```
All output .xlsx files will appear in the Output/ folder.

---

## ⚙️ Optional: `config.yaml`
You can optionally create a `config.yaml` to manually map `.d` filenames to character names and/or for merging dialogue:
```yaml
# config.yaml
DIA_Diego.d: Diego
DIA_BDT_1081_Wache.d: Wache
```

## 💡 Tip
Use Google Drive’s import option to convert .xlsx files into Google Sheets format for online collaboration.