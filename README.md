# 📄 Bulk PowerPoint (.pptx) to PDF Converter

This script batch converts all `.pptx` files in a directory and its subfolders into `.pdf` files, saving each PDF in the **same location as the original PowerPoint file**.

### ⚙️ How It Works

- Searches for `.pptx` files recursively in all subdirectories.
- Uses **Microsoft PowerPoint** (via COM automation) to convert each presentation to a `.pdf`.
- Saves the `.pdf` in the same folder as the original `.pptx`.

---

## 🛠 Requirements

- ✅ **Windows OS**
- ✅ **Microsoft PowerPoint installed** (Office 2010 or later)
- ❌ Will not work on macOS or Linux (due to COM dependency)

---

## 📦 Files Included

| File                    | Description                                 |
|-------------------------|---------------------------------------------|
| `convert_pptx_to_pdf.bat` | Main batch script to start the conversion   |
| `pptx_to_pdf.vbs`         | VBScript used to automate PowerPoint export |

---

## 🚀 How to Use

1. Place both `convert_pptx_to_pdf.bat` and `pptx_to_pdf.vbs` in the **root folder** where your `.pptx` files (and subfolders) are located.
2. Double-click `convert_pptx_to_pdf.bat` to begin.
3. All `.pptx` files will be converted to `.pdf` and saved alongside the originals.

---

## 📝 Example

Assume your folder looks like:

```
📁 Presentations
├──📁 Reports
│   └── Q1_Report.pptx
├──📁 Meetings
│   └── Team_Meeting.pptx
└── convert_pptx_to_pdf.bat
```

After running the script:

```
📁 Presentations
├──📁 Reports
│   ├── Q1_Report.pptx
│   └── Q1_Report.pdf
├──📁 Meetings
│   ├── Team_Meeting.pptx
│   └── Team_Meeting.pdf
```

---

## ❗ Notes

- All existing `.pdf` files with the same name will be overwritten.
- Make sure PowerPoint is **not running** while running this script for better stability.
- Large folders may take some time to process.

---

## 📃 License

This script is released under the [MIT License](LICENSE). Feel free to modify or distribute it.

---
