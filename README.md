
# PPTX-Templater  

🚀 **Python tool for automated PowerPoint template processing**  
*Replace placeholders in PPTX templates while preserving all formatting*  

---

## Features  
✔ Replace `{{placeholders}}` in PowerPoint slides  
✔ Maintain original fonts, colors and styles  
✔ Works with both slides and speaker notes  
✔ Simple Python API  

---

## Quick Start  

1. First install the required package:
```bash
pip install python-pptx
```

2. Create your PowerPoint template (`template.pptx`) with placeholders:
```
{{title}}
{{presenter}}
{{date}}
```

3. Download `app.py` and edit the data dictionary:
```python
data = {
    "title": "Quarterly Report Q3 2023",
    "presenter": "Jane Doe", 
    "date": "September 30, 2023"
}
```

4. Run the script:
```bash 
python app.py
```

This will generate `output.pptx` with your content.

---

## How It Works

The script:
1. Scans all text boxes in your template
2. Replaces any `{{placeholder}}` with your values
3. Preserves all original formatting
4. Saves the customized presentation

---

## Customization

Edit these variables in `app.py`:
```python
# Path to your template
TEMPLATE_PATH = "template.pptx"  

# Where to save the result  
OUTPUT_PATH = "output.pptx"

# Your data to insert
data = {
    "placeholder1": "Value 1",
    "placeholder2": "Value 2"
}
```

---

## Requirements
- Python 3.6+
- python-pptx library

---

## License
MIT License - free for both personal and commercial use.

---

⭐ **If you find this useful, please star the repository!**
