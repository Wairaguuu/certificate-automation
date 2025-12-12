Markdown

# Certificate Automation Tool 🎓

A Python automation script that generates bulk PDF certificates using names from an Excel sheet and placing them onto a PNG template.

## 🚀 Features
* **Bulk Generation:** Creates hundreds of certificates in seconds.
* **Smart Centering:** Automatically calculates the center of the image based on the length of each name.
* **Custom Typography:** Supports custom `.otf` or `.ttf` fonts with adjustable kerning (letter spacing).
* **PDF Output:** Converts high-quality PNG templates into ready-to-print PDF files.

## 🛠️ Built With
* **Python 3.x**
* **Pandas** (Data management)
* **Pillow (PIL)** (Image processing)

## 📂 Project Structure
```text
├── generate_certs.py      # The main script
├── list.xlsx              # Excel file containing names (Column header: "Name")
├── certificate_template.png   # Blank certificate design
├── texgyretermes-italic.otf   # Custom font file
└── output/                # Folder where PDFs are saved
⚙️ How to Run
Install Dependencies:

Bash

pip install pandas openpyxl pillow
Prepare your Data:

Add names to list.xlsx under a column named "Name".

Ensure your template image is named certificate_template.png.

Run the Script:

Bash

python generate_certs.py
Check Output:

The certificates will appear in the output/ folder.

📝 Customization
You can adjust the configuration at the top of generate_certs.py to fit your specific template design:

Python

FONT_SIZE = 150            # Size of the text
TEXT_Y_POSITION = 651      # Vertical position of the name
LETTER_SPACING = 10        # Space between letters
