import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

# --- 1. CONFIGURATION (The Control Panel) ---
EXCEL_FILE = 'list.xlsx'
TEMPLATE_FILE = 'certificate_template.png'
OUTPUT_FOLDER = 'output'

# UPDATED: Pointing to the Italic version of the font
FONT_PATH = 'texgyretermes-italic.otf' 

FONT_SIZE = 80           
TEXT_COLOR = (0, 0, 0)      # Black
TEXT_Y_POSITION = 590     # Vertical height (Change this based on your design)
LETTER_SPACING = 13        # Space between letters

# --- 2. SETUP (Preparing the environment) ---
# Create the output folder if it doesn't exist
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# Load the Excel data
print("Reading Excel file...")
df = pd.read_excel(EXCEL_FILE)

# Load the Font
try:
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
except IOError:
    print(f"Error: Could not find {FONT_PATH}. Please make sure the file is in the folder.")
    # Fallback to default if font is missing
    font = ImageFont.load_default()

# --- 3. HELPER FUNCTION (Custom Logic) ---
def get_text_width_with_spacing(text, font, spacing):
    """
    Calculates the total width of the text including the custom spacing.
    """
    total_width = 0
    for char in text:
        # Get the bounding box of the character
        bbox = font.getbbox(char)
        # Width = Right edge - Left edge
        char_width = bbox[2] - bbox[0]
        total_width += char_width + spacing
    
    # We remove the spacing from the very last letter because it's not needed
    return total_width - spacing

# --- 4. THE MAIN LOOP (Doing the work) ---
print("Starting certificate generation...")

for index, row in df.iterrows():
    # Get the name and clean it up (remove extra spaces)
    name = str(row['Name']).strip()
    
    # 1. Open the blank template freshly for this person
    image = Image.open(TEMPLATE_FILE)
    draw = ImageDraw.Draw(image)
    
    # 2. Calculate where to start writing so it is CENTERED
    image_width, image_height = image.size
    text_width = get_text_width_with_spacing(name, font, LETTER_SPACING)
    
    # The Math: (Total Image Width - Name Width) / 2 = Starting X Position
    current_x = (image_width - text_width) / 2
    
    # 3. Draw the name letter by letter
    for char in name:
        draw.text((current_x, TEXT_Y_POSITION), char, fill=TEXT_COLOR, font=font)
        
        # Move the "cursor" to the right for the next letter
        bbox = font.getbbox(char)
        char_width = bbox[2] - bbox[0]
        current_x += char_width + LETTER_SPACING
    
    # 4. Save as PDF
    # Convert from RGBA (Image) to RGB (Document)
    image_rgb = image.convert('RGB')
    
    # Create a safe filename (remove symbols that might crash Windows)
    clean_filename = "".join([c for c in name if c.isalpha() or c.isdigit() or c==' ']).rstrip()
    save_path = os.path.join(OUTPUT_FOLDER, f"{clean_filename}.pdf")
    
    image_rgb.save(save_path)
    print(f"Generated: {save_path}")

print("All Done!")