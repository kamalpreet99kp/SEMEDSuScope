import os
import glob
from PIL import Image
from tkinter import Tk, filedialog

# --- Select folder with a popup ---
root = Tk()
root.withdraw()  # hide main tkinter window
folder = filedialog.askdirectory(title="Select Folder Containing Images")
if not folder:
    raise SystemExit("❌ No folder selected. Exiting...")

# --- Create output folder ---
output_folder = os.path.join(folder, "rotated")
os.makedirs(output_folder, exist_ok=True)

# --- Rotate all JPG & PNG images by 135° ---
count = 0
for ext in ("*.jpg", "*.jpeg", "*.png"):
    for file in glob.glob(os.path.join(folder, ext)):
        img = Image.open(file)
        rotated = img.rotate(45, expand=True)
        rotated.save(os.path.join(output_folder, os.path.basename(file)))
        count += 1

print(f"✅ Done! {count} images rotated by 135° and saved to:\n{output_folder}")
