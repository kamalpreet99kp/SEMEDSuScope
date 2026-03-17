import cv2
import numpy as np
import os
from skimage.metrics import structural_similarity as ssim
import shutil

# ==== Paths ====
input_folder = r"H:\Test\Duplicates in Au Analysis\Au Grains\fixed_images"
unique_folder = r"H:\Test\Duplicates in Au Analysis\Au Grains\unique"
duplicates_folder = r"H:\Test\Duplicates in Au Analysis\Au Grains\duplicates"

os.makedirs(unique_folder, exist_ok=True)
os.makedirs(duplicates_folder, exist_ok=True)

# ==== BGR range for Au (from your values) ====
lower_bgr = np.array([15, 90, 215], dtype=np.uint8)
upper_bgr = np.array([255, 255, 255], dtype=np.uint8)

# Similarity threshold (0.0 to 1.0)
SIMILARITY_THRESHOLD = 0.20


# ==== Helper: Extract Au mask ====
def get_au_mask(image):
    mask = cv2.inRange(image, lower_bgr, upper_bgr)
    return mask


# ==== Helper: Compare two images ====
def compare_images(img1, img2):
    # Resize to same size
    h = max(img1.shape[0], img2.shape[0])
    w = max(img1.shape[1], img2.shape[1])
    img1_resized = cv2.resize(img1, (w, h))
    img2_resized = cv2.resize(img2, (w, h))

    score = ssim(img1_resized, img2_resized)
    return score


# ==== Main process ====
processed = set()
files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.tif'))]

for i, file1 in enumerate(files):
    if file1 in processed:
        continue

    path1 = os.path.join(input_folder, file1)
    img1 = cv2.imread(path1)
    mask1 = get_au_mask(img1)

    group = [file1]
    processed.add(file1)

    for j, file2 in enumerate(files):
        if file2 in processed or file1 == file2:
            continue

        path2 = os.path.join(input_folder, file2)
        img2 = cv2.imread(path2)
        mask2 = get_au_mask(img2)

        similarity = compare_images(mask1, mask2)

        if similarity > SIMILARITY_THRESHOLD:
            group.append(file2)
            processed.add(file2)

    # Save the first image in the group as unique
    best_img = group[0]
    shutil.copy(os.path.join(input_folder, best_img), os.path.join(unique_folder, best_img))

    # Save the rest as duplicates
    for dup in group[1:]:
        shutil.copy(os.path.join(input_folder, dup), os.path.join(duplicates_folder, dup))

print("✅ DONE")
print(f"Unique images saved: {len(os.listdir(unique_folder))}")
print(f"Duplicates saved: {len(os.listdir(duplicates_folder))}")
