import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

crop_value = 0.9  # Fraction of original width/height to keep
SUPPORTED_EXTENSIONS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp")
SAMPLE_CLICKS_PER_SUBFOLDER = 3


class BatchSubfolderCropper:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Crop to Center - All Images at once")
        self.root.geometry("1100x850")

        self.base_folder = filedialog.askdirectory(title="Select parent folder containing subfolders")
        if not self.base_folder:
            print("No parent folder selected. Exiting.")
            self.root.destroy()
            return

        self.selected_subfolders = self.select_subfolders_gui()
        if not self.selected_subfolders:
            print("No subfolders selected. Exiting.")
            self.root.destroy()
            return

        self.subfolder_jobs = self.build_subfolder_jobs(self.selected_subfolders)
        if not self.subfolder_jobs:
            messagebox.showwarning("No images", "No supported image files were found in selected subfolders.")
            self.root.destroy()
            return

        self.crop_percent_w = crop_value
        self.crop_percent_h = crop_value

        self.canvas = tk.Canvas(self.root, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self.on_click)

        self.info_label = tk.Label(self.root, text="", anchor="w", justify="left")
        self.info_label.pack(fill=tk.X)

        # Click collection state
        self.folder_index = 0
        self.sample_index = 0
        self.last_click_per_folder = {}  # subfolder_path -> (x, y)

        self.original_image = None
        self.display_image = None
        self.tk_image = None

        self.show_current_sample_image()
        self.root.mainloop()

    def select_subfolders_gui(self):
        subfolders = []
        for name in sorted(os.listdir(self.base_folder)):
            p = os.path.join(self.base_folder, name)
            if os.path.isdir(p):
                subfolders.append(p)

        if not subfolders:
            messagebox.showwarning("No subfolders", "No subfolders found in selected folder.")
            return []

        dialog = tk.Toplevel(self.root)
        dialog.title("Select subfolders to process")
        dialog.geometry("600x550")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = tk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        canvas = tk.Canvas(frame)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        vars_by_folder = {}
        for folder in subfolders:
            var = tk.BooleanVar(value=False)
            vars_by_folder[folder] = var
            tk.Checkbutton(inner, text=os.path.basename(folder), variable=var, anchor="w", justify="left").pack(fill=tk.X, padx=4, pady=2)

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)

        def select_all():
            for v in vars_by_folder.values():
                v.set(True)

        def clear_all():
            for v in vars_by_folder.values():
                v.set(False)

        tk.Button(btn_frame, text="Select All", command=select_all).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_frame, text="Clear All", command=clear_all).pack(side=tk.LEFT, padx=4)

        result = {"folders": []}

        def confirm():
            result["folders"] = [f for f, v in vars_by_folder.items() if v.get()]
            dialog.destroy()

        tk.Button(btn_frame, text="Process Selected", command=confirm).pack(side=tk.RIGHT, padx=4)

        self.root.wait_window(dialog)
        return result["folders"]

    def build_subfolder_jobs(self, folders):
        jobs = []
        for folder in folders:
            image_files = [
                f for f in sorted(os.listdir(folder))
                if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith(SUPPORTED_EXTENSIONS)
            ]
            if not image_files:
                continue
            sample_count = min(SAMPLE_CLICKS_PER_SUBFOLDER, len(image_files))
            jobs.append({
                "folder": folder,
                "files": image_files,
                "sample_files": image_files[:sample_count],
                "output_dir": os.path.join(folder, "cropped")
            })
        return jobs

    def show_current_sample_image(self):
        if self.folder_index >= len(self.subfolder_jobs):
            self.run_batch_crop()
            messagebox.showinfo("Done", "Cropping completed for all selected subfolders.")
            self.root.quit()
            return

        job = self.subfolder_jobs[self.folder_index]
        sample_files = job["sample_files"]

        if self.sample_index >= len(sample_files):
            # Safety: if sample loop exhausted unexpectedly, move on
            self.folder_index += 1
            self.sample_index = 0
            self.show_current_sample_image()
            return

        current_filename = sample_files[self.sample_index]
        img_path = os.path.join(job["folder"], current_filename)

        self.original_image = Image.open(img_path)

        self.root.update_idletasks()
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10 or canvas_h < 10:
            canvas_w, canvas_h = (900, 700)

        self.display_image = self.original_image.copy()
        self.display_image.thumbnail((canvas_w, canvas_h))
        self.tk_image = ImageTk.PhotoImage(self.display_image)

        self.canvas.delete("all")
        self.canvas.config(scrollregion=(0, 0, self.tk_image.width(), self.tk_image.height()))
        self.canvas.create_image(0, 0, image=self.tk_image, anchor="nw")

        info = (
            f"Subfolder {self.folder_index + 1}/{len(self.subfolder_jobs)}: {os.path.basename(job['folder'])}\n"
            f"Sample image {self.sample_index + 1}/{len(sample_files)}: {current_filename}\n"
            f"Click center. Last click in this subfolder becomes final center for all images in that subfolder."
        )
        self.info_label.config(text=info)
        print(info.replace("\n", " | "))

    def on_click(self, event):
        if self.folder_index >= len(self.subfolder_jobs):
            return

        # Map displayed click to original image coordinates
        disp_w, disp_h = self.display_image.size
        orig_w, orig_h = self.original_image.size

        scale_x = orig_w / disp_w
        scale_y = orig_h / disp_h

        real_x = int(event.x * scale_x)
        real_y = int(event.y * scale_y)

        job = self.subfolder_jobs[self.folder_index]
        self.last_click_per_folder[job["folder"]] = (real_x, real_y)

        print(f"Captured click for {os.path.basename(job['folder'])}: ({real_x}, {real_y}) on sample {self.sample_index + 1}")

        # Advance sample image; after last sample for this folder move to next folder
        self.sample_index += 1
        if self.sample_index >= len(job["sample_files"]):
            self.folder_index += 1
            self.sample_index = 0

        self.show_current_sample_image()

    def run_batch_crop(self):
        print("\nStarting batch crop using final click per subfolder...")
        for job in self.subfolder_jobs:
            folder = job["folder"]
            image_files = job["files"]
            output_dir = job["output_dir"]
            os.makedirs(output_dir, exist_ok=True)

            if folder not in self.last_click_per_folder:
                print(f"Skipping {os.path.basename(folder)}: no click center captured.")
                continue

            center_x, center_y = self.last_click_per_folder[folder]
            print(f"Processing {os.path.basename(folder)} with center ({center_x}, {center_y})")

            for filename in image_files:
                img_path = os.path.join(folder, filename)
                original_image = Image.open(img_path)
                orig_w, orig_h = original_image.size

                crop_w = int(orig_w * self.crop_percent_w)
                crop_h = int(orig_h * self.crop_percent_h)

                left = center_x - crop_w // 2
                top = center_y - crop_h // 2
                right = left + crop_w
                bottom = top + crop_h

                if left < 0:
                    left = 0
                if top < 0:
                    top = 0
                if right > orig_w:
                    right = orig_w
                if bottom > orig_h:
                    bottom = orig_h

                cropped_img = original_image.crop((left, top, right, bottom))
                save_path = os.path.join(output_dir, filename)
                cropped_img.save(save_path)

            print(f"Saved cropped images in: {output_dir}")


if __name__ == "__main__":
    BatchSubfolderCropper()
