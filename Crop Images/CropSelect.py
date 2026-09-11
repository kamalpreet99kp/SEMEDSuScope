import tkinter as tk
from PIL import Image, ImageTk
import numpy as np
import os
from tkinter import filedialog
import cv2


class Application(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.canvas = tk.Canvas(self, width=400, height=400, cursor="cross")
        self.canvas.pack(side="top", fill="both", expand=True)
        self.canvas.bind("<Button-1>", self.start_square)
        self.canvas.bind("<B1-Motion>", self.update_square)
        self.canvas.bind("<ButtonRelease-1>", self.end_square)

        self.directory = filedialog.askdirectory(
            title="Select the folder containing microscope JPG images",
            initialdir=os.getcwd(),
        )
        self.image_files = [os.path.join(self.directory, f) for f in os.listdir(self.directory) if f.endswith('.jpg')]
        self.current_image = 0
        self.square_points = []

        self.load_image()

        self.prev_button = tk.Button(self, text="Previous", command=self.prev_image)
        self.prev_button.pack(side="left")
        self.next_button = tk.Button(self, text="Next", command=self.next_image)
        self.next_button.pack(side="left")

        self.image_name_label = tk.Label(self, text=os.path.basename(self.image_files[self.current_image]))
        self.image_name_label.pack(side="left")

        if self.current_image == len(self.image_files) - 1:
            self.next_button.configure(state="disabled")

        self.bind_all("<space>", self.next_image_keyboard)

    def load_image(self):
        self.im = Image.open(self.image_files[self.current_image])
        self.width, self.height = self.im.size
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.ratio_w = (screen_width-100) / self.width
        self.ratio_h = (screen_height-177) / self.height
        self.ratio = min(self.ratio_w, self.ratio_h)
        self.im = self.im.resize((int(self.width * self.ratio), int(self.height * self.ratio)), Image.ANTIALIAS)
        self.tk_im = ImageTk.PhotoImage(self.im)
        self.canvas.configure(width=int(self.width * self.ratio), height=int(self.height * self.ratio))
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_im)

    def next_image(self):
        self.current_image += 1
        if self.current_image >= len(self.image_files):
            self.current_image = len(self.image_files) - 1
        self.load_image()
        self.image_name_label.config(text=os.path.basename(self.image_files[self.current_image]))

        if self.current_image == len(self.image_files) - 1:
            self.next_button.configure(state="disabled")
        else:
            self.next_button.configure(state="normal")

    def prev_image(self):
        self.current_image -= 1
        if self.current_image < 0:
            self.current_image = 0
        self.load_image()
        self.image_name_label.config(text=os.path.basename(self.image_files[self.current_image]))

        if self.current_image == len(self.image_files) - 1:
            self.next_button.configure(state="disabled")
        else:
            self.next_button.configure(state="normal")

    def start_square(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.current_rectangle = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y, outline='black'
        )

    def update_square(self, event):
        self.canvas.coords(self.current_rectangle, self.start_x, self.start_y, event.x, event.y)

    def end_square(self, event):
        self.end_x = min(event.x, self.canvas.winfo_width())
        self.end_y = min(event.y, self.canvas.winfo_height())
        self.canvas.coords(self.current_rectangle, self.start_x, self.start_y, self.end_x, self.end_y)
        self.cropout_image()

    def cropout_image(self):
        scale_x = self.im.size[0] / self.canvas.winfo_width()
        scale_y = self.im.size[1] / self.canvas.winfo_height()
        start_x, end_x = sorted([self.start_x, self.end_x])
        start_y, end_y = sorted([self.start_y, self.end_y])

        crop_area = (
            int(start_x * scale_x),
            int(start_y * scale_y),
            int(end_x * scale_x),
            int(end_y * scale_y),
        )
        print(f"Start X: {self.start_x}, Start Y: {self.start_y}, End X: {self.end_x}, End Y: {self.end_y}")
        print(f"Scale X: {scale_x}, Scale Y: {scale_y}")
        print(f"Image Size: {self.im.size}")
        print(f"Crop Area: {crop_area}")

        cropped_image = self.im.crop(crop_area)
        cropped_image_dir = self.directory + '/cropout'
        if not os.path.exists(cropped_image_dir):
            os.makedirs(cropped_image_dir)

        cropped_filename = os.path.join("cropouts", os.path.basename(self.image_files[self.current_image]))
        cropped_filename = os.path.join(cropped_image_dir, os.path.basename(self.image_files[self.current_image]))
        print(f"Saving cropped image to {cropped_filename}")
        cropped_image.save(cropped_filename, 'JPEG', quality=95)

    def next_image_keyboard(self, event=None):
        self.next_image()


if __name__ == "__main__":
    app = Application()
    app.mainloop()
