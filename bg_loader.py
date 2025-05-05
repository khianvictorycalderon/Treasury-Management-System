from tkinter import Canvas
from PIL import Image, ImageTk
import base64
import io

# Prerequisite:
# - Pillow library

def load_bg_image(root, image_path_or_base64, size="cover"):
    # Check if the input is a Base64 string or a file path
    if image_path_or_base64.startswith('data:image'):  # Check if it's a Base64 string
        # Decode the Base64 string
        image_data = base64.b64decode(image_path_or_base64.split(',')[1])  # Strip the data:image part
        original_image = Image.open(io.BytesIO(image_data))
    else:
        # Otherwise, treat it as a file path
        original_image = Image.open(image_path_or_base64)

    # Create a Canvas widget to hold the background image
    canvas = Canvas(root, highlightthickness=0)
    canvas.place(x=0, y=0, relwidth=1.0, relheight=1.0)  # Use .place() to cover the entire window
    canvas.lower("all")

    # Store references to avoid garbage collection
    root._bg_data = {
        'original_image': original_image,
        'tk_image': None,
        'canvas': canvas  # Store canvas for later use
    }

    def resize_bg(event=None):
        # Get the current size of the window
        win_w, win_h = root.winfo_width(), root.winfo_height()

        # Get the size of the original image
        img_w, img_h = original_image.size

        if size == "cover":
            # Scale the image to cover the entire window while maintaining the aspect ratio
            scale = max(win_w / img_w, win_h / img_h)
            new_size = (int(img_w * scale), int(img_h * scale))

        elif size == "contain":
            # Scale the image to fit within the window without cropping
            scale = min(win_w / img_w, win_h / img_h)
            new_size = (int(img_w * scale), int(img_h * scale))

        elif isinstance(size, str) and "%" in size:
            # Custom percentage (e.g., "25%" -> 25% of the original size)
            percent = float(size.strip('%')) / 100
            new_size = (int(img_w * percent), int(img_h * percent))

        else:
            # Default to "cover" if an invalid size string is provided
            scale = max(win_w / img_w, win_h / img_h)
            new_size = (int(img_w * scale), int(img_h * scale))

        # Resize the image using LANCZOS filter for high-quality resizing
        resized = original_image.resize(new_size, Image.Resampling.LANCZOS)

        # Update the Tkinter image reference
        tk_image = ImageTk.PhotoImage(resized)
        root._bg_data['tk_image'] = tk_image

        # Clear any previous image and display the resized background
        canvas.delete("all")  # Delete previous canvas items (if any)
        x = (win_w - new_size[0]) // 2
        y = (win_h - new_size[1]) // 2
        canvas.create_image(x, y, image=tk_image, anchor='nw')

    # Bind the resize event to the function to handle window resizing
    root.bind('<Configure>', resize_bg)

    # Return the canvas so it can be used in `main.py` for placing widgets on top
    return canvas
