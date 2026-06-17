from PIL import Image
import os

def make_white_transparent(img_path, out_path, tolerance=50):
    img = Image.open(img_path).convert("RGBA")
    data = img.getdata()
    
    new_data = []
    for item in data:
        # Check if the pixel is white or near-white
        if item[0] > 255 - tolerance and item[1] > 255 - tolerance and item[2] > 255 - tolerance:
            # Change all white to transparent
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
            
    img.putdata(new_data)
    img.save(out_path, "PNG")
    print(f"Saved transparent image to {out_path}")

if __name__ == "__main__":
    make_white_transparent("logo.png", "logo_transparent.png")
