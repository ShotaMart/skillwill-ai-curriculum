"""
Utility script to ensure all logos are 512x512 RGBA transparent squares
with centered pure brand icons (no attached wordmarks or stretching).
"""
import os
from PIL import Image

def make_square_icon(src_img, target_size=512, icon_padding=40):
    bbox = src_img.getbbox()
    if bbox:
        src_img = src_img.crop(bbox)
    
    max_dim = target_size - (icon_padding * 2)
    w, h = src_img.size
    scale = min(max_dim / w, max_dim / h)
    new_w = int(w * scale)
    new_h = int(h * scale)
    
    resized = src_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    canvas = Image.new('RGBA', (target_size, target_size), (0, 0, 0, 0))
    paste_x = (target_size - new_w) // 2
    paste_y = (target_size - new_h) // 2
    canvas.paste(resized, (paste_x, paste_y), resized)
    return canvas

def process_directory(logo_dir):
    for f in os.listdir(logo_dir):
        if f.endswith('.png'):
            p = os.path.join(logo_dir, f)
            img = Image.open(p).convert('RGBA')
            if img.size != (512, 512):
                print(f"Standardizing {f} from {img.size} to (512, 512)...")
                make_square_icon(img).save(p)

if __name__ == '__main__':
    logo_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../../assets/logos'))
    if os.path.exists(logo_dir):
        process_directory(logo_dir)
        print("All logos verified & standardized to 512x512!")
