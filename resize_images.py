import sys
import subprocess

def install_pillow():
    try:
        import PIL
    except ImportError:
        print("Installing Pillow...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])

install_pillow()

from PIL import Image
import os

def resize_and_crop(image_path, target_size=(1920, 1280)):
    if not os.path.exists(image_path):
        print(f"File not found: {image_path}")
        return
        
    img = Image.open(image_path)
    img_ratio = img.width / img.height
    target_ratio = target_size[0] / target_size[1]
    
    if img_ratio > target_ratio:
        # Image is wider than target, crop width
        new_width = int(target_ratio * img.height)
        offset = (img.width - new_width) / 2
        crop_box = (offset, 0, img.width - offset, img.height)
        img = img.crop(crop_box)
    elif img_ratio < target_ratio:
        # Image is taller than target, crop height
        new_height = int(img.width / target_ratio)
        offset = (img.height - new_height) / 2
        crop_box = (0, offset, img.width, img.height - offset)
        img = img.crop(crop_box)
        
    img = img.resize(target_size, Image.Resampling.LANCZOS)
    img.save(image_path)
    print(f"Resized {image_path} to {target_size}")

images = [
    r"c:\Users\moca1\.gemini\antigravity\scratch\ai-consultant-lp\assets\ai_tools_hero.png",
    r"c:\Users\moca1\.gemini\antigravity\scratch\ai-consultant-lp\assets\ai_tools_illustration.png",
    r"c:\Users\moca1\.gemini\antigravity\scratch\ai-consultant-lp\assets\ai_tools_illustration_2.png"
]

for img_path in images:
    try:
        resize_and_crop(img_path)
    except Exception as e:
        print(f"Error on {img_path}: {e}")
