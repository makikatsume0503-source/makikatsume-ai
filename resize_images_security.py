import sys
import subprocess
try:
    import PIL
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'Pillow'])

from PIL import Image
import os
import glob

def resize_and_crop(image_path, target_size=(1920, 1280)):
    if not os.path.exists(image_path): return
    img = Image.open(image_path)
    img_ratio = img.width / img.height
    target_ratio = target_size[0] / target_size[1]
    
    if img_ratio > target_ratio:
        new_width = int(target_ratio * img.height)
        offset = (img.width - new_width) / 2
        img = img.crop((offset, 0, img.width - offset, img.height))
    elif img_ratio < target_ratio:
        new_height = int(img.width / target_ratio)
        offset = (img.height - new_height) / 2
        img = img.crop((0, offset, img.width, img.height - offset))
        
    img = img.resize(target_size, Image.Resampling.LANCZOS)
    img.save(image_path)
    print(f'Resized {image_path}')

# Find the generated images in the brain folder
brain_path = r'C:\Users\moca1\.gemini\antigravity\brain\c514e011-91cb-48ce-ba9a-627cae40d0cc'
hero = glob.glob(os.path.join(brain_path, 'article_ai_security_hero*.png'))
body1 = glob.glob(os.path.join(brain_path, 'article_ai_security_body1*.png'))
body2 = glob.glob(os.path.join(brain_path, 'article_ai_security_body2*.png'))

for img_list in [hero, body1, body2]:
    if img_list:
        resize_and_crop(img_list[0])