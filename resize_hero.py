from PIL import Image

src_path = r'C:\Users\moca1\.gemini\antigravity\brain\9dd91cc6-3aac-4831-9d10-1f4f6df77b4d\article_gemini_investment_hero_1772775646499.png'
dst_path = r'C:\Users\moca1\.gemini\antigravity\scratch\ai-consultant-lp\assets\article_gemini_investment_hero.png'

try:
    img = Image.open(src_path)
    target_ratio = 1920 / 1280
    current_ratio = img.width / img.height

    if current_ratio > target_ratio:
        new_width = int(img.height * target_ratio)
        left = (img.width - new_width) // 2
        top = 0
        right = left + new_width
        bottom = img.height
    else:
        new_height = int(img.width / target_ratio)
        left = 0
        top = (img.height - new_height) // 2
        right = img.width
        bottom = top + new_height

    img = img.crop((left, top, right, bottom))
    img = img.resize((1920, 1280), Image.Resampling.LANCZOS)

    img.save(dst_path)
    print(f'Successfully saved resized image to {dst_path}')
except Exception as e:
    print(f'Error: {e}')
