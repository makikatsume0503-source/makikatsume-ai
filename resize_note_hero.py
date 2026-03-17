from PIL import Image

src_path = r'C:\Users\moca1\.gemini\antigravity\brain\9dd91cc6-3aac-4831-9d10-1f4f6df77b4d\article_gemini_investment_hero_very_1772776258923.png'
dst_path = r'C:\Users\moca1\.gemini\antigravity\brain\9dd91cc6-3aac-4831-9d10-1f4f6df77b4d\note_eyecatch_gemini_investment_very.png'

try:
    img = Image.open(src_path)
    # note.com eyecatch size: 1280 x 670
    target_ratio = 1280 / 670
    current_ratio = img.width / img.height

    if current_ratio > target_ratio:
        # crop sides
        new_width = int(img.height * target_ratio)
        left = (img.width - new_width) // 2
        top = 0
        right = left + new_width
        bottom = img.height
    else:
        # crop top
        new_height = int(img.width / target_ratio)
        left = 0
        top = (img.height - new_height) // 2
        right = img.width
        bottom = top + new_height

    img = img.crop((left, top, right, bottom))
    img = img.resize((1280, 670), Image.Resampling.LANCZOS)

    img.save(dst_path)
    print(f'Successfully saved resized note eyecatch image to {dst_path}')
except Exception as e:
    print(f'Error: {e}')
