import os
from bs4 import BeautifulSoup

html_path = r'c:\Users\moca1\.gemini\antigravity\scratch\ai-consultant-lp\article-gemini-investment.html'
md_path = r'C:\Users\moca1\.gemini\antigravity\scratch\ai-consultant-lp\note_draft_gemini_investment.md'

with open(html_path, 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

article_body = soup.find('div', class_='article-body')

md_lines = []
title = soup.find('h1', class_='article-title').text.strip()
md_lines.append(f"# {title}\n")

for elem in article_body.children:
    if elem.name == 'p':
        text = elem.get_text()
        text = text.replace('\n', '')
        md_lines.append(text + "\n")
    elif elem.name == 'h2':
        md_lines.append(f"## {elem.text.strip()}\n")
    elif elem.name == 'h3':
        if '特徴' not in elem.text and 'メリット' not in elem.text: 
            # Skipping specific LP styling headers if any
            md_lines.append(f"### {elem.text.strip()}\n")
    elif elem.name == 'ul':
        for li in elem.find_all('li', recursive=False):
            md_lines.append(f"- {li.text.strip()}")
        md_lines.append("\n")
    elif elem.name == 'div' and 'info-box' in elem.get('class', []):
        md_lines.append("```text")
        md_lines.append(elem.get_text(separator='\n').strip())
        md_lines.append("```\n")

# filter out empty lines
final_lines = []
for line in md_lines:
    if line.strip() == "目次": continue
    if "この記事を書いた人" in line or "AI導入・活用でお困りですか" in line:
        break # Skip LP specific footer
    final_lines.append(line)

with open(md_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(final_lines))

print(f"Created Markdown draft at {md_path}")
