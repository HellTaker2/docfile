#!/usr/bin/env python3
"""Convert docx to markdown: markitdown for text + manual image extraction."""
import os
import sys
import io
import datetime
import re
import zipfile
import xml.etree.ElementTree as ET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from markitdown import MarkItDown

BASE_DIR = r"d:\1mydict\java_study\aishoping\docfile"
IMAGE_DIR = os.path.join(BASE_DIR, "images")
os.makedirs(IMAGE_DIR, exist_ok=True)

timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def extract_images_in_order(docx_path):
    """Extract images from docx in document appearance order. Returns list of saved filenames."""
    saved = []

    with zipfile.ZipFile(docx_path, 'r') as z:
        # 1. Parse rels to map rId -> media path
        rels_xml = z.read('word/_rels/document.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        rId_to_path = {}
        for rel in rels_root:
            rId = rel.get('Id')
            target = rel.get('Target')
            if rId and target and '/image' in rel.get('Type', ''):
                # Normalize path: target is like "media/image1.png"
                rId_to_path[rId] = target

        # 2. Parse document.xml to find image order
        doc_xml = z.read('word/document.xml')
        doc_root = ET.fromstring(doc_xml)

        # Find all drawings and extract r:embed ids
        img_rIds = []
        for drawing in doc_root.iter(f'{{{WP_NS}}}inline'):
            blip = drawing.find(f'.//{{{A_NS}}}blip')
            if blip is not None:
                embed = blip.get(f'{{{R_NS}}}embed')
                if embed:
                    img_rIds.append(embed)

        # Also check for anchor drawings
        for drawing in doc_root.iter(f'{{{WP_NS}}}anchor'):
            blip = drawing.find(f'.//{{{A_NS}}}blip')
            if blip is not None:
                embed = blip.get(f'{{{R_NS}}}embed')
                if embed:
                    img_rIds.append(embed)

        # 3. Extract images in order, saving with new names
        for i, rId in enumerate(img_rIds):
            media_path = rId_to_path.get(rId)
            if media_path is None:
                print(f"    Warning: rId {rId} not found in rels")
                continue

            # Try different path formats
            try_paths = [
                media_path,
                'word/' + media_path,
                'word/media/' + os.path.basename(media_path),
            ]
            data = None
            for tp in try_paths:
                try:
                    data = z.read(tp)
                    break
                except KeyError:
                    continue

            if data is None:
                print(f"    Warning: media file not found: {media_path}")
                continue

            ext = os.path.splitext(media_path)[1] or '.png'
            new_name = f"{timestamp}-{i+1:04d}{ext}"
            dst_path = os.path.join(IMAGE_DIR, new_name)
            with open(dst_path, 'wb') as f:
                f.write(data)
            saved.append(new_name)

    return saved


def convert_docx_with_images(docx_path, md_path, doc_name):
    # Extract images first
    print(f"  Extracting images from {doc_name}...")
    saved_images = extract_images_in_order(docx_path)
    print(f"    Found {len(saved_images)} images in document order")

    # Convert with markitdown
    print(f"  Converting {doc_name} with markitdown...")
    md = MarkItDown()
    result = md.convert(docx_path)
    content = result.text_content

    # Replace truncated base64 references with actual image files
    # The markitdown output has: ![](data:image/png;base64...)
    # We need to replace these in order with the saved images
    img_idx = [0]

    def replace_image_ref(match):
        if img_idx[0] < len(saved_images):
            img_name = saved_images[img_idx[0]]
            img_idx[0] += 1
            alt = match.group(1) or ''
            if alt:
                return f'![{alt}](images/{img_name})'
            return f'![](images/{img_name})'
        return match.group(0)

    # Pattern: ![](data:image/...;base64...) or ![alt](data:image/...;base64...)
    pattern = re.compile(r'!\[([^\]]*)\]\(data:image/[^;]+;base64\.\.\.\)')
    content = pattern.sub(replace_image_ref, content)

    # Remove excessive blank lines
    content = re.sub(r'\n{3,}', '\n\n', content)

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content.rstrip('\n') + '\n')

    print(f"  {doc_name}: {len(content.splitlines())} lines -> {md_path}")
    if img_idx[0] < len(saved_images):
        print(f"    Warning: {len(saved_images) - img_idx[0]} images not matched in markdown")
    if img_idx[0] > 0:
        print(f"    Matched {img_idx[0]} image references")


if __name__ == '__main__':
    print(f"Image timestamp: {timestamp}\n")
    convert_docx_with_images(
        os.path.join(BASE_DIR, "AI化电商平台需求分析.docx"),
        os.path.join(BASE_DIR, "AI化电商平台需求分析.md"),
        "需求分析"
    )
    convert_docx_with_images(
        os.path.join(BASE_DIR, "AI化电商平台原型设计.docx"),
        os.path.join(BASE_DIR, "AI化电商平台原型设计.md"),
        "原型设计"
    )
    print("\nDone!")
