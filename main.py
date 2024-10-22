#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
from pathlib import Path

dir_in = "EBT2"
path_in = Path("in") / dir_in
path_out = Path("out") / dir_in

if not path_out.exists():
    os.mkdir(path_out)

ppt_in = [file for file in path_in.iterdir() if file.name.endswith(".pptx")]
list(ppt_in)


# In[24]:


from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def extract(ppt):
    prs = Presentation(ppt)
    lines = []
    
    dir_out = path_out / ppt.stem
    if not dir_out.exists():
        os.mkdir(dir_out)
    
    for i, slide in enumerate(prs.slides):
        for j, shape in enumerate(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                filename = f"{i}-{j}_{shape.image.filename}"
                img_out = dir_out / filename
                if not img_out.exists():
                    with open(img_out, "wb") as file:
                        file.write(shape.image.blob)
                lines.append(f"![img]({filename})")
                continue
            
            if not shape.has_text_frame:
                continue
                
            for paragraph in shape.text_frame.paragraphs:
                line = ""
                for run in paragraph.runs:
                    text = run.text.strip()
                    if text == "": continue;
                    font = run.font
                    if font.bold: text = f"**{text}**";
                    if font.italic: text = f"*{text}*";
                    if font.underline: text = f"<u>{text}</u>";
                    if line != "": line += " ";
                    line += text
                if line == "": continue;
                lines.append(line)
                    
    with open(dir_out / "README.md", "w") as file:
        file.write("\n\n".join(lines))


# In[25]:


for ppt in ppt_in:
    extract(ppt)


# In[ ]:




