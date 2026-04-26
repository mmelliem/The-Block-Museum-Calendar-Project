#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
csv_path = filedialog.askopenfilename(title = "Select CSV file", filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
df = pd.read_csv(csv_path, parse_dates=['start_dt', 'end_dt'])


# In[207]:


import sys
import numpy as np
import pandas as pd
from docx import Document
from datetime import datetime


# In[208]:


df = pd.read_csv('03.01.2026-04.01.2026.CSV')

cols = ['Subject', 'Start Date', 'Start Time', 'End Time', 'Description', 'Location']
df = df[cols]

df['start_dt'] = pd.to_datetime(df['Start Date'] + ' ' + df['Start Time'])
df['end_dt'] = pd.to_datetime(df['Start Date'] + ' ' + df['End Time'])
df = df.drop(['Start Date', 'Start Time', 'End Time'], axis = 1)

df = df.sort_values('start_dt')

df['Subject'] = df['Subject'].str.replace('\n', ' ').str.strip()
mask = ~df['Subject'].str.strip().str.upper().str.startswith(('CANCELED', 'AUTOMATED'))
df = df[mask]

df['Description'] = df['Description'].str.replace('\r\n', ' ').str.strip()
zoom_strings = ['urldefense', 'zoom.com', 'zoom.us']
pattern = '|'.join(zoom_strings) # what?
df['Description'] = df['Description'].mask(df['Description'].str.contains(pattern, case=False), '')
df['Location'] = df['Location'].mask(df['Location'].str.contains(pattern, case=False), '')


# In[209]:


"""
import textwrap

for idx, desc in enumerate(df['Subject'].head(30)):
    print(f"--- Row {idx} ---")
    print(textwrap.fill(desc, width=80))
    print()
"""


# In[210]:


# filter for week?
"""
start = '2026-02-05'
end = '2026-02-12'

df = df[(df['start_dt'] >= start) & (df['start_dt'] < end)]
"""

df['date_str'] = df['start_dt'].dt.strftime('%A, %B, %d')
df['time_str'] = (df['start_dt']).dt.strftime('%-I:%M %p') + ' - ' + df['end_dt'].dt.strftime('%-I:%M %p')


# In[211]:


req_fields = [
    'Date:',
    'Time:',
    'Type of Event:',
    'Name of Group:',
    'Location:',
    'Internal Contact Person(s):',
    'External Contact Person(s)',
    'Guests Expected:',
    'Taking Attendance (Y/N):',
    'Attendance Total:',
    'Notes:'
]

for field in req_fields:
    df['Description'] = df['Description'].str.replace(field, '\n' + field)


# In[212]:


# make a function(s) that fixes date and time mistakes (make sure they match the real date and real time), have correct 
# formatting, also make function for location make sure it matches actual location


# In[213]:


import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re


def add_horizontal_line(para):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)

def set_spacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(before))
    spacing.set(qn("w:after"), str(after))
    pPr.append(spacing)

def fmt_time(dt):
    return dt.strftime("%I:%M %p").lstrip("0")

doc = Document()

for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.gutter = Inches(0)

style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(12)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.space_after = Pt(0)

p = doc.add_paragraph()
p.add_run().font.size = Pt(12)

# title
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("Weekly Events Attendance Totals")
run.font.name = "Arial"
run.font.size = Pt(12)
run.italic = True

doc.add_paragraph() # one blank line

for _, row in df.iterrows():
    # date header (size 14, left indent 0.13")
    date_label = row["start_dt"].strftime("%A, %B %d, %Y")
    date_para = doc.add_paragraph()
    date_para.paragraph_format.left_indent = Inches(0.13)
    set_spacing(date_para)
    run = date_para.add_run(date_label)
    run.font.name = "Segoe UI Semibold"
    run.font.size = Pt(14)
    run.bold = True

    # horizontal line
    line_para = doc.add_paragraph()
    line_para.paragraph_format.left_indent = Inches(0.13)
    line_para.paragraph_format.space_before = Pt(0)
    line_para.paragraph_format.space_after = Pt(0)
    line_para.paragraph_format.line_spacing = Pt(1)
    set_spacing(line_para)
    add_horizontal_line(line_para)

    # size-9 event line
    time_str = f"{fmt_time(row['start_dt'])} - {fmt_time(row['end_dt'])}"
    event_para = doc.add_paragraph()
    event_para.paragraph_format.left_indent = Inches(0.13)
    set_spacing(event_para)
    run = event_para.add_run(f"{time_str}    {str(row['Subject']).strip()} — {str(row['Location']).strip()}")
    run.font.name = "Gadugi"
    run.font.size = Pt(9)
    run.bold = True

    # description lines (italic, "Date" bolded, first line has 7.4pt space before)
    desc_lines = [l for l in str(row["Description"]).split("\n") if l.strip()]
    for i, line in enumerate(desc_lines):
        desc_para = doc.add_paragraph()
        desc_para.paragraph_format.left_indent = Inches(0.31)
        set_spacing(desc_para, before=(148 if i == 0 else 0)) # 7.4pt = 148 twips
        for part in re.split(r'\b(Date)\b', line):
            run = desc_para.add_run(part)
            run.font.name = "Arial"
            run.font.size = Pt(12)
            run.italic = True
            run.bold = (part == "Date")

    p1 = doc.add_paragraph()
    p1.add_run().font.size = Pt(9)
    p2 = doc.add_paragraph()
    p2.add_run().font.size = Pt(9)


section = doc.sections[0]
sectPr = section._sectPr

for hdr_ref in sectPr.findall(qn("w:headerReference")):
    sectPr.remove(hdr_ref)

# margins
pgMar = sectPr.find(qn("w:pgMar"))
if pgMar is None:
    pgMar = OxmlElement("w:pgMar")
    sectPr.append(pgMar)
pgMar.set(qn("w:top"),    "0")
pgMar.set(qn("w:bottom"), str(int(0.19 * 1440)))   # 274
pgMar.set(qn("w:left"),   str(int(0.25 * 1440)))   # 360
pgMar.set(qn("w:right"),  str(int(1 * 1440)))      # 1440
pgMar.set(qn("w:header"), "0")
pgMar.set(qn("w:footer"), "0")
pgMar.set(qn("w:gutter"), "0")


doc.save("output.docx")


# In[ ]:


output_path = filedialog.asksaveasfilename(title="Save Word document as", defaultextension='.docx', filetypes=['Word files','*.docx'])
doc.save(output_path)

