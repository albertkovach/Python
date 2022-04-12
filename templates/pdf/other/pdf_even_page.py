#!/bin/usr/env python
# Usage: python pdf_even_page.py

import os
import PyPDF2
import re


def make_even_page(in_fpath, out_fpath):
    reader = PyPDF2.PdfFileReader(in_fpath)
    writer = PyPDF2.PdfFileWriter()
    for i in range(reader.getNumPages()):
        writer.addPage(reader.getPage(i))
    if reader.getNumPages() % 2 == 1:
        _, _, w, h = reader.getPage(0)['/MediaBox']
        writer.addBlankPage(w, h)
    with open(out_fpath, 'wb') as fd:
        writer.write(fd)

# Output directory: ./output
if not os.path.exists('output'):
    os.mkdir('output')

# Iterating all dirs below here, make all pdf files so that their pages be even
pdf_pattern = re.compile(".+(\\.pdf)$", flags=re.IGNORECASE)

for root, dirs, files in os.walk("."):
    if root.startswith('./output'):
        continue

    for fname in files:
        if pdf_pattern.match(fname):
            in_fpath = os.path.join(root, fname)
            out_fname = fname[:-4] + "_out.pdf"
            out_dir = os.path.join('./output', root).replace('/.', '')
            out_fpath = os.path.join('./output', root, out_fname).replace('/.', '')

            if not os.path.exists(out_dir):
                os.makedirs(out_dir)

            try:
                print (in_fpath + "\t -> \t" + out_fpath)
                make_even_page(in_fpath, out_fpath)
            except:
                # Some pdfs will fail
                print ('\tFAILED!! ' + in_fpath)
                pass
