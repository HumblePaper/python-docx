from __future__ import print_function

from docx.html2docx import html2docx

content = open('simple.html').read()
with open('simple.docx', 'w') as f:
	print('converting html to docx...', end="")
	f.write(html2docx(content))
	print('done')

