"""
    CM1: ListNumber can be
        1. 0 level -> ListParagraph
        2. 1 level -> ListParagraph
        3. 2 level -> ListParagraph
"""

from __future__ import print_function

from copy import copy
import docx
from bs4 import BeautifulSoup
import bs4
from path import path
import tempfile
import logging
logger = logging.getLogger('html2docx')
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)

ignore = ['hr']
document = None

class NumberingScheme(object):

    defaultscheme = "decimal;decimal;decimal;decimal;decimal"

    def __init__(self, scheme, levels, prefix, suffix):
        self.scheme = [''] # as levels start with 1
        self.scheme.extend(scheme)
        self.levels = levels
        self.default = self.scheme[1]
        self.prefix = prefix
        self.suffix = suffix

    def get_scheme(self, level):
        if type(level) == str:
            level = int(level)
        if level <= self.levels:
            return self.scheme[level]
        return self.default

    @staticmethod
    def get_default():
        prefix, suffix = '', ''
        scheme = NumberingScheme.defaultscheme.split(';')
        return NumberingScheme(scheme, len(scheme), prefix, suffix)

class Style(object):

    def __init__(self):
        self.level = 0
        self.listtype = None
        self.bold = False
        self.italic = False
        self.para = None
        self.tag = '';
        self.underline = False
        self.tab = False
        self.firstli = False

    def set_class_property(self, cssclass):
        if cssclass == 'pydocx-underline':
            self.underline = True
        elif cssclass == 'pydocx-strike':
            self.strike = True
        elif cssclass == 'pydocx-tab':
            self.tab = True

    @staticmethod
    def get_defaults():
        return {}

    def __str__(self):
        return 'level: %d, bold: %r, italic: %r' % (
                                                    self.level,
                                                    self.bold,
                                                    self.italic
                                                    )


def dfs(root, depth, style):

    if type(root) == bs4.element.NavigableString:

        if not root or root == ' ': return

        para = style.para
        if not para:
            para = document.add_paragraph('')

        run = para.add_run('')
        if style.tab: run.add_tab()
        run.add_text(' '+root.strip())
        run.bold = style.bold
        run.italic = style.italic
        run.underline = style.underline
        return

    if not root.name and type(root) == bs4.element.Comment: return

    newstyle = copy(style)
    newstyle.level = style.level
    classes = root.attrs.get('class', None)
    if classes:
        for cssclass in classes:
            style.set_class_property(cssclass)

    if root.name in ignore: return

    if root.name == 'br':
        newstyle.line_break = True
        if newstyle.para:
            newstyle.para.add_run('').add_break()
        else:
            para = document.add_paragraph(text=r"\n")

        return
    elif root.name == 'ol' or root.name == 'ul':
        newstyle.level += 1
        root.attrs['level'] = newstyle.level
        newstyle.listtype = root.name
        firstlevel = True
        for parent in root.parents:
            if parent.name == 'ol':
                firstlevel = False
        if firstlevel:
            document.add_paragraph('') # add empty paragraph
    elif root.name == 'span':
        inlinestyle = root.attrs.get('style', None)
        newstyle = style
        if inlinestyle and 'margin-left' in inlinestyle.split(';'):
            newstyle.tag = True
    elif root.name == 'strong':
        newstyle.bold = True
    elif root.name == 'i':
        newstyle.italic = True
    elif root.name == 'p':
        if not style.para:
            newstyle.para = document.add_paragraph('')
    elif root.name and root.name.startswith('h'):
        newstyle.para = document.add_heading('', int(root.name[-1]))
    elif root.name == 'li':
        level = '' if style.level == 1 else str(style.level)
        if root.parent.name == 'ol':
            if not newstyle.firstli:
                style.firstli = True
                document.new_list()
            newstyle.para = document.add_paragraph('', style='ListNumber'+level)
        elif root.parent.name == 'ul':
            newstyle.para = document.add_paragraph('', style='ListBullet'+level)
    elif root.name == 'table':
        HandleTable(root, style)
        return
    elif root.name == 'pagebreak':
        document.add_page_break()
        return

    for tag in root.contents:
        if tag:
            dfs(tag, depth+1, newstyle)

def HandleTable(tag, style):
    "Add header and rows to the table"

    nrows= 1
    ncols = len(tag.tr.find_all('th'))
    if not ncols:
        ncols = len(tag.tr.find_all('td'))
        if not ncols:
            return
    table = document.add_table(rows=1, cols=ncols)

    # if header is present, add it
    if tag.thead:
        hdr_cells = table.rows[0].cells
        for i, head in enumerate(tag.thead.tr.find_all('th')):
            hdr_cells[i].text = head.text

    # add rows
    for tr in tag.findChildren('tr', recursive=False):
        row_cells = table.add_row().cells
        for i, td in enumerate(tr.findChildren('td')):
            row_cells[i].text = td.text

def getlevels(html):
    soup = BeautifulSoup(html)
    dfs(soup.body, 0, Style())

    return soup

def getNumberingScheme(body):
    # Get scheme from first ol present in doc

    if body.ol:
        scheme = body.ol.attrs.get('scheme',
                                    NumberingScheme.defaultscheme).split(';')
        prefix = body.ol.attrs.get('prefix', '')
        suffix = body.ol.attrs.get('suffix', '')
        ns = NumberingScheme(scheme, len(scheme), prefix, suffix)
        return ns

    return NumberingScheme.get_default()

def html2docx(htmlcontent):
    "Exposed API"

    import docx
    global document
    document = docx.Document()
    soup = BeautifulSoup(htmlcontent.replace('\n', ''))
    body = soup.body
    ns = getNumberingScheme(body)
    style = Style()
    dfs(body, 0, style)

    try:
        docx = tempfile.NamedTemporaryFile()
        document.save(docx.name)
        docxcontent = docx.read()
    except:
        logger.error('Something went wrong')
    finally:
        docx.close()

    return docxcontent


def convert(filepath):
    global document
    # logger.info('converting: '+filepath)
    document = docx.Document()
    content = open(filepath, 'r').read()
    soup = BeautifulSoup(content.replace('\n', ''))
    body = soup.body
    ns = getNumberingScheme(body)

    style = Style()
    dfs(body, 0, style)

if __name__=='__main__':

    document = docx.Document()
    convert('simple.html')
    document.save('simple.docx')

    # for html in path('html').files("*.html"):
    #     convert(html)
    #     docxpath = path('docx').joinpath(html.basename()[:-5]+'.docx')
    #     document.save(docxpath)
    #     logger.info('Document saved: ' + docxpath)

