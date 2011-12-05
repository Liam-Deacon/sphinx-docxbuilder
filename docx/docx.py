# -*- coding: utf-8 -*-
'''
  Microsoft Word 2007 Document Composer

  Copyright 2011 by haraisao at gmail dot com

  This software based on 'python-docx' which developed by Mike MacCana.

'''
'''
  Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

  Part of Python's docx module - http://github.com/mikemaccana/python-docx
  See LICENSE for licensing information.
'''

from lxml import etree
import Image
import zipfile
import shutil
import re
import time
import os
from os.path import join
import tempfile
import sys
import codecs


# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily. 
nsprefixes = {
    # Text Content
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
    'dc':"http://purl.org/dc/elements/1.1/", 
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr':'http://schemas.openxmlformats.org/package/2006/relationships',
    # xml 
    'xml':'http://www.w3.org/XML/1998/namespace'
    }

Enum_Types = {
    'arabic':'decimal',
    'loweralpha':'lowerLetter',
    'upperalpha':'upperLetter',
    'lowerroman':'lowerRoman',
    'upperroman':'upperRoman'
    }

#####################
def norm_name(tagname, namespaces=nsprefixes):
    '''
       Convert the 'tagname' to a formal expression.
          'ns:tag' --> '{namespace}tag'
          'tag' --> 'tag'
    '''
    ns_name = tagname.split(':', 1)
    if len(ns_name) >1 :
      tagname = "{%s}%s" % (namespaces[ns_name[0]], ns_name[1])
    return tagname

def get_elements(xml, path, ns=nsprefixes):
    '''
       Get elements from a Element tree with 'path'.
    '''
    result = []
    try:
      result = xml.xpath(path, namespaces=ns)
    except:
       pass
    return result

def find_file(filename, child_dir=None):
    '''
       Find file...
    '''
    fname = filename
    if not os.access( filename ,os.F_OK):
      for pth in sys.path:
        if child_dir :
          pth = join(pth, child_dir)
        fname = join(pth, filename)
        if os.access(fname, os.F_OK):
          break
        else:
          fname=None
    return fname

def get_enumerate_type(typ):
  try:
    typ=Enum_Types[typ]
  except:
    typ="decimal"
    pass
  return  typ
#
#  DocxDocument class
#   This class for analizing docx-file
#
class DocxDocument:
  def __init__(self, docxfile=None):
    '''
      Constructor
    '''
    self.title = ""
    self.subject = ""
    self.creator = "Python:DocDocument"
    self.company = ""
    self.category = ""
    self.descriptions = ""
    self.keywords = []
    self.stylenames = {}

    self.set_document(docxfile)
    self.docxfile = docxfile

  def set_document(self, fname):
    '''
      set docx document 
    '''
    if fname :
      self.docxfile = fname
      self.docx = zipfile.ZipFile(fname)

      self.document = self.get_xmltree('word/document.xml')
      self.docbody = get_elements(self.document, '/w:document/w:body')[0]

      self.numbering = self.get_xmltree('word/numbering.xml')
      self.styles = self.get_xmltree('word/styles.xml')
      self.extract_stylenames()
      self.paragraph_style_id = self.stylenames['Normal']
      self.character_style_id = self.stylenames['Default Paragraph Font']

    return self.document

  def get_xmltree(self, fname):
    '''
      Extract a document tree from the docx file
    '''
    try:
      return etree.fromstring(self.docx.read(fname))
    except:
      return None
    
  def extract_stylenames(self):
    '''
      Extract a stylenames from the docx file
    '''
    style_elems = get_elements(self.styles, 'w:style')

    for style_elem in style_elems:
        aliases_elems = get_elements(style_elem, 'w:aliases')
        if aliases_elems:
            name = aliases_elems[0].attrib[norm_name('w:val')]
        else:
            name_elem = get_elements(style_elem,'w:name')[0]
            name = name_elem.attrib[norm_name('w:val')]
        value = style_elem.attrib[norm_name('w:styleId')]
        self.stylenames[name] = value
    return self.stylenames

  def get_paper_info(self):
    self.paper_info = get_elements(self.document,'/w:document/w:body/w:sectPr')[0]
    return self.paper_info
    

  def extract_file(self,fname, outname=None, pprint=True):
    '''
      Extract file from docx 
    '''
    try:
      filelist = self.docx.namelist()

      if filelist.index(fname) >= 0 :
        xmlcontent = self.docx.read(fname)
        document = etree.fromstring(xmlcontent)
        xmlcontent = etree.tostring(document, encoding='UTF-8', pretty_print=pprint)
        if outname == None : outname = os.path.basename(fname)
        f = codecs.open(outname, 'w','UTF-8')
        f.write(xmlcontent)
        f.close()
    except:
        print "Error in extract_document: %s" % fname
        print filelist
    return


  def extract_files(self,to_dir):
    '''
      Extract all files from docx 
    '''
    try:
      if not os.access(to_dir, os.F_OK) :
        os.mkdir(to_dir)

      filelist = self.docx.namelist()
      for fname in filelist :
        xmlcontent = self.docx.read(fname)
        file_name = join(to_dir,fname)
        if not os.path.exists(os.path.dirname(file_name)) :
          os.makedirs(os.path.dirname(file_name)) 
        f = open(file_name, 'w')
        f.write(xmlcontent)
        f.close()
    except:
      print "Error in extract_files ..."
      return False
    return True

  def restruct_docx(self, docx_dir, docx_filename, files_to_skip=[]):
    '''
       This function is copied and modified the 'savedocx' function contained 'python-docx' library
      Restruct docx file from files in 'doxc_dir'
    '''
    if not os.access( docx_dir ,os.F_OK):
      print "Can't found docx directory: %s" % docx_dir
      return

    docxfile = zipfile.ZipFile(docx_filename, mode='w', compression=zipfile.ZIP_DEFLATED)

    prev_dir = os.path.abspath('.')
    os.chdir(docx_dir)

    # Add & compress support files
    files_to_ignore = ['.DS_Store'] # nuisance from some os's
    for dirpath,dirnames,filenames in os.walk('.'):
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath,filename)
            archivename = os.path.normpath(templatefile)
            archivename = '/'.join(archivename.split(os.sep))
            if archivename in files_to_skip:
                continue
            #print 'Saving: '+archivename          
            docxfile.write(templatefile, archivename)

    os.chdir(prev_dir) # restore previous working dir
    return docxfile

  def get_filelist(self):
      '''
         Extract file names from docx file
      '''
      filelist = self.docx.namelist()
      return filelist

  def search(self, search):
    '''
      This function is copied from 'python-docx' library
      Search a document for a regex, return success / fail result
    '''
    result = False
    text_tag = norm_name('w:t')
    searchre = re.compile(search)
    for element in self.docbody.iter():
        if element.tag == text_tag :
            if element.text:
                if searchre.search(element.text):
                    result = True
    return result

  def replace(self, search,replace):
    '''
      This function copied from 'python-docx' library
      Replace all occurences of string with a different string, return updated document
    '''
    text_tag = norm_name('w:t')
    newdocument = self.docbody
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == text_tag :
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search,replace,element.text)
    return newdocument

  def get_numbering_left(self, style):
    '''
       get numbering indeces
    '''
    abstractNums=get_elements(self.numbering, 'w:abstractNum')

    indres=[0]

    for x in abstractNums :
      styles=get_elements(x, 'w:lvl/w:pStyle')
      if styles :
        pstyle_name = styles[0].get(norm_name('w:val') )
        if pstyle_name == style :
          ind=get_elements(x, 'w:lvl/w:pPr/w:ind')
	  if ind :
            indres=[]
	    for indx in ind :
              indres.append(int(indx.get(norm_name('w:left'))))
          return indres
    return indres

############
##  Numbering
  def get_numbering_style_id(self, style):
    try:
      style_elems = get_elements(self.styles, '/w:styles/w:style')
      for style_elem in style_elems:
        name_elem = get_elements(style_elem,'w:name')[0]
        name = name_elem.attrib[norm_name('w:val')]
	if name == style :
            numPr = get_elements(style_elem,'w:pPr/w:numPr/w:numId')[0]
            value = numPr.attrib[norm_name('w:val')]
            return value
    except: 
      pass
    return '0'

  def get_numbering_ids(self):
      num_elems = get_elements(self.numbering, '/w:numbering/w:num')
      result = []
      for num_elem in num_elems :
        nid = num_elem.attrib[norm_name('w:numId')]
        result.append( nid )
      return result

  def get_max_numbering_id(self):
      max_id = 0
      num_ids = self.get_numbering_ids()
      for x in num_ids :
	if int(x) > max_id :  max_id = int(x)
      return max_id


##########

  def getdocumenttext(self):
    '''
      This function copied from 'python-docx' library
      Return the raw text of a document, as a list of paragraphs.
    '''
    paragraph_tag == norm_nama('w:p')
    text_tag == norm_nama('w:text')
    paratextlist=[]   
    # Compile a list of all paragraph (p) elements
    paralist = []
    for element in self.document.iter():
        # Find p (paragraph) elements
        if element.tag == paragraph_tag:
            paralist.append(element)    
    # Since a single sentence might be spread over multiple text elements, iterate through each 
    # paragraph, appending all text (t) children to that paragraphs text.     
    for para in paralist:      
        paratext=u''  
        # Loop through each paragraph
        for element in para.iter():
            # Find t (text) elements
            if element.tag == text_tag:
                if element.text:
                    paratext = paratext+element.text
        # Add our completed paragraph text to the list of paragraph text    
        if not len(paratext) == 0:
            paratextlist.append(paratext)                    
    return paratextlist        

#
# DocxComposer Class
#
class DocxComposer:
  def __init__(self, stylefile=None):
    '''
       Constructor
    '''
    self._coreprops=None
    self._appprops=None
    self._contenttypes=None
    self._websettings=None
    self._wordrelationships=None
    self.breakbefore = False
    self.last_paragraph = None
    self.stylenames = {}
    self.title = ""
    self.subject = ""
    self.creator = "Python:DocDocument"
    self.company = ""
    self.category = ""
    self.descriptions = ""
    self.keywords = []


    if stylefile == None :
      self.template_dir = None
    else :
      self.new_document(stylefile)

  def set_style_file(self, stylefile):
    '''
       Set style file 
    '''
    fname = find_file(stylefile, 'sphinx-docxbuilder/docx')

    if fname == None:
      print "Error: style file( %s ) not found" % stylefile
      return None
      
    self.styleDocx = DocxDocument(fname)

    self.template_dir = tempfile.mkdtemp(prefix='docx-')
    result = self.styleDocx.extract_files(self.template_dir)

    if not result :
      print "Unexpected error in copy_docx_to_tempfile"
      shutil.rmtree(temp_dir, True)
      self.template_dir = None
      return 

    self.stylenames = self.styleDocx.extract_stylenames()
    self.paper_info = self.styleDocx.get_paper_info()
    self.bullet_list_indents = self.get_numbering_left('ListBullet')
    self.bullet_list_numId = self.styleDocx.get_numbering_style_id('ListBullet')
    self.number_list_indent = self.get_numbering_left('ListNumber')[0]
    self.number_list_numId = self.styleDocx.get_numbering_style_id('ListNumber')

    return

  def delete_template(self):
    '''
       Delete the temporary directory which we use compose a new document. 
    '''
    shutil.rmtree(self.template_dir, True)

  def get_numbering_left(self, style):
    '''
       Get numbering indeces...
    '''
    return self.styleDocx.get_numbering_left(style)

  def new_document(self, stylefile):
    '''
       Preparing a new document
    '''
    self.set_style_file(stylefile)
    self.document = self.makeelement('w:document')
    self.document.append(self.makeelement('w:body'))
    self.docbody = get_elements(self.document, '/w:document/w:body')[0]

    self.relationships = self.relationshiplist()

    return self.document

  def set_props(self, title, subject, creator, company='', category='', descriptions='', keywords=[]):
    '''
      Set document's properties: title, subject, creatro, company, category, descriptions, keywrods.
    '''
    self.title = title
    self.subject = subject
    self.creator = creator
    self.company = company
    self.category = category
    self.descriptions = descriptions
    self.keywords = keywords

  def save(self, docxfilename):
    '''
      Save the composed document to the docx file 'docxfilename'.
    '''
    assert os.path.isdir(self.template_dir)

    self.coreproperties()
    self.appproperties()
    self.contenttypes()
    self.websettings()

    self.wordrelationships()

    self.docbody.append(self.paper_info)

    # Serialize our trees into out zip file
    treesandfiles = {self.document:'word/document.xml',
                     self._coreprops:'docProps/core.xml',
                     self._appprops:'docProps/app.xml',
                     self._contenttypes:'[Content_Types].xml',
                     self.styleDocx.numbering:'word/numbering.xml',
                     self.styleDocx.styles:'word/styles.xml',
                     self._websettings:'word/webSettings.xml',
                     self._wordrelationships:'word/_rels/document.xml.rels'}

    docxfile = self.styleDocx.restruct_docx(self.template_dir, docxfilename, treesandfiles.values())

    for tree in treesandfiles:
        if tree != None:
            #print 'Saving: '+treesandfiles[tree]    
            treestring =  etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone='yes')
            docxfile.writestr(treesandfiles[tree],treestring)
    
    print 'Saved new file to: '+docxfilename
    shutil.rmtree(self.template_dir)
    return
    
  def get_child(self, element, tag):
    '''
      Get child elements...
    '''
    return element.xpath(tag, namespaces=nsprefixes)

  def make_element(self,tagname,tagtext=None):
    '''
      Make an element without attributes
    '''
    newele = etree.Element(norm_name(tagname), nsmap=nsprefixes)
    if tagtext :
      newele.text = tagtext
    return newele

  def set_attribute(self, element, key, value):
    '''
      Set an attribute of the element
    '''
    element.set(norm_name(key), str(value))

  def set_attributes(self, ele, attributes):
    '''
      Set attributes of the element
    '''
    if not attributes :
      return

    for attr in attributes:
      ele.set(norm_name(attr), attributes[attr])

  def makeelement(self,tagname,tagtext=None,attributes=None):
    '''
      Make an element with attributes
    '''
    newelement = self.make_element(tagname, tagtext)
    if attributes :
      self.set_attributes(newelement, attributes)
    return newelement

  def append(self, para):
    '''
      Append paragraph to document
    '''
    self.docbody.append(para)
    self.last_paragraph = para
    return para

  def pagebreak(self,type='page', orient='portrait'):
    '''
      Insert a break, default 'page'.
      See http://openxmldeveloper.org/forums/thread/4075.aspx
      Return our page break element.

      This method is copied from 'python-docx' library
    '''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if type not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))
    pagebreak = self.makeelement('w:p')
    if type == 'page':
        run = self.makeelement('w:r')
	br = self.makeelement('w:br',attributes={'w:type':type})
        run.append(br)
        pagebreak.append(run)
    elif type == 'section':
        pPr = self.makeelement('w:pPr')
	sectPr = self.makeelement('w:sectPr')
        if orient == 'portrait':
            pgSz = self.makeelement('w:pgSz',attributes={'w:w':'12240','w:h':'15840'})
        elif orient == 'landscape':
            pgSz = self.makeelement('w:pgSz',attributes={'w:h':'12240','w:w':'15840', 'w:orient':'landscape'})
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)

    self.docbody.append(pagebreak)
    self.breakbrefore = True
    return pagebreak    

  def paragraph(self, paratext, style='BodyText', block_level=0, create_only=False):
    '''
      Make a new paragraph element, containing a run, and some text. 
      Return the paragraph element.
    '''
    isliteralblock=False
    if style == 'LiteralBlock' :
      paratext = paratext[0].splitlines()
      isliteralblock=True

    # Make paragraph elements
    paragraph = self.makeelement('w:p')
    self.insert_paragraph_property(paragraph, style)
                
    run = None

    # Insert lastRenderedPageBreak for assistive technologies like
    # document narrators to know when a page break occurred.
    if self.breakbefore :
        run = self.makeelement('w:r')    
	lastRenderedPageBreak = self.makeelement('w:lastRenderedPageBreak')
        run.append(lastRenderedPageBreak)
        paragraph.append(run)    

    #  Insert a text run
    if paratext != None:
        self.make_runs(paragraph, paratext, isliteralblock)

    if block_level > 0 :
        ind = self.number_list_indent * block_level
        self.set_indent(paragraph, ind)

    #  if the 'create_only' flag is True, append paragraph to the document
    if not create_only :
        self.docbody.append(paragraph)
        self.last_paragraph = paragraph

    return paragraph

  def insert_paragraph_property(self, paragraph, style='BodyText'):
    '''
       Insert paragraph property element with style.
    '''
    pPr = self.makeelement('w:pPr')
    if style not in self.stylenames :
      self.new_paragraph_style(style)
    style = self.stylenames.get(style, 'BodyText')
    pStyle = self.makeelement('w:pStyle',attributes={'w:val':style})
    pPr.append(pStyle)

    paragraph.append(pPr) 
    return paragraph

  def get_paragraph_style(self, paragraph):
    '''
       Get stylename of the paragraph
    '''

    pStyle = self.get_child(paragraph, 'w:pPr/w:pStyle')
    if not pStyle :
      return 'BodyText'
    return pStyle[0].attrib[norm_name('w:val')]

  def get_numbering_indent(self, style='ListBullet', lvl=0, nId=0):
    '''
       Get indenent value
    '''
    result = 0

    if style == 'ListBullet' or nId == 0 :
      if len(self.bullet_list_indents) > lvl :
        result = self.bullet_list_indents[lvl]
      else:
        result = self.bullet_list_indents[-1]
    else:
      result = self.number_list_indent * (lvl+1)

    return result
     
  def insert_numbering_property(self, paragraph, lvl=0, nId=0, start=1, enum_prefix=None, enum_type=None):
    '''
       Insert paragraph property element with style.
    '''
    style=self.get_paragraph_style(paragraph)
    pPr = self.get_child(paragraph, 'w:pPr')
    if not pPr :
      self.insert_paragraph_property(paragraph)
      pPr = self.get_child(paragraph, 'w:pPr')

    numPr = self.makeelement('w:numPr')
    if style == 'ListNumber':
      ilvl = self.makeelement('w:ilvl',attributes={'w:val': '0'})
    else:
      ilvl = self.makeelement('w:ilvl',attributes={'w:val': str(lvl)})
    numPr.append(ilvl)

    lvl_text='%1.'
    if nId <= 0 :
      if nId == 0 :
        num_id = '0'
      else :
        num_id = self.styleDocx.get_numbering_style_id(style)
    else :
      num_id = str(nId)
      if num_id not in self.styleDocx.get_numbering_ids() :
	if enum_prefix : lvl_text=enum_prefix
        newid = self.styleDocx.get_max_numbering_id()+1
	if newid < nId : newid = nId
        num_id = str(self.new_ListNumber_style(newid, start, lvl_text, enum_type))

    numId = self.makeelement('w:numId',attributes={'w:val': num_id})
    numPr.append(numId)

    pPr[0].append(numPr)

    sty = self.get_paragraph_style(paragraph)
    ind = self.get_numbering_indent(sty, lvl, nId)
    self.set_indent(paragraph, ind)

    #print ">>> numId: indent=%s, nId=%d, num_id=%s, (%s) [%s]" % (ind, nId, num_id, enum_type, lvl_text)
    return pPr

  def set_indent(self, paragraph, lskip):
    '''
       Set indent of paragraph
    '''
    pPr = self.get_child(paragraph, 'w:pPr')
    if not pPr :
      self.insert_paragraph_property(paragraph)
      pPr = get_child(paragraph, 'w:pPr')

    ind = self.get_child(pPr[0], 'w:ind')
    if not ind :
      ind = self.makeelement('w:ind',attributes={'w:left': str(lskip)})
      pPr[0].append(ind)
    else:
      self.set_attribute(ind[0], 'w:left', lskip)

    return pPr

  def make_runs(self, paragraph, targettext, literal_block=False):
    '''
      Make new runs with text.
    '''
    if isinstance(targettext, (list)) :
        for i,x in enumerate(targettext) :
            if isinstance(x, (list)) :
                run = self.make_run(x[0], style=x[1])
            else:
                run = self.make_run(x)
            paragraph.append(run) 
	    if literal_block and i+1 <  len(targettext) :
                paragraph.append( self.make_run(':br') )
    else:
        run = self.make_run(targettext)
        paragraph.append(run)    
    return paragraph

  def make_run(self, txt, style='Normal', create_only=True):
    '''
      Make a new styled run from text.
    '''
    # Make run element
    run = self.makeelement('w:r')  

    if txt == ":br" :
      text = self.makeelement('w:cr')
      run.append(text)
    else:
      text = self.makeelement('w:t',tagtext=txt)

      # if the txt contain spaces, we should add an attribute 'xml:space="preserve"' to w:text-tag.
      if txt.find(' ') != -1 :
        self.set_attribute(text, 'xml:space','preserve')

      if style != 'Normal' :
        if style not in self.stylenames :
          self.new_character_style(style)

        style = self.stylenames.get(style, 'Normal')
	rPr = self.makeelement('w:rPr')
	rStyle = self.makeelement('w:rStyle',attributes={'w:val':style})
        rPr.append(rStyle)
        run.append(rPr)    
                
      # Add the text the run
      run.append(text)    

    if not create_only :
      if self.last_paragraph == None:
        self.paragraph(None)
      self.last_paragraph.append(run)

    return run

  def add_br(self):
    '''
      append line break in current paragraph
    '''
    run = self.makeelement('w:r')    
    text = self.makeelement('w:br','')
    run.append(text)    

    if self.last_paragraph == None:
        self.paragraph(None)

    self.last_paragraph.append(run)    
    return run

  def add_space(self, style='Normal'):
    '''
      append a space in current paragraph
    '''
    # Make rum element
    run = self.makeelement('w:r')    
    text = self.makeelement('w:t',' ', attributes={'xml:space':'preserve'})

    if style != 'Normal' :
      rPr = self.makeelement('w:rPr')
      style = self.stylenames.get(style, 'Normal')
      rStyle = self.makeelement('w:rStyle',attributes={'w:val':style})
      rPr.append(rStyle)
      run.append(rPr)    
                
    # Add the text the run
    run.append(text)    

    if self.last_paragraph == None:
        self.paragraph(None)

    # append the run to last paragraph
    self.last_paragraph.append(run)    
    return run

  def heading(self, headingtext, headinglevel):
    '''
      Make a heading
    '''
    # Make paragraph element
    paragraph = self.makeelement('w:p')
    self.insert_paragraph_property(paragraph, 'Heading'+str(headinglevel))

    self.make_runs(paragraph, headingtext)

    self.last_paragraph = paragraph
    self.docbody.append(paragraph)

    return paragraph   

  def list_item(self, itemtext, style='ListBullet', lvl=1, nid=0, enum_prefix=None, enum_prefix_type=None, start=1):
    '''
      Make a new list paragraph
    '''
    # Make paragraph element
    paragraph = self.makeelement('w:p')

    self.insert_paragraph_property( paragraph, style)
    self.insert_numbering_property(paragraph, lvl-1, nid, start, enum_prefix, enum_prefix_type)
    self.make_runs(paragraph, itemtext)

    self.last_paragraph = paragraph
    self.docbody.append(paragraph)

    return paragraph   

  def get_max_numbering_id(self):
    return self.styleDocx.get_max_numbering_id()

  def new_ListNumber_style(self, nId, start_val=1, lvl_txt='%1.', typ=None):
    '''
      create new List Number style 
    '''
    orig_numid = self.number_list_numId
    #newid = int(max(self.styleDocx.get_numbering_ids()))+1
    newid = nId

    num = self.makeelement('w:num', attributes={'w:numId':str(newid)})
    abstNum = self.makeelement('w:abstrctNumId', attributes={'w:val':orig_numid})
    lvlOverride = self.makeelement('w:lvlOverride', attributes={'w:ilvl':'0'})
    start = self.makeelement('w:startOverride', attributes={'w:val':str(start_val)})
    lvlOverride.append(start)

    lvl = self.makeelement('w:lvl', attributes={'w:ilvl':'0'})
    lvlText = self.makeelement('w:lvlText', attributes={'w:val': lvl_txt})
    lvl.append(lvlText)
    typ =  get_enumerate_type(typ)
#    print ">>> %d: %s  %s  %d <<<" % (newid, typ, lvl_txt, start_val)
    numFmt = self.makeelement('w:numFmt', attributes={'w:val': typ})
    lvl.append(numFmt)
#    lvlJc = self.makeelement('w:lvlJc', attributes={'w:val': "start"})
#    lvl.append(lvlJc)
    lvlOverride.append(lvl)

    num.append(abstNum)
    num.append(lvlOverride)

    self.styleDocx.numbering.append(num)
    return  newid

  def new_character_style(self, styname):
    newstyle = self.makeelement('w:style', attributes={'w:type':'character','w:customStye':'1', 'w:styleId': styname})
    name = self.makeelement('w:name', attributes={'w:val': styname})
    base = self.makeelement('w:basedOn', attributes={'w:val': self.styleDocx.character_style_id})
    rPr = self.makeelement('w:rPr')
    clr = self.makeelement('w:color', attributes={'w:val': 'FF0000'})
    rPr.append(clr)

    newstyle.append(name)
    newstyle.append(base)
    newstyle.append(rPr)

    self.styleDocx.styles.append(newstyle)
    self.stylenames[styname] = styname
    return styname

  def new_paragraph_style(self, styname):
    newstyle = self.makeelement('w:style', attributes={'w:type':'paragraph','w:customStye':'1', 'w:styleId': styname})
    name = self.makeelement('w:name', attributes={'w:val': styname})
    base = self.makeelement('w:basedOn', attributes={'w:val': self.styleDocx.paragraph_style_id})
    qF = self.makeelement('w:qFormat')

    newstyle.append(name)
    newstyle.append(base)
    newstyle.append(qF)

    self.styleDocx.styles.append(newstyle)
    self.stylenames[styname] = styname
    return styname

  def table(self, contents):
    '''
      Get a list of lists, return a table
      This function is copied from 'python-docx' library
    '''
    table = self.makeelement('w:tbl')
    columns = len(contents[0][0])    
    # Table properties
    tableprops = self.makeelement('w:tblPr')
    tablestyle = self.makeelement('w:tblStyle',attributes={'w:val':'ColorfulGrid-Accent1'})
    tablewidth = self.makeelement('w:tblW',attributes={'w:w':'0','w:type':'auto'})
    tablelook = self.makeelement('w:tblLook',attributes={'w:val':'0400'})
    for tableproperty in [tablestyle,tablewidth,tablelook]:
        tableprops.append(tableproperty)
    table.append(tableprops)    
    # Table Grid    
    tablegrid = self.makeelement('w:tblGrid')
    for _ in range(columns):
        tablegrid.append(self.makeelement('w:gridCol',attributes={'w:w':'2390'}))
    table.append(tablegrid)     
    # Heading Row    
    row = self.makeelement('w:tr')
    rowprops = self.makeelement('w:trPr')
    cnfStyle = self.makeelement('w:cnfStyle',attributes={'w:val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    for heading in contents[0]:
        cell = self.makeelement('w:tc')  
        # Cell properties  
        cellprops = self.makeelement('w:tcPr')
        cellwidth = self.makeelement('w:tcW',attributes={'w:w':'2390','w:type':'dxa'})
        cellstyle = self.makeelement('w:shd',attributes={'w:val':'clear','w:color':'auto','w:fill':'548DD4','w:themeFill':'text2','w:themeFillTint':'99'})
        cellprops.append(cellwidth)
        cellprops.append(cellstyle)
        cell.append(cellprops)        
        # Paragraph (Content)
        cell.append(self.paragraph(heading, create_only=True))
        row.append(cell)
    table.append(row)            
    # Contents Rows   
    for contentrow in contents[1:]:
        row = self.makeelement('w:tr')     
        for content in contentrow:   
            cell = self.makeelement('w:tc')
            # Properties
	    cellprops = self.makeelement('w:tcPr')
	    cellwidth = self.makeelement('w:tcW',attributes={'w:type':'dxa'})
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            cell.append(self.paragraph(content, create_only=True))
            row.append(cell)    
        table.append(row)   

    self.docbody.append(table)
    return table                 

  def picture(self, picname, picdescription, pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True, align='center'):
    '''
      Take a relationshiplist, picture file name, and return a paragraph containing the image
      and an updated relationshiplist
      
      This function is copied from 'python-docx' library
    '''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''  
    # Copy the file into the media dir
    media_dir = join(self.template_dir,'word','media')
    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
    picpath, picname = os.path.abspath(picname), os.path.basename(picname)
    shutil.copyfile(picpath, join(media_dir,picname))
    relationshiplist = self.relationships

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth,pixelheight = Image.open(picpath).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs            
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)   
    
    # Set relationship ID to the first available  
    picid = '2'    
    picrelid = 'rId'+str(len(relationshiplist)+1)
    relationshiplist.append([
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])
    
    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = self.makeelement('pic:blipFill')
    blipfill.append(self.makeelement('a:blip',attributes={'r:embed':picrelid}))
    stretch = self.makeelement('a:stretch')
    stretch.append(self.makeelement('a:fillRect'))
    blipfill.append(self.makeelement('a:srcRect'))
    blipfill.append(stretch)
    
    # 2. The non visual picture properties 
    nvpicpr = self.makeelement('pic:nvPicPr')
    cnvpr = self.makeelement('pic:cNvPr', attributes={'id':'0','name':'Picture 1','descr':picname}) 
    nvpicpr.append(cnvpr) 
    cnvpicpr = self.makeelement('pic:cNvPicPr')                           
    cnvpicpr.append(self.makeelement('a:picLocks',
                    attributes={'noChangeAspect':str(int(nochangeaspect)),
                    'noChangeArrowheads':str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)
        
    # 3. The Shape properties
    sppr = self.makeelement('pic:spPr',attributes={'bwMode':'auto'})
    xfrm = self.makeelement('a:xfrm')
    xfrm.append(self.makeelement('a:off',attributes={'x':'0','y':'0'}))
    xfrm.append(self.makeelement('a:ext',attributes={'cx':width,'cy':height}))
    prstgeom = self.makeelement('a:prstGeom',attributes={'prst':'rect'})
    prstgeom.append(self.makeelement('a:avLst'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    a_nofill = self.makeelement('a:noFill')
    a_ln = self.makeelement('a:ln')
    a_nofill2 = self.makeelement('a:noFill')
    a_ln.append(a_nofill2)
    sppr.append(a_nofill)
    sppr.append(a_ln)
    
    # Add our 3 parts to the picture element
    pic = self.makeelement('pic:pic')    
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)
    
    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = self.makeelement('a:graphicData',
        attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = self.makeelement('a:graphic')
    graphic.append(graphicdata)

    framelocks = self.makeelement('a:graphicFrameLocks',attributes={'noChangeAspect':'1'})    
    framepr = self.makeelement('wp:cNvGraphicFramePr')
    framepr.append(framelocks)
    docpr = self.makeelement('wp:docPr',
        attributes={'id':picid,'name':'Picture 1','descr':picdescription})
    effectextent = self.makeelement('wp:effectExtent',
        attributes={'l':'25400','t':'0','r':'0','b':'0'})
    extent = self.makeelement('wp:extent',attributes={'cx':width,'cy':height})
    inline = self.makeelement('wp:inline',
        attributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"})
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = self.makeelement('w:drawing')
    drawing.append(inline)
    run = self.makeelement('w:r')
    rPr = self.makeelement('w:rPr')
    noProof = self.makeelement('w:noProof')
    rPr.append(noProof)
    run.append(rPr)
    run.append(drawing)
    paragraph = self.makeelement('w:p')
    pPr = self.makeelement('w:pPr')
    jc = self.makeelement('w:jc', attributes={'w:val':align})
    pPr.append(jc)
    paragraph.append(pPr)
    paragraph.append(run)

    self.relationships = relationshiplist
    self.docbody.append(paragraph)

    self.last_paragraph = None
    return paragraph


  def contenttypes(self):
    '''
       create [Content_Types].xml 
       This function copied from 'python-docx' library
    '''
    prev_dir = os.getcwd() # save previous working dir
    os.chdir(self.template_dir)

    filename = '[Content_Types].xml'
    if not os.path.exists(filename):
        raise RuntimeError('You need %r file in template' % filename)

    parts = dict([
        (x.attrib['PartName'], x.attrib['ContentType'])
        for x in etree.fromstring(open(filename).read()).xpath('*')
        if 'PartName' in x.attrib
    ])

    # FIXME - doesn't quite work...read from string as temp hack...
    #types = self.makeelement('Types',nsprefix='ct')
    types = etree.fromstring('''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
    for part in parts:
        types.append(self.makeelement('Override',attributes={'PartName':part,'ContentType':parts[part]}))
    # Add support for filetypes
    filetypes = {'rels':'application/vnd.openxmlformats-package.relationships+xml',
                 'xml':'application/xml',
                 'jpeg':'image/jpeg',
                 'jpg':'image/jpeg',
                 'gif':'image/gif',
                 'png':'image/png'}

    for extension in filetypes:
        types.append(self.makeelement('Default',attributes={'Extension':extension,'ContentType':filetypes[extension]}))

    os.chdir(prev_dir)
    self._contenttypes = types
    return types

  def coreproperties(self,lastmodifiedby=None):
    '''
      Create core properties (common document properties referred to in the 'Dublin Core' specification).
      See appproperties() for other stuff.
       This function copied from 'python-docx' library
    '''

    coreprops = self.makeelement('cp:coreProperties')    
    coreprops.append(self.makeelement('dc:title',tagtext=self.title))
    coreprops.append(self.makeelement('dc:subject',tagtext=self.subject))
    coreprops.append(self.makeelement('dc:creator',tagtext=self.creator))
    coreprops.append(self.makeelement('cp:keywords',tagtext=','.join(self.keywords)))    
    if not lastmodifiedby:
        lastmodifiedby = self.creator
    coreprops.append(self.makeelement('cp:lastModifiedBy',tagtext=lastmodifiedby))
    coreprops.append(self.makeelement('cp:revision',tagtext='1'))
    coreprops.append(self.makeelement('cp:category',tagtext=self.category))
    coreprops.append(self.makeelement('dc:description',tagtext=self.descriptions))
    currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')

    # Document creation and modify times
    # Prob here: we have an attribute who name uses one namespace, and that 
    # attribute's value uses another namespace.
    # We're creating the lement from a string as a workaround...
    for doctime in ['created','modified']:
        coreprops.append(etree.fromstring('''<dcterms:'''+doctime+''' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">'''+currenttime+'''</dcterms:'''+doctime+'''>'''))
        pass

    self._coreprops = coreprops
    return coreprops

  def appproperties(self):
    '''
       Create app-specific properties. See docproperties() for more common document properties.
       This function copied from 'python-docx' library
    '''
    appprops = self.makeelement('ep:Properties')
    appprops = etree.fromstring(
    '''<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>''')
    props = {
            'Template':'Normal.dotm',
            'TotalTime':'6',
            'Pages':'1',  
            'Words':'83',   
            'Characters':'475', 
            'Application':'Microsoft Word 12.0.0',
            'DocSecurity':'0',
            'Lines':'12', 
            'Paragraphs':'8',
            'ScaleCrop':'false', 
            'LinksUpToDate':'false', 
            'CharactersWithSpaces':'583',  
            'SharedDoc':'false',
            'HyperlinksChanged':'false',
            'AppVersion':'12.0000',    
            'Company':self.company,    
            }
    for prop in props:
        appprops.append(self.makeelement(prop,tagtext=props[prop]))

    self._appprops = appprops
    return appprops


  def websettings(self):
    '''
      Generate websettings
      This function copied from 'python-docx' library
    '''
    web = self.makeelement('w:webSettings')
    web.append(self.makeelement('w:allowPNG'))
    web.append(self.makeelement('w:doNotSaveAsSingleFile'))

    self._websettings = web
    return web

  def relationshiplist(self):
    prev_dir = os.getcwd() # save previous working dir
    os.chdir(self.template_dir)

    filename = 'word/_rels/document.xml.rels'
    if not os.path.exists(filename):
        raise RuntimeError('You need %r file in template' % filename)

    relationships = etree.fromstring(open(filename).read())
    relationshiplist = [
            [x.attrib['Type'], x.attrib['Target']]
            for x in relationships.xpath('*')
    ]

    os.chdir(prev_dir)

    return relationshiplist

  def wordrelationships(self):
    '''
      Generate a Word relationships file
      This function copied from 'python-docx' library
    '''
    # Default list of relationships
    # FIXME: using string hack instead of making element
    #relationships = self.makeelement('pr:Relationships',nsprefix='pr')    

    relationships = etree.fromstring(
    '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>'''    
    )
    count = 0
    for relationship in self.relationships:
        # Relationship IDs (rId) start at 1.
        relationships.append(self.makeelement('Relationship',attributes={'Id':'rId'+str(count+1),
        'Type':relationship[0],'Target':relationship[1]}))
        count += 1

    self._wordrelationships = relationships
    return relationships    

