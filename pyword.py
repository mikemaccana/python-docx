#!/usr/bin/env python26
'''
Open and modify Microsoft Word 2007 docx files (called 'Office OpenXML' by Microsoft)

TODO:
- Return all text in document
- return document properties dict
- Read word XML reference 
- Functions to recieve dict and put into table
key is left column, data is right column
- Package for easy_install
- rest converter??
'''

from lxml import etree
import zipfile
#import ipdb

wordnamespaces = {
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
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    }

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)    
    return document

def makeelement(tagname,tagattributes=None,tagtext=None):
    '''Create an element & return it'''
    namespace = '{'+wordnamespaces['w']+'}'    
    newelement = etree.Element(namespace+tagname)
    # Add attributes with namespaces
    if tagattributes:
        for tagattribute in tagattributes:
            newelement.set('{'+wordnamespaces['w']+'}'+tagattribute, tagattributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement
    

def addparagraph(paratext):
    '''Make a new paragraph element, containing a run, and some text. Return the paragraph element.'''
    global document
    # Make our elements
    paragraph = makeelement('p',tagattributes={'rsidR':'008D6863','rsidRDefault':'00590D07'})
    run = makeelement('r')
    text = makeelement('t',tagtext=paratext)
    # Add the text the run, and the run to the paragraph
    run.append(text)
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph
    
def search(phrase):
    '''Recieve a search, return the results'''
    results = False
    return results
    
def savedocx(oldfilename,document,newfilename=None):
    '''Save a modified document'''
    return    
    
if __name__ == '__main__':        
    #document = opendocx('Hello world.docx')
    document = etree.parse('sample.xml')
    #testelement = makeelement('success')
    #ipdb.set_trace()

    # This location is where most document content lives 
    documentbody = document.xpath('/w:document/w:body', namespaces=wordnamespaces)
    
    # Attach a paragraph element to our document
    newpara = addparagraph(paratext='success')    
    documentbody[0].insert(1, newpara)
    
    

    resultsfile = open('results','w')
    newxml = etree.tostring(document, pretty_print=True)
    print newxml
    resultsfile.write(newxml)    