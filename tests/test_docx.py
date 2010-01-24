#!/usr/bin/env python2.6
'''
Test docx module
'''
import os
import lxml
from nose import with_setup
from docx import *

TEST_FILE = 'Short python-docx test.docx'

def setup_func():
    '''Set up test fixtures'''
    testnewdocument()

def teardown_func():
    '''Tear down test fixtures'''
    if TEST_FILE in os.listdir('.'):
        os.remove(TEST_FILE)

def testunsupportedpagebreak():
    '''Ensure unsupported page break types are trapped'''
    document = newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    try:
        docbody.append(pagebreak(type='unsup'))
    except ValyueError:
        return # passed
    assert False # failed

def testnewdocument():
    '''Test that a new document can be created'''
    relationships = relationshiplist()
    document = newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    docbody.append(heading('Heading 1',1)  )   
    docbody.append(heading('Heading 2',2))
    docbody.append(paragraph('Paragraph 1'))
    for point in ['List Item 1','List Item 2','List Item 3']:
        docbody.append(paragraph(point,style='ListNumber'))
    docbody.append(pagebreak(type='page', orient='portrait'))
    docbody.append(paragraph('Paragraph 2')) 
    docbody.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))
    docbody.append(paragraph('Paragraph 3'))
    properties = docproperties('Python docx testnewdocument','A short example of making docx from Python','Alan Brooks',['python','Office Open XML','Word'])
    savedocx(document, properties, contenttypes(), websettings(), wordrelationships(relationships), TEST_FILE)

@with_setup(setup_func, teardown_func)
def testopendocx():
    '''Ensure an etree element is returned'''
    if isinstance(opendocx(TEST_FILE),lxml.etree._Element):
        pass
    else:
        assert False

def testmakeelement():
    '''Ensure custom elements get created'''
    testelement = makeelement('testname',attributes={'testattribute':'testvalue'},tagtext='testtagtext')
    assert testelement.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testname'
    assert testelement.attrib == {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testattribute': 'testvalue'}
    assert testelement.text == 'testtagtext'

def testparagraph():
    '''Ensure paragraph creates p elements'''
    testpara = paragraph('paratext',style='BodyText')
    assert testpara.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    pass
    
def testtable():
    '''Ensure tables make sense'''
    testtable = table([['A1','A2'],['B1','B2'],['C1','C2']])
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    assert testtable.xpath('/ns0:tbl/ns0:tr[2]/ns0:tc[2]/ns0:p/ns0:r/ns0:t',namespaces={'ns0':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0].text == 'B2'