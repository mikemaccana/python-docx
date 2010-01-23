Adding Features
===============

# Recommended reading

- The [LXML tutorial](http://codespeak.net/lxml/tutorial.html) covers the basics of XML etrees, which we create, append and insert to make XML documents. LXML also provides XPath, which we use to specify locations in the document. 
- If you're stuck. check out the [OpenXML specs and videos](http://openxmldeveloper.org). In particular, the is [OpenXML ECMA spec] [] is well worth a read.
- Learning about [XML namespaces](http://www.w3schools.com/XML/xml_namespaces.asp)
- The [Namespaces section of Dive into Python](http://diveintopython3.org/xml.html)
- Microsoft's [introduction to the Office (2007) Open XML File Formats](http://msdn.microsoft.com/en-us/library/aa338205.aspx)

# How can I contribute?

Fork the project on github, then send the main project a [pull request](http://github.com/guides/pull-requests). The project will then accept your pull (in most cases), which will show your changes part of the changelog for the main project, along with your name and picture.

# A note about namespaces and LXML

LXML doesn't use namespace prefixes. It just uses the actual namespaces, and wants you to set a namespace on each tag. For example, rather than making an element with the 'w' namespace prefix, you'd make an element with the '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' prefix. 

To make this easier:

- The most common namespace, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' (prefix 'w') is automatically added by makeelement()
- You can specify other namespaces with 'nsprefix', which maps the prefixes Word files use to the actual namespaces, eg:

<pre>makeelement('coreProperties',nsprefix='cp')</pre>

will generate:

    <ns0:coreProperties xmlns:ns0="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">

which is the same as what Word generates:

    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">

The namespace prefixes are different, but that's irrelevant as the namespaces themselves are the same.

There's also a cool side effect - you can ignore setting 'xmlns' attributes that aren't used directly in the current element, since there's no need. Eg, you can make the equivalent of this from a Word file:

	<cp:coreProperties 
	xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
	xmlns:dc="http://purl.org/dc/elements/1.1/" 
	xmlns:dcterms="http://purl.org/dc/terms/" 
	xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	</cp:coreProperties>

With the following code:
  
	docprops = makeelement('coreProperties',nsprefix='cp')

We only need to specify the 'cp' prefix because that's what this element uses. The other 'xmlns' attributes are used to specify the prefixes for child elements. We don't need to specify them here because each child element will have its namespace specified when we make that child.

# Coding Style 

Basically just look at what's there. But if you need something more specific:

- Functional - every function should take some inputs, return something, and not use any globals.
- [Google Python Style Guide style](http://code.google.com/p/soc/wiki/PythonStyleGuide)

# Unit Testing

After adding code, open **tests/test_docx.py** and add a test that calls your function and checks its output.

- Use **easy_install** to fetch the **nose** and **coverage** modules
- Run 

<pre>nosetests --with-coverage</pre>

to run all the doctests. They should all pass.

# Tips

## If Word complains about files:

First, determine whether Word can recover the files:
- If Word cannot recover the file, you most likely have a problem with your zip file
- If Word can recover the file, you most likely have a problem with your XML

### Common Zipfile issues

- Ensure the same file isn't included twice in your zip archive. Zip supports this, Word doesn't.
- Ensure that all media files have an entry for their file type in [Content_Types].xml
- Ensure that files in zip file file have leading '/'s removed. 

### Common XML issues

- Ensure the _rels, docProps, word, etc directories are in the top level of your zip file.
- Check your namespaces - on both the tags, and the attributes
- Check capitalization of tag names
- Ensure you're not missing any attributes
- If images or other embedded content is shown with a large red X, your relationships file is missing data.

#### One common debugging technique we've used before

- Re-save the document in Word will produced a fixed version of the file
- Unzip and grabbing the serialized XML out of the fixed file
- Use etree.fromstring() to turn it into an element, and include that in your code.
- Check that a correct file is generated
- Remove an element from your string-created etree (including both opening and closing tags)
- Use element.append(makelement()) to add that element to your tree
- Open the doc in Word and see if it still works
- Repeat the last three steps until you discover which element is causing the prob

[OpenXML ECMA spec]: http://www.ecma-international.org/publications/files/ECMA-ST/Office%20Open%20XML%201st%20edition%20Part%204%20(DOCX).zip