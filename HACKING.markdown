Adding Features
===============

# Recommended reading

- The LXML tutorial at http://codespeak.net/lxml/tutorial.html covers the basics of XML etrees, which we create append and insert to make XML documents
- The OpenXML WordML specs and videos at http://openxmldeveloper.org (if you're stuck)
- Learning about XML namespaces http://www.w3schools.com/XML/xml_namespaces.asp
- The Namespaces section of http://diveintopython3.org/xml.html

# A note about namespaces and LXML

LXML doesn't use namespace prefixes. It just uses the actual namespaces, and wants you to set a namespace on each tag. For example, rather than making an element with the 'w' namespace prefix, you'd make an element with the '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' prefix. 

To make this easier:

- The most common namespace, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' (prefix 'w') is automatically added by makeelement()
- You can specify other namespaces with getns(), which maps the prefixes Word files use to the actual namespaces, eg:
	makeelement('coreProperties',tagnamespace=getns(propns,'cp'))

There's also a cool sideeffect - you can ignore setting all 'xmlns' attributes, since there's no need. Eg, this in word:
	<cp:coreProperties 
	xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
	xmlns:dc="http://purl.org/dc/elements/1.1/" 
	xmlns:dcterms="http://purl.org/dc/terms/" 
	xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	</cp:coreProperties>
  
	docprops = makeelement('coreProperties',tagnamespace=getns(propns,'cp'))

# Coding Style 

Basically just look at what's there. But if you need something more specific:

- Functional - every function should take some inputs, return something, and not use any globals.
- Google style - http://code.google.com/p/soc/wiki/PythonStyleGuide
- Unit tests are handled with nose / coverage

# Tips

## If Word complains about files:

- Your zip file or XML file has a problem
- Ensure the same file isn't included twice in your zip archive. Zip supports this, Word doesn't.
- Ensure the _rels, docProps, word, etc directories are in the top level of your zip file.
- Check your namespaces - on both the tags, and the attributes
- Check capitalization of tag names
- Ensure you're not missing any attributes

## One common debugging technique we've used before

- Re-save the document in Word will produced a fixed version of the file
- Unzip and grabbing the serialized XML out of the fixed file
- Use etree.fromstring() to turn it into an element, and include that in your code.
- Check that a correct file is generated
- Remove an element from your string-created etree (including both opening and closing tags)
- Use element.append(makelement()) to add that element to your tree
- Open the doc in Word and see if it still works
- Repeat the last three steps until you discover which element is causing the prob
