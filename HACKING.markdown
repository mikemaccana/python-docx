Adding Features
===============

# Recommended reading: 
- The LXML tutorial at http://codespeak.net/lxml/tutorial.html covers the basics of XML etrees, which we create append and insert to make XML documents
- The OpenXML WordML specs and videos at http://openxmldeveloper.org (if you're stuck)

# Coding Style 
Basically just look at what's there. But uf you need something more specific:
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
