Introduction
============

A set of utilities to transform to and from WordML.

An HTML to WordML transform is the only transform currently available.

It can be used both on the command line or called directly from your
Python application.

Command line usage
==================

usage: html2wordml.py [-h] [-c] -p BASEPATH htmlfile

Convert HTML to WordML

positional arguments:
  htmlfile              /path/to/htmlfile

optional arguments:
  -h, --help            show this help message and exit
  -c, --create-package  Create WordML package
  -p BASEPATH, --basepath BASEPATH
                        Base path for relative urls

For examle:

./html2wordml.py -c -p ./tests ./tests/test.html > test.docx


From Python
===========

>>> from upfront.wordmlutils.html2wordml import transform
>>> from StringIO import StringIO
>>> html = """<html>
<body>
<p>Hallo world!</p>
</body>
</html>"""
>>> basepath = '/'
>>> output = StringIO()
>>> transform(basepath, html, outfile=output)
>>> docxfile = output.getvalue()

The transform method above also take an optional image_resolver
argument. The image_resolver argument must be an instance of a class
that has a get_images method and returns a list of images where each
element in the list is the image data itself. Look at the default
get_images method in html2wordml.py as reference.
