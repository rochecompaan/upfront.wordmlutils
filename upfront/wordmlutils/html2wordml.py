#!/usr/bin/env python

import os
import sys
import urllib
import zipfile
from cStringIO import StringIO
from lxml import etree, html
from PIL import Image

dirname = os.path.dirname(__file__)

def get_images(doc):
    images = []
    for img in doc.xpath('//img'):
        url = img.get('src')
        image = urllib.urlopen(url)
        images.append((url, StringIO(image.read())))
    return images
    
def convertPixelsToEMU(px):
    points = px * 72.0 / 96.0
    inches = points / 72.0
    emu = inches * 914400
    return int(emu)

def transform(htmlfile, xslfile):
    xslt_root = etree.XML(xslfile.read())
    transform = etree.XSLT(xslt_root)
    doc = html.parse(htmlfile)
    images = get_images(doc)
    result_tree = transform(doc)
    wordml = etree.tostring(result_tree)
    wordml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>' + \
        wordml

    template = zipfile.ZipFile(os.path.join(dirname, 'template.docx'))

    output = StringIO()
    zf = zipfile.ZipFile(output, 'w')
    namelist = template.namelist()
    docindex = namelist.index('word/document.xml')
    for count, img in enumerate(images):
        url, data = img
        # insert image before document
        namelist.insert(docindex, 'word/media/image%s' % count)

        # insert image sizes in the wordml
        img = Image.open(data)
        width, height = img.size

        # convert to EMU (English Metric Unit) 
        width = convertPixelsToEMU(width)
        height = convertPixelsToEMU(height)

        widthattr = '%s-$width' % url
        heightattr = '%s-$height' % url
        ridattr = '%s-$rid' % url
        wordml = wordml.replace(widthattr, str(width))
        wordml = wordml.replace(heightattr, str(height))
        wordml = wordml.replace(ridattr, 'rImageId%s' % count)

    print wordml

    for filename in namelist:
        if filename == 'word/document.xml':
            zf.writestr('word/document.xml', wordml)
        elif filename.startswith('word/media'):
            index = int(filename[-1])
            filename = 'word/media/image%s' % index
            filecontent = images[index][-1].read()
            zf.writestr(filename, filecontent)
        else:
            content = template.read(filename)
            zf.writestr(filename, content)
        
    template.close()
    zf.close()
    zipcontent = output.getvalue()
    # sys.stdout.write(zipcontent)

def main():
    htmlfilename = sys.argv[1]
    htmlfile = urllib.urlopen(htmlfilename)
    xslfile = open(os.path.join(dirname, 'xsl/html2wordml.xsl'))

    transform(htmlfile, xslfile)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print "Usage: html2wordml.py /path/to/htmlfile"
        sys.exit(1)

    main()
