#!/usr/bin/env python

import os
import sys
import urllib
import zipfile
import argparse
from cStringIO import StringIO
from lxml import etree, html
from PIL import Image

dirname = os.path.dirname(__file__)

def get_images(doc):
    images = {}
    for img in doc.xpath('//img'):
        url = img.get('src')
        image = urllib.urlopen(url)
        filename = url.split('/')[-1]
        images[filename] = (url, StringIO(image.read()))
    return images
    
def convertPixelsToEMU(px):
    points = px * 72.0 / 96.0
    inches = points / 72.0
    emu = inches * 914400
    return int(emu)

def transform(htmlfile, xslfile, create_package=True):
    xslt_root = etree.XML(xslfile.read())
    transform = etree.XSLT(xslt_root)
    doc = html.parse(htmlfile)
    images = get_images(doc)
    result_tree = transform(doc)
    wordml = etree.tostring(result_tree)
    wordml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>' + \
        wordml

    template = zipfile.ZipFile(os.path.join(dirname, 'template.docx'))

    # read and parse relations from template so that we can update it
    # with links to images
    rels = template.read('word/_rels/document.xml.rels')
    rels = etree.parse(StringIO(rels)).getroot()

    output = StringIO()
    zf = zipfile.ZipFile(output, 'w')
    namelist = template.namelist()
    docindex = namelist.index('word/document.xml')
    for filename, img in images.items():
        url, data = img
        # insert image before document
        namelist.insert(docindex, 'word/media/%s' % filename)

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
        wordml = wordml.replace(ridattr, filename)
        relxml = """<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>""" % (
            filename, filename)
        rels.append(etree.fromstring(relxml))

    relsxml = etree.tostring(rels)

    for filepath in namelist:
        if filepath == 'word/document.xml':
            zf.writestr(filepath, wordml)
        elif filepath.startswith('word/media'):
            filename = filepath.split('/')[-1]
            filecontent = images[filename][-1].read()
            zf.writestr(filepath, filecontent)
        elif filepath.startswith('word/_rels/document.xml.rels'):
            zf.writestr(filepath, relsxml)
        else:
            content = template.read(filepath)
            zf.writestr(filepath, content)
        
    template.close()
    zf.close()
    zipcontent = output.getvalue()
    if create_package:
        sys.stdout.write(zipcontent)
    else:
        sys.stdout.write(wordml)

def main():
    parser = argparse.ArgumentParser(description='Convert HTML to WordML')
    parser.add_argument('-p', '--create-package', action='store_true',
        help='Create WordML package') 
    parser.add_argument('htmlfile', help='/path/to/htmlfile') 
    args = parser.parse_args()

    htmlfile = urllib.urlopen(args.htmlfile)

    xslfile = open(os.path.join(dirname, 'xsl/html2wordml.xsl'))

    transform(htmlfile, xslfile, create_package=args.create_package)

if __name__ == '__main__':

    main()
