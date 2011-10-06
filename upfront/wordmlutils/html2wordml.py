#!/usr/bin/env python

import os
import sys
import urllib
import zipfile
import argparse
from cStringIO import StringIO
from lxml import etree, html
from lxml.html import soupparser
from PIL import Image

dirname = os.path.dirname(__file__)

leftmargin = 1440 # 1 inch
righmarg = 1440 # 1 inch
# 8.5 inches = 12240 twips
# default left and right margin is 1 inch each
# pagewidth is 6.5 inches = 9360 twips
pagewidth = 9360

def get_images(basepath, doc):
    images = []
    for img in doc.xpath('//img'):
        url = img.get('src')
        image = urllib.urlopen('%s/%s' % (basepath, url))
        images.append(StringIO(image.read()))
    return images

def normalize_image_urls(doc):
    for count, img in enumerate(doc.xpath('//img')):
        img.set('src', 'image%s' % count)

def convertPixelsToEMU(px):
    points = px * 72.0 / 96.0
    inches = points / 72.0
    emu = inches * 914400
    return int(emu)

def normalize_width(width):
    if width.endswith('%'):
        percentage = int(width[:-1])/100.0
        return int(pagewidth * percentage)
    elif width.endswidth('px'):
        pixels = int(width[:-2])
        inches = pixels/72.0
        return inches * 1440
    else:
        raise RuntimeErorr('Unit not known: %s' % width)

def tablewidthspec(table):
    # get the max number of columns
    maxcolumns = 0
    widthspec = {}

    # compute column widths. iterate in reverse to ensure the top most
    # width spec wins
    rows = table.xpath('thead/tr|tbody/tr|tr')
    rows.reverse()
    for tr in rows:
        colcount = len(tr.getchildren())
        if colcount > maxcolumns:
            maxcolumns = colcount
        for index, col in enumerate(tr.getchildren()):
            if col.get('width'):
                widthspec[index] = normalize_width(col.get('width'))

    # calculate widths for columns than don't have width specified
    if len(widthspec) < maxcolumns:
        divisor = maxcolumns - len(widthspec) 
        spaceleft = pagewidth
        for width in widthspec.values():
            spaceleft = spaceleft - width
        width = spaceleft / divisor
        for i in range(maxcolumns):
            if not widthspec.has_key(i):
                widthspec[i] = width

    return widthspec

def gridcolwidth(context, colindex):
    """ calculate the width of a column
    """
    table = context.context_node
    colindex = int(colindex)
    widthspec = tablewidthspec(table)

    return widthspec.get(colindex, pagewidth)

def tcwidth(context, colspan):
    """ calculate the width of a column
    """
    td = context.context_node
    table = td.xpath('ancestor::table')[0]
    colindex = td.getparent().getchildren().index(td)
    if colspan:
        colspan = int(colspan[0])
    else:
        colspan = 1
    widthspec = tablewidthspec(table)

    # make list of column widths in column order
    wlist = []
    for i in range(len(widthspec)):
        wlist.append(widthspec.get(i))

    columnwidth = 0
    for w in wlist[colindex:colindex+colspan]:
        columnwidth += w

    return columnwidth

def imgsize(context, size, dimension):
    """ convert the size
    """
    if size and size[0].endswith('px'):
        size = size[0][:-2]
        return convertPixelsToEMU(int(size))
    else:
        return '%s-$%s' % (context.context_node.get('src'), dimension)


def transform(basepath, htmlfile, image_resolver=None,
        create_package=True, outfile=sys.stdout):

    """ transform html to wordml
        
        image_resolver needs to be an instance of a class with a
        get_images method accepting `basepath` and `doc` as args. it should
        return images in the same format as the get_images method above.
    """
    # register columnwidth extension
    ns = etree.FunctionNamespace('http://upfrontsystems.co.za/wordmlutils')
    ns.prefix = 'upy'
    ns['gridcolwidth'] = gridcolwidth
    ns['tcwidth'] = tcwidth
    ns['imgsize'] = imgsize

    xslfile = open(os.path.join(dirname, 'xsl/html2wordml.xsl'))

    xslt_root = etree.XML(xslfile.read())
    transform = etree.XSLT(xslt_root)

    doc = soupparser.fromstring(htmlfile)
    if image_resolver:
        images = image_resolver.get_images(basepath, doc)
    else:
        images = get_images(basepath, doc)

    normalize_image_urls(doc)
    result_tree = transform(doc)
    wordml = etree.tostring(result_tree)

    # XXX: Don't hardcode encoding and check if doc has xml processing
    # instruction before adding it
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
    relmap = {}
    for count, img in enumerate(images):
        relid = 'image%s' % count

        # insert image sizes in the wordml
        img = Image.open(img)
        width, height = img.size
        filename = relid + '.' + img.format.lower()

        # insert image before document
        filepath = 'word/media/%s' % filename
        namelist.insert(docindex, filepath)
        relmap[filepath] = count

        # convert to EMU (English Metric Unit) 
        width = convertPixelsToEMU(width)
        height = convertPixelsToEMU(height)

        widthattr = '%s-$width' % relid
        heightattr = '%s-$height' % relid
        idattr = '%s-$id' % relid
        ridattr = '%s-$rid' % relid
        wordml = wordml.replace(widthattr, str(width))
        wordml = wordml.replace(heightattr, str(height))
        wordml = wordml.replace(ridattr, relid)
        wordml = wordml.replace(idattr, str(count))
        relxml = """<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>""" % (
            relid, filename)
        try:
            rels.append(etree.fromstring(relxml))
        except:
            raise str(relxml)

    relsxml = etree.tostring(rels)

    for filepath in namelist:
        if filepath == 'word/document.xml':
            zf.writestr(filepath, wordml)
        elif filepath.startswith('word/media'):
            index = relmap[filepath]
            filecontent = images[index]
            filecontent.seek(0)
            zf.writestr(filepath, filecontent.read())
        elif filepath.startswith('word/_rels/document.xml.rels'):
            zf.writestr(filepath, relsxml)
        else:
            content = template.read(filepath)
            zf.writestr(filepath, content)
        
    template.close()
    zf.close()
    zipcontent = output.getvalue()
    if create_package:
        outfile.write(zipcontent)
    else:
        outfile.write(wordml)

def main():
    parser = argparse.ArgumentParser(description='Convert HTML to WordML')
    parser.add_argument('-c', '--create-package', action='store_true',
        help='Create WordML package') 
    parser.add_argument('-p', '--basepath', help='Base path for relative urls',
        required=True)
    parser.add_argument('htmlfile', help='/path/to/htmlfile') 
    args = parser.parse_args()

    htmlfile = urllib.urlopen(args.htmlfile)
    basepath = args.basepath

    transform(basepath, htmlfile, create_package=args.create_package)

if __name__ == '__main__':

    main()
