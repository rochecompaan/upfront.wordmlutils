import os
import re
import glob
import unittest
from StringIO import StringIO

from upfront.wordmlutils.html2wordml import transform

dirname = os.path.dirname(__file__)

class TestTransform(unittest.TestCase):

    def setUp(self):
        self.cwd = os.getcwd()
        os.chdir(dirname)

    def tearDown(self):
        os.chdir(self.cwd)

    def test_transform(self):
        for filename in glob.glob('*.html'):
            html = open(filename).read()
            wordmlfilename = os.path.splitext(filename)[0] + '.xml'
            if not os.path.exists(wordmlfilename):
                continue
            wordml = open(wordmlfilename).read()
            # remove newlines and indentation
            wordml = wordml.replace('\n', '')
            wordml = re.sub('>\s+<', '><', wordml)
            basepath = '.'
            output = StringIO()
            transform(basepath, html, create_package=False, outfile=output)
            self.assertEqual(wordml, output.getvalue())

def test_suite():
    from unittest import TestSuite, makeSuite
    suite = TestSuite()
    suite.addTest(makeSuite(TestTransform))
    return suite
