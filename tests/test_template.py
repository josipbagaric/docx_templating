# -*- coding: utf-8 -*-

import context
from template import Report
from docx import Document

import unittest, os, sys, datetime


class ReportTestSuite(unittest.TestCase):
	"""Advanced test cases."""

	def test_front_page(self):
		""" Checks if the front page of the report is generated properly. """

		report = Report(title='Title', subtitle="Sub title", version="0.1")
		report.save('tests/test.docx')

		document = Document('tests/test.docx')

		title, subtitle, version, date = False, False, False, False

		for paragraph in document.paragraphs:
			if 'Title' in paragraph.text:
				title = True
			elif 'Sub title' in paragraph.text:
				subtitle = True
			elif '0.1' in paragraph.text:
				version = True

			try:
				datetime.datetime.strptime(paragraph.text, "%d %B %Y")
				date = True
			except ValueError:
				continue

		self.assertEqual([title, subtitle, version, date], [True, True, True, True])


if __name__ == '__main__':
	unittest.main()