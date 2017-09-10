<?php

namespace Decaseal\XlsxCreator\Xml\Book\Workbook;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\ListXml;
use XMLWriter;

class WorkbookXml extends BaseXml{
	private const FILE_VERSION_XML = '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>';
	private const WORKBOOK_PR_XML = '<workbookPr defaultThemeVersion="164011" filterPrivacy="1"/>';
	private const BOOK_VIEWS_XML = '<bookViews>
		<workbookView xWindow="0" yWindow="0" windowWidth="12000" windowHeight="24000"/>
	</bookViews>';
	private const CALC_PR_XML = '<calcPr calcId="171027"/>';

	function render(XMLWriter $xml, $model = null){
		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('workbook');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$xml->writeAttribute('mc:Ignorable', 'x15');
		$xml->writeAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

		$xml->writeRaw(WorkbookXml::FILE_VERSION_XML);
		$xml->writeRaw(WorkbookXml::WORKBOOK_PR_XML);
		$xml->writeRaw(WorkbookXml::BOOK_VIEWS_XML);

		(new ListXml('sheets', new SheetXml()))->render($xml, $model);

		$xml->writeRaw(WorkbookXml::CALC_PR_XML);

		$xml->endElement();
		$xml->endDocument();
	}
}