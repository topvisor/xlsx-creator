<?php

namespace Topvisor\XlsxCreator\Xml\Book;

use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\ListXml;
use XMLWriter;

class WorkbookXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('workbook');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$xml->writeAttribute('mc:Ignorable', 'x15');
		$xml->writeAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

		$xml->startElement('fileVersion');
		$xml->writeAttribute('appName', 'xl');
		$xml->writeAttribute('lastEdited', 5);
		$xml->writeAttribute('lowestEdited', 5);
		$xml->writeAttribute('rupBuild', 9303);
		$xml->endElement();

		$xml->startElement('workbookPr');
		$xml->writeAttribute('defaultThemeVersion', 164011);
		$xml->writeAttribute('filterPrivacy', 1);
		$xml->endElement();

//		if ($model['view'] ?? false) (new ListXml('bookViews', new WorkbookView()))->render($model['view']);

		(new ListXml('sheets', new SheetXml()))->render($xml, $model);

		$xml->startElement('calcPr');
		$xml->writeAttribute('calcId', 171027);
		$xml->endElement();

		$xml->endElement();
		$xml->endDocument();
	}
}
