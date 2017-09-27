<?php

namespace XlsxCreator\Xml\Core;

use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class ContentTypesXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (is_null($model)) return;

		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('Types');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

		$xml->startElement('Default');
		$xml->writeAttribute('Extension', 'rels');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-package.relationships+xml');
		$xml->endElement();

		$xml->startElement('Default');
		$xml->writeAttribute('Extension', 'xml');
		$xml->writeAttribute('ContentType', 'application/xml');
		$xml->endElement();

		$xml->startElement('Override');
		$xml->writeAttribute('PartName', '/xl/workbook.xml');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
		$xml->endElement();

		foreach ($model as $worksheet) {
			$xml->startElement('Override');
			$xml->writeAttribute('PartName', '/' . $worksheet['partName']);
			$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
			$xml->endElement();
		}

		$xml->startElement('Override');
		$xml->writeAttribute('PartName', '/xl/theme/theme1.xml');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml');
		$xml->endElement();

		$xml->startElement('Override');
		$xml->writeAttribute('PartName', '/xl/styles.xml');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
		$xml->endElement();

		$xml->startElement('Override');
		$xml->writeAttribute('PartName', '/xl/sharedStrings.xml');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
		$xml->endElement();

		$xml->startElement('Override');
		$xml->writeAttribute('PartName', '/docProps/core.xml');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml');
		$xml->endElement();

		$xml->startElement('Override');
		$xml->writeAttribute('PartName', '/docProps/app.xml');
		$xml->writeAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
		$xml->endElement();

		$xml->endElement();
		$xml->endDocument();
	}
}