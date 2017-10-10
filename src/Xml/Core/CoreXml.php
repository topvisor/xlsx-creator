<?php

namespace Topvisor\XlsxCreator\Xml\Core;

use DateTimeZone;
use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\Simple\DateXml;
use Topvisor\XlsxCreator\Xml\Simple\StringXml;
use XMLWriter;

class CoreXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (is_null($model)) return;

		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('cp:coreProperties');

		$xml->writeAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
		$xml->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
		$xml->writeAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
		$xml->writeAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
		$xml->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

		(new StringXml('dc:creator'))->render($xml, [$model['creator']]);
		(new StringXml('dc:title'))->render($xml, null);
		(new StringXml('dc:subject'))->render($xml, null);
		(new StringXml('dc:description'))->render($xml, null);
		(new StringXml('dc:identifier'))->render($xml, null);
		(new StringXml('dc:language'))->render($xml, null);
		(new StringXml('cp:keywords'))->render($xml, null);
		(new StringXml('cp:category'))->render($xml, null);
		(new StringXml('cp:lastModifiedBy'))->render($xml, [$model['lastModifiedBy']]);
		(new StringXml('cp:lastPrinted'))->render($xml, null);
		(new StringXml('cp:revision'))->render($xml, null);
		(new StringXml('cp:contentStatus'))->render($xml, null);
		(new StringXml('dcterms:created', ['xsi:type' => 'dcterms:W3CDTF']))
			->render($xml, [$model['created']->setTimezone(new DateTimeZone('UTC'))->format('Y-m-d\TH:i:s\Z')]);
		(new StringXml('dcterms:modified', ['xsi:type' => 'dcterms:W3CDTF']))
			->render($xml, [$model['modified']->setTimezone(new DateTimeZone('UTC'))->format('Y-m-d\TH:i:s\Z')]);

		$xml->endElement();
		$xml->endDocument();
	}
}