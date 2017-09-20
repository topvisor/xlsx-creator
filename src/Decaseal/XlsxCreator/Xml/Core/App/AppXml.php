<?php

namespace Decaseal\XlsxCreator\Xml\Core\App;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Simple\StringXml;
use XMLWriter;

class AppXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (is_null($model)) return;

		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('Properties');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
		$xml->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

		$xml->writeElement('Application', 'Microsoft Excel');
		$xml->writeElement('DocSecurity', '0');
		$xml->writeElement('ScaleCrop', 'false');

		(new AppHeadingPairsXml())->render($xml, $model['worksheets']);
		(new AppTitlesOfPartsXml())->render($xml, $model['worksheets']);
		(new StringXml('Company'))->render($xml, [$model['company']]);
		(new StringXml('Manager'))->render($xml, [$model['manager']]);

		$xml->writeElement('LinksUpToDate', 'false');
		$xml->writeElement('SharedDoc', 'false');
		$xml->writeElement('HyperlinksChanged', 'false');
		$xml->writeElement('AppVersion', '16.0300');

		$xml->endElement();
		$xml->endDocument();
	}
}