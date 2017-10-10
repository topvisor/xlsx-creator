<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class PageMargins extends BaseXml {
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('pageMargins');

		if ($model['left'] ?? false) $xml->writeAttribute('left', $model['left']);
		if ($model['right'] ?? false) $xml->writeAttribute('right', $model['right']);
		if ($model['top'] ?? false) $xml->writeAttribute('top', $model['top']);
		if ($model['bottom'] ?? false) $xml->writeAttribute('bottom', $model['bottom']);
		if ($model['header'] ?? false) $xml->writeAttribute('header', $model['header']);
		if ($model['footer'] ?? false) $xml->writeAttribute('footer', $model['footer']);

		$xml->endElement();
	}
}