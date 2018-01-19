<?php

namespace Topvisor\XlsxCreator\Xml\Strings;

use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\Styles\Font\FontXml;
use XMLWriter;

class RichTextXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$fontXml = new FontXml('rPr', 'rFont');
		$textXml = new TextXml();

		foreach ($model['value'] as $value) {
			$xml->startElement('r');

			if ($value['font'] ?? false) $fontXml->render($xml, $value['font']);
			$textXml->render($xml, ['value' => $this->prepareText($value['text'])]);

			$xml->endElement();
		}
	}
}