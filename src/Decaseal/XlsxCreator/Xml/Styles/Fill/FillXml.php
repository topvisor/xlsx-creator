<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Fill;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class FillXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('fill');

		switch ($model['type']) {
			case 'pattern': (new PatternFillXml())->render($xml, $model); break;
			case 'gradient': (new GradientFillXml())->render($xml, $model); break;
		}

		$xml->endElement();
	}
}