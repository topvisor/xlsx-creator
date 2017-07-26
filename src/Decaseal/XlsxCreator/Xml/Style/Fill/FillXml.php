<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class FillXml extends BaseXml {
	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;
		if($model['type'] != 'pattern' && $model['type'] != 'gradient') return;

		$xml->startElement('fill');

		switch ($model['type']) {
			case 'pattern': (new PatternFillXml())->render($xml, $model); break;
			case 'gradient': (new PatternFillXml())->render($xml, $model); break;
		}

		$xml->endElement();
	}
}