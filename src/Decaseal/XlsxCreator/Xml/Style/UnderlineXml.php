<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class UnderlineXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement('u');

		switch ($model) {
			case 'double': $attributes = ['val' => 'double']; break;
			case 'singleAccounting': $attributes = ['val' => 'singleAccounting']; break;
			case 'doubleAccounting': $attributes = ['val' => 'doubleAccounting']; break;
			default: $attributes = []; break;
		}
		foreach ($attributes as $name => $value) $xml->writeAttribute($name, $value);

		$xml->endElement();
	}
}