<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class ColorXml extends BaseXml{
	private $tag;

	public function __construct(string $tag = 'color'){
		$this->tag = $tag;
	}

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement($this->tag);

		if (isset($model['rgb'])) {
			$xml->writeAttribute('rgb', $model['rgb']);
		} elseif (isset($model['theme'])) {
			$xml->writeAttribute('theme', $model['theme']);
			if (isset($model['tint'])) $xml->writeAttribute('tint', $model['tint']);
		} elseif(isset($model['indexed'])) {
			$xml->writeAttribute('indexed', $model['indexed']);
		} else {
			$xml->writeAttribute('auto', 1);
		}

		$xml->endElement();
	}
}