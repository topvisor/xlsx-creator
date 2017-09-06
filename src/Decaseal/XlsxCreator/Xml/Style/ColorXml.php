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
		if (!$model) return;

		$xml->startElement($this->tag);

		switch (true) {
			case is_string($model):
				$xml->writeAttribute('rgb', $model);
				break;

			case isset($model['theme']):
				$xml->writeAttribute('theme', $model['theme']);
				if (isset($model['tint'])) $xml->writeAttribute('tint', $model['tint']);
				break;

			case isset($model['indexed']):
				$xml->writeAttribute('indexed', $model['indexed']);
				break;

			default:
				$xml->writeAttribute('auto', 1);
				break;
		}

		$xml->endElement();
	}
}