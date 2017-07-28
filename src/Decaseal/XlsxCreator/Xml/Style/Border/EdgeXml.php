<?php

namespace Decaseal\XlsxCreator\Xml\Style\Border;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\ColorXml;
use XMLWriter;

class EdgeXml extends BaseXml{
	private $tag;
	private $defaultColor;

	function __construct(string $tag, array $defaultColor = null){
		$this->tag = $tag;
		$this->defaultColor = $defaultColor;
	}

	function render(XMLWriter $xml, $model = null){
		$xml->startElement($this->tag);

		if ($model && isset($model['style'])) {
			$xml->writeAttribute('style', $model['style']);
			if ($color = $model['color'] ?? $this->defaultColor) (new ColorXml())->render($xml, $color);
		}

		$xml->endElement();
	}
}