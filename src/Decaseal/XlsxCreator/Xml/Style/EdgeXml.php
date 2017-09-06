<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class EdgeXml extends BaseXml {
	private $tag;
	private $defaultColor;

	function __construct(string $tag, string $defaultColor = null){
		$this->tag = $tag;
		$this->defaultColor = $defaultColor;
	}

	function render(XMLWriter $xml, $model = null){
		$colorXml = new ColorXml();
		$colorModel = $model || $this->defaultColor;

		$xml->startElement($this->tag);

		if ($model && $model['style'] ?? false) {
			$xml->writeAttribute('style', $model['style']);
			if ($colorModel) $colorXml->render($xml, $colorModel);
		}

		$xml->endElement();
	}
}