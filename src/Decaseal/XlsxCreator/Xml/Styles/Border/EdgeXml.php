<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Border;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class EdgeXml extends BaseXml {
	private $tag;
	private $defaultColor;

	function __construct(string $tag, array $defaultColor = null){
		$this->tag = $tag;
		$this->defaultColor = $defaultColor;
	}

	function render(XMLWriter $xml, array $model = null){
		$colorXml = new ColorXml();
		$colorModel = $model || $this->defaultColor;

		$xml->startElement($this->tag);

		if ($model && ($model['style'] ?? false)) {
			$xml->writeAttribute('style', $model['style']);
			if ($colorModel) $colorXml->render($xml, $colorModel);
		}

		$xml->endElement();
	}
}