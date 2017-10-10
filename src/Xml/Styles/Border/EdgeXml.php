<?php

namespace Topvisor\XlsxCreator\Xml\Styles\Border;

use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\Styles\ColorXml;
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
		$colorModel = $model['color'] || $this->defaultColor;

		$xml->startElement($this->tag);

		if ($model && ($model['style'] ?? false)) {
			$xml->writeAttribute('style', $model['style']);
			if ($colorModel) $colorXml->render($xml, $colorModel);
		}

		$xml->endElement();
	}
}