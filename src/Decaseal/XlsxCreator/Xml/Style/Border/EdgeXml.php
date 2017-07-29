<?php

namespace Decaseal\XlsxCreator\Xml\Style\Border;

use Decaseal\XlsxCreator\XlsxCreator;
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

		if ($model && isset($model[XlsxCreator::BORDER_STYLE])) {
			$xml->writeAttribute(XlsxCreator::BORDER_STYLE, $model[XlsxCreator::BORDER_STYLE]);
			if ($color = $model[XlsxCreator::BORDER_COLOR] ?? $this->defaultColor) (new ColorXml())->render($xml, $color);
		}

		$xml->endElement();
	}
}