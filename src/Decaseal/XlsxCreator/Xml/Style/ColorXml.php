<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class ColorXml extends BaseXml{
	const DEFAULT_TAG = 'color';
	const DEFAULT_AUTO = 1;

	const RGB = 'rgb';
	const THEME = 'theme';
	const TINT = 'tint';
	const INDEXED = 'indexed';
	const AUTO = 'auto';

	private $tag;

	public function __construct(string $tag = ColorXml::DEFAULT_TAG){
		$this->tag = $tag;
	}

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement($this->tag);

		if (isset($model[XlsxCreator::COLOR_ARGB])) {
			$xml->writeAttribute(ColorXml::RGB, $model[XlsxCreator::COLOR_ARGB]);
		} elseif (isset($model[ColorXml::THEME])) {
			$xml->writeAttribute(ColorXml::THEME, $model[ColorXml::THEME]);
			if (isset($model[ColorXml::TINT])) $xml->writeAttribute(ColorXml::TINT, $model[ColorXml::TINT]);
		} elseif(isset($model[ColorXml::INDEXED])) {
			$xml->writeAttribute(ColorXml::INDEXED, $model[ColorXml::INDEXED]);
		} else {
			$xml->writeAttribute(ColorXml::AUTO, ColorXml::DEFAULT_AUTO);
		}

		$xml->endElement();
	}
}