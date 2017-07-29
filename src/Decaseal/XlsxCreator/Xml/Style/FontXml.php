<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Simple\BoolXml;
use Decaseal\XlsxCreator\Xml\Simple\StringXml;
use XMLWriter;

class FontXml extends BaseXml{
	const DEFAULT_TAG = 'font';

	const VAL = 'val';

	private $tag;
	private $fontNameTag;

	private $subNodes;

	public function __construct(string $tag = FontXml::DEFAULT_TAG, string $fontNameTag = XlsxCreator::FONT_NAME){
		$this->tag = $tag;
		$this->fontNameTag = $fontNameTag;

		$this->subNodes = [
			XlsxCreator::FONT_BOLD => new BoolXml(XlsxCreator::FONT_BOLD),
			XlsxCreator::FONT_ITALIC => new BoolXml(XlsxCreator::FONT_ITALIC),
			XlsxCreator::FONT_UNDERLINE => new UnderlineXml(),
			XlsxCreator::FONT_CHARSET => new StringXml(XlsxCreator::FONT_CHARSET, [], FontXml::VAL),
			XlsxCreator::FONT_COLOR => new ColorXml(),
			XlsxCreator::FONT_CONDENSE => new BoolXml(XlsxCreator::FONT_CONDENSE),
			XlsxCreator::FONT_EXTEND => new BoolXml(XlsxCreator::FONT_EXTEND),
			XlsxCreator::FONT_FAMILY => new StringXml(XlsxCreator::FONT_FAMILY, [], FontXml::VAL),
			XlsxCreator::FONT_OUTLINE => new BoolXml(XlsxCreator::FONT_OUTLINE),
			XlsxCreator::FONT_SCHEME => new StringXml(XlsxCreator::FONT_SCHEME, [], FontXml::VAL),
			XlsxCreator::FONT_SHADOW => new BoolXml(XlsxCreator::FONT_SHADOW),
			XlsxCreator::FONT_STRIKE => new BoolXml(XlsxCreator::FONT_STRIKE),
			XlsxCreator::FONT_SIZE => new StringXml(XlsxCreator::FONT_SIZE, [], FontXml::VAL),
			XlsxCreator::FONT_NAME => new StringXml($this->fontNameTag, [], FontXml::VAL)
		];
	}

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement($this->tag);

		foreach ($this->subNodes as $prop => $node) if (isset($model[$prop])) $node->render($xml, $model[$prop]);

		$xml->endElement();
	}
}