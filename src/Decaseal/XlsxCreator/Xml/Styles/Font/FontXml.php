<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Font;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Simple\BoolXml;
use Decaseal\XlsxCreator\Xml\Simple\StringXml;
use Decaseal\XlsxCreator\Xml\Styles\ColorXml;
use XMLWriter;

class FontXml extends BaseXml{
	private $tag;
	private $fontNameTag;

	function __construct(string $tag = 'font', string $fontNameTag = 'name'){
		$this->tag = $tag;
		$this->fontNameTag = $fontNameTag;
	}

	function render(XMLWriter $xml, array $model = null){
		$xml->startElement($this->tag);

		if ($model) {
			(new BoolXml('b'))->render($xml, [$model['b'] ?? false]);
			(new BoolXml('i'))->render($xml, [$model['i'] ?? false]);
			(new UnderlineXml())->render($xml, [$model['u'] ?? false]);
			(new StringXml('charset', [], 'val'))->render($xml, [$model['charset'] ?? false]);
			(new ColorXml())->render($xml, $model['color'] ?? false);
			(new BoolXml('condense'))->render($xml, [$model['condense'] ?? false]);
			(new BoolXml('extend'))->render($xml, [$model['extend'] ?? false]);
			(new StringXml('family', [], 'val'))->render($xml, [$model['family'] ?? false]);
			(new BoolXml('outline'))->render($xml, [$model['outline'] ?? false]);
			(new StringXml('scheme', [], 'val'))->render($xml, [$model['scheme'] ?? false]);
			(new BoolXml('shadow'))->render($xml, [$model['shadow'] ?? false]);
			(new BoolXml('strike'))->render($xml, [$model['strike'] ?? false]);
			(new StringXml('sz', [], 'val'))->render($xml, [$model['sz'] ?? false]);
			(new StringXml($this->fontNameTag, [], 'val'))->render($xml, [$model['name'] ?? false]);
		}

		$xml->endElement();
	}
}