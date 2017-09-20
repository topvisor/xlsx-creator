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
			(new BoolXml('b'))->render($xml, [$model['b']]);
			(new BoolXml('i'))->render($xml, [$model['i']]);
			(new UnderlineXml())->render($xml, [$model['u']]);
			(new StringXml('charset', [], 'val'))->render($xml, [$model['charset']]);
			(new ColorXml())->render($xml, $model['color']);
			(new BoolXml('condense'))->render($xml, [$model['condense']]);
			(new BoolXml('extend'))->render($xml, [$model['extend']]);
			(new StringXml('family', [], 'val'))->render($xml, [$model['family']]);
			(new BoolXml('outline'))->render($xml, [$model['outline']]);
			(new StringXml('scheme', [], 'val'))->render($xml, [$model['scheme']]);
			(new BoolXml('shadow'))->render($xml, [$model['shadow']]);
			(new BoolXml('strike'))->render($xml, [$model['strike']]);
			(new StringXml('sz', [], 'val'))->render($xml, [$model['sz']]);
			(new StringXml($this->fontNameTag, [], 'val'))->render($xml, [$model['name']]);
		}

		$xml->endElement();
	}
}