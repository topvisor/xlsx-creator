<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Simple\BoolXml;
use Decaseal\XlsxCreator\Xml\Simple\StringXml;
use XMLWriter;

class FontXml extends BaseXml{
	private $tag;
	private $fontNameTag;

	private $subNodes;

	public function __construct(string $tag = 'font', string $fontNameTag = 'name'){
		$this->tag = $tag;
		$this->fontNameTag = $fontNameTag;

		$this->subNodes = [
			'bold' => new BoolXml('b'),
			'italic' => new BoolXml('i'),
			'underline' => new UnderlineXml(),
			'charset' => new StringXml('charset', [], 'val'),
			'color' => new ColorXml(),
			'condense' => new BoolXml('condense'),
			'extend' => new BoolXml('extend'),
			'family' => new StringXml('family', [], 'val'),
			'outline' => new BoolXml('outline'),
			'scheme' => new StringXml('scheme', [], 'val'),
			'shadow' => new BoolXml('shadow'),
			'strike' => new BoolXml('strike'),
			'size' => new StringXml('sz', [], 'val'),
			'name' => new StringXml($this->fontNameTag, [], 'val')
		];
	}

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement($this->tag);

		foreach ($this->subNodes as $prop => $node) if (isset($model[$prop])) $node->render($xml, $model[$prop]);

		$xml->endElement();
	}
}