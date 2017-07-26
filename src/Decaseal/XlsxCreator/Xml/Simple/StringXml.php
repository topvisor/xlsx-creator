<?php

namespace Decaseal\XlsxCreator\Xml\Simple;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class StringXml extends BaseXml{
	private $tag;
	private $attributes;
	private $attribute;

	public function __construct(string $tag, array $attributes = [], string $attribute = ''){
		$this->tag = $tag;
		$this->attributes = $attributes;
		$this->attribute = $attribute;
	}

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement($this->tag);

		if ($this->attributes) foreach ($this->attributes as $name => $value) $xml->writeAttribute($name, $value);

		if ($this->attribute) $xml->writeAttribute($this->attribute, $model);
		else $xml->text($model);

		$xml->endElement();
	}
}