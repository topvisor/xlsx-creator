<?php

namespace XlsxCreator\Xml\Simple;

use XlsxCreator\Xml\BaseXml;
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

	function render(XMLWriter $xml, array $model = null){
		if(!$model || !isset($model[0]) || is_null($model[0]) || $model[0] === false) return;

		$xml->startElement($this->tag);

		if ($this->attributes) foreach ($this->attributes as $name => $value) $xml->writeAttribute($name, $value);

		if ($this->attribute) $xml->writeAttribute($this->attribute, $model[0]);
		else $xml->text($model[0]);

		$xml->endElement();
	}
}