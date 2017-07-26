<?php

namespace Decaseal\XlsxCreator\Xml;

use XMLWriter;

class ListXml extends BaseXml{
	private $tag;
	private $baseXml;
	private $attributes;
	private $isEmpty;
	private $isCount;
	private $countAttributeName;

	function __construct(string $tag, BaseXml $baseXml, array $attributes = [], bool $isEmpty = false, bool $isCount = false, string $countAttributeName = 'count'){
		$this->tag = $tag;
		$this->baseXml = $baseXml;
		$this->attributes = $attributes;
		$this->isEmpty = $isEmpty;
		$this->isCount = $isCount;
		$this->countAttributeName = $countAttributeName;
	}

	function render(XMLWriter $xml, $model = null){
		if ($this->isEmpty) {
			$xml->writeElement($this->tag);
		} else {
			if(is_null($model)) return;

			$xml->startElement($this->tag);

			foreach ($this->attributes as $name => $value) $xml->writeAttribute($name, $value);
			if ($this->isCount) $xml->writeAttribute($this->countAttributeName, count($model));

			foreach ($model as $childModel) $this->baseXml->render($xml, $childModel);

			$xml->endElement();
		}
	}
}