<?php

namespace Decaseal\XlsxCreator\Xml\Simple;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BoolXml extends BaseXml{
	private $tag;

	function __construct(string $tag = ''){
		$this->tag = $tag;
	}

	function setTag(string $tag){
		$this->tag = $tag;
	}

	function render(XMLWriter $xml, $model = null){
		if ($model) $xml->writeElement($this->tag);
	}
}