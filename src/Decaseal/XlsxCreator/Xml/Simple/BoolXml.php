<?php

namespace Decaseal\XlsxCreator\Xml\Simple;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BoolXml extends BaseXml{
	private $tag;

	function __construct(string $tag = ''){
		$this->tag = $tag;
	}

	function render(XMLWriter $xml, array $model = null){
		if ($model && ($model[0] ?? false)) $xml->writeElement($this->tag);
	}
}