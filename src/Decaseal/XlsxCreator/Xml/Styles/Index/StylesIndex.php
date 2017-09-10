<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Index;

use Decaseal\XlsxCreator\Xml\BaseXml;

class StylesIndex{
	protected $baseXml;

	protected $indexes;
	protected $xmls;

	function __construct(BaseXml $baseXml){
		$this->baseXml = $baseXml;

		$this->indexes = [];
		$this->xmls = [];
	}

	function addIndex($model){
		$xml = $this->baseXml->toXml($model);

		if (isset($this->indexes[$xml])) return $this->indexes[$xml];

		$index = count($this->xmls);

		$this->indexes[$xml] = $index;
		$this->xmls[] = $xml;

		return $index;
	}

	function getXmls() : array{
		return $this->xmls;
	}
}