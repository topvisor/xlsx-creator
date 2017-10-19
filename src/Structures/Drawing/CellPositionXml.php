<?php

namespace Topvisor\XlsxCreator\Structures\Drawing;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class CellPositionXml extends BaseXml{
	private $tag;

	function __construct(string $tag){
		$this->tag = $tag;
	}

	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement($this->tag);

		$col = $model['col'] - 1;
		$row = $model['row'] - 1;

		$xml->writeElement('xdr:col', $intCol = (int) floor($col));
		$xml->writeElement('xdr:colOff', (int) floor(($col - $intCol) * 640000));
		$xml->writeElement('xdr:row', $intRow = (int) floor(($row)));
		$xml->writeElement('xdr:rowOff', (int) floor(($row - $intRow) * 180000));

		$xml->endElement();
	}
}