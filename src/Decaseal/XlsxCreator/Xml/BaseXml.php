<?php

namespace Decaseal\XlsxCreator\Xml;

use XMLWriter;

abstract class BaseXml{
	abstract function render(XMLWriter $xml, array $model = null);

	function toXml($model = null) {
		$xml = new XMLWriter();
		$xml->openMemory();

		$this->render($xml, $model);

		return $xml->outputMemory();
	}
}