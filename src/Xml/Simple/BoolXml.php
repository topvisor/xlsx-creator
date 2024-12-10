<?php

namespace Topvisor\XlsxCreator\Xml\Simple;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BoolXml extends BaseXml {
	private $tag;

	public function __construct(string $tag = '') {
		$this->tag = $tag;
	}

	public function render(XMLWriter $xml, ?array $model = null) {
		if ($model && ($model[0] ?? false)) $xml->writeElement($this->tag);
	}
}
