<?php

namespace Decaseal\XlsxCreator;


use XMLWriter;

class SheetRels{
	private $worksheet;

	private $hyperlinks;

	private $filename;
	private $xml;

	function __construct(Worksheet $worksheet){
		$this->worksheet = $worksheet;

		$this->hyperlinks = [];

		$this->filename = $this->worksheet->getWorkbook()->genTempFilename();
		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);
	}
}