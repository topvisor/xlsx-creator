<?php

namespace Decaseal\XlsxCreator;


use XMLWriter;

class SheetRels{
	private const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
	private const TARGET_MODE = 'External';

	private $worksheet;

	private $hyperlinks;
	private $committed;

	private $filename;
	private $xml;

	function __construct(Worksheet $worksheet){
		$this->worksheet = $worksheet;

		$this->hyperlinks = [];
		$this->committed = false;

		$this->filename = $this->worksheet->getWorkbook()->genTempFilename();
		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->startSheetRels();
	}

	public function __destruct(){
		if (file_exists($this->filename)) unlink($this->filename);
	}

	function getHyperlinks() : array{
		return $this->hyperlinks;
	}

	function getFilename() : string{
		return $this->filename;
	}

	function getLocalname() : string{
		return '/xl/worksheets/_rels/sheet' . $this->worksheet->getId() . '.xml.rels';
	}

	function addHyperlink($target, $address){
		$this->hyperlinks = [
			'address' => $address,
			'rId' => $this->writeRelationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $target, 'External')
		];
	}

	function commit() {
		if ($this->committed) return;
		$this->committed = true;

		$this->endSheetRels();

		$this->xml->flush();
		unset($this->xml);
	}

	private function startSheetRels(){
		$this->xml->startDocument('1.0', 'UTF-8', 'yes');
		$this->xml->startElement('Relationships');
		$this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
	}

	private function endSheetRels(){
		$this->xml->endElement();
		$this->xml->endDocument();
	}

	private function writeRelationship(string $type, string $target, string $targetMode = null) : string{
		$rId = 'rId' . (count($this->hyperlinks) + 1);

		$this->xml->startElement('Relationship');

		$this->xml->writeAttribute('Id', $rId);
		$this->xml->writeAttribute('Type', $type);
		$this->xml->writeAttribute('Target', $target);
		if ($targetMode) $this->xml->writeAttribute('TargetMode', $targetMode);

		$this->xml->endElement();

		return $rId;
	}
}