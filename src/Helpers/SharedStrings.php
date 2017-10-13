<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Exceptions\EmptyObjectException;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Values\SharedStringValue;
use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Workbook;
use Topvisor\XlsxCreator\Xml\Strings\SharedStringXml;
use TypeError;
use XMLWriter;

class SharedStrings{
	private $workbook;

	private $committed;
	private $nextId;
	private $sharedStrings;

	private $filename;
	private $xml;

	function __construct(Workbook $workbook){
		$this->workbook = $workbook;

		$this->committed = false;
		$this->nextId = 0;
		if ($workbook->getUseSharedStrings()) $this->sharedStrings = [];
	}

	public function __destruct(){
		unset($this->workbook);
		unset($this->xml);

		if ($this->filename && file_exists($this->filename)) unlink($this->filename);
	}

	/**
	 * @return null|string - путь к временному файлу связей
	 */
	function getFilename(){
		return $this->filename;
	}

	function add($value) : SharedStringValue{
		$this->checkCommitted();
		if (!$value) throw new InvalidValueException('$value must be');
		if (!($value instanceof Value)) $value = Value::parse($value);

		if (!$this->xml) $this->startSharedStrings();

		switch ($value->getType()) {
			case Value::TYPE_STRING: break;
			case Value::TYPE_RICH_TEXT: break;
			default: throw new TypeError();
		}

		if ($this->sharedStrings) {
			if (isset($this->sharedStrings[$value->getValue()])) {
				$id = $this->sharedStrings[$value->getValue()];
			} else {
				$id = $this->addToXml($value);
			}
		} else {
			$id = $this->addToXml($value);
		}

		return new SharedStringValue($id);
	}

	function isCommitted() : bool{
		return $this->committed;
	}

	/**
	 * Зафиксировать файл общих строк.
	 *
	 * @throws ObjectCommittedException
	 */
	function commit() {
		$this->checkCommitted();
		$this->committed = true;

		$this->endSharedStrings();
	}

	private function addToXml(Value $value) : int{
		(new SharedStringXml())->render($this->xml, ['type' => $value->getType(), 'value' => $value->getValue()]);
		return $this->nextId++;
	}

	private function checkCommitted(){
		if ($this->committed) throw new ObjectCommittedException();
	}

	private function startSharedStrings(){
		$this->filename = $this->workbook->genTempFilename();

		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->xml->startDocument('1.0', 'UTF-8', 'yes');
		$this->xml->startElement('sst');
		$this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
	}

	private function endSharedStrings(){
		if (!$this->xml) return;

		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);
	}
}