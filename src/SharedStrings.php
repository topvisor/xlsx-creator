<?php

namespace Topvisor\XlsxCreator;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Values\SharedStringValue;
use Topvisor\XlsxCreator\Structures\Values\Value;
use TypeError;

class SharedStrings{
	private $workbook;

	private $committed;
	private $nextId;
	private $sharedStrings;

	private $filename;
	private $xml;

	function __construct(Workbook $workbook, bool $checkDuplicate){
		$this->workbook = $workbook;

		$this->nextId = 0;
		if ($checkDuplicate) $this->sharedStrings = [];
	}

	public function __destruct(){
		unset($this->workbook);
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
		$id = $this->nextId++;

		return $id;
	}

	private function checkCommitted(){
		if (!$this->xml || $this->committed) throw new ObjectCommittedException();
	}

	private function startSharedStrings(){

	}

	private function endSharedStrings(){

	}
}