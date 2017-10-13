<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Values\RichTextValue;
use Topvisor\XlsxCreator\Structures\Values\SharedStringValue;
use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Workbook;
use Topvisor\XlsxCreator\Xml\Strings\SharedStringXml;
use XMLWriter;

/**
 * Class SharedStrings. Служит для добавления общих строк.
 *
 * @package Topvisor\XlsxCreator\Helpers
 */
class SharedStrings{
	private $committed;
	private $nextId;
	private $sharedStrings;

	private $filename;
	private $xml;

	/**
	 * SharedStrings constructor.
	 *
	 * @param Workbook $workbook
	 */
	function __construct(Workbook $workbook){
		$this->committed = false;
		$this->nextId = 0;
		if ($workbook->getUseSharedStrings()) $this->sharedStrings = [];

		$this->filename = $workbook->genTempFilename();

		$this->xml = new XMLWriter();
		$this->xml->openURI($this->filename);

		$this->startSharedStrings();
	}

	public function __destruct(){
		unset($this->xml);

		if ($this->filename && file_exists($this->filename)) unlink($this->filename);
	}

	/**
	 * @return null|string - путь к временному файлу общих строк
	 */
	function getFilename(){
		return $this->filename;
	}

	/**
	 * @param string|RichTextValue $value - значение
	 * @return SharedStringValue - общая строка
	 * @throws InvalidValueException
	 */
	function add($value) : SharedStringValue{
		$this->checkCommitted();
		if (!$value) throw new InvalidValueException('$value must be');
		if (!($value instanceof Value)) $value = Value::parse($value);

		switch ($value->getType()) {
			case Value::TYPE_STRING: break;
			case Value::TYPE_RICH_TEXT: break;
			default: throw new InvalidValueException('$value must be string or rich text');
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

	/**
	 * @param Value $value - добавить значение в файл общих строк
	 * @return int - id значения
	 */
	private function addToXml(Value $value) : int{
		(new SharedStringXml())->render($this->xml, ['type' => $value->getType(), 'value' => $value->getValue()]);
		return $this->nextId++;
	}

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommitted(){
		if ($this->committed) throw new ObjectCommittedException();
	}

	/**
	 * Начать файл общих строк.
	 */
	private function startSharedStrings(){
		$this->xml->startDocument('1.0', 'UTF-8', 'yes');
		$this->xml->startElement('sst');
		$this->xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
	}

	/**
	 * Завершить файл общих строк.
	 */
	private function endSharedStrings(){
		$this->xml->endElement();
		$this->xml->endDocument();

		$this->xml->flush();
		unset($this->xml);
	}
}