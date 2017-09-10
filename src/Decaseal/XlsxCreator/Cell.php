<?php

namespace Decaseal\XlsxCreator;

use DateTime;

class Cell{
	const TYPE_NULL = 0;
	const TYPE_MERGE = 1;
	const TYPE_NUMBER = 2;
	const TYPE_STRING = 3;
	const TYPE_DATE = 4;
	const TYPE_HYPERLINK = 5;
	const TYPE_FORMULA = 6;
	const TYPE_RICH_TEXT = 7;
	const TYPE_BOOL = 8;
	const TYPE_ERROR = 9;
	const TYPE_JSON = 10;

	private $row;
	private $col;

	private $value;
	private $type;
	private $master;

	function __construct(Row $row, int $col){
		$this->row = $row;
		$this->col = $col;

		$this->value = null;
		$this->type = Cell::TYPE_NULL;
		$this->master = null;
	}

	function setValue($value){
		if ($this->type === Cell::TYPE_MERGE && !is_null($this->master)) {
			$this->master->setValue($value);
		} else {
			$this->value = $value;
			$this->type = Cell::getValueType($value);
		}
	}

	function getType() : int{
		return $this->type;
	}

	private static function getValueType($value) : int{
		switch (true) {
			case is_null($value): return Cell::TYPE_NULL;
			case is_string($value): return Cell::TYPE_STRING;
			case is_numeric($value): return Cell::TYPE_NUMBER;
			case is_bool($value): return Cell::TYPE_BOOL;
			case ($value instanceof DateTime): return Cell::TYPE_DATE;
			case is_array($value):
				switch (true) {
					case ($value['text'] ?? false && $value['hyperlink'] ?? false): return Cell::TYPE_HYPERLINK;
					case ($value['formula'] ?? false): return Cell::TYPE_FORMULA;
					case ($value['richText'] ?? false): return Cell::TYPE_RICH_TEXT;
					case ($value['error'] ?? false): return Cell::TYPE_ERROR;
				}
		}

		return Cell::TYPE_JSON;
	}
}