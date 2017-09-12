<?php

namespace Decaseal\XlsxCreator;

class Row{
	private $worksheet;
	private $number;

	private $cells;

	function __construct(Worksheet $worksheet, int $number){
		$this->worksheet = $worksheet;
		$this->number = $number;

		$this->cells = [];
	}

	function setValues(array $values = null){
		$this->cells = [];

		foreach ($values as $index => $value) {
			$cell = new Cell($this, $index);
			$this->cells[] = $cell;
			if ($value) $cell->setValue($value);
		}
	}

	function hasValues(){
		foreach ($this->cells as $cell) if ($cell->getType() !== Cell::TYPE_NULL) return true;
		return false;
	}
}