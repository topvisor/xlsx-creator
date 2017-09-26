<?php

namespace Decaseal\XlsxCreator;

class Row{
	private $worksheet;
	private $number;
	private $style;
	private $height;
	private $hidden;
	private $outlineLevel;

	private $cells;

	function __construct(Worksheet $worksheet, int $number, array $style = [], int $height = null, bool $hidden = false, int $outlineLevel = 0){
		$this->worksheet = $worksheet;
		$this->number = $number;
		$this->style = $style;
		$this->height = $height;
		$this->hidden = $hidden;
		$this->outlineLevel = $outlineLevel;

		$this->cells = [];
	}

	function __destruct(){
		unset($this->worksheet);
	}

	function getWorksheet() : Worksheet{
		return $this->worksheet;
	}

	function getNumber() : int{
		return $this->number;
	}

	function setValues(array $values = null){
		$this->cells = [];

		if ($values) foreach ($values as $index => $value) {
			$cell = new Cell($this, $index + 1);
			$this->cells[] = $cell;
			if ($value) $cell->setValue($value);
		}
	}

	function commit(){
		$this->worksheet->commitRow($this);
	}

	function hasValues(){
		foreach ($this->cells as $cell) if ($cell->getType() !== Cell::TYPE_NULL) return true;
		return false;
	}

	function genModel(){
		$cellsModels = [];
		$min = 0;
		$max = 0;

		foreach ($this->cells as $cell) {
			$cellsModels[] = $cell->genModel();

			$cellCol = $cell->getCol();
			if (!$min || $min > $cellCol) $min = $cellCol;
			if ($max < $cellCol) $max = $cellCol;
		}

		$styleId = $this->worksheet->getWorkbook()->getStyles()->addStyle($this->style);
		$collapsed = (bool) ($this->outlineLevel && $this->outlineLevel >= $this->worksheet->getOutlineLevelRow());

		return $cellsModels ? [
			'cells' => $cellsModels,
			'number' => $this->number,
			'min' => $min,
			'max' => $max,
			'height' => $this->height,
			'style' => $this->style,
			'styleId' => $styleId,
			'hidden' => $this->hidden,
			'outlineLevel' => $this->outlineLevel,
			'collapsed' => $collapsed
		] : null;
	}
}