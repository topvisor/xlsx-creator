<?php

namespace XlsxCreator;

/**
 * Class Row. Содержит методы для работы со строкой.
 *
 * @package XlsxCreator
 */
class Row{
	private $worksheet;
	private $number;
	private $style;
	private $height;
	private $hidden;
	private $outlineLevel;

	private $cells;

	/**
	 * Row constructor.
	 *
	 * @param Worksheet $worksheet - таблица
	 * @param int $number - номер строки
	 */
	function __construct(Worksheet $worksheet, int $number){
		$this->worksheet = $worksheet;
		$this->number = $number;

		$this->style = [];
		$this->height = null;
		$this->hidden = false;
		$this->outlineLevel = 0;

		$this->cells = [];
	}

	function __destruct(){
		unset($this->worksheet);
	}

	/**
	 * @return Worksheet - таблица
	 */
	function getWorksheet() : Worksheet{
		return $this->worksheet;
	}

	/**
	 * @return int - номер строки
	 */
	function getNumber() : int{
		return $this->number;
	}

	/**
	 * @see Row::setStyle() Параметры $style
	 *
	 * @return array - стили
	 */
	function getStyle() : array{
		return $this->style;
	}

	/**
	 * @param array $style - стили
	 * @return Row - $this
	 */
	function setStyle(array $style) : Row{
		$this->style = $style;
		return $this;
	}

	/**
	 * @return null|int - высота строки
	 */
	function getHeight(){
		return $this->height;
	}

	/**
	 * @param int|null $height - высота строки
	 * @return Row - $this
	 */
	function setHeight(int $height = null) : Row{
		$this->height = $height;
		return $this;
	}

	/**
	 * @return bool - скрытая ли строка
	 */
	function isHidden() : bool{
		return $this->hidden;
	}

	/**
	 * @param bool $hidden - скрыть строку
	 * @return Row - $this
	 */
	function setHidden(bool $hidden) : Row{
		$this->hidden = $hidden;
		return $this;
	}

	/**
	 * @return int - row outline level
	 */
	function getOutlineLevel() : int{
		return $this->outlineLevel;
	}

	/**
	 * @param int $outlineLevel - row outline level
	 * @return Row - $this
	 */
	function setOutlineLevel(int $outlineLevel) : Row{
		$this->outlineLevel = $outlineLevel;
		return $this;
	}

	/**
	 * @param array|null $values - значения ячеек строки
	 * @return Row - $this
	 */
	function setValues(array $values = null) : Row{
		$this->cells = [];

		if ($values) foreach ($values as $index => $value) {
			$cell = new Cell($this, $index + 1);
			$this->cells[] = $cell;
			if ($value) $cell->setValue($value);
		}

		return $this;
	}

	/**
	 * @return bool - есть ли в строке ячейки
	 */
	function hasValues(){
		foreach ($this->cells as $cell) if ($cell->getType() !== Cell::TYPE_NULL) return true;
		return false;
	}

	/**
	 *	Зафиксировать строку (и предыдущие).
	 */
	function commit(){
		$this->worksheet->commitRow($this);
	}

	/**
	 * @return array|null - модель строки
	 */
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

		return ($cellsModels || $this->height) ? [
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