<?php

namespace Topvisor\XlsxCreator;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Helpers\Validator;
use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\Styles\Alignment\Alignment;
use Topvisor\XlsxCreator\Structures\Styles\Borders\Borders;
use Topvisor\XlsxCreator\Structures\Styles\Font;
use Topvisor\XlsxCreator\Structures\Styles\Style;
use Topvisor\XlsxCreator\Structures\Values\Value;

/**
 * Class Row. Содержит методы для работы со строкой.
 *
 * @package  Topvisor\XlsxCreator
 */
class Row extends Style{
	private $worksheet;
	private $number;
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

		$this->height = null;
		$this->hidden = false;
		$this->outlineLevel = 0;

		$this->cells = [];
	}

	function __destruct(){
		parent::__destruct();

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
	 * @param string|null $numFmt - формат чисел ячейки
	 * @return Style - $this
	 */
	function setNumFmt(string $numFmt = null) : Style{
		parent::setNumFmt($numFmt);
		foreach ($this->cells as $cell) $cell->setNumFmt($numFmt);

		return $this;
	}

	/**
	 * @param Font|null $font - шрифт
	 * @return Style - $this
	 */
	function setFont(Font $font = null) : Style{
		parent::setFont($font);
		foreach ($this->cells as $cell) $cell->setFont($font);

		return $this;
	}

	/**
	 * @param Color|null $color - заливка ячейки
	 * @return Style - $this
	 */
	function setFill(Color $color = null) : Style{
		parent::setFill($color);
		foreach ($this->cells as $cell) $cell->setFill($color);

		return $this;
	}

	/**
	 * @param Borders|null $borders - границы ячейки
	 * @return Style - $this
	 */
	function setBorders(Borders $borders = null) : Style{
		parent::setBorders($borders);
		foreach ($this->cells as $cell) $cell->setBorders($borders);

		return $this;
	}

	/**
	 * @param Alignment|null $alignment - выравнивание текста
	 * @return Style - $this
	 */
	function setAlignment(Alignment $alignment = null) : Style{
		parent::setAlignment($alignment);
		foreach ($this->cells as $cell) $cell->setAlignment($alignment);

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
	function setHeight(int $height = null) : self{
		Validator::validateInRange($height, 0, 409, '$height');

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
	function setHidden(bool $hidden) : self{
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
	function setOutlineLevel(int $outlineLevel) : self{
		Validator::validateInRange($outlineLevel, 0, 409, '$outlineLevel');

		$this->outlineLevel = $outlineLevel;
		return $this;
	}

	function getCell(int $col) : Cell{
		Validator::validateInRange($col, 1, 16384, '$col');

		if (count($this->cells) < $col)
			for ($i = count($this->cells); $i < $col; $i++)
				$this->cells[$i] = new Cell($this, $i + 1);

		return $this->cells[$col - 1];
	}

	/**
	 * Заменяет все ячеки строки на $values
	 *
	 * @param array|null $values - значения ячеек строки
	 * @return Row - $this
	 * @throws InvalidValueException
	 */
	function setCells(array $values = null) : self{
		$this->cells = [];

		if ($values)
			foreach ($values as $index => $value)
				$this->setCell($value, $index + 1);

		return $this;
	}

	/**
	 * Добавить ячейку в конец строки
	 *
	 * @param $value - значение ячейки
	 * @return Cell - ячейка
	 * @throws InvalidValueException
	 */
	function addCell($value) : Cell{
		return $this->setCell($value, count($this->cells) + 1);
	}

	/**
	 * @return bool - есть ли в строке ячейки
	 */
	function hasValues(){
		foreach ($this->cells as $cell) if ($cell->getType() !== Value::TYPE_NULL) return true;
		return false;
	}

	/**
	 *	Зафиксировать строку (и предыдущие).
	 */
	function commit(){
		$this->worksheet->commitRows($this);
	}

	/**
	 * @return array|null - модель строки
	 */
	function getModel(){
		$cellsModels = [];
		$min = 0;
		$max = 0;

		foreach ($this->cells as $cell) {
			$cellsModels[] = $cell->getModel();

			$cellCol = $cell->getCol();
			if (!$min || $min > $cellCol) $min = $cellCol;
			if ($max < $cellCol) $max = $cellCol;
		}

		return ($this->isDefault()) ? null : [
			'cells' => $cellsModels,
			'number' => $this->number,
			'min' => $min,
			'max' => $max,
			'style' => [],
			'styleId' => 0,
			'height' => $this->height,
			'hidden' => $this->hidden,
			'outlineLevel' => $this->outlineLevel,
			'collapsed' => (bool) ($this->outlineLevel && $this->outlineLevel >= $this->worksheet->getOutlineLevelRow())
		];
	}

	/**
	 * Устанавливает значение ячейки. Функция НЕ ПРОВЕРЯЕТ (!!!) наличие предыдущих ячеек
	 *
	 * @param $value - значение ячейки
	 * @param int $col - колонка ячейки
	 * @return Cell - ячейка
	 */
	private function setCell($value, int $col) : Cell{
		$cell = new Cell($this, $col);

		$cell->setValue($value);
		$cell->setNumFmt($this->numFmt);
		$cell->setFont($this->font);
		$cell->setFill($this->fill);
		$cell->setBorders($this->borders);
		$cell->setAlignment($this->alignment);

		$this->cells[$col - 1] = $cell;

		return $cell;
	}

	/**
	 * @return bool - является ли строка не измененной
	 */
	private function isDefault() : bool{
		if ($this->cells) return false;
		if ($this->height) return false;
		if ($this->hidden) return false;
		if ($this->outlineLevel) return false;
		if ($this->getStyleModel()) return false;

		return true;
	}
}