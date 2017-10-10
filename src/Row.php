<?php

namespace Topvisor\XlsxCreator;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Structures\Values\Value;

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
	function setStyle(array $style) : self{
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
	function setHeight(int $height = null) : self{
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
		$this->outlineLevel = $outlineLevel;
		return $this;
	}

	function getCell(int $col) : Cell{
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
		$this->cells[$col] = $cell;

		return $cell;
	}
}