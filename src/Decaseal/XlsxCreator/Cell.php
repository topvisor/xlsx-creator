<?php

namespace Decaseal\XlsxCreator;

use DateTime;
use Exception;

/**
 * Class Cell. Содержит методы для работы c ячейкой.
 *
 * @package Decaseal\XlsxCreator
 */
class Cell{
	const TYPE_NULL = 0;
	const TYPE_MERGE = 1;
	const TYPE_NUMBER = 2;
	const TYPE_STRING = 3;
	const TYPE_DATE = 4;
	const TYPE_HYPERLINK = 5;
	const TYPE_FORMULA = 6;
	const TYPE_RICH_TEXT = 8;
	const TYPE_BOOL = 9;
	const TYPE_ERROR = 10;
	const TYPE_JSON = 11;

	private $row;
	private $col;
	private $style;

	private $value;
	private $type;
	private $master;

	/**
	 * Cell constructor.
	 *
	 * @param Row $row - строка
	 * @param int $col - номер колонки
	 */
	function __construct(Row $row, int $col){
		$this->row = $row;
		$this->col = $col;
		$this->style = [];

		$this->value = null;
		$this->type = Cell::TYPE_NULL;
		$this->master = null;
	}

	function __destruct(){
		unset($this->row);
		unset($this->value);
		unset($this->master);
	}

	/**
	 * @param $value - значение ячейки
	 */
	function setValue($value){
		if ($this->type === Cell::TYPE_MERGE && !is_null($this->master)) {
			$this->master->setValue($value);
		} else {
			$this->value = $value;
			$this->type = Cell::genValueType($value);
		}
	}

	/**
	 * @return Row - строка
	 */
	function getRow() : Row{
		return $this->row;
	}

	/**
	 * @return int - колонка
	 */
	function getCol() : int{
		return $this->col;
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
	 * @see Row::setStyle() Параметры $style
	 *
	 * @param array $style - стили
	 * @return Cell - $this
	 */
	function setStyle(array $style) : Cell{
		$this->style = $style;
		return $this;
	}

	/**
	 * @return int - тип значения ячейки
	 */
	function getType() : int{
		return $this->type;
	}

	/**
	 * @return array - модель ячейки
	 */
	function genModel() : array{
		$model = [
			'address' => $this->getAddress(),
			'type' => $this->type,
			'style' => $this->style,
			'styleId' => $this->row->getWorksheet()->getWorkbook()->getStyles()->addStyle($this->style),
		];

		switch ($this->type) {
			case Cell::TYPE_NUMBER:
			case Cell::TYPE_STRING:
			case Cell::TYPE_DATE:
			case Cell::TYPE_RICH_TEXT:
			case Cell::TYPE_BOOL:
			case Cell::TYPE_ERROR:
				$model['value'] = $this->value;
				break;

			case Cell::TYPE_HYPERLINK:
				$model['text'] = $this->value['text'] ?? '';
				$model['hyperlink'] = $this->value['hyperlink'] ?? '';

				$this->row->getWorksheet()->getSheetRels()->addHyperlink($model['hyperlink'], $model['address']);
				break;

			case Cell::TYPE_MERGE:
				$model['master'] = $this->master;
				break;

			case Cell::TYPE_FORMULA:
				$model['formula'] = $this->value['formula'] ?? null;
				$model['result'] = $this->value['result'] ?? null;
				break;

			case Cell::TYPE_JSON:
				$model['type'] = Cell::TYPE_STRING;
				$model['value'] = json_encode($this->value);
				$model['rawValue'] = $this->value;
				break;
		}

		return $model;
	}

	/**
	 * @return string - адрес ячейки ('A1', 'D23')
	 */
	function getAddress() : string{
		return Cell::genAddress($this->col, $this->row->getNumber());
	}

	/**
	 * Возвращает строку колонки по ее номеру. Например, 1 - A, 3 - C.
	 *
	 * @param int $col - номер колонки
	 * @return string - строка колонки
	 * @throws Exception - ошибочный номер колонки
	 */
	static function genColStr(int $col) : string{
		if ($col < 1 || $col > 16384) throw new Exception("$col is out of bounds. Excel supports columns from 1 to 16384");
		if ($col > 26) return Cell::genColStr(($col - 1) / 26) . chr(($col % 26 ? $col % 26 : 26) + 64);
		return chr($col + 64);
	}

	/**
	 * Возвращает номер колонки по ее строке. Например, A - 1, AA - 27.
	 *
	 * @param string $col - строка колонки
	 * @return int - номер колонки
	 * @throws Exception - ошибочная строка колонки
	 */
	static function genColNum(string $col) : int{
		$len = strlen($col);
		if ($len < 1 || $len > 3) throw new Exception("Out of bounds. Invalid column $col");

		$result = 0;
		for ($i = 0; $i < $len; $i++){
			$charCode = ord(substr($col, -$i - 1, 1));
			if ($charCode < 65 || $charCode > 90) throw new Exception("Out of bounds. Invalid column $col");

			$result += ($charCode - 64) * pow(26, $i);
		}

		return $result;
	}

	/**
	 * @param int $col - номер колонки
	 * @param int $row - номер строки
	 * @return string - адрес ячейки ('A1', 'D23')
	 * @throws Exception - ошибочный номер колонки/строки
	 */
	static function genAddress(int $col, int $row) : string{
		if ($row < 1 || $col > 1048576) throw new Exception("$row is out of bounds. Excel supports rows from 1 to 1048576");
		return self::genColStr($col) . $row;
	}

	/**
	 * @param $value - значение ячейки
	 * @return int - тип значения ячейки
	 */
	private static function genValueType($value) : int{
		switch (true) {
			case is_null($value): return Cell::TYPE_NULL;
			case is_string($value): return Cell::TYPE_STRING;
			case is_numeric($value): return Cell::TYPE_NUMBER;
			case is_bool($value): return Cell::TYPE_BOOL;
			case ($value instanceof DateTime): return Cell::TYPE_DATE;
			case is_array($value):
				switch (true) {
					case (($value['text'] ?? false) && ($value['hyperlink'] ?? false)): return Cell::TYPE_HYPERLINK;
					case ($value['formula'] ?? false): return Cell::TYPE_FORMULA;
					case ($value['richText'] ?? false): return Cell::TYPE_RICH_TEXT;
					case ($value['error'] ?? false): return Cell::TYPE_ERROR;
				}
		}

		return Cell::TYPE_JSON;
	}
}