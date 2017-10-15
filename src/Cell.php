<?php

namespace Topvisor\XlsxCreator;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Helpers\Comments;
use Topvisor\XlsxCreator\Helpers\SheetRels;
use Topvisor\XlsxCreator\Helpers\Validator;
use Topvisor\XlsxCreator\Structures\Styles\Style;
use Topvisor\XlsxCreator\Structures\Values\HyperlinkValue;
use Topvisor\XlsxCreator\Structures\Values\RichText\RichTextValue;
use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Xml\Styles\StylesXml;

/**
 * Class Cell. Содержит методы для работы c ячейкой.
 *
 * @package  Topvisor\XlsxCreator
 */
class Cell extends Style{
	private $row;
	private $col;
	private $style;

	private $comment;
	private $commentWidth;
	private $commentHeight;

	private $value;
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

		$this->commentWidth = 1;
		$this->commentHeight = 1;

		$this->value = Value::parse(null);
	}

	function __destruct(){
		parent::__destruct();

		unset($this->row);
		unset($this->value);
		unset($this->master);
	}

	/**
	 * @return Value - значение ячейки
	 */
	function getValue() : Value{
		return $this->value;
	}

	/**
	 * @param $value - значение ячейки
	 * @throws InvalidValueException
	 * @return Cell - $this
	 */
	function setValue($value) : self{
		if (!($value instanceof Value)) $value = Value::parse($value);

		if ($this->master) $this->master->setValue($value);
		else $this->value = $value;

		return $this;
	}

	/**
	 * @return int - ширина комментария
	 */
	function getCommentWidth() : int{
		return $this->commentWidth;
	}

	/**
	 * @param int $width - ширина комментария
	 * @return Cell - $this
	 */
	function setCommentWidth(int $width) : self{
		Validator::validateInRange($width, 1, 409, '$width');

		if ($this->master) $this->master->setCommentWidth($width);
		else $this->commentWidth = $width;

		return $this;
	}

	/**
	 * @return int - высота комментария
	 */
	public function getCommentHeight(): int{
		return $this->commentHeight;
	}

	/**
	 * @param int $height - высота комментария
	 * @return Cell - $this
	 */
	function setCommentHeight(int $height) : self{
		Validator::validateInRange($height, 1, 409, '$height');

		if ($this->master) $this->master->setCommentHeight($height);
		else $this->commentHeight = $height;

		return $this;
	}

	/**
	 * @return Value|null - комментарий
	 */
	function getComment(){
		return $this->comment;
	}

	/**
	 * @param string|RichTextValue|null $comment - комментарий
	 * @return Cell - $this
	 * @throws InvalidValueException
	 */
	function setComment($comment) : self{
		if (!is_null($comment) && !($comment instanceof Value)){
			$comment = Value::parse($comment);
			if ($comment->getType() !== Value::TYPE_STRING && $comment->getType() !== Value::TYPE_RICH_TEXT)
				throw new InvalidValueException('$comment must be string or rich text');
		}

		if ($this->master) $this->master->setComment($comment);
		else $this->comment = $comment;

		return $this;
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
	 * @return int - тип значения ячейки
	 */
	function getType() : int{
		return $this->value->getType();
	}

	/**
	 * @param StylesXml $styles - стили xlsx
	 * @param SheetRels $sheetRels - связи таблицы
	 * @param Comments $comments - комментарии таблицы
	 * @return array - модель ячейки
	 */
	function prepareToCommit(StylesXml $styles, SheetRels $sheetRels, Comments $comments) : array{
		$workbook = $this->row->getWorksheet()->getWorkbook();
		$value = $this->value;

		if ($workbook->getUseSharedStrings()){
			switch ($value->getType()) {
				case Value::TYPE_STRING:
					$value = $workbook->addSharedString($value);
					break;

				case Value::TYPE_HYPERLINK:
					$valueModel = $this->getValue();
					if (!isset($valueModel['ssId']))
						$value = new HyperlinkValue($valueModel['hyperlink'], $workbook->addSharedString($valueModel['text']));
			}
		}

		if ($value->getType() === Value::TYPE_RICH_TEXT) $value = $workbook->addSharedString($value);

		$model = [
			'address' => $this->getAddress(),
			'value' => $value->getValue(),
			'type' => $value->getType(),
			'styleId' => $styles->addStyle($this, $this->getType())
		];

		if ($this->master) $model['master'] = $this->master->getModel();

		if ($this->value instanceof HyperlinkValue) $sheetRels->addHyperlink(
			$model['value']['hyperlink'],
			$model['address']
		);

		if ($this->comment) $comments->addComment($this);

		return $model;
	}

	/**
	 * @return string - адрес ячейки ('A1', 'D23')
	 */
	function getAddress() : string{
		return Cell::genAddress($this->col, $this->row->getNumber());
	}

	/**
	 * Внутренняя функция. Назначает главную ячейку
	 *
	 * @param Cell|null $master - главная ячейка
	 */
	function setMaster(Cell $master = null){
		$this->master = $master;
	}

	/**
	 * Возвращает строку колонки по ее номеру. Например, 1 - A, 3 - C.
	 *
	 * @param int $col - номер колонки
	 * @return string - строка колонки
	 * @throws InvalidValueException - ошибочный номер колонки
	 */
	static function genColStr(int $col) : string{
		if ($col < 1 || $col > 16384) throw new InvalidValueException("$col is out of bounds. Excel supports columns from 1 to 16384");
		if ($col > 26) return Cell::genColStr(($col - 1) / 26) . chr(($col % 26 ? $col % 26 : 26) + 64);
		return chr($col + 64);
	}

	/**
	 * Возвращает номер колонки по ее строке. Например, A - 1, AA - 27.
	 *
	 * @param string $col - строка колонки
	 * @return int - номер колонки
	 * @throws InvalidValueException - ошибочная строка колонки
	 */
	static function genColNum(string $col) : int{
		$len = strlen($col);
		if ($len < 1 || $len > 3) throw new InvalidValueException("Out of bounds. Invalid column $col");

		$result = 0;
		for ($i = 0; $i < $len; $i++){
			$charCode = ord(substr($col, -$i - 1, 1));
			if ($charCode < 65 || $charCode > 90) throw new InvalidValueException("Out of bounds. Invalid column $col");

			$result += ($charCode - 64) * pow(26, $i);
		}

		return $result;
	}

	/**
	 * @param int $col - номер колонки
	 * @param int $row - номер строки
	 * @return string - адрес ячейки ('A1', 'D23')
	 * @throws InvalidValueException - ошибочный номер колонки/строки
	 */
	static function genAddress(int $col, int $row) : string{
		if ($row < 1 || $row > 1048576) throw new InvalidValueException("$row is out of bounds. Excel supports rows from 1 to 1048576");
		return self::genColStr($col) . $row;
	}
}