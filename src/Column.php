<?php

namespace Topvisor\XlsxCreator;

use Topvisor\XlsxCreator\Helpers\Validator;
use Topvisor\XlsxCreator\Structures\Styles\Style;

/**
 * Class Column. Содержит методы для работы с колонкой.
 *
 * @package Topvisor\XlsxCreator
 */
class Column extends Style{
	private $worksheet;
	private $number;

	private $width;
	private $hidden;
	private $outlineLevel;

	public function __construct(Worksheet $worksheet, int $number){
		$this->worksheet = $worksheet;
		$this->number = $number;
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
	 * @return null|int - ширина столбца
	 */
	function getWidth(){
		return $this->width;
	}

	/**
	 * @param int|null $width - ширина столбца
	 * @return Column - $this
	 */
	function setWidth(int $width = null) : self{
		Validator::validateInRange($width, 0, 409, '$width');

		$this->width = $width;
		return $this;
	}

	/**
	 * @return bool - скрытый ли столбец
	 */
	function isHidden() : bool{
		return $this->hidden;
	}

	/**
	 * @param bool $hidden - скрыть столбец
	 * @return Column - $this
	 */
	function setHidden(bool $hidden) : self{
		$this->hidden = $hidden;
		return $this;
	}

	/**
	 * @return int - column outline level
	 */
	function getOutlineLevel() : int{
		return $this->outlineLevel;
	}

	/**
	 * @param int $outlineLevel - column outline level
	 * @return Column - $this
	 */
	function setOutlineLevel(int $outlineLevel) : self{
		Validator::validateInRange($outlineLevel, 0, 409, '$outlineLevel');

		$this->outlineLevel = $outlineLevel;
		return $this;
	}

	/**
	 * @return array|null - модель
	 */
	function getModel(){
		$styleId = $this->worksheet->getWorkbook()->getStyles()->addStyle($this);
		$collapsed = (bool) ($this->outlineLevel && $this->outlineLevel > $this->worksheet->getOutlineLevelCol());

		return ($this->isDefault()) ? null : [
			'min' => $this->number,
			'max' => $this->number,
			'width' => $this->width,
			'styleId' => $styleId,
			'hidden' => $this->hidden,
			'outlineLevel' => $this->outlineLevel,
			'collapsed' => $collapsed
		];
	}

	/**
	 * @return bool - является ли колонка не измененной
	 */
	private function isDefault() : bool{
		if ($this->isCustomWidth()) return false;
		if ($this->hidden) return false;
		if ($this->outlineLevel) return false;
		if ($this->getStyleModel()) return false;

		return true;
	}

	/**
	 * @return bool - высота колонки изменена
	 */
	private function isCustomWidth() : bool{
		return (bool) ($this->width && $this->width != 8);
	}
}