<?php

namespace Topvisor\XlsxCreator\Structures\Styles;

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\Styles\Alignment\Alignment;
use Topvisor\XlsxCreator\Structures\Styles\Borders\Borders;

/**
 * class Style. Управляет стилями.
 *
 * @package Topvisor\XlsxCreator\Structures\Styles
 */
class Style{
	protected $numFmt;
	protected $font;
	protected $fill;
	protected $borders;
	protected $alignment;

	function __destruct(){
		unset($this->font);
		unset($this->fill);
		unset($this->borders);
		unset($this->alignment);
	}

	/**
	 * @return string|null - формат чисел ячейки
	 */
	function getNumFmt(){
		return $this->numFmt;
	}

	/**
	 * @param string|null $numFmt - формат чисел ячейки
	 * @return Style - $this
	 */
	function setNumFmt(string $numFmt = null) : self{
		$this->numFmt = $numFmt;
		return $this;
	}

	/**
	 * @return Font|null - шрифт
	 */
	function getFont(){
		return $this->font;
	}

	/**
	 * @param Font|null $font - шрифт
	 * @return Style - $this
	 */
	function setFont(Font $font = null) : self{
		$this->font = $font;
		return $this;
	}

	/**
	 * @return Color|null - заливка ячейки
	 */
	function getFill(){
		return $this->fill;
	}

	/**
	 * @param Color|null $color - заливка ячейки
	 * @return Style - $this
	 */
	function setFill(Color $color = null) : self{
		$this->fill = $color;
		return $this;
	}

	/**
	 * @return Borders|null - границы ячейки
	 */
	function getBorders(){
		return $this->borders;
	}

	/**
	 * @param Borders|null $borders - границы ячейки
	 * @return Style - $this
	 */
	function setBorders(Borders $borders = null) : self{
		$this->borders = $borders;
		return $this;
	}

	/**
	 * @return Alignment|null - выравнивание текста
	 */
	function getAlignment(){
		return $this->alignment ?? null;
	}

	/**
	 * @param Alignment|null $alignment - выравнивание текста
	 * @return Style - $this
	 */
	function setAlignment(Alignment $alignment = null) : self{
		$this->alignment = $alignment;
		return $this;
	}

	function isDefaultStyle() : bool{
		return is_null($this->numFmt) && !$this->font && !$this->fill && !$this->borders && !$this->alignment;
	}

	/**
	 * @return array - модель
	 */
	function getStyleModel() : array{
		if ($this->isDefaultStyle()) return [];
		
		return [
			'numFmt' => $this->numFmt,
			'font' => $this->font ? $this->font->getModel() : null,
			'fill' => $this->fill ? [
				'type' => 'pattern',
				'pattern' => 'solid',
				'fgColor' => $this->fill->getModel(),
				'bgColor' => $this->fill->getModel()
			] : null,
			'border' => $this->borders ? $this->borders->getModel() : null,
			'alignment' => $this->alignment ? $this->alignment->getModel() : null,
		];
	}
}