<?php

namespace Topvisor\XlsxCreator\Structures\Styles;

use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\Styles\Alignment\Alignment;
use Topvisor\XlsxCreator\Structures\Styles\Borders\Borders;

/**
 * Class Style. Описывает стили ячейки
 *
 * @package Topvisor\XlsxCreator\Structures\Styles
 */
class Style{
	private $numFmt;
	private $font;
	private $fill;
	private $borders;
	private $alignment;

	public function __destruct(){
		unset($this->font);
		unset($this->fill);
		unset($this->borders);
		unset($this->alignment);
	}

	function getNumFmt(){
		return $this->numFmt;
	}

	function setNumFmt(string $numFmt = null) : self{
		$this->numFmt = $numFmt;
		return $this;
	}

	function getFont(){
		return $this->font;
	}

	function setFont(Font $font = null) : self{
		$this->font = $font;
		return $this;
	}

	function getFill(){
		return $this->fill;
	}

	function setFill(Color $color = null){
		$this->fill = $color;
		return $this;
	}

	function getBorders(){
		return $this->borders;
	}

	function setBorders(Borders $borders = null) : self{
		$this->borders = $borders;
		return $this;
	}

	function getAlignment(){
		return $this->alignment ?? null;
	}

	function setAlignment(Alignment $alignment = null) : self{
		$this->alignment = $alignment;
		return $this;
	}

	function getModel() : array{
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