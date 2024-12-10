<?php

namespace Topvisor\XlsxCreator;

use Topvisor\XlsxCreator\Helpers\Styles;
use Topvisor\XlsxCreator\Helpers\Validator;
use Topvisor\XlsxCreator\Structures\Styles\Style;

/**
 * Class Column. Содержит методы для работы с колонкой.
 *
 * @package Topvisor\XlsxCreator
 */
class Column extends Style {
	private $worksheet;
	private $number;

	private $width;
	private $hidden;
	private $outlineLevel;

	public function __construct(Worksheet $worksheet, int $number) {
		$this->width = 8;
		$this->worksheet = $worksheet;
		$this->number = $number;
	}

	public function __destruct() {
		parent::__destruct();

		unset($this->worksheet);
	}

	/**
	 * @return Worksheet - таблица
	 */
	public function getWorksheet(): Worksheet {
		return $this->worksheet;
	}

	/**
	 * @return int - номер строки
	 */
	public function getNumber(): int {
		return $this->number;
	}

	/**
	 * @return int - ширина столбца
	 */
	public function getWidth(): int {
		return $this->width;
	}

	/**
	 * @param int $width - ширина столбца
	 * @return Column - $this
	 */
	public function setWidth(int $width = 8): self {
		Validator::validateInRange($width, 0, 409, '$width');

		$this->width = $width;

		return $this;
	}

	/**
	 * @return bool - скрытый ли столбец
	 */
	public function isHidden(): bool {
		return $this->hidden;
	}

	/**
	 * @param bool $hidden - скрыть столбец
	 * @return Column - $this
	 */
	public function setHidden(bool $hidden): self {
		$this->hidden = $hidden;

		return $this;
	}

	/**
	 * @return int - column outline level
	 */
	public function getOutlineLevel(): int {
		return $this->outlineLevel;
	}

	/**
	 * @param int $outlineLevel - column outline level
	 * @return Column - $this
	 */
	public function setOutlineLevel(int $outlineLevel): self {
		Validator::validateInRange($outlineLevel, 0, 409, '$outlineLevel');

		$this->outlineLevel = $outlineLevel;

		return $this;
	}

	/**
	 * @param Styles $styles - стили xlsx
	 * @return array|null - модель
	 */
	public function prepareToCommit(Styles $styles): array {
		$styleId = $styles->addStyle($this);
		$collapsed = (bool) ($this->outlineLevel && $this->outlineLevel > $this->worksheet->getOutlineLevelCol());

		return [
			'min' => $this->number,
			'max' => $this->number,
			'width' => $this->width,
			'styleId' => $styleId,
			'hidden' => $this->hidden,
			'outlineLevel' => $this->outlineLevel,
			'collapsed' => $collapsed,
		];
	}
}
