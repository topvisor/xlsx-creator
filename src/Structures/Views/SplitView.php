<?php

namespace Topvisor\XlsxCreator\Structures\Views;

use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class SplitView. Представление разделено на 4 секции с независимой прокруткой.
 *
 * @package  Topvisor\XlsxCreator\Structures\Views
 */
class SplitView extends View {
	public const VALID_ACTIVE_PANE = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'];

	public function __construct(int $xSplit, int $ySplit) {
		$this->model['state'] = 'split';

		$this->model['xSplit'] = $xSplit;
		$this->model['ySplit'] = $ySplit;
	}

	/**
	 * @return int - Количество точек слева до границы
	 */
	public function getXSplit(): int {
		return $this->model['xSplit'];
	}

	/**
	 * @param int $xSplit - Количество точек слева до границы
	 * @return SplitView - $this
	 */
	public function setXSplit(int $xSplit): self {
		Validator::validatePositive($xSplit, '$xSplit');

		$this->model['xSplit'] = $xSplit;

		return $this;
	}

	/**
	 * @return int - Количество точек сверху до границы
	 */
	public function getYSplit(): int {
		return $this->model['ySplit'];
	}

	/**
	 * @param int $ySplit - Количество точек сверху до границы
	 * @return SplitView - $this
	 */
	public function setYSplit(int $ySplit): self {
		Validator::validatePositive($ySplit, '$ySplit');

		$this->model['ySplit'] = $ySplit;

		return $this;
	}

	/**
	 * @return string|null - Левая-верхняя ячейка в нижней правой панели
	 */
	public function getTopLeftCell() {
		return $this->model['topLeftCell'] ?? null;
	}

	/**
	 * @param string|null $topLeftCell - Левая-верхняя ячейка в нижней правой панели (например, 'A1', 'B10', и т.д.)
	 * @return SplitView - $this
	 */
	public function setTopLeftCell(?string $topLeftCell = null): self {
		if (!is_null($topLeftCell)) Validator::validateAddress($topLeftCell);

		$this->model['topLeftCell'] = $topLeftCell;

		return $this;
	}

	/**
	 * @return string|null - активная панель
	 */
	public function getActivePane() {
		return $this->model['activePane'] ?? null;
	}

	/**
	 * @param string|null $activePane - активная панель
	 * @return SplitView - $this
	 */
	public function setActivePane(?string $activePane = null): self {
		if (!is_null($activePane)) Validator::validate($activePane, '$activePane', self::VALID_ACTIVE_PANE);

		$this->model['activePane'] = $activePane;

		return $this;
	}
}
