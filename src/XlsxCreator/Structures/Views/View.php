<?php

namespace XlsxCreator\Structures\Views;

use XlsxCreator\Validator;

/**
 * Class View. Описывает представление worksheet.
 *
 * @package XlsxCreator\Structures\Views
 */
abstract class View{
	const VALID_VIEW = ['pageBreakPreview', 'pageLayout'];

	protected $model = [];

	/**
	 * @return bool - ориентация справа на лево
	 */
	function getRightToLeft() : bool{
		return $this->model['rightToLeft'] ?? false;
	}

	/**
	 * @param bool $rightToLeft - ориентация справа на лево
	 * @return View - $this
	 */
	function setRightToLeft(bool $rightToLeft) : self{
		$this->model['rightToLeft'] = $rightToLeft;
		return $this;
	}

	/**
	 * @return string|null - Адрес выбраной ячейки
	 */
	function getActiveCell(){
		return $this->model['activeCell'] ?? null;
	}

	/**
	 * @param string|null $activeCell - адрес выбранной ячейки (например, 'A1', 'B10', и т.д.)
	 * @return View - $this
	 */
	function setActiveCell(string $activeCell = null) : self{
		if (!is_null($activeCell)) Validator::validateAddress($activeCell);

		$this->model['activeCell'] = $activeCell;
		return $this;
	}

	/**
	 * @return bool - показывать линейку в макете страницы
	 */
	function getShowRuler() : bool{
		return $this->model['showRuler'] ?? false;
	}

	/**
	 * @param bool $showRuler - показывать линейку в макете страницы
	 * @return View - $this
	 */
	function setShowRuler(bool $showRuler) : self{
		$this->model['showRuler'] = $showRuler;
		return $this;
	}

	/**
	 * @return bool - показывать заголовки строк и столбцов
	 */
	function getShowRowColHeaders() : bool{
		return $this->model['showRowColHeaders'] ?? false;
	}

	/**
	 * @param bool $showRowColHeaders - показывать заголовки строк и столбцов (например, A1, B1 вверху и 1,2,3 слева)
	 * @return View - $this
	 */
	function setShowRowColHeaders(bool $showRowColHeaders) : self{
		$this->model['showRowColHeaders'] = $showRowColHeaders;
		return $this;
	}

	/**
	 * @return bool - показывать линии сетки
	 */
	function getShowGridLines() : bool{
		return $this->model['showGridLines'] ?? false;
	}

	/**
	 * @param bool $showGridLines - показывать линии сетки
	 * @return View - $this
	 */
	function setShowGridLines(bool $showGridLines) : self{
		$this->model['showGridLines'] = $showGridLines;
		return $this;
	}

	/**
	 * @return int|null - процент увеличения
	 */
	function getZoomScale(){
		return $this->model['zoomScale'] ?? null;
	}

	/**
	 * @param int|null $zoomScale - процент увеличения
	 * @return View - $this
	 */
	function setZoomScale(int $zoomScale = null) : self{
		if (!is_null($zoomScale)) Validator::validatePositive($zoomScale, '$zoomScale');

		$this->model['zoomScale'] = $zoomScale;
		return $this;
	}

	/**
	 * @return int|null - нормальное увеличение
	 */
	function getZoomScaleNormal(){
		return $this->model['zoomScaleNormal'] ?? null;
	}

	/**
	 * @param int|null $zoomScaleNormal - нормальное увеличение
	 * @return View - $this
	 */
	function setZoomScaleNormal(int $zoomScaleNormal = null) : self{
		if (!is_null($zoomScaleNormal)) Validator::validatePositive($zoomScaleNormal, '$zoomScaleNormal');

		$this->model['zoomScaleNormal'] = $zoomScaleNormal;
		return $this;
	}

	/**
	 * @return string|null - cтиль отображения
	 */
	function getView(){
		return $this->model['view'] ?? null;
	}

	/**
	 * @param string|null $view - cтиль отображения ('pageBreakPreview', 'pageLayout')
	 * @return View - $this
	 */
	function setView(string $view = null) : self{
		if (!is_null($view)) Validator::validate($view, '$view', self::VALID_VIEW);

		$this->model['view'] = $view;
		return $this;
	}

	/**
	 * @return array - модель
	 */
	function getModel() : array{
		return $this->model;
	}
}