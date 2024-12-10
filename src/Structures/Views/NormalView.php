<?php

namespace Topvisor\XlsxCreator\Structures\Views;

/**
 * Class NormalView. Обычное представление.
 *
 * @package  Topvisor\XlsxCreator\Structures\Views
 */
class NormalView extends View {
	public function __construct() {
		$this->model['state'] = 'normal';
	}
}
