<?php

namespace XlsxCreator\Structures\Views;

/**
 * Class NormalView. Обычное представление.
 *
 * @package XlsxCreator\Structures\Views
 */
class NormalView extends View{
	public function __construct(){
		$this->model['state'] = 'normal';
	}
}