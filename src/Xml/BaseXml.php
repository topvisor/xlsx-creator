<?php

namespace Topvisor\XlsxCreator\Xml;

use XMLWriter;

/**
 * Class BaseXml. Класс, и все наследующиеся от него, являются оберткой для создания xml кода.
 *
 * @package  Topvisor\XlsxCreator\Xml
 */
abstract class BaseXml{
	/**
	 * Создает xml код согласно $model используя XMLWriter $xml
	 *
	 * @param XMLWriter $xml - поток для записи xml кода
	 * @param array|null $model - модель, согласно которой генерируется xml код
	 */
	abstract function render(XMLWriter $xml, array $model = null);

	/**
	 * Создает xml код согласно $model
	 *
	 * @param null $model - модель, согласно которой генерируется xml код
	 * @return string - xml код
	 */
	function toXml($model = null) : string{
		$xml = new XMLWriter();
		$xml->openMemory();

		$this->render($xml, $model);

		return $xml->outputMemory();
	}

	/**
	 * Обрабатывает текст перед записью в xlsx файл
	 * Экранирует excel utf-8 символы
	 *
	 * @param string $text
	 * @return string
	 */
	protected function prepareText(string $text) : string{
		return preg_replace('/_x\d{4}_/', '_x005F$0', $text);
	}
}