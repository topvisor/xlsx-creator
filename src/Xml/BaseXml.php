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
		$text = preg_replace('/_x\d{4}_/', '_x005F$0', $text);
		$preparedText = "";

		for ($i = 0; $i < mb_strlen($text); $i++) {
			$chr = mb_substr($text, $i, 1);
			$ord = ord($chr);

			$encode = true;
			if (strlen($chr) > 1) $encode = false;
			if ($encode && $ord > 0x8 && $ord < 0xb) $encode = false;
			if ($encode && $ord > 0xc && $ord < 0xe) $encode = false;
			if ($encode && $ord > 0x1f && $ord < 0x7f) $encode = false;
			if ($encode && $ord > 0x9f) $encode = false;

			if ($encode)
				$preparedText .= sprintf("_x%04x_", $ord);
			else
				$preparedText .= $chr;
		}

		return $preparedText;
	}
}