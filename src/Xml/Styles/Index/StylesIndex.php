<?php

namespace Topvisor\XlsxCreator\Xml\Styles\Index;

use Topvisor\XlsxCreator\Xml\BaseXml;

/**
 * Class StylesIndex. Индексирует модели стилей.
 *
 * @package XlsxCreator\Xml\Styles\Index
 */
class StylesIndex{
	protected $baseXml;

	protected $indexes;
	protected $xmls;

	/**
	 * StylesIndex constructor.
	 *
	 * @param BaseXml $baseXml - класс, модели которого будут индексироваться
	 */
	function __construct(BaseXml $baseXml){
		$this->baseXml = $baseXml;

		$this->indexes = [];
		$this->xmls = [];
	}

	/**
	 * Добавить модель в индекс
	 *
	 * @param $model - модель
	 * @return int - индекс
	 */
	function addIndex($model) : int{
		$xml = $this->baseXml->toXml($model);

		if (isset($this->indexes[$xml])) return $this->indexes[$xml];

		$index = count($this->xmls);

		$this->indexes[$xml] = $index;
		$this->xmls[] = $xml;

		return $index;
	}

	/**
	 * Возвращает весь сгенерированный из моделей xml код
	 *
	 * @return array - xml код
	 */
	function getXmls() : array{
		return $this->xmls;
	}
}