<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\ListXml;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class StylesXml extends BaseXml{
	private $index;
	private $model;

	function __construct(){
		$this->index = [
			'style' => [],
			'numFmt' => [],
			'numFmtNextId' => 164,
			'font' => [],
			'border' => [],
			'fill' => [],
		];

		$this->model = [
			'styles' => [],
			'numFmts' => [],
			'fonts' => [],
			'borders' => [],
			'fills' => []
		];

		$this->addFont(['size' => 11, 'color' => ['theme' => 1], 'name' => 'Calibri', 'family' => 2, 'scheme' => 'minor']);
	}

	function render(XMLWriter $xml, $model = null){

	}

	private function addFont(array $fontModel) : int{
		$xml = (new FontXml())->toXml($fontModel);

		if(!isset($this->index['font'][$xml])){
			$this->index['font'][$xml] = count($this->model['fonts']);
			$this->model['fonts'][] = $xml;
		}

		return $this->index['font'][$xml];
	}
}