<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\Border\BorderXml;
use XMLWriter;

class StylesXml extends BaseXml{
	private const DEFAULT_NUM_FMTS = [
		'General' => 0,
		'0' => 1,
		'0.00' => 2,
		'#,##0' => 3,
		'#,##0.00' => 4,
		'0%' => 9,
		'0.00%' => 10,
		'0.00E+00' => 11,
		'# ?/?' => 12,
		'# ??/??' => 13,
		'mm-dd-yy' => 14,
		'd-mmm-yy' => 15,
		'd-mmm' => 16,
		'mmm-yy' => 17,
		'h:mm AM/PM' => 18,
		'h:mm:ss AM/PM' => 19,
		'h:mm' => 20,
		'h:mm:ss' => 21,
		'm/d/yy "h":mm' => 22,
		'#,##0 ;(#,##0)' => 37,
		'#,##0 ;[Red](#,##0)' => 38,
		'#,##0.00 ;(#,##0.00)' => 39,
		'#,##0.00 ;[Red](#,##0.00)' => 40,
		'mm:ss' => 45,
		'[h]:mm:ss' => 46,
		'mmss.0' => 47,
		'##0.0E+0' => 48,
		'@' => 49
	];

	private $models;

	function __construct(){
		$this->models = [];

		foreach (['numFmts', 'fonts', 'fills', 'borders', 'styles'] as $key) {
			$this->models[$key] = [
				'models' => []
			];
		}

		$this->models['numFmts']['nextIndex'] = 164;

		$this->addFont(['size' => 11, 'color' => ['theme' => 1], 'name' => 'Calibri', 'family' => 2, 'scheme' => 'minor']);
	}

	function render(XMLWriter $xml, $model = null){

	}

	private function addIndex(string $key, BaseXml $baseXml, array $model) : int{
		if ($key === 'numFmts') return $this->addNumFmtsIndex($baseXml, $model);
	}

	private function addNumFmtsIndex(BaseXml $baseXml, array $model) :int{

	}

//	private function addFont(array $fontModel) : int{
//		$xml = (new FontXml())->toXml($fontModel);
//
//		if(!isset($this->index['font'][$xml])){
//			$this->index['font'][$xml] = count($this->model['fonts']);
//			$this->model['fonts'][] = $xml;
//		}
//
//		return $this->index['font'][$xml];
//	}
//
//	private function addBorder(array $borderModel) : int{
//		$xml = (new BorderXml())->toXml($borderModel);
//
//	}
}