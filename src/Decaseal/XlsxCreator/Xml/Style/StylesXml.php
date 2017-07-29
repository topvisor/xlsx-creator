<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\Border\BorderXml;
use Decaseal\XlsxCreator\Xml\Style\Fill\FillXml;
use Decaseal\XlsxCreator\Xml\Style\Index\Index;
use Decaseal\XlsxCreator\Xml\Style\Index\NumFmtsIndex;
use XMLWriter;

class StylesXml extends BaseXml{
	const INDEX = 'index';
	const BASE_XML = 'baseXml';

	private $indexes;

	function __construct(){
		$this->indexes = [];

		$this->indexes[NumFmtXml::class] = [StylesXml::BASE_XML => new NumFmtXml(), StylesXml::INDEX => new NumFmtsIndex()];
		$this->indexes[StyleXml::class] = [StylesXml::BASE_XML => new StyleXml(true), StylesXml::INDEX => new Index()];

		foreach ([FontXml::class, FillXml::class, BorderXml::class] as $key)
			$this->indexes[$key] = [StylesXml::BASE_XML => new $key(), StylesXml::INDEX => new Index()];

		$this->addIndex(FontXml::class, [
			XlsxCreator::FONT_SIZE => 11,
			XlsxCreator::FONT_COLOR => [ColorXml::THEME => 1],
			XlsxCreator::FONT_NAME => 'Calibri',
			XlsxCreator::FONT_FAMILY => 2,
			XlsxCreator::FONT_SCHEME => XlsxCreator::FONT_SCHEME_MINOR
		]);

		$this->addIndex(BorderXml::class, []);
	}

	function render(XMLWriter $xml, $model = null){

	}

	private function addIndex(string $key, $model){
		return $this->indexes[$key][StylesXml::INDEX]->addIndex($this->indexes[$key][StylesXml::BASE_XML], $model);
	}
}