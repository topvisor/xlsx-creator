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

	const TAG = 'styleSheet';
	const STYLESHEET_ATTRIBUTES = [
		'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
		'xmlns:mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006',
		'mc:Ignorable' => 'x14ac x16r2',
		'xmlns:x14ac' => 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
		'xmlns:x16r2' => 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main'
	];

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

		$this->addIndex(FillXml::class, [XlsxCreator::FILL_TYPE => XlsxCreator::FILL_PATTERN, XlsxCreator::FILL_PATTERN => XlsxCreator::FILL_PATTERN_NONE]);
		$this->addIndex(FillXml::class, [XlsxCreator::FILL_TYPE => XlsxCreator::FILL_PATTERN, XlsxCreator::FILL_PATTERN => XlsxCreator::FILL_PATTERN_GRAY_125]);
	}

	function render(XMLWriter $xml, $model = null){
		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('styleSheet');

		foreach (StylesXml::STYLESHEET_ATTRIBUTES as $name => $value) $xml->writeAttribute($name, $value);



		$xml->endElement();
		$xml->endDocument();
	}

	private function addIndex(string $key, $model){
		return $this->indexes[$key][StylesXml::INDEX]->addIndex($this->indexes[$key][StylesXml::BASE_XML], $model);
	}
}