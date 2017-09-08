<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\Index\StylesIndex;
use XMLWriter;

class StylesXml extends BaseXml{
	private $fontIndex;
	private $borderIndex;
	private $styleIndex;

	function __construct(){
		$this->fontIndex = new StylesIndex(new FontXml());
		$this->borderIndex = new StylesIndex(new BorderXml());
		$this->styleIndex = new StylesIndex(new StyleXml());

		$this->fontIndex->addIndex(['sz' => 11, 'color' => ['theme' => 1], 'name' => 'Calibri', 'family' => 2, 'scheme' => 'minor']);
		$this->borderIndex->addIndex([]);
		$this->styleIndex->addIndex(['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'borderId' => 0, 'xfId' => 0]);
	}

	function render(XMLWriter $xml, $model = null){
		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('styleSheet');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$xml->writeAttribute('mc:Ignorable', 'x14ac x16r2');
		$xml->writeAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
		$xml->writeAttribute('xmlns:x16r2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main');



		$xml->endElement();
		$xml->endDocument();
	}
}