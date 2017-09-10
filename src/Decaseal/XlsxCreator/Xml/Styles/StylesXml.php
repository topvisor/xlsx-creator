<?php

namespace Decaseal\XlsxCreator\Xml\Styles;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\ListXml;
use Decaseal\XlsxCreator\Xml\Styles\Border\BorderXml;
use Decaseal\XlsxCreator\Xml\Styles\Fill\FillXml;
use Decaseal\XlsxCreator\Xml\Styles\Font\FontXml;
use Decaseal\XlsxCreator\Xml\Styles\Index\StylesIndex;
use Decaseal\XlsxCreator\Xml\Styles\Index\StylesNumFmtIndex;
use Decaseal\XlsxCreator\Xml\Styles\Style\StyleXml;
use XMLWriter;

class StylesXml extends BaseXml{
	private $fontIndex;
	private $borderIndex;
	private $styleIndex;
	private $fillIndex;
	private $numFmtIndex;

	function __construct(){
		$this->fontIndex = new StylesIndex(new FontXml());
		$this->borderIndex = new StylesIndex(new BorderXml());
		$this->styleIndex = new StylesIndex(new StyleXml());
		$this->fillIndex = new StylesIndex(new FillXml());
		$this->numFmtIndex = new StylesNumFmtIndex(new NumFmtXml());

		$this->fontIndex->addIndex(['sz' => 11, 'color' => ['theme' => 1], 'name' => 'Calibri', 'family' => 2, 'scheme' => 'minor']);
		$this->borderIndex->addIndex([]);
		$this->styleIndex->addIndex(['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'borderId' => 0, 'xfId' => 0]);
		$this->fillIndex->addIndex(['type' => 'pattern', 'pattern' => 'none']);
		$this->fillIndex->addIndex(['type' => 'pattern', 'pattern' => 'gray125']);
	}

	function render(XMLWriter $xml, $model = null){
		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('styleSheet');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$xml->writeAttribute('mc:Ignorable', 'x14ac x16r2');
		$xml->writeAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
		$xml->writeAttribute('xmlns:x16r2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main');

		if ($rawXmls = $this->numFmtIndex->getXmls() ?? false) $this->addIndexToXml($xml, 'numFmts', $rawXmls);
		if ($rawXmls = $this->fontIndex->getXmls() ?? false) $this->addIndexToXml($xml, 'fonts', $rawXmls);
		if ($rawXmls = $this->fillIndex->getXmls() ?? false) $this->addIndexToXml($xml, 'fills', $rawXmls);
		if ($rawXmls = $this->borderIndex->getXmls() ?? false) $this->addIndexToXml($xml, 'borders', $rawXmls);

		(new ListXml('cellStyleXfs', new StyleXml(false), [], false, true))
			->render($xml, ['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'borderId' => 0]);

		if ($rawXmls = $this->styleIndex->getXmls() ?? false) $this->addIndexToXml($xml, 'cellXfs', $rawXmls);



		$xml->endElement();
		$xml->endDocument();
	}

	function addIndexToXml(XMLWriter $xml, string $tag, array $rawXmls){
		$xml->startElement($tag);
		$xml->writeAttribute('count', count($rawXmls));

		foreach ($rawXmls as $rawXml) $xml->writeRaw($rawXml);

		$xml->endElement();
	}
}