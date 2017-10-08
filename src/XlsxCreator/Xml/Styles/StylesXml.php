<?php

namespace XlsxCreator\Xml\Styles;

use XlsxCreator\Structures\Values\Value;
use XlsxCreator\Xml\BaseXml;
use XlsxCreator\Xml\ListXml;
use XlsxCreator\Xml\Styles\Border\BorderXml;
use XlsxCreator\Xml\Styles\Fill\FillXml;
use XlsxCreator\Xml\Styles\Font\FontXml;
use XlsxCreator\Xml\Styles\Index\StylesIndex;
use XlsxCreator\Xml\Styles\Index\StylesNumFmtIndex;
use XlsxCreator\Xml\Styles\Style\StyleXml;
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

	function render(XMLWriter $xml, array $model = null){
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
			->render($xml, [['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'borderId' => 0]]);

		if ($rawXmls = $this->styleIndex->getXmls() ?? false) $this->addIndexToXml($xml, 'cellXfs', $rawXmls);

		$this->writeStatic($xml);

		$xml->endElement();
		$xml->endDocument();
	}

	function addStyle(array $model, int $cellType = null) : int{
		if (!$model) return 0;

		$cellType = $cellType ?? Value::TYPE_NUMBER;
		$styleModel = [];

		if ($model['numFmt'] ?? false) {
			$styleModel['numFmtId'] = $this->numFmtIndex->addIndex($styleModel['numFmt']);
		} else {
			switch ($cellType) {
				case Value::TYPE_NUMBER: $styleModel['numFmtId'] = $this->numFmtIndex->addIndex('General'); break;
				case Value::TYPE_DATE: $styleModel['numFmtId'] = $this->numFmtIndex->addIndex('mm-dd-yy'); break;
			}
		}

		if ($model['font'] ?? false) $styleModel['fontId'] = $this->fontIndex->addIndex($model['font']);
		if ($model['fill'] ?? false) $styleModel['fillId'] = $this->fillIndex->addIndex($model['fill']);
		if ($model['border'] ?? false) $styleModel['borderId'] = $this->fillIndex->addIndex($model['border']);
		if ($model['alignment'] ?? false) $styleModel['alignment'] = $model['alignment'];

		return $this->styleIndex->addIndex($styleModel);
	}

	private function addIndexToXml(XMLWriter $xml, string $tag, array $rawXmls){
		$xml->startElement($tag);
		$xml->writeAttribute('count', count($rawXmls));

		foreach ($rawXmls as $rawXml) $xml->writeRaw($rawXml);

		$xml->endElement();
	}

	private function writeStatic(XMLWriter $xml){
		$xml->startElement('cellStyles');
		$xml->writeAttribute('count', 1);

		$xml->startElement('cellStyle');
		$xml->writeAttribute('name', 'Normal');
		$xml->writeAttribute('xfId', 0);
		$xml->writeAttribute('builtinId', 0);

		$xml->endElement();
		$xml->endElement();

		$xml->startElement('dxfs');
		$xml->writeAttribute('count', 0);
		$xml->endElement();

		$xml->startElement('tableStyles');
		$xml->writeAttribute('count', 0);
		$xml->writeAttribute('defaultTableStyle', 'TableStyleMedium2');
		$xml->writeAttribute('defaultPivotStyle', 'PivotStyleLight16');
		$xml->endElement();

		$xml->startElement('extLst');

		$xml->startElement('ext');
		$xml->writeAttribute('uri', '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}');
		$xml->writeAttribute('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main');

		$xml->startElement('x14:slicerStyles');
		$xml->writeAttribute('defaultSlicerStyle', 'SlicerStyleLight1');

		$xml->endElement();
		$xml->endElement();

		$xml->startElement('ext');
		$xml->writeAttribute('uri', '{9260A510-F301-46a8-8635-F512D64BE5F5}');
		$xml->writeAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

		$xml->startElement('x15:timelineStyles');
		$xml->writeAttribute('defaultTimelineStyle', 'TimeSlicerStyleLight1');

		$xml->endElement();
		$xml->endElement();

		$xml->endElement();
	}
}