<?php

namespace Topvisor\XlsxCreator\Helpers;

use Serializable;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Structures\Styles\Alignment\Alignment;
use Topvisor\XlsxCreator\Structures\Styles\Borders\Borders;
use Topvisor\XlsxCreator\Structures\Styles\Font;
use Topvisor\XlsxCreator\Structures\Color;
use Topvisor\XlsxCreator\Structures\Styles\Style;
use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Workbook;
use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\ListXml;
use Topvisor\XlsxCreator\Xml\Styles\Border\BorderXml;
use Topvisor\XlsxCreator\Xml\Styles\Fill\FillXml;
use Topvisor\XlsxCreator\Xml\Styles\Font\FontXml;
use Topvisor\XlsxCreator\Xml\Styles\NumFmtXml;
use Topvisor\XlsxCreator\Xml\Styles\Style\StyleXml;
use XMLWriter;

class Styles{
	const DEFAULT_NUM_FMT = [
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

	const NUM_FMT_START_INDEX = 164;
	const FONT_START_INDEX = 1;
	const FILL_START_INDEX = 2;
	const STYLE_START_INDEX = 1;

	private $fontIndex;
	private $borderIndex;
	private $styleIndex;
	private $fillIndex;
	private $numFmtIndex;

	function __construct(){
		$this->fontIndex = [];
		$this->borderIndex = [];
		$this->fillIndex = [];
		$this->numFmtIndex = [];
		$this->styleIndex = [];

		$this->addIndex(new Borders(), $this->borderIndex);
	}

	function writeToFile(string $filename){
		$xml = new XMLWriter();
		$xml->openURI($filename);

		$xml->startDocument('1.0', 'UTF-8', 'yes');
		$xml->startElement('styleSheet');

		$xml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$xml->writeAttribute('mc:Ignorable', 'x14ac x16r2');
		$xml->writeAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
		$xml->writeAttribute('xmlns:x16r2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main');

		$this->renderNumFmts($xml);
		$this->renderFonts($xml);
		$this->renderFills($xml);
		$this->renderBorders($xml);

		(new ListXml('cellStyleXfs', new StyleXml(false), [], false, true))
			->render($xml, [['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'borderId' => 0]]);

		$this->renderStyles($xml);

		$this->writeStatic($xml);

		$xml->endElement();
		$xml->endDocument();

		$xml->flush();
		unset($xml);
	}

	function addStyle(Style $style, int $cellType = null) : int{
		if (!$style->getNumFmt()) {
			switch ($cellType) {
				case null:
				case Value::TYPE_NUMBER: $style->setNumFmt('General'); break;
				case Value::TYPE_DATE: $style->setNumFmt('mm-dd-yy'); break;
			}
		}

		if (!$style->isDefaultStyle()) return 0;

		$styleKey =
			($style->getNumFmt()
				? $this->addIndex($style->getNumFmt(), $this->numFmtIndex, self::NUM_FMT_START_INDEX, self::DEFAULT_NUM_FMT)
				: ''
			) . ';' .
			($style->getFont() ? $this->addIndex($style->getFont(), $this->fontIndex, self::FONT_START_INDEX) : '') . ';' .
			($style->getFill() ? $this->addIndex($style->getFill(), $this->fillIndex, self::FILL_START_INDEX) : '') . ';' .
			($style->getBorders() ? $this->addIndex($style->getBorders(), $this->borderIndex) : '') . ';' .
			($style->getAlignment() ? str_replace(';', urlencode(';'), $style->getAlignment()->serialize()) : '');

		return $this->addIndex($styleKey, $this->styleIndex, self::STYLE_START_INDEX);
	}

	private function addIndex($key, array &$indexes, int $startIndex = 0, array $defaults = []) : int{
		if ($key instanceof Serializable) $key = $key->serialize();
		elseif (!is_string($key)) throw new InvalidValueException('$key must be string or Serializable');

		if (isset($defaults[$key])) return $defaults[$key];
		if (isset($indexes[$key])) return $indexes[$key];

		$index = count($indexes) + $startIndex;
		$indexes[$key] = $index;

		return $index;
	}

	private function renderNumFmts(XMLWriter $xml){
		$xml->startElement('numFmts');
		$xml->writeAttribute('count', count($this->numFmtIndex));

		foreach (array_keys($this->numFmtIndex) as $key) (new NumFmtXml())->render($xml, [
			'formatCode' => $key,
			'id' => $this->numFmtIndex[$key]
		]);

		$xml->endElement();
	}

	private function renderFonts(XMLWriter $xml){
		$xml->startElement('fonts');
		$xml->writeAttribute('count', count($this->fontIndex) + self::FONT_START_INDEX);

		$fontXml = new FontXml();
		$fontXml->render($xml, ['sz' => 11, 'color' => ['theme' => 1], 'name' => 'Calibri', 'family' => 2, 'scheme' => 'minor']);

		$font = new Font();
		foreach (array_keys($this->fontIndex) as $key){
			$font->unserialize($key);
			$fontXml->render($xml, $font->getModel());
		}

		$xml->endElement();
	}

	private function renderFills(XMLWriter $xml){
		$xml->startElement('fills');
		$xml->writeAttribute('count', count($this->fillIndex) + self::FILL_START_INDEX);

		$fillXml = new FillXml();
		$fillXml->render($xml, ['type' => 'pattern', 'pattern' => 'none']);
		$fillXml->render($xml, ['type' => 'pattern', 'pattern' => 'gray125']);

		$color = Color::fromHex();
		foreach (array_keys($this->fillIndex) as $key) {
			$color->unserialize($key);
			$fillXml->render($xml, [
				'type' => 'pattern',
				'pattern' => 'solid',
				'fgColor' => $color->getModel(),
				'bgColor' => $color->getModel()
			]);
		}

		$xml->endElement();
	}

	private function renderBorders(XMLWriter $xml){
		$xml->startElement('borders');
		$xml->writeAttribute('count', count($this->borderIndex));

		$borderXml = new BorderXml();
		$border = new Borders();
		foreach (array_keys($this->borderIndex) as $key){
			$border->unserialize($key);
			$borderXml->render($xml, $border->getModel());
		}

		$xml->endElement();
	}

	private function renderStyles(XMLWriter $xml){
		$xml->startElement('cellXfs');
		$xml->writeAttribute('count', count($this->styleIndex) + self::STYLE_START_INDEX);

		$styleXml = new StyleXml();
		$styleXml->render($xml, ['numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'borderId' => 0, 'xfId' => 0]);

		$alignment = new Alignment();
		foreach (array_keys($this->styleIndex) as $key){
			$params = explode(';', $key);
			$model = [
				'numFmtId' => (int) $params[0],
				'fontId' => (int) $params[1],
				'fillId' => (int) $params[2],
				'borderId' => (int) $params[3]
			];

			if ($params[4]) {
				$alignment->unserialize(str_replace(urlencode(';'), ';', $params[4]));
				$model['alignment'] = $alignment->getModel();
			}

			$styleXml->render($xml, $model);
		}

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