<?php

/**
 * Библиотека для создания xlsx файлов
 *
 * @author decaseal <decaseal@gmail.com>
 * @version 1.0
 */

namespace Decaseal\XlsxCreator;

use DateTime;
use Decaseal\XlsxCreator\Xml\Core\AppXml;
use Decaseal\XlsxCreator\Xml\Core\ContentTypesXml;
use Decaseal\XlsxCreator\Xml\Core\RelationshipsXml;
use Decaseal\XlsxCreator\Xml\Style\StylesXml;
use ZipArchive;

/**
 * Class XlsxCreator. Используйте его для создания xlsx файла.
 *
 * @package Decaseal\XlsxCreator
 */
class XlsxCreator{
	private $tempdir;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $company;
	private $manager;

	private $stylesXml;
	private $worksheets;
	private $committed;
	private $nextId;
	private $zip;

	/**
	 * XlsxCreator constructor
	 *
	 * @param string $filename - путь к xlsx файлу
	 * @param string|null $tempdir - путь к директории для хранения временных файлов
	 * @param DateTime|null $created - время создания файла
	 * @param DateTime|null $modified - время изменения файла
	 * @param string|null $creator - автор файла
	 * @param string|null $lastModifiedBy - автор последнего изменения файла
	 * @param string|null $company - компания
	 * @param string|null $manager - менеджер
	 */
	function __construct(string $filename, string $tempdir = null, DateTime $created = null, DateTime $modified = null, string $creator = null,
						 string $lastModifiedBy = null, string $company = null, string $manager = null){

		$this->tempdir = $tempdir ?? sys_get_temp_dir();
		$this->created = $created ?? new DateTime();
		$this->modified = $modified ?? $this->created;
		$this->creator = $creator ?? 'XlsxWriter';
		$this->lastModifiedBy = $lastModifiedBy ?? $this->creator;
		$this->company = $company ?? '';
		$this->manager = $manager ?? null;

		$this->stylesXml = new StylesXml();
		$this->worksheets = [];
		$this->committed = false;
		$this->nextId = 1;

		$this->zip = new ZipArchive();
		$this->zip->open($filename, ZipArchive::CREATE | ZipArchive::OVERWRITE);
	}

	/**
	 * Добавить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @param string|null $tabColor - цвет вкладки в формате argb (#FFFFFF00)
	 * @param int $defaultRowHeight - высота строки в px
	 * @return Worksheet - таблица
	 */
	function addWorksheet(string $name, string $tabColor = null, int $defaultRowHeight = 15) : Worksheet{
		$worksheet = new Worksheet($this->nextId(), $name, $tabColor, $defaultRowHeight);
		$worksheets[$name] = $worksheet;

		return $worksheet;
	}

	/**
	 * Получить таблицу
	 *
	 * @param string $name - имя таблицы
	 * @return Worksheet - таблица
	 */
	function getWorksheet(string $name) : Worksheet{
		return $this->worksheets[$name];
	}

	/**
	 * Получить все таблицы
	 *
	 * @return array - массив таблиц
	 */
	function getWorksheets() : array{
		return $this->worksheets;
	}

	function isCommitted() : bool{
		return $this->committed;
	}

	function commit(){
		if ($this->isCommitted() || !count($this->getWorksheets())) return;

		foreach ($this->worksheets as $worksheet) $worksheet->commit();

		$this->zip->addFile('./Xml/Static/theme1.xml', 'xl/theme/theme1.xml');
		$this->zip->addFromString('/_rels/.rels', (new RelationshipsXml())->toXml([
			['Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target' => 'xl/workbook.xml'],
			['Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'Target' => 'docProps/core.xml'],
			['Id' => 'rId3', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'Target' => 'docProps/app.xml']
		]));
		$this->zip->addFromString('[Content_Types].xml', (new ContentTypesXml())->toXml(['worksheets' => $this->getWorksheets()]));
		$this->zip->addFromString('docProps/app.xml', (new AppXml())->toXml([
			'worksheets' => $this->getWorksheets(),
			'company' => $this->company,
			'manager' => $this->manager
		]));
	}

	private function nextId() : int{
		return $this->nextId++;
	}

	private function finalize(){

	}
}

//	const BORDER_LEFT = 'left';
//	const BORDER_RIGHT = 'right';
//	const BORDER_TOP = 'top';
//	const BORDER_BOTTOM = 'bottom';
//
//	const BORDER_DIAGONAL = 'diagonal';
//	const BORDER_DIAGONAL_UP = 'up';
//	const BORDER_DIAGONAL_DOWN = 'down';
//
//	const BORDER_COLOR = 'color';
//	const BORDER_STYLE = 'style';
//
//	const BORDER_STYLE_THIN = 'thin';
//	const BORDER_STYLE_DOTTED = 'dotted';
//	const BORDER_STYLE_DASHDOT = 'dashdot';
//	const BORDER_STYLE_HAIR = 'hair';
//	const BORDER_STYLE_DASHDOTDOT = 'dashdotdot';
//	const BORDER_STYLE_SLANTDASHDOT = 'slantdashdot';
//	const BORDER_STYLE_MEDIUMDASHED = 'mediumdashed';
//	const BORDER_STYLE_MEDIUMDASHDOTDOT = 'mediumdashdotdot';
//	const BORDER_STYLE_MEDIUMDASHDOT = 'mediumdashdot';
//	const BORDER_STYLE_MEDIUM = 'medium';
//	const BORDER_STYLE_DOUBLE = 'double';
//	const BORDER_STYLE_THICK = 'thick';
//
//	const FILL_TYPE = 'type';
//
//	const FILL_GRADIENT = 'gradient';
//	const FILL_GRADIENT_ANGLE = 'angle';
//	const FILL_GRADIENT_PATH = 'path';
//	const FILL_DEGREE = 'degree';
//	const FILL_CENTER = 'center';
//	const FILL_CENTER_LEFT = 'left';
//	const FILL_CENTER_RIGHT = 'right';
//	const FILL_CENTER_TOP = 'top';
//	const FILL_CENTER_BOTTOM = 'bottom';
//	const FILL_STOPS = 'stops';
//	const FILL_STOP_POSITION = 'position';
//	const FILL_STOP_COLOR = 'color';
//
//	const FILL_PATTERN = 'pattern';
//	const FILL_FG_COLOR = 'fgColor';
//	const FILL_BG_COLOR = 'bgColor';
//
//	const FILL_PATTERN_NONE = 'none';
//	const FILL_PATTERN_SOLID = 'solid';
//	const FILL_PATTERN_DARK_GRAY = 'darkGray';
//	const FILL_PATTERN_MEDIUM_GRAY = 'mediumGray';
//	const FILL_PATTERN_LIGHT_GRAY = 'lightGray';
//	const FILL_PATTERN_GRAY_125 = 'gray125';
//	const FILL_PATTERN_GRAY_0625 = 'gray0625';
//	const FILL_PATTERN_DARK_HORIZONTAL = 'darkHorizontal';
//	const FILL_PATTERN_DARK_VERTICAL = 'darkVertical';
//	const FILL_PATTERN_DARK_DOWN = 'darkDown';
//	const FILL_PATTERN_DARK_UP = 'darkUp';
//	const FILL_PATTERN_DARK_GRID = 'darkGrid';
//	const FILL_PATTERN_DARK_TRELLIS = 'darkTrellis';
//	const FILL_PATTERN_LIGHT_HORIZONTAL = 'lightHorizontal';
//	const FILL_PATTERN_LIGHT_VERTICAL = 'lightVertical';
//	const FILL_PATTERN_LIGHT_DOWN = 'lightDown';
//	const FILL_PATTERN_LIGHT_UP = 'lightUp';
//	const FILL_PATTERN_LIGHT_TRELLIS = 'lightTrellis';
//	const FILL_PATTERN_LIGHT_GRID = 'lightGrid';
//
//	const COLOR_ARGB = 'argb';
//
//	const FONT_BOLD = 'b';
//	const FONT_ITALIC = 'i';
//	const FONT_UNDERLINE = 'u';
//	const FONT_CHARSET = 'charset';
//	const FONT_COLOR = 'color';
//	const FONT_CONDENSE = 'condense';
//	const FONT_EXTEND = 'extend';
//	const FONT_FAMILY = 'family';
//	const FONT_OUTLINE = 'outline';
//	const FONT_SCHEME = 'scheme';
//	const FONT_SHADOW = 'shadow';
//	const FONT_STRIKE = 'strike';
//	const FONT_SIZE = 'sz';
//	const FONT_NAME = 'name';
//
//	const FONT_SCHEME_MINOR = 'minor';
//	const FONT_SCHEME_MAJOR = 'major';
//	const FONT_SCHEME_NONE = 'none';
//
//	const FONT_UNDERLINE_SINGLE = 'single';
//	const FONT_UNDERLINE_DOUBLE = 'double';
//	const FONT_UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting';
//	const FONT_UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting';
//
//	const NUM_FMT = 'numFmt';
//
//	const ALIGNMENT_HORIZONTAL = 'horizontal';
//	const ALIGNMENT_HORIZONTAL_LEFT = 'left';
//	const ALIGNMENT_HORIZONTAL_CENTER = 'center';
//	const ALIGNMENT_HORIZONTAL_RIGHT = 'right';
//	const ALIGNMENT_HORIZONTAL_FILL = 'fill';
//	const ALIGNMENT_HORIZONTAL_JUSTIFY = 'justify';
//	const ALIGNMENT_HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous';
//	const ALIGNMENT_HORIZONTAL_DISTRIBUTED = 'distributed';
//
//	const ALIGNMENT_VERTICAL = 'vertical';
//	const ALIGNMENT_VERTICAL_TOP = 'top';
//	const ALIGNMENT_VERTICAL_CENTER = 'center';
//	const ALIGNMENT_VERTICAL_BOTTOM = 'bottom';
//	const ALIGNMENT_VERTICAL_DISTRIBUTED = 'distributed';
//	const ALIGNMENT_VERTICAL_JUSTIFY = 'justify';
//
//	const ALIGNMENT_WRAP_TEXT = 'wrapText';
//	const ALIGNMENT_SHRINK_TO_FIT = 'shrinkToFit';
//	const ALIGNMENT_INDENT = 'indent';
//
//	const ALIGNMENT_READING_ORDER = 'readingOrder';
//	const ALIGNMENT_READING_ORDER_LTR = '1';
//	const ALIGNMENT_READING_ORDER_RTL = '2';
//
//	const ALIGNMENT_TEXT_ROTATION = 'textRotation';
//	const ALIGNMENT_TEXT_ROTATION_VERTICAL = '255';