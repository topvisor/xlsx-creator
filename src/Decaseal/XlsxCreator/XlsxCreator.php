<?php

namespace Decaseal\XlsxCreator;

use DateTime;

class XlsxCreator{
	const BORDER_LEFT = 'left';
	const BORDER_RIGHT = 'right';
	const BORDER_TOP = 'top';
	const BORDER_BOTTOM = 'bottom';

	const BORDER_DIAGONAL = 'diagonal';
	const BORDER_DIAGONAL_UP = 'up';
	const BORDER_DIAGONAL_DOWN = 'down';

	const BORDER_COLOR = 'color';
	const BORDER_STYLE = 'style';
	
	const BORDER_STYLE_THIN = 'thin';
	const BORDER_STYLE_DOTTED = 'dotted';
	const BORDER_STYLE_DASHDOT = 'dashdot';
	const BORDER_STYLE_HAIR = 'hair';
	const BORDER_STYLE_DASHDOTDOT = 'dashdotdot';
	const BORDER_STYLE_SLANTDASHDOT = 'slantdashdot';
	const BORDER_STYLE_MEDIUMDASHED = 'mediumdashed';
	const BORDER_STYLE_MEDIUMDASHDOTDOT = 'mediumdashdotdot';
	const BORDER_STYLE_MEDIUMDASHDOT = 'mediumdashdot';
	const BORDER_STYLE_MEDIUM = 'medium';
	const BORDER_STYLE_DOUBLE = 'double';
	const BORDER_STYLE_THICK = 'thick';

	const FILL_TYPE = 'type';

	const FILL_GRADIENT = 'gradient';
	const FILL_GRADIENT_ANGLE = 'angle';
	const FILL_GRADIENT_PATH = 'path';
	const FILL_DEGREE = 'degree';
	const FILL_CENTER = 'center';
	const FILL_CENTER_LEFT = 'left';
	const FILL_CENTER_RIGHT = 'right';
	const FILL_CENTER_TOP = 'top';
	const FILL_CENTER_BOTTOM = 'bottom';
	const FILL_STOPS = 'stops';
	const FILL_STOP_POSITION = 'position';
	const FILL_STOP_COLOR = 'color';

	const FILL_PATTERN = 'pattern';
	const FILL_FG_COLOR = 'fgColor';
	const FILL_BG_COLOR = 'bgColor';

	const FILL_PATTERN_NONE = 'none';
	const FILL_PATTERN_SOLID = 'solid';
	const FILL_PATTERN_DARK_GRAY = 'darkGray';
	const FILL_PATTERN_MEDIUM_GRAY = 'mediumGray';
	const FILL_PATTERN_LIGHT_GRAY = 'lightGray';
	const FILL_PATTERN_GRAY_125 = 'gray125';
	const FILL_PATTERN_GRAY_0625 = 'gray0625';
	const FILL_PATTERN_DARK_HORIZONTAL = 'darkHorizontal';
	const FILL_PATTERN_DARK_VERTICAL = 'darkVertical';
	const FILL_PATTERN_DARK_DOWN = 'darkDown';
	const FILL_PATTERN_DARK_UP = 'darkUp';
	const FILL_PATTERN_DARK_GRID = 'darkGrid';
	const FILL_PATTERN_DARK_TRELLIS = 'darkTrellis';
	const FILL_PATTERN_LIGHT_HORIZONTAL = 'lightHorizontal';
	const FILL_PATTERN_LIGHT_VERTICAL = 'lightVertical';
	const FILL_PATTERN_LIGHT_DOWN = 'lightDown';
	const FILL_PATTERN_LIGHT_UP = 'lightUp';
	const FILL_PATTERN_LIGHT_TRELLIS = 'lightTrellis';
	const FILL_PATTERN_LIGHT_GRID = 'lightGrid';

	const COLOR_ARGB = 'argb';

	const FONT_BOLD = 'b';
	const FONT_ITALIC = 'i';
	const FONT_UNDERLINE = 'u';
	const FONT_CHARSET = 'charset';
	const FONT_COLOR = 'color';
	const FONT_CONDENSE = 'condense';
	const FONT_EXTEND = 'extend';
	const FONT_FAMILY = 'family';
	const FONT_OUTLINE = 'outline';
	const FONT_SCHEME = 'scheme';
	const FONT_SHADOW = 'shadow';
	const FONT_STRIKE = 'strike';
	const FONT_SIZE = 'sz';
	const FONT_NAME = 'name';

	const FONT_SCHEME_MINOR = 'minor';
	const FONT_SCHEME_MAJOR = 'major';
	const FONT_SCHEME_NONE = 'none';

	const FONT_UNDERLINE_SINGLE = 'single';
	const FONT_UNDERLINE_DOUBLE = 'double';
	const FONT_UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting';
	const FONT_UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting';

	const NUM_FMT = 'numFmt';


	private $stream;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $lastPrinted;

	function __construct(resource &$stream, DateTime $created = null, DateTime $modified = null, string $creator = null, string $lastModifiedBy = null, DateTime $lastPrinted = null){
		$this->stream = $stream;
		$this->created = $created ?? new DateTime();
		$this->modified = $modified ?? $this->created;
		$this->creator = $creator ?? 'XlsxWriter';
		$this->lastModifiedBy = $lastModifiedBy ?? $this->creator;
		$this->lastPrinted = $lastPrinted;
	}
}