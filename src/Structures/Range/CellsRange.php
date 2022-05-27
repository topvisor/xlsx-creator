<?php

namespace Topvisor\XlsxCreator\Structures\Range;

use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;

class CellsRange extends Range{

	/**
	 * CellsRange constructor.
	 *
	 * @param int $row1 - номер первой строки
	 * @param int $col1 - номер первого столбца
	 * @param int $row2 - номер второй строки
	 * @param int $col2 - номер второго столбца
	 * @throws InvalidValueException
	 */
	function __construct(int $row1, int $col1, int $row2, int $col2){
		parent::__construct($row1, $col1, $row2, $col2);
	}

	/**
	 * @return string
	 * @throws InvalidValueException
	 */
	function __toString(){
		return Cell::genColStr((int)$this->getTopLeftCol()) . $this->getTopLeftRow() . ':' .
			Cell::genColStr((int)$this->getBottomRightCol()) . $this->getBottomRightRow();
	}
}
