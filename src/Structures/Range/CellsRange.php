<?php

namespace Topvisor\XlsxCreator\Structures\Range;

use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;

class CellsRange extends Range{
	function __construct(int $row1, int $col1, int $row2, int $col2){
		parent::__construct($row1, $col1, $row2, $col2);
	}

	function __toString(){
		return Cell::genColStr($this->left) . $this->top . ':' . Cell::genColStr($this->right) . $this->bottom;
	}
}