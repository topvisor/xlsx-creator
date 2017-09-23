<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Index;

class StylesNumFmtIndex extends StylesIndex{
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

	function addIndex($model) : int{
		if (isset(StylesNumFmtIndex::DEFAULT_NUM_FMT[$model])) return StylesNumFmtIndex::DEFAULT_NUM_FMT[$model];
		if (isset($this->indexes[$model])) return $this->indexes[$model];

		$index = StylesNumFmtIndex::NUM_FMT_START_INDEX + count($this->xmls);
		$this->xmls[] = $this->baseXml->toXml(['id' => $index, 'formatCode' => $model]);

		return $index;
	}
}