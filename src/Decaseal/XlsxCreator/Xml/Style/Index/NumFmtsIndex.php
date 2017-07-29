<?php

namespace Decaseal\XlsxCreator\Xml\Style\Index;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Decaseal\XlsxCreator\Xml\Style\NumFmtXml;

class NumFmtsIndex extends Index{
	const INDEX = 'index';
	const XML = 'xml';
	const DEFAULT_NUM_FMTS = [
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
	const NUM_FMTS_BASE = 164;

	function addIndex(BaseXml $baseXml, $model): int{
		if (isset(NumFmtsIndex::DEFAULT_NUM_FMTS[$model])) {
			$index = NumFmtsIndex::DEFAULT_NUM_FMTS[$model];
		} elseif (isset($this->indexes[$model])) {
			$index = $this->indexes[$model][NumFmtsIndex::INDEX];
		} else {
			$index = count($this->indexes) + NumFmtsIndex::NUM_FMTS_BASE;
			$xml = $baseXml->toXml([NumFmtXml::FORMATE_CODE => $model, NumFmtXml::NUM_FMT_ID => $index]);
			$this->indexes[$model] = [NumFmtsIndex::XML => $xml, NumFmtsIndex::INDEX => $index];
		}

		return $index;
	}

	public function current(){
		return $this->indexes[parent::current()][NumFmtsIndex::XML];
	}

	public function key(){
		return $this->indexes[parent::current()][NumFmtsIndex::INDEX];
	}
}