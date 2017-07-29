<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class NumFmtXml extends BaseXml{
	const TAG = 'numFmt';

	const NUM_FMT_ID = 'numFmtId';
	const FORMATE_CODE = 'formatCode';

	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement(NumFmtXml::TAG);

		$xml->writeAttribute(NumFmtXml::NUM_FMT_ID, $model[NumFmtXml::NUM_FMT_ID]);
		$xml->writeAttribute(NumFmtXml::FORMATE_CODE, $model[NumFmtXml::FORMATE_CODE]);

		$xml->endElement();
	}
}