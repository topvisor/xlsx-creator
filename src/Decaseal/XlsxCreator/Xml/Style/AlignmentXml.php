<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\XlsxCreator;
use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class AlignmentXml extends BaseXml{
	const TAG = 'alignment';

	function render(XMLWriter $xml, $model = null){
		if (is_null($model)) return;

		$xml->startElement(AlignmentXml::TAG);

		if (isset($model[XlsxCreator::ALIGNMENT_WRAP_TEXT]) && $model[XlsxCreator::ALIGNMENT_WRAP_TEXT]) $model[XlsxCreator::ALIGNMENT_WRAP_TEXT] = '1';
		if (isset($model[XlsxCreator::ALIGNMENT_SHRINK_TO_FIT]) && $model[XlsxCreator::ALIGNMENT_SHRINK_TO_FIT]) $model[XlsxCreator::ALIGNMENT_SHRINK_TO_FIT] = '1';

		foreach ($model as $name => $value) if ($value) $xml->writeAttribute($name, $value);

		$xml->endElement();
 	}
}