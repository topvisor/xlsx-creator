<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class AlignmentXml extends BaseXml{
	const TAG = 'alignment';

	function render(XMLWriter $xml, $model = null){
		if (is_null($model)) return;

		$xml->startElement(AlignmentXml::TAG);



		$xml->endElement();
 	}
}