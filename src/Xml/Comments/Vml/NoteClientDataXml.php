<?php

namespace Topvisor\XlsxCreator\Xml\Comments\Vml;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class NoteClientDataXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model) return;

		$xml->startElement('x:ClientData');
		$xml->writeAttribute('ObjectType', 'Note');

		$xml->writeElement('x:MoveWithCells');
		$xml->writeElement('x:SizeWithCells');
		$xml->writeElement('x:Anchor', implode(', ', [
			$model['col'],
			15,
			$model['row'],
			10,
			3 + $model['col'] + $model['width'],
			15,
			1 + $model['row'] + $model['height'],
			4,
		]));
		$xml->writeElement('x:AutoFill', 'False');
		$xml->writeElement('x:Row', $model['row'] - 1);
		$xml->writeElement('x:Column', $model['col'] - 1);

		$xml->endElement();
	}
}
