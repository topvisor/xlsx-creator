<?php

namespace Topvisor\XlsxCreator\Xml\Comments\Vml;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class ShapeXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('v:shape');

		$xml->writeAttribute('id', '_x0000_s000' . $model['id']);
		$xml->writeAttribute('style', 'visibility:hidden');
		$xml->writeAttribute('fillcolor', '#ffffe1');
		$xml->writeAttribute('type', '#_x0000_t202');

		$xml->startElement('v:fill');
		$xml->writeAttribute('angle', 0);
		$xml->writeAttribute('color2', '#ffffe1');
		$xml->endElement();

		$xml->startElement('v:shadow');
		$xml->writeAttribute('color', 'black');
		$xml->writeAttribute('obscured', 't');
		$xml->writeAttribute('on', 't');
		$xml->endElement();

		$xml->writeElement('v:textbox');

		(new NoteClientDataXml())->render($xml, $model);

		$xml->endElement();
	}
}