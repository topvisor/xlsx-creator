<?php

namespace Topvisor\XlsxCreator\Structures\Drawing\Pic;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class BlipFillXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('xdr:blipFill');

		$xml->startElement('a:blip');
		$xml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$xml->writeAttribute('r:embed', $model['rId']);

		$xml->startElement('a:extLst');

		$xml->startElement('a:ext');
		$xml->writeAttribute('uri', '{28A0092B-C50C-407E-A947-70E740481C1C}');

		$xml->startElement('a14:useLocalDpi');
		$xml->writeAttribute('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main');
		$xml->writeAttribute('val', 0);
		$xml->endElement();

		$xml->endElement();
		$xml->endElement();

		$xml->endElement();

		$xml->startElement('a:stretch');
		$xml->writeElement('a:fillRect');
		$xml->endElement();

		$xml->endElement();
	}
}