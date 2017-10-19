<?php

namespace Topvisor\XlsxCreator\Structures\Drawing\Pic;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class PicXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		$xml->startElement('xdr:pic');

		(new PicPropertiesXml())->render($xml, $model);
		(new BlipFillXml())->render($xml, $model);

		$xml->startElement('xdr:spPr');
		$xml->startElement('a:xfrm');
		$xml->startElement('a:off');

		$xml->writeAttribute('x', 0);
		$xml->writeAttribute('y', 0);

		$xml->endElement();

		$xml->startElement('a:ext');
		$xml->writeAttribute('cx', 2057400);
		$xml->writeAttribute('cy', 528034);
		$xml->endElement();

		$xml->endElement();

		$xml->startElement('a:prstGeom');
		$xml->writeAttribute('prst', 'rect');

		$xml->writeElement('a:avLst');

		$xml->endElement();
		$xml->endElement();
		$xml->endElement();
	}
}