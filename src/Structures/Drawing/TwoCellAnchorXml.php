<?php

namespace Topvisor\XlsxCreator\Structures\Drawing;

use Topvisor\XlsxCreator\Structures\Drawing\Pic\PicXml;
use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class TwoCellAnchorXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		$xml->startElement('xdr:twoCellAnchor');
		$xml->writeAttribute('editAs', 'oneCell');

		(new CellPositionXml('xdr:from'))->render($xml, ['row' => $model['position']['top'], 'col' => $model['position']['left']]);
		(new CellPositionXml('xdr:to'))->render($xml, ['row' => $model['position']['bottom'], 'col' => $model['position']['right']]);

		(new PicXml())->render($xml, $model['pic']);

		$xml->writeElement('xdr:clientData');

		$xml->endElement();
	}
}