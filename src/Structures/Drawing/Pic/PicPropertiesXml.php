<?php

namespace Topvisor\XlsxCreator\Structures\Drawing\Pic;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class PicPropertiesXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('xdr:nvPicPr');

		$xml->startElement('xdr:cNvPr');
		$xml->writeAttribute('id', $model['id']);
		$xml->writeAttribute('name', $model['name']);
		$xml->endElement();

		$xml->startElement('xdr:cNvPicPr');
		$xml->startElement('a:picLocks');

		$xml->writeAttribute('noChangeAspect', 1);

		$xml->endElement();
		$xml->endElement();

		$xml->endElement();
	}
}