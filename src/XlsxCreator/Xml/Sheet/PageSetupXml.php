<?php

namespace XlsxCreator\Xml\Sheet;

use XlsxCreator\Xml\BaseXml;
use XMLWriter;

class PageSetupXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (!$model) return;

		$xml->startElement('pageSetup');

		if ($model['paperSize'] ?? false) $xml->writeAttribute('paperSize', $model['paperSize']);
		if ($model['orientation'] ?? false) $xml->writeAttribute('orientation', $model['orientation']);
		if ($model['horizontalDpi'] ?? false) $xml->writeAttribute('horizontalDpi', $model['horizontalDpi']);
		if ($model['verticalDpi'] ?? false) $xml->writeAttribute('verticalDpi', $model['verticalDpi']);
		if ($model['pageOrder'] ?? false) $xml->writeAttribute('pageOrder', $model['pageOrder']);
		if ($model['blackAndWhite'] ?? false) $xml->writeAttribute('blackAndWhite', $model['blackAndWhite']);
		if ($model['draft'] ?? false) $xml->writeAttribute('draft', $model['draft']);
		if ($model['cellComments'] ?? false) $xml->writeAttribute('cellComments', $model['cellComments']);
		if ($model['errors'] ?? false) $xml->writeAttribute('errors', $model['errors']);
		if ($model['scale'] ?? false) $xml->writeAttribute('scale', $model['scale']);
		if ($model['fitToWidth'] ?? false) $xml->writeAttribute('fitToWidth', $model['fitToWidth']);
		if ($model['fitToHeight'] ?? false) $xml->writeAttribute('fitToHeight', $model['fitToHeight']);
		if ($model['firstPageNumber'] ?? false) $xml->writeAttribute('firstPageNumber', $model['firstPageNumber']);
		if ($model['useFirstPageNumber'] ?? false) $xml->writeAttribute('useFirstPageNumber', $model['useFirstPageNumber']);
		if ($model['usePrinterDefaults'] ?? false) $xml->writeAttribute('usePrinterDefaults', $model['usePrinterDefaults']);
		if ($model['copies'] ?? false) $xml->writeAttribute('copies', $model['copies']);

		$xml->endElement();
	}
}