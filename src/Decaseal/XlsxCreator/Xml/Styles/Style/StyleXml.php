<?php

namespace Decaseal\XlsxCreator\Xml\Styles\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class StyleXml extends BaseXml{
	private $isXfId;

	public function __construct(bool $isXfId = true){
		$this->isXfId = $isXfId;
	}

	function render(XMLWriter $xml, array $model = null){
		$model = $model ?? [];

		$xml->startElement('xf');

		$xml->writeAttribute('numFmtId', $model['numFmtId'] ?? 0);
		$xml->writeAttribute('fontId', $model['fontId'] ?? 0);
		$xml->writeAttribute('fillId', $model['fillId'] ?? 0);
		$xml->writeAttribute('borderId', $model['borderId'] ?? 0);

		if ($this->isXfId) $xml->writeAttribute('xfId', $model['xfId'] ?? 0);

		if (isset($model['numFmtId'])) $xml->writeAttribute('applyNumberFormat', 1);
		if (isset($model['fontId'])) $xml->writeAttribute('applyFont', 1);
		if (isset($model['fillId'])) $xml->writeAttribute('applyFill', 1);
		if (isset($model['borderId'])) $xml->writeAttribute('applyBorder', 1);

		if ($model['alignment'] ?? false) {
			$xml->writeAttribute('applyAlignment', 1);
			(new AlignmentXml())->render($xml, $model['alignment']);
		}

		$xml->endElement();
	}
}