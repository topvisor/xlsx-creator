<?php

namespace Decaseal\XlsxCreator\Xml\Style;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class StyleXml extends BaseXml{
	const TAG = 'xf';

	const NUM_FMT_ID = 'numFmtId';
	const FONT_ID = 'fontId';
	const FILL_ID = 'fillId';
	const BORDER_ID = 'borderId';
	const XF_ID = 'xfId';
	const ALIGNMENT = 'alignment';

	const APPLY_NUMBER_FORMAT = 'applyNumberFormat';
	const APPLY_FONT = 'applyFont';
	const APPLY_FILL = 'applyFill';
	const APPLY_BORDER = 'applyBorder';
	const APPLY_ALIGNMENT = 'applyAlignment';

	private $isXfId;

	public function __construct(bool $isXfId = false){
		$this->isXfId = $isXfId;
	}

	function render(XMLWriter $xml, $model = null){
		if (is_null($model)) return;

		$xml->startElement(StyleXml::TAG);

		$xml->writeAttribute(StyleXml::NUM_FMT_ID, (int) ($model[StyleXml::NUM_FMT_ID] ?? 0));
		$xml->writeAttribute(StyleXml::FONT_ID, (int) ($model[StyleXml::FONT_ID] ?? 0));
		$xml->writeAttribute(StyleXml::FILL_ID, (int) ($model[StyleXml::FILL_ID] ?? 0));
		$xml->writeAttribute(StyleXml::BORDER_ID, (int) ($model[StyleXml::BORDER_ID] ?? 0));

		if ($this->isXfId) $xml->writeAttribute(StyleXml::XF_ID, (int) ($model[StyleXml::XF_ID] ?? 0));

		if ($model[StyleXml::NUM_FMT_ID] ?? false) $xml->writeAttribute(StyleXml::APPLY_NUMBER_FORMAT, 1);
		if ($model[StyleXml::FONT_ID] ?? false) $xml->writeAttribute(StyleXml::APPLY_FONT, 1);
		if ($model[StyleXml::FILL_ID] ?? false) $xml->writeAttribute(StyleXml::APPLY_FILL, 1);
		if ($model[StyleXml::BORDER_ID] ?? false) $xml->writeAttribute(StyleXml::APPLY_BORDER, 1);

		if ($model[StyleXml::ALIGNMENT] ?? false) {
			$xml->writeAttribute(StyleXml::APPLY_ALIGNMENT, 1);
			(new AlignmentXml())->render($xml, $model[StyleXml::ALIGNMENT]);
		}

		$xml->endElement();
	}
}