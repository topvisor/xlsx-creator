<?php

namespace Topvisor\XlsxCreator\Xml\Sheet;

use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class SheetViewsXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		if (!$model || count($model) === 1 && ($model['state'] ?? '') == 'normal') return;

		$xml->startElement('sheetViews');
		$xml->startElement('sheetView');

		$xml->writeAttribute('workbookViewId', $model['workbookViewId'] ?? 0);

		if ($model['rightToLeft'] ?? false) $xml->writeAttribute('rightToLeft', 1);
		if ($model['tabSelected'] ?? false) $xml->writeAttribute('tabSelected', 1);
		if (!($model['showRuler'] ?? true)) $xml->writeAttribute('showRuler', 0);
		if (!($model['showRowColHeaders'] ?? true)) $xml->writeAttribute('showRowColHeaders', 0);
		if ($model['zoomScale'] ?? false) $xml->writeAttribute('zoomScale', $model['zoomScale']);
		if ($model['zoomScaleNormal'] ?? false) $xml->writeAttribute('zoomScaleNormal', $model['zoomScaleNormal']);
		if ($model['view'] ?? false) $xml->writeAttribute('view', $model['view']);

		$xSplit = (int) ($model['xSplit'] ?? 0);
		$ySplit = (int) ($model['ySplit'] ?? 0);
		$topLeftCell = $model['topLeftCell'] ?? null;
		$activeCell = $model['activeCell'] ?? "";
		$activePane = "";

		switch ($model['state'] ?? 'normal') {
			case 'frozen':
				$topLeftCell ??= Cell::genAddress($xSplit + 1, $ySplit + 1);

				switch (true) {
					case ($xSplit && $ySplit): $activePane = 'bottomRight';

					break;
					case $xSplit: $activePane = 'topRight';

					break;
					default: $activePane = 'bottomLeft';

					break;
				}

				$xml->startElement('pane');

				if ($xSplit) $xml->writeAttribute('xSplit', $xSplit);
				if ($ySplit) $xml->writeAttribute('ySplit', $ySplit);

				$xml->writeAttribute('topLeftCell', $topLeftCell);
				$xml->writeAttribute('activePane', $activePane);
				$xml->writeAttribute('state', 'frozen');

				$xml->endElement();

				break;

			case 'split':
				if (($model['activePane'] ?? '') === 'topLeft') $model['activePane'] = false;

				$xml->startElement('pane');

				if ($xSplit) $xml->writeAttribute('xSplit', $xSplit);
				if ($ySplit) $xml->writeAttribute('ySplit', $ySplit);
				if ($topLeftCell) $xml->writeAttribute('topLeftCell', $topLeftCell);
				if ($activePane) $xml->writeAttribute('activePane', $activePane);

				$xml->endElement();

				break;

			default:
				$activePane = "";

				break;
		}

		$xml->startElement('selection');

		if ($activePane) $xml->writeAttribute('pane', $activePane);

		if ($activeCell) {
			$xml->writeAttribute('activeCell', $activeCell);
			$xml->writeAttribute('sqref', $activeCell);
		}

		$xml->endElement();
		$xml->endElement();
		$xml->endElement();
	}
}
