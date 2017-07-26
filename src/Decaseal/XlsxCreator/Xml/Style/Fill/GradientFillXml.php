<?php

namespace Decaseal\XlsxCreator\Xml\Style\Fill;

use Decaseal\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class GradientFillXml extends BaseXml{
	function render(XMLWriter $xml, $model = null){
		if(is_null($model)) return;

		$xml->startElement('gradientFill');

		switch ($model['gradient']) {
			case 'angle':
				$xml->writeAttribute('degree', $model['degree']);
				break;

			case 'path':
				$xml->writeAttribute('type', 'path');

				if ($model['left']) {
					$xml->writeAttribute('left', $model['left']);
					if(!$model['right']) $xml->writeAttribute('right', $model['left']);
				}

				if ($model['top']) {
					$xml->writeAttribute('top', $model['top']);
					if(!$model['bottom']) $xml->writeAttribute('bottom', $model['top']);
				}

				if($model['right']) $xml->writeAttribute('right', $model['right']);

				if($model['bottom']) $xml->writeAttribute('bottom', $model['bottom']);

				break;
		}

		$stopXml = new StopXml();
		foreach ($model['stops'] as $stopModel) $stopXml->render($xml, $stopModel);

		$xml->endElement();
	}
}