<?php

namespace Topvisor\XlsxCreator\Xml\Core\Relationships;

use Topvisor\XlsxCreator\Xml\BaseXml;
use XMLWriter;

class RelationshipXml extends BaseXml{
	function render(XMLWriter $xml, array $model = null){
		if (is_null($model)) return;

		$xml->startElement('Relationship');

		$xml->writeAttribute('Id', $model['id']);
		$xml->writeAttribute('Type', $model['type']);
		$xml->writeAttribute('Target', $model['target']);
		if ($model['targetMode'] ?? false) $xml->writeAttribute('TargetMode', $model['targetMode']);

		foreach ($model as $name => $value) $xml->writeAttribute($name, $value);

		$xml->endElement();
	}
}