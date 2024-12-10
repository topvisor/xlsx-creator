<?php

namespace Topvisor\XlsxCreator\Xml\Comments;

use Topvisor\XlsxCreator\Xml\BaseXml;
use Topvisor\XlsxCreator\Xml\Strings\SharedStringXml;
use XMLWriter;

class CommentXml extends BaseXml {
	public function render(XMLWriter $xml, ?array $model = null) {
		$xml->startElement('comment');

		$xml->writeAttribute('authorId', 0);
		$xml->writeAttribute('ref', $model['address']);

		(new SharedStringXml('text'))->render($xml, $model);

		$xml->endElement();
	}
}
