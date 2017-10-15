<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Values\RichText\RichTextValue;
use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Worksheet;
use Topvisor\XlsxCreator\Xml\Strings\SharedStringXml;
use XMLWriter;

class Comments{
	private $worksheet;

	private $empty;
	private $committed;
	private $nextShapeId;

	private $commentsFilename;
	private $commentsXml;

	private $vmlFilename;
	private $vmlXml;

	function __construct(Worksheet $worksheet){
		$this->worksheet = $worksheet;

		$this->empty = true;
		$this->committed = false;
		$this->nextShapeId = 1;
	}

	public function __destruct(){
		unset($this->worksheet);
		unset($this->commentsXml);
		unset($this->vmlXml);

		if ($this->commentsFilename && file_exists($this->commentsFilename)) unlink($this->commentsFilename);
		if ($this->vmlFilename && file_exists($this->vmlFilename)) unlink($this->vmlFilename);
	}

	function getCommentsFilename(){
		if ($this->commentsXml) $this->commentsXml->flush();
		return $this->commentsFilename;
	}

	function getVmlFilename(){
		if ($this->vmlXml) $this->vmlXml->flush();
		return $this->vmlFilename;
	}

	/**
	 * @param Cell $cell - ячейка
	 */
	function addComment(Cell $cell){
		if ($this->empty) $this->startComments();

		$this->addCommentToComments($cell->getAddress(), $cell->getComment());
		$this->addCommentToVml($cell->getCol(), $cell->getRow()->getNumber(), $cell->getCommentWidth(), $cell->getCommentHeight());
	}

	function isEmpty() : bool{
		return $this->empty;
	}

	function commit(){
		$this->checkCommitted();
		$this->committed = true;

		$this->endComments();
	}

	private function addCommentToComments(string $address, Value $comment){
		if (!$comment) return;

		$this->commentsXml->startElement('comment');

		$this->commentsXml->writeAttribute('authorId', 0);
		$this->commentsXml->writeAttribute('ref', $address);

		(new SharedStringXml('text'))->render($this->commentsXml, [
			'type' => $comment->getType(),
			'value' => $comment->getValue()
		]);

		$this->commentsXml->endElement();
	}

	private function addCommentToVml(int $col, int $row, int $width, int $height){
		$this->vmlXml->startElement('v:shape');

		$this->vmlXml->writeAttribute('id', '_x0000_s000' . $this->nextShapeId++);
		$this->vmlXml->writeAttribute('style', 'visibility:hidden');
		$this->vmlXml->writeAttribute('fillcolor', '#ffffe1');
		$this->vmlXml->writeAttribute('type', '#_x0000_t202');

		$this->vmlXml->startElement('v:fill');
		$this->vmlXml->writeAttribute('angle', 0);
		$this->vmlXml->writeAttribute('color2', '#ffffe1');
		$this->vmlXml->endElement();

		$this->vmlXml->startElement('v:shadow');
		$this->vmlXml->writeAttribute('color', 'black');
		$this->vmlXml->writeAttribute('obscured', 't');
		$this->vmlXml->writeAttribute('on', 't');
		$this->vmlXml->endElement();

		$this->vmlXml->writeElement('v:textbox');

		$this->vmlXml->startElement('x:ClientData');

		$this->vmlXml->writeElement('x:MoveWithCells');
		$this->vmlXml->writeElement('x:SizeWithCells');
		$this->vmlXml->writeElement('x:Anchor', implode(', ', [$col, 15, $row, 10, 3 + $col + $width, 15, 1 + $row + $height, 4]));
		$this->vmlXml->writeElement('x:AutoFill', 'False');
		$this->vmlXml->writeElement('x:Row', $row - 1);
		$this->vmlXml->writeElement('x:Column', $col - 1);

		$this->vmlXml->endElement();
		$this->vmlXml->endElement();
	}

	private function startComments(){
		$this->empty = false;

		$this->startCommentsXml();
		$this->startVmlXml();
	}

	private function startCommentsXml(){
		$this->commentsFilename = $this->worksheet->getWorkbook()->genTempFilename();

		$this->commentsXml = new XMLWriter();
		$this->commentsXml->openURI($this->commentsFilename);

		$this->commentsXml->startDocument('1.0', 'UTF-8', 'yes');

		$this->commentsXml->startElement('comments');
		$this->commentsXml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$this->commentsXml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

		$this->commentsXml->writeElement('authors', '<author></author>');

		$this->commentsXml->startElement('commentList');
	}

	private function startVmlXml(){
		$this->vmlFilename = $this->worksheet->getWorkbook()->genTempFilename();

		$this->vmlXml = new XMLWriter();
		$this->vmlXml->openURI($this->vmlFilename);

		$this->vmlXml->startDocument('1.0', 'UTF-8', 'yes');

		$this->vmlXml->startElement('xml');
		$this->vmlXml->writeAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml');
		$this->vmlXml->writeAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office');
		$this->vmlXml->writeAttribute('xmlns:x', 'urn:schemas-microsoft-com:office:excel');

		$this->vmlXml->startElement('o:shapelayout');

		$this->vmlXml->startElement('o:idmap');
		$this->vmlXml->writeAttribute('data', 1);
		$this->vmlXml->writeAttribute('v:ext', 'edit');
		$this->vmlXml->endElement();

		$this->vmlXml->endElement();

		$this->vmlXml->startElement('v:shapetype');
		$this->vmlXml->writeAttribute('id', '_x0000_t202');
		$this->vmlXml->writeAttribute('coordsize', '21600,21600');
		$this->vmlXml->writeAttribute('path', 'm,l,21600r21600,l21600,xe');
		$this->vmlXml->endElement();
	}

	private function endComments(){
		if ($this->empty) return;

		$this->commentsXml->endElement();
		$this->commentsXml->endElement();
		$this->commentsXml->endDocument();

		$this->commentsXml->flush();
		unset($this->commentsXml);

		$this->vmlXml->endElement();
		$this->vmlXml->endDocument();

		$this->vmlXml->flush();
		unset($this->vmlXml);
	}

	private function checkCommitted(){
		if ($this->committed) throw new ObjectCommittedException();
	}
}