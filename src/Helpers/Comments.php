<?php

namespace Topvisor\XlsxCreator\Helpers;

use Topvisor\XlsxCreator\Cell;
use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Exceptions\ObjectCommittedException;
use Topvisor\XlsxCreator\Structures\Values\RichText\RichTextValue;
use Topvisor\XlsxCreator\Structures\Values\Value;
use Topvisor\XlsxCreator\Worksheet;
use Topvisor\XlsxCreator\Xml\Comments\CommentXml;
use Topvisor\XlsxCreator\Xml\Comments\Vml\ShapeXml;
use Topvisor\XlsxCreator\Xml\Strings\SharedStringXml;
use XMLWriter;

/**
 * Class Comments. Комментарии таблицы
 *
 * @package Topvisor\XlsxCreator\Helpers
 */
class Comments{
	private $worksheet;

	private $empty;
	private $committed;
	private $nextShapeId;

	private $commentsFilename;
	private $commentsXml;

	private $vmlFilename;
	private $vmlXml;

	/**
	 * Comments constructor.
	 *
	 * @param Worksheet $worksheet - таблица
	 */
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

	/**
	 * @return string|null - имя файла с текстом комментариев
	 */
	function getCommentsFilename(){
		if ($this->commentsXml ?? false) $this->commentsXml->flush();
		return $this->commentsFilename;
	}

	/**
	 * @return string|null - имя файла с положением и размерами комментариев
	 */
	function getVmlFilename(){
		if ($this->vmlXml ?? false) $this->vmlXml->flush();
		return $this->vmlFilename;
	}

	/**
	 * @param Cell $cell - ячейка
	 */
	function addComment(Cell $cell){
		if ($this->empty) $this->startComments();

		(new CommentXml())->render($this->commentsXml, [
			'address' => $cell->getAddress(),
			'type' => $cell->getComment()->getType(),
			'value' => $cell->getComment()->getValue()
		]);

		(new ShapeXml())->render($this->vmlXml, [
			'id' => $this->nextShapeId++,
			'col' => $cell->getCol(),
			'row' => $cell->getRow()->getNumber(),
			'width' => $cell->getCommentWidth(),
			'height' => $cell->getCommentHeight()
		]);
	}

	/**
	 * @return bool - комментариев нет
	 */
	function isEmpty() : bool{
		return $this->empty;
	}

	/**
	 * Зафиксировать комментарии.
	 */
	function commit(){
		$this->checkCommitted();
		$this->committed = true;

		$this->endComments();
	}

	/**
	 * Начать файлы комментариев.
	 */
	private function startComments(){
		$this->empty = false;

		$this->startCommentsXml();
		$this->startVmlXml();
	}

	/**
	 * Начать файл с текстом комментариев.
	 */
	private function startCommentsXml(){
		$this->commentsFilename = $this->worksheet->getWorkbook()->genTempFilename();

		$this->commentsXml = new XMLWriter();
		$this->commentsXml->openURI($this->commentsFilename);

		$this->commentsXml->startDocument('1.0', 'UTF-8', 'yes');

		$this->commentsXml->startElement('comments');
		$this->commentsXml->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$this->commentsXml->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

		$this->commentsXml->startElement('authors');
		$this->commentsXml->writeElement('author');
		$this->commentsXml->endElement();

		$this->commentsXml->startElement('commentList');
	}

	/**
	 * Начать файл с положением и размерами комментариев.
	 */
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

	/**
	 * Закончить файлы комментариев.
	 */
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

	/**
	 * @throws ObjectCommittedException
	 */
	private function checkCommitted(){
		if ($this->committed) throw new ObjectCommittedException();
	}
}