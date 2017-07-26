<?php

namespace Decaseal\XlsxCreator;

use DateTime;

class WorkbookWriter{
	private $stream;
	private $created;
	private $modified;
	private $creator;
	private $lastModifiedBy;
	private $lastPrinted;

	private $is_commited = false;

	function __construct(resource &$stream, DateTime $created = null, DateTime $modified = null, string $creator = null, string $lastModifiedBy = null, DateTime $lastPrinted = null){
		$this->stream = $stream;
		$this->created = $created ?? new DateTime();
		$this->modified = $modified ?? $this->created;
		$this->creator = $creator ?? 'XlsxWriter';
		$this->lastModifiedBy = $lastModifiedBy ?? $this->creator;
		$this->lastPrinted = $lastPrinted;
	}

	function getCreated(): DateTime{
		return $this->created;
	}

	function getModified(): DateTime{
		return $this->modified;
	}

	function getCreator(): string{
		return $this->creator;
	}

	function getLastModifiedBy(): string{
		return $this->lastModifiedBy;
	}

	function getLastPrinted(){
		return $this->lastPrinted;
	}

	function setCreated(DateTime $created){
		if ($this->is_commited) throw new ElementIsCommittedException('Workbook is committed');

		$this->created = $created;
	}

	function setModified(DateTime $modified){
		if ($this->is_commited) throw new ElementIsCommittedException('Workbook is committed');
		if ($modified < $this->created) throw new WrongValueException('The $modified must be less than $created');

		$this->modified = $modified;
	}

	function setCreator(string $creator){
		if ($this->is_commited) throw new ElementIsCommittedException('Workbook is committed');

		$this->creator = $creator;
	}

	function setLastModifiedBy(string $lastModifiedBy){
		if ($this->is_commited) throw new ElementIsCommittedException('Workbook is committed');

		$this->lastModifiedBy = $lastModifiedBy;
	}

	function setLastPrinted(DateTime $lastPrinted = null){
		if ($this->is_commited) throw new ElementIsCommittedException('Workbook is committed');

		$this->lastPrinted = $lastPrinted;
	}
}