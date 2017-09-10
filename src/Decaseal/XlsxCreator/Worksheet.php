<?php

namespace Decaseal\XlsxCreator;

class Worksheet{
	private $id;
	private $name;
	private $tabColor;
	private $defaultRowHeight;

	private $committed;
	private $rows;
	private $startedData;

	private $rId;

	function __construct(int $id, string $name, string $tabColor = null, int $defaultRowHeight = 15){
		$this->id = $id;
		$this->name = $name;
		$this->tabColor = $tabColor;
		$this->defaultRowHeight = $defaultRowHeight;

		$this->committed = false;
		$this->rows = [];
		$this->startedData = false;
	}

	function getId() : int{
		return $this->id;
	}

	function getName() : string{
		return $this->name;
	}

	function isCommitted() : bool{
		return $this->committed;
	}

	function setRId(string $rId){
		$this->rId = $rId;
	}

	function getRId() : string{
		return $this->rId ?? '';
	}

	function addRow($values = null) : Row{
		$row = new Row($this, count($this->rows) + 1);
		if (!is_null($values)) $row->setValues($values);

		$this->rows[] = $row;

		return $row;
	}

	function commit(){
		if ($this->isCommitted()) return;
		$this->committed = true;
	}

	private function writeRows(){ ###
		if (!$this->startedData) {
			$this->writeColumns();
			$this->writeOpenSheetData();
			$this->startedData = true;
		}

		foreach ($this->rows as $row) {
//			if ($row->hasValues())
		}
	}

	private function writeColumns(){

	}

	private function writeOpenSheetData(){

	}
}