<?php

namespace Decaseal\XlsxCreator;

class Worksheet{
	private $id;
	private $name;
	private $tabColor;
	private $defaultRowHeight;

	private $committed;

	function __construct(int $id, string $name, string $tabColor = null, int $defaultRowHeight = 15){
		$this->id = $id;
		$this->name = $name;
		$this->tabColor = $tabColor;
		$this->defaultRowHeight = $defaultRowHeight;

		$this->committed = false;
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

	function commit(){
		if ($this->isCommitted()) return;
	}
}