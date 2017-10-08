<?php

namespace XlsxCreator;

use Exception;
use XlsxCreator\Exceptions\InvalidValueException;

class Validator{
	/**
	 * Проверка, присутствует ли значение в массиве
	 *
	 * @param $val - проверяемое значение
	 * @param string $varName - название проверяемой переменной
	 * @param array $validArray - массив правильных значений
	 * @throws InvalidValueException
	 */
	static function validate($val, string $varName, array $validArray){
		if (!in_array($val, $validArray)) throw new InvalidValueException(self::genMustBeInErrorMessage($varName, $validArray));
	}

	/**
	 * Проверка, является ли строка адресом ячейки
	 *
	 * @param string $address - адрес ячейки
	 * @throws InvalidValueException
	 */
	static function validateAddress(string $address){
		if (!preg_match('^[A-Z]{1,3}(\d{1,5})$', $address, $matches)) throw new InvalidValueException('Unavailable address format');
		if ($matches[1] < 1 || $matches[1] > 1048576) throw new InvalidValueException('Row is out of bounds. Excel supports rows from 1 to 1048576');
	}

	/**
	 * Проверка на положительность
	 *
	 * @param $positive - проверяемое число
	 * @param string $varName - имя переменной
	 * @throws InvalidValueException
	 */
	static function validatePositive($positive, string $varName){
		if ($positive < 0) throw new InvalidValueException("$varName must be a positive");
	}

	/**
	 * Проверка на поподание в диапазон
	 *
	 * @param $numeric - число
	 * @param $from - нижняя граница
	 * @param $to - верхняя граница
	 * @param string $varName - имя переменной
	 * @throws InvalidValueException
	 */
	static function validateInRange($numeric, $from, $to, string $varName){
		if ($numeric < $from || $numeric > $to) throw new InvalidValueException("$varName must be in [$from;$to]");
	}

	/**
	 * @param string $var - название переменной
	 * @param array $in - массив с правильными значениями
	 * @return string - сообщение об ошибке
	 */
	static function genMustBeInErrorMessage(string $var, array $in) : string{
		return "$var must be in ['" . implode("','", $in) . "']";
	}
}