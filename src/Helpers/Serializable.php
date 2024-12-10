<?php

namespace Topvisor\XlsxCreator\Helpers;

/**
 * Interface Serializable. Создан как замена интерфейсу \Serializable, объявленному deprecated в php 8.0
 *
 * @package Topvisor\XlsxCreator\Helpers
 */
interface Serializable {
	public function serialize();
	public function unserialize($data);
}
