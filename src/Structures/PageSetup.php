<?php

namespace Topvisor\XlsxCreator\Structures;

use Topvisor\XlsxCreator\Validator;

/**
 * Class PageSetup. Параметры печати.
 *
 * @package  Topvisor\XlsxCreator\Structures
 */
class PageSetup{
    const VALID_ORIENTATION = ['portrait', 'landscape'];
    const VALID_PAGE_ORDER = ['downThenOver', 'overThenDown'];
    const VALID_CELL_COMMENTS = ['atEnd', 'asDisplayed', 'None'];
    const VALID_ERRORS = ['dash', 'blank', 'NA', 'displayed'];

    private $model;

    function __construct(){
        $this->model = [
            'margins' => [
                'left' => 0.7,
                'right' => 0.7,
                'top' => 0.75,
                'bottom' => 0.75,
                'header' => 0.3,
                'footer' => 0.3
            ],
            'orientation' => 'portrait',
            'horizontalDpi' => 4294967295,
            'verticalDpi' => 4294967295,
            'fitToPage' => false,
            'blackAndWhite' => false,
            'draft' => false,
            'scale' => 100,
            'fitToWidth' => 1,
            'fitToHeight' => 1,
            'paperSize' => null,
            'showRowColHeaders' => false,
            'showGridLines' => false,
            'horizontalCentered' => false,
            'verticalCentered' => false,
            'rowBreaks' => null,
            'colBreaks' => null
        ];
    }

    /**
     * @return float - отступ от границы страницы
     */
    function getMarginLeft() : float{
        return $this->model['margins']['left'];
    }

    /**
     * @param float $marginLeft - отступ от границы страницы
     * @return PageSetup - $this
     */
    function setMarginLeft(float $marginLeft) : self{
        Validator::validateInRange($marginLeft, 0, 1, '$marginLeft');

        $this->model['margins']['left'] = $marginLeft;
        return $this;
    }

    /**
     * @return float - отступ от границы страницы
     */
    function getMarginRight() : float{
        return $this->model['margins']['right'];
    }

    /**
     * @param float $marginRight - отступ от границы страницы
     * @return PageSetup - $this
     */
    function setMarginRight(float $marginRight) : self{
        Validator::validateInRange($marginRight, 0, 1, '$marginRight');

        $this->model['margins']['right'] = $marginRight;
        return $this;
    }

    /**
     * @return float - отступ от границы страницы
     */
    function getMarginTop() : float{
        return $this->model['margins']['top'];
    }

    /**
     * @param float $marginTop - отступ от границы страницы
     * @return PageSetup - $this
     */
    function setMarginTop(float $marginTop) : self{
        Validator::validateInRange($marginTop, 0, 1, '$marginTop');

        $this->model['margins']['top'] = $marginTop;
        return $this;
    }

    /**
     * @return float - отступ от границы страницы
     */
    function getMarginBottom() : float{
        return $this->model['margins']['bottom'];
    }

    /**
     * @param float $marginBottom - отступ от границы страницы
     * @return PageSetup - $this
     */
    function setMarginBottom(float $marginBottom) : self{
        Validator::validateInRange($marginBottom, 0, 1, '$marginBottom');

        $this->model['margins']['bottom'] = $marginBottom;
        return $this;
    }

    /**
     * @return float - отступ от границы страницы
     */
    function getMarginHeader() : float{
        return $this->model['margins']['header'];
    }

    /**
     * @param float $marginHeader - отступ от границы страницы
     * @return PageSetup - $this
     */
    function setMarginHeader(float $marginHeader) : self{
        Validator::validateInRange($marginHeader, 0, 1, '$marginHeader');

        $this->model['margins']['header'] = $marginHeader;
        return $this;
    }

    /**
     * @return float - отступ от границы страницы
     */
    function getMarginFooter() : float{
        return $this->model['margins']['footer'];
    }

    /**
     * @param float $marginFooter - отступ от границы страницы
     * @return PageSetup - $this
     */
    function setMarginFooter(float $marginFooter) : self{
        Validator::validateInRange($marginFooter, 0, 1, '$marginFooter');

        $this->model['margins']['footer'] = $marginFooter;
        return $this;
    }

    /**
     * @return string - ориентация страницы
     */
    function getOrientation() : string{
        return $this->model['orientation'];
    }

    /**
     * @param string $orientation - ориентация страницы ('portrait', 'landscape')
     * @return PageSetup - $this
     */
    function setOrientation(string $orientation) : self{
        Validator::validate($orientation, '$orientation', self::VALID_ORIENTATION);

        $this->model['orientation'] = $orientation;
        return $this;
    }

    /**
     * @return int - точек на дюйм по горизонтали
     */
    function getHorizontalDpi() : int{
        return $this->model['horizontalDpi'];
    }

    /**
     * @param int $horizontalDpi - точек на дюйм по горизонтали
     * @return PageSetup - $this
     */
    function setHorizontalDpi(int $horizontalDpi) : self{
        Validator::validateInRange($horizontalDpi, 1, 4294967295, '$horizontalDpi');

        $this->model['horizontalDpi'] = $horizontalDpi;
        return $this;
    }

    /**
     * @return int - точек на дюйм по вертикали
     */
    function getVerticalDpi() : int{
        return $this->model['verticalDpi'];
    }

    /**
     * @param int $verticalDpi - точек на дюйм по вертикали
     * @return PageSetup - $this
     */
    function setVerticalDpi(int $verticalDpi) : self{
        Validator::validateInRange($verticalDpi, 1, 4294967295, '$verticalDpi');

        $this->model['verticalDpi'] = $verticalDpi;
        return $this;
    }

    /**
     * @return bool - использовать ли настройки fitToWidth и fitToHeight или scale
     */
    function getFitToPage() : bool{
        return $this->model['fitToPage'];
    }

    /**
     * @param bool $fitToPage - использовать ли настройки fitToWidth и fitToHeight или scale
     * @return PageSetup - $this
     */
    function setFitToPage(bool $fitToPage) : self{
        $this->model['fitToPage'] = $fitToPage;
        return $this;
    }

    /**
     * @return string - порядок печати страниц ('downThenOver', 'overThenDown')
     */
    function getPageOrder() : string{
        return $this->model['pageOrder'];
    }

    /**
     * @param string $pageOrder - порядок печати страниц ('downThenOver', 'overThenDown')
     * @return PageSetup - $this
     */
    function setPageOrder(string $pageOrder) : self{
        Validator::validate($pageOrder, '$pageOrder', self::VALID_ORIENTATION);

        $this->model['pageOrder'] = $pageOrder;
        return $this;
    }

    /**
     * @return bool - печать без цвета
     */
    function getBlackAndWhite() : bool{
        return $this->model['blackAndWhite'];
    }

    /**
     * @param bool $blackAndWhite - печать без цвета
     * @return PageSetup - $this
     */
    function setBlackAndWhite(bool $blackAndWhite) : self{
        $this->model['blackAndWhite'] = $blackAndWhite;
        return $this;
    }

    /**
     * @return bool - печать с меньшим качеством (и чернилами)
     */
    function getDraft() : bool{
        return $this->model['draft'];
    }

    /**
     * @param bool $draft - печать с меньшим качеством (и чернилами)
     * @return PageSetup - $this
     */
    function setDraft(bool $draft) : self{
        $this->model['draft'] = $draft;
        return $this;
    }

    /**
     * @return string - где разместить комментарии ('atEnd', 'asDisplayed', 'None')
     */
    function getCellComments() : string{
        return $this->model['cellComments'];
    }

    /**
     * @param string $cellComments - где разместить комментарии ('atEnd', 'asDisplayed', 'None')
     * @return PageSetup - $this
     */
    function setCellComments(string $cellComments) : self{
        Validator::validate($cellComments, '$cellComments', self::VALID_ORIENTATION);

        $this->model['cellComments'] = $cellComments;
        return $this;
    }

    /**
     * @return string - где показывать ошибки ('dash', 'blank', 'NA', 'displayed')
     */
    function getErrors() : string{
        return $this->model['errors'];
    }

    /**
     * @param string $errors - где показывать ошибки ('dash', 'blank', 'NA', 'displayed')
     * @return PageSetup - $this
     */
    function setErrors(string $errors) : self{
        Validator::validate($errors, '$errors', self::VALID_ERRORS);

        $this->model['errors'] = $errors;
        return $this;
    }

    /**
     * @return int - процент увеличения/уменьшения размеров печати
     */
    function getScale() : int{
        return $this->model['scale'];
    }

    /**
     * @param int $scale - процент увеличения/уменьшения размеров печати
     * @return PageSetup - $this
     */
    function setScale(int $scale) : self{
        Validator::validatePositive($scale, '$scale');

        $this->model['scale'] = $scale;
        return $this;
    }

    /**
     * @return int - сколько страниц должно помещаться на листе по ширине
     */
    function getFitToWidth() : int{
        return $this->model['fitToWidth'];
    }

    /**
     * @param int $fitToWidth - сколько страниц должно помещаться на листе по ширине
     * @return PageSetup - $this
     */
    function setFitToWidth(int $fitToWidth) : self{
        Validator::validatePositive($fitToWidth, '$fitToWidth');

        $this->model['fitToWidth'] = $fitToWidth;
        return $this;
    }

    /**
     * @return int - сколько страниц должно помещаться на листе по высоте
     */
    function getFitToHeight() : int{
        return $this->model['fitToHeight'];
    }

    /**
     * @param int $fitToHeight - сколько страниц должно помещаться на листе по высоте
     * @return PageSetup - $this
     */
    function setFitToHeight(int $fitToHeight) : self{
        Validator::validatePositive($fitToHeight, '$fitToHeight');

        $this->model['fitToHeight'] = $fitToHeight;
        return $this;
    }

    /**
     * @return int - какой размер бумаги использовать (9 - А4)
     */
    function getPaperSize() : int{
        return $this->model['paperSize'];
    }

    /**
     * @param int $paperSize - какой размер бумаги использовать (9 - А4)
     * @return PageSetup - $this
     */
    function setPaperSize(int $paperSize) : self{
        Validator::validatePositive($paperSize, '$paperSize');

        $this->model['paperSize'] = $paperSize;
        return $this;
    }

    /**
     * @return bool - показывать номера строк и столбцов
     */
    function getShowRowColHeaders() : bool{
        return $this->model['showRowColHeaders'];
    }

    /**
     * @param bool $showRowColHeaders - показывать номера строк и столбцов
     * @return PageSetup - $this
     */
    function setShowRowColHeaders(bool $showRowColHeaders) : self{
        $this->model['showRowColHeaders'] = $showRowColHeaders;
        return $this;
    }

    /**
     * @return int - какой номер использовать для первой страницы
     */
    function getFirstPageNumber() : int{
        return $this->model['firstPageNumber'];
    }

    /**
     * @param int $firstPageNumber - какой номер использовать для первой страницы
     * @return PageSetup - $this
     */
    function setFirstPageNumber(int $firstPageNumber) : self{
        Validator::validateInRange($firstPageNumber, 1, PHP_INT_MAX, '$firstPageNumber');

        $this->model['firstPageNumber'] = $firstPageNumber;
        return $this;
    }

    /**
     * @return array - модель
     */
    function getModel() : array{
        return $this->model;
    }
}