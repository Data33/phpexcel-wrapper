<?php
namespace Data33\ExcelWrapper;

use PHPExcel_Style_Border;
use Data33\ExcelWrapper\Exceptions\ExcelException;

class ExcelStyle {
    static $defaultStyles = [
        'default' => [
            'font' => [
                'size' => 10,
                'name' => 'Arial'
            ]
        ],
        'title' => [
            'font' => [
                'size' => 16,
                'name' => 'Calibri',
                'bold' => true
            ]
        ],
        'header' => [
            'font' => [
                'size' => 11,
                'name' => 'Calibri',
                'bold' => true
            ],
            'borders' => [
                'bottom' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN
                ]
            ]
        ]
    ];

    /**
     * Add a style to the document or modify an existing one
     *
     * @throws ExcelException
     *
     * @param string $name The name of the new style. If using an existing name that style will be overwritten
     *
     * @param array $style An array containing PHPExcel style formatting directives
     */
    public static function setStyle($name, array $style){
        if (!is_string($name)){
            throw new ExcelException('The supplied style name is invalid!', 0);
        }

        self::$defaultStyles[$name] = $style;
    }

    /**
     * Fetch style array from name
     *
     * @param string $name The name of the new style. If using an existing name that style will be overwritten
     *
     * @param array $style An array containing PHPExcel style formatting directives
     */
    public static function style($style){
        $style = strtolower($style);
        return isset(self::$defaultStyles[$style]) ? self::$defaultStyles[$style] : self::$defaultStyles['default'];
    }
} 