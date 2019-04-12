<?php

// autoload_static.php @generated by Composer

namespace Composer\Autoload;

class ComposerStaticInit1af2954bb030c7c2a6d9305a345e674a
{
    public static $prefixLengthsPsr4 = array (
        'Z' => 
        array (
            'Zend\\Escaper\\' => 13,
        ),
        'P' => 
        array (
            'PhpOffice\\PhpWord\\' => 18,
            'PhpOffice\\Common\\' => 17,
        ),
        'H' => 
        array (
            'HTMLtoOpenXML\\' => 14,
        ),
    );

    public static $prefixDirsPsr4 = array (
        'Zend\\Escaper\\' => 
        array (
            0 => __DIR__ . '/..' . '/zendframework/zend-escaper/src',
        ),
        'PhpOffice\\PhpWord\\' => 
        array (
            0 => __DIR__ . '/..' . '/phpoffice/phpword/src/PhpWord',
        ),
        'PhpOffice\\Common\\' => 
        array (
            0 => __DIR__ . '/..' . '/phpoffice/common/src/Common',
        ),
        'HTMLtoOpenXML\\' => 
        array (
            0 => __DIR__ . '/..' . '/rkorebrits/htmltoopenxml/src',
        ),
    );

    public static $classMap = array (
        'PclZip' => __DIR__ . '/..' . '/pclzip/pclzip/pclzip.lib.php',
        'om\\Freq' => __DIR__ . '/..' . '/om/icalparser/src/Freq.php',
        'om\\IcalParser' => __DIR__ . '/..' . '/om/icalparser/src/IcalParser.php',
        'om\\Recurrence' => __DIR__ . '/..' . '/om/icalparser/src/Recurrence.php',
    );

    public static function getInitializer(ClassLoader $loader)
    {
        return \Closure::bind(function () use ($loader) {
            $loader->prefixLengthsPsr4 = ComposerStaticInit1af2954bb030c7c2a6d9305a345e674a::$prefixLengthsPsr4;
            $loader->prefixDirsPsr4 = ComposerStaticInit1af2954bb030c7c2a6d9305a345e674a::$prefixDirsPsr4;
            $loader->classMap = ComposerStaticInit1af2954bb030c7c2a6d9305a345e674a::$classMap;

        }, null, ClassLoader::class);
    }
}