<?php

ini_set('display_errors', '1');
ini_set('display_startup_errors', '1');
ini_set('error_reporting', E_ALL);
ini_set('log_errors','on');
ini_set('error_log',__DIR__.'/errors/errors.log');

const __ROOT = __DIR__.'/../public';
const PATH_TO_FILES = __DIR__;

$loader = require_once __ROOT.'/vendor/autoload.php';

$loader->addPsr4('Model\\', __ROOT.'/classes/Model');
$loader->addPsr4('Core\\', __ROOT.'/classes/Core');
$loader->addPsr4('Parsers\\', __DIR__.'/classes/Parsers');
$loader->addPsr4('Exceptions\\', __DIR__.'/classes/Exceptions');

use Websm\Framework\Db\Config;
use Parsers\Room\RoomParser;

Config::init(include __ROOT.'/admin/config.php');

$roomParser = new RoomParser(
    PATH_TO_FILES.'/import.xlsx'
);

$roomParser->parse();
