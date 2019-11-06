<?php

namespace Exceptions;

use Exception;

class MainException extends Exception
{
    public function printError(){
        print_r("Error:\n");
        print_r($this->getMessage()."\n");
        print_r("Trace:\n");
        print_r($this->getTraceAsString());
    }
}
