PHPExcel in CodeIgniter
========

PHPExcel is a pure PHP library for reading and writing spreadsheet files and CodeIgniter is one of the well known PHP MVC Framework.

Step 1: Download and setup CodeIgniter.(download it here: https://www.codeigniter.com/)

Step 2: Download PHPExcel.(download it here: https://github.com/PHPOffice/PHPExcel)

Step 3: Unzip or extract the downloaded PHPExcel lib files and copy Class directory inside files to application/third-party directory(folder).



#### Create one file called Excel.php in  [ application/libraries ]
```php
<?php

if (!defined('BASEPATH')) exit('No direct script access allowed');  
 
require_once APPPATH."/third_party/PHPExcel-1.8/Classes/PHPExcel.php";
require_once APPPATH."/third_party/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php";
 
class Excel extends PHPExcel {
    public function __construct() {
        parent::__construct();
    }
}
?>
```


#### Save Excel.php library [ application->config/autoload.php ]
```php
<?php
defined('BASEPATH') OR exit('No direct script access allowed');  
 
$autoload['libraries'] = array('database','session','form_validation','excel');
?>
``` 
### **Your PHP version must be less than 7.2 to use the phpexcel library. If your PHP version is higher than 7.2, please try <a href="https://phpspreadsheet.readthedocs.io/en/latest/">PhpSpreadsheet.</a>
