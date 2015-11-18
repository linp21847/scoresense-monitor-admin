<?php

header('Access-Control-Allow-Origin: *');
var_dump($_POST['data']);
$params = mysql_real_escape_string($_POST['data']);
echo $params;

?>