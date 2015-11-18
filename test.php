<?php

header('Access-Control-Allow-Origin: *');
$params = mysql_real_escape_string($_POST['data']);
echo $params;

?>