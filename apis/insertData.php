<?php
header('Access-Control-Allow-Origin: *');
include '../includes/config.php';

if (!isset($_POST) || empty($_POST)) {
	echo json_encode(array('status' => 'bad'));
} else {
	$params = mysql_real_escape_string($_POST['data']);
	$now = date('H:i:s m-d-Y');
	$mysqli->query("INSERT INTO `data`(`data`, `create_at`) VALUES ('{$params}','{$now}')");
	// $insertResult = mysql_query($insertQuery);

	$id = $mysqli->insert_id;
	echo json_encode(array('status' => 'ok', 'msg' => "Successfully Added.", 'id' => $id));
}
mysqli_close($mysqli);
?>