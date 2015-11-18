<?php
header('Access-Control-Allow-Origin: *');
include '../includes/config.php';

if (!isset($_GET) || !empty($_GET)) {
	$id = $_GET['id'];
} else if (!isset($_POST) || !empty($_POST)) {
	echo json_encode(array('status' => 'bad'));
	exit;
} else {
	$id = $_POST['id'];
}
	$result = mysqli_query($mysqli, "SELECT `data` FROM `data` WHERE `id`=$id");
	$rows = array();
	while($r = mysqli_fetch_array($result)) {
		$rows[] = $r;
	}
	echo $rows[0][0];

	mysqli_close($mysqli);
	exit;
?>