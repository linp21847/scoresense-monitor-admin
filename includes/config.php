<?php
	include 'env.php';

	$mysqli = new mysqli($db_host, $db_username, $db_password, $db_name);

	/* check connection */
	if (mysqli_connect_errno()) {
		printf("Connect failed: %s\n", mysqli_connect_error());
		exit();
	}
?>