<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Credit Score Sensor Export</title>
	<link rel="stylesheet" type="text/css" href="assets/css/style.css">

	<!--Required scripts-->
	<script src="assets/js/jquery.min.js"></script>
	<!-- External files for exporting -->
	<script src="http://www.igniteui.com/js/external/FileSaver.js"></script>
	<script src="http://www.igniteui.com/js/external/Blob.js"></script>

	<!-- Ignite UI Loader Script -->
	<script src="http://cdn-na.infragistics.com/igniteui/2015.2/latest/js/infragistics.loader.js"></script>

	<script>
		$.ig.loader({
			scriptPath: "http://cdn-na.infragistics.com/igniteui/2015.2/latest/js/",
			cssPath: "http://cdn-na.infragistics.com/igniteui/2015.2/latest/css/",
			resources: 'modules/infragistics.util.js,' +
						'modules/infragistics.documents.core.js,' +
						'modules/infragistics.excel.js'
		});
	</script>

	

</head>
<?php
	
	if (isset($_GET['id']) && !empty($_GET['id'])) {
		$id = $_GET['id'];
	}
	if (!empty($id)) {
?>
	 	<script src="assets/js/common.js"></script>
	 	<script type="text/javascript">
	 	var data = null;
	 	window.onload = function() {
	 		var request = new XMLHttpRequest();
			request.onreadystatechange = function() {
				if (request.readyState === 4) {
					if (request.status === 200) {
						$.ig.loader({
							scriptPath: "http://cdn-na.infragistics.com/igniteui/2015.2/latest/js/",
							cssPath: "http://cdn-na.infragistics.com/igniteui/2015.2/latest/css/",
							resources: 'modules/infragistics.util.js,' +
										'modules/infragistics.documents.core.js,' +
										'modules/infragistics.excel.js'
						});
						data = JSON.parse(request.responseText);
						CreditReportExtractor.cluster = data.cluster;
						CreditReportExtractor.fraud = data.fraud;
						CreditReportExtractor.personal = data.personal;
						CreditReportExtractor.inquiries = data.inquiries;
						CreditReportExtractor.scores = data.scores;
						CreditReportExtractor.public = data.public;
						CreditReportExtractor.createWorkbook(function() {
							localStorage.setItem("export_time", JSON.stringify((new Date()).getTime()));
						});
					} else {
						document.body.className = 'error';
					}
				}
			};
			request.open("GET", "/apis/getData.php?id=<?php echo $id; ?>" , true);
			request.setRequestHeader('Content-Type', 'application/json; charset=UTF-8');
			request.send(null);
	 	};
	 	</script>
	 	<?php
	}
?>
<body>
	<h2>Wait for it</h2>
</body>


</html>