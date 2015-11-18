<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>

    <style>
        .right {
            float: right;
        }

        #exportButton {
            float: left;
        }
    </style>
    <!--Required scripts-->
    <script src="http://code.jquery.com/jquery-1.9.1.min.js"></script>
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
    <script>

        function createFormulasWorkbook() {

            var workbook = new $.ig.excel.Workbook($.ig.excel.WorkbookFormat.excel2007);
            var sheet = workbook.worksheets().add('Sheet1');
            sheet.columns(0).setWidth(180, $.ig.excel.WorksheetColumnWidthUnit.pixel);
            sheet.columns(1).setWidth(116, $.ig.excel.WorksheetColumnWidthUnit.pixel);
            sheet.columns(2).setWidth(124, $.ig.excel.WorksheetColumnWidthUnit.pixel);

            // Create some rows and columns of data
            sheet.getCell('A1').value('Project/Client');
            sheet.getCell('B1').value('Status');
            sheet.getCell('C1').value('Est. Duration (Mos.)');

            sheet.getCell('A2').value('Bank of Winniwonka');
            sheet.getCell('B2').value('In Progress');
            sheet.getCell('C2').value(7);

            sheet.getCell('A3').value('Arvalis, Inc.');
            sheet.getCell('B3').value('Negotiation');
            sheet.getCell('C3').value(3);

            sheet.getCell('A4').value('Panamoramic Studios');
            sheet.getCell('B4').value('Negotiation');
            sheet.getCell('C4').value(15);

            sheet.getCell('A5').value('Werkz.me');
            sheet.getCell('B5').value('Change Request');
            sheet.getCell('C5').value(24);

            // Add a formula to average one fo the columns of data
            sheet.getCell('C6').applyFormula("=SUM(C2:C5)");

            // Alternatively, the formula can be parsed and applied manually, like so:
            //var formula = $.ig.excel.Formula.parse('=AVERAGE(C2:C5)', $.ig.excel.CellReferenceMode.a1);
            //formula.applyTo(sheet.getCell('C6'));

            // Save the workbook
            saveWorkbook(workbook, "Formulas.xlsx");
        }

        function saveWorkbook(workbook, name) {
            workbook.save({ type: 'blob' }, function (data) {
                saveAs(data, name);
            }, function (error) {
                alert('Error exporting: : ' + error);
            });
        }

    </script>

</head>
<body>
    <br />
    <button id="exportButton" onclick="createFormulasWorkbook()">Create File</button>
    <br />
    <img alt="Result in Excel" src="http://www.igniteui.com/images/samples/client-side-excel-library/excel-formulas.png" />
</body>
</html>