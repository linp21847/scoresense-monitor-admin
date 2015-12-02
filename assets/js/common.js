var stateCodes = [["Alabama", "al", "49802"],
					["Alaska", "ak", "76816"],
					["Arizona", "az", "55371"],
					["Arkansas", "ar", "47047"],
					["California", "ca", "63544"],
					["Colorado", "co", "66870"],
					["Connecticut", "ct", "73424"],
					["Delaware", "de", "62771"],
					["District of Columbia", "dc", "84730"],
					["Florida", "fl", "52125"],
					["Georgia", "ga", "53127"],
					["Hawaii", "hi", "66588"],
					["Idaho", "id", "51499"],
					["Illinois", "il", "61450"],
					["Indiana", "in", "53050"],
					["Iowa", "ia", "59977"],
					["Kansas", "ks", "58548"],
					["Kentucky", "ky", "47977"],
					["Louisiana", "la", "49159"],
					["Maine", "me", "54351"],
					["Maryland", "md", "76248"],
					["Massachusetts", "ma", "69395"],
					["Michigan", "mi", "53450"],
					["Minnesota", "mn", "65067"],
					["Mississippi", "ms", "44523"],
					["Missouri", "mo", "52188"],
					["Montana", "mt", "55262"],
					["Nebraska", "ne", "60334"],
					["Nevada", "nv", "56071"],
					["New Hampshire", "nh", "67677"],
					["New Jersey", "nj", "70933"],
					["New Mexico", "nm", "52166"],
					["New York", "ny", "60502"],
					["North Carolina", "nc", "52151"],
					["North Dakota", "nd", "63989"],
					["Ohio", "oh", "53960"],
					["Oklahoma", "ok", "52279"],
					["Oregon", "or", "56398"],
					["Pennsylvania", "pa", "56505"],
					["Rhode Island", "ri", "62724"],
					["South Carolina", "sc", "50493"],
					["South Dakota", "sd", "57630"],
					["Tennessee", "tn", "49552"],
					["Texas", "tx", "57174"],
					["Utah", "ut", "58110"],
					["Vermont", "vt", "61812"],
					["Virginia", "va", "66585"],
					["Washington", "wa", "65215"],
					["West Virginia", "wv", "46103"],
					["Wisconsin", "wi", "59064"],
					["Wyoming", "wy", "64902"]];

var fontSizeMapping = {
	11: 220,
	12: 240,
	13: 260,
	14: 280,
	15: 300,
	16: 320
};

$.ig.loader({
	scriptPath: "http://cdn-na.infragistics.com/igniteui/2015.2/latest/js/",
	cssPath: "http://cdn-na.infragistics.com/igniteui/2015.2/latest/css/",
	resources: 'modules/infragistics.util.js,' +
				'modules/infragistics.documents.core.js,' +
				'modules/infragistics.excel.js'
});

var CreditReportExtractor = {
	scores: JSON.parse(localStorage.getItem("scores") || JSON.stringify({})),

	personal: JSON.parse(localStorage.getItem("personal") || JSON.stringify({})),

	cluster: JSON.parse(localStorage.getItem("cluster") || JSON.stringify({bank:[], closed: [], installment: []})),

	createWorkbook: function(callback) {
		console.log("Creating workbook...");

		var self = this,
			personName = self.personal.name,
			workbook = new $.ig.excel.Workbook($.ig.excel.WorkbookFormat.excel2007),
			calculatorWorksheet = workbook.worksheets().add("Calculator"),
			verificationCallWorksheet = workbook.worksheets().add("Verification Call"),
			summaryWorksheet = workbook.worksheets().add("Summary"),
			stateCodesWorksheet = workbook.worksheets().add("State Codes");

		calculatorWorksheet = self.createCalculatorWorksheet(calculatorWorksheet);
		verificationCallWorksheet = self.createVerificationCallWorksheet(verificationCallWorksheet);
		summaryWorksheet = self.createSummaryWorksheet(summaryWorksheet);
		stateCodesWorksheet = self.createStateCodesWorksheet(stateCodesWorksheet);

		workbook.save({ type: 'blob' }, function(data) {
			console.log(data);
			saveAs(data, personName.substr(0, 30) + ".xlsx");

			if (typeof callback === "function")
				callback();
		},
		function(error) {
			console.log(error);
		});
	},

	createCalculatorWorksheet: function(worksheet) {
		var self = this,
			bankAccounts = self.cluster.bank,
			retailCards = self.cluster.retail,
			closedAccounts = self.cluster.closed,
			authorized = self.cluster.authorized,
			installmentAccounts = self.cluster.installment,
			fraud = self.fraud,
			curRowIndex = 0,
			personal = self.personal,
			publicRecords = self.public,
			inquiries = self.inquiries,
			formattedString = new $.ig.excel.FormattedString( "Formatted String" ),
			setCurrencyModeToCell = function(cell, value) {
				balanceCellFormat = cell.cellFormat();
				balanceCellFormat.formatString("$0");
				cell.value(parseInt(value));
			},
			getAccountStatus = function(remarkString) {
				if (remarkString.toLowerCase().indexOf("paid") > -1)
					return "Paid";
				else if (remarkString.toLowerCase().indexOf("closed") > -1)
					return "Closed";
				else
					return "Paid";
			},
			setTitleModeToCell = function(cell, value) {
				cellFormat = cell.cellFormat();
				cellFormat.alignment($.ig.excel.HorizontalCellAlignment.center);
				cellFormat.font().bold(true);
				cellFormat.font().height(fontSizeMapping['12']);
				cell.value(value);
			},
			setTableHeadModeToCell = function(cell, value) {
				format = cell.cellFormat();
				format.alignment($.ig.excel.HorizontalCellAlignment.left);
				format.font().bold(true);
				format.fill($.ig.excel.CellFill.createSolidFill('#8DB4E3'));
				format.topBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.rightBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.bottomBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.leftBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				cell.value(value);
			},
			fillGreenToCell = function(cell) {
				format = cell.cellFormat();
				format.fill($.ig.excel.CellFill.createSolidFill('#DBEEF3'));
				format.topBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.rightBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.bottomBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.leftBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				return cell;
			},
			fillYellowToCell = function(cell) {
				format = cell.cellFormat();
				format.fill($.ig.excel.CellFill.createSolidFill('#FFFF00'));
				format.topBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.rightBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.bottomBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				format.leftBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
				return cell;
			},
			drawBorderToCells = function(r1, c1, r2, c2) {
				for (var i = r1; i < r2 + 1; i ++) {
					for (var j = c1; j < c2 + 1; j ++) {
						format = worksheet.rows(i).cells(j).cellFormat();

						format.topBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
						format.rightBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
						format.bottomBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
						format.leftBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
					}
				}
			},
			getLatePaymentCommonValue = function(param) {
				var result = {
						30: "",
						60: "",
						90: ""
					},
					indArr = ['30', '60', '90'],
					indTempArr = ['TransUnion', 'Experian', 'Equifax'];

				for (var i = 0; i < indArr.length; i++) {
					for (var j = 0; j < indTempArr.length; j++) {
						if (param[indTempArr[j]][indArr[i]])
							result[indArr[i]] = param[indTempArr[j]][indArr[i]];
					}
				}

				return result;
			},
			createInqueryTable = function(rowIndex, inquiries) {
				var inqueryTableStartIndex = rowIndex,
					tempIndex = rowIndex;

				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(8), "Inquiries");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(9), "Date");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(10), "Experian");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(11), "Equifax");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(12), "Transunion");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(13), "Type of Inquiry");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(14), "60 Days Late");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(15), "90 Days Late");
				setTableHeadModeToCell(worksheet.rows(tempIndex).cells(16), "120 Days Late");
				tempIndex++;

				for(var i = 0; i < inquiries.length; i++) {
					item = inquiries[i];
					worksheet.rows(tempIndex).cells(8).value(item.name);
					worksheet.rows(tempIndex).cells(9).value(item.date);

					switch(item.creditBureau) {
						case "Experian":
							worksheet.rows(tempIndex).cells(10).value("X");
							worksheet.rows(tempIndex).cells(10).cellFormat().alignment($.ig.excel.HorizontalCellAlignment.center);
							break;

						case "Equifax":
							worksheet.rows(tempIndex).cells(11).value("X");
							worksheet.rows(tempIndex).cells(11).cellFormat().alignment($.ig.excel.HorizontalCellAlignment.center);
							break;
							
						case "TransUnion":
							worksheet.rows(tempIndex).cells(12).value("X");
							worksheet.rows(tempIndex).cells(12).cellFormat().alignment($.ig.excel.HorizontalCellAlignment.center);
							break;
							
						default:
							consoel.log("Unknown credit bureau found.");
							break;
					}
					tempIndex++;
				}
				drawBorderToCells(inqueryTableStartIndex, 8, tempIndex - 1, 16);
			};

		//	Column Width Config
			worksheet.columns(0).setWidth(17.14, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(1).setWidth(10.29, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(2).setWidth(17.29, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(3).setWidth(15.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(4).setWidth(12.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(5).setWidth(14, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(6).setWidth(18, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(7).setWidth(7.57, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(8).setWidth(15.43, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(9).setWidth(12.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(10).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(11).setWidth(11, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(12).setWidth(12.14, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(13).setWidth(10.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(14).setWidth(10, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(15).setWidth(10, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(16).setWidth(10, $.ig.excel.WorksheetColumnWidthUnit.character);

		//	Printing person name
			worksheet.rows(curRowIndex).cells(0).value("Person Name");
			tmpCell = worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 7);
			tmpCell.value(personal.name);
			format = tmpCell.cellFormat();
			format.font().height(fontSizeMapping[16]);
			format.font().bold(true);
			drawBorderToCells(curRowIndex, 0, curRowIndex, 7);
			curRowIndex++;

		//	Rows 0 - Bank Accounts Section...
			bankCardsTitle = worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4);
			setTitleModeToCell(bankCardsTitle,"Bank Cards");
			curRowIndex++;

			//	Rows 1
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Account Name");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Limit");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(3), "Debt to Credit Ratio");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(4), "Amount to Pay");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(5), "New Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(6), "Account Number");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(8), "High Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(9), "Highest Balance Held Ratio");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(11), "Date Opened");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(12), "Age");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(13), "30 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(14), "60 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(15), "90 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(16), "120 Days Late");
			curRowIndex++;

			//	From Rows 2, Bank accounts
			bankAccountStartIndex = curRowIndex + 1;
			for (var i = 0; i < bankAccounts.length; i++) {
				worksheet.rows(curRowIndex).cells(0).value(bankAccounts[i].name);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(1), bankAccounts[i].balance);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(2), bankAccounts[i].limit);
				tmpCell = worksheet.getCell('D' + (curRowIndex+1));
				tmpCell.cellFormat().formatString("0%");
				tmpCell.applyFormula("=B" + (curRowIndex+1) + "/C" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(4).applyFormula("=IF(C" + (curRowIndex+1) + "<=1000,B" + (curRowIndex+1) + ",IF(D" + (curRowIndex+1) + "<0.4,0,B" + (curRowIndex+1) + "-(C" + (curRowIndex+1) + "*0.4)))");
				fillGreenToCell(worksheet.rows(curRowIndex).cells(5)).applyFormula("=B" + (curRowIndex+1) + "-E" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(6).value(bankAccounts[i].accountNumber);
				worksheet.rows(curRowIndex).cells(8).value(bankAccounts[i].highBalance);
				tmpCell = worksheet.rows(curRowIndex).cells(9);
				tmpCell.cellFormat().formatString("0%");
				tmpCell.applyFormula("=I" + (curRowIndex+1) + "/C" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(11).value(bankAccounts[i].opened);
				worksheet.rows(curRowIndex).cells(12).applyFormula('=DATEDIF(L' + (curRowIndex+1) + ',TODAY(),"Y")');
				// worksheet.rows(curRowIndex).cells(13).value(bankAccounts[i].latePayments['30']);
				// worksheet.rows(curRowIndex).cells(14).value(bankAccounts[i].latePayments['60']);
				// worksheet.rows(curRowIndex).cells(15).value(bankAccounts[i].latePayments['90']);
				lateDates = getLatePaymentCommonValue(bankAccounts[i].latePaymentDates);
				worksheet.rows(curRowIndex).cells(13).value(lateDates['30']);
				worksheet.rows(curRowIndex).cells(14).value(lateDates['60']);
				worksheet.rows(curRowIndex).cells(15).value(lateDates['90']);
				curRowIndex++;
			}
			bankAccountEndIndex = curRowIndex;

			drawBorderToCells(bankAccountStartIndex - 2, 0, bankAccountEndIndex - 1, 6);
			drawBorderToCells(bankAccountStartIndex - 2, 8, bankAccountEndIndex - 1, 9);
			drawBorderToCells(bankAccountStartIndex - 2, 11, bankAccountEndIndex - 1, 16);

		//	Rows 1+(bank accounts count) - Retail Cards Section...
			bankCardsTitle = worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4);
			setTitleModeToCell(bankCardsTitle, "Retail Cards");
			curRowIndex++;

			retailCardStartIndex = curRowIndex;

			//	Rows 2+(bank accounts count)
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Account Name");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Limit");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(3), "Debt to Credit Ratio");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(4), "Amount to Pay");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(5), "New Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(6), "Account Number");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(8), "High Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(9), "Highest Balance Held Ratio");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(11), "Date Opened");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(12), "Age");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(13), "30 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(14), "60 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(15), "90 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(16), "120 Days Late");
			curRowIndex++;

			for (var i = 0; i < retailCards.length; i++) {
				worksheet.rows(curRowIndex).cells(0).value(retailCards[i].name);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(1), retailCards[i].balance);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(2), retailCards[i].limit);
				tmpCell = worksheet.getCell('D' + (curRowIndex+1));
				tmpCell.cellFormat().formatString("0%");
				tmpCell.applyFormula("=B" + (curRowIndex+1) + "/C" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(4).applyFormula("=IF(C" + (curRowIndex+1) + "<=1000,B" + (curRowIndex+1) + ",IF(D" + (curRowIndex+1) + "<0.4,0,B" + (curRowIndex+1) + "-(C" + (curRowIndex+1) + "*0.4)))");
				fillGreenToCell(worksheet.rows(curRowIndex).cells(5)).applyFormula("=B" + (curRowIndex+1) + "-E" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(6).value(retailCards[i].accountNumber);
				worksheet.rows(curRowIndex).cells(8).value(retailCards[i].highBalance);
				tmpCell = worksheet.rows(curRowIndex).cells(9);
				tmpCell.cellFormat().formatString("0%");
				tmpCell.applyFormula("=I" + (curRowIndex+1) + "/C" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(11).value(retailCards[i].opened);
				worksheet.rows(curRowIndex).cells(12).applyFormula('=DATEDIF(L' + (curRowIndex+1) + ',TODAY(),"Y")');
				// worksheet.rows(curRowIndex).cells(13).value(retailCards[i].latePayments['30']);
				// worksheet.rows(curRowIndex).cells(14).value(retailCards[i].latePayments['60']);
				// worksheet.rows(curRowIndex).cells(15).value(retailCards[i].latePayments['90']);
				lateDates = getLatePaymentCommonValue(retailCards[i].latePaymentDates);
				worksheet.rows(curRowIndex).cells(13).value(lateDates['30']);
				worksheet.rows(curRowIndex).cells(14).value(lateDates['60']);
				worksheet.rows(curRowIndex).cells(15).value(lateDates['90']);
				curRowIndex++;
			}
			retailCardEndIndex = curRowIndex;

		//	Adding a blank row
		curRowIndex++;

		//	Summary Line
			summaryLineIndex = curRowIndex + 1;
			setTitleModeToCell(worksheet.rows(curRowIndex).cells(0), "SUM:");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(1)).applyFormula("=SUM(B3:B" + bankAccountEndIndex + ",B" + retailCardStartIndex + ":B" + retailCardEndIndex + ")");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(2)).applyFormula("=SUM(C3:C" + bankAccountEndIndex + ",C" + retailCardStartIndex + ":C" + retailCardEndIndex + ")");
			setTitleModeToCell(worksheet.rows(curRowIndex).cells(3), "Total Amt to Pay:");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(4)).applyFormula("=SUM(E3:E" + bankAccountEndIndex + ",E" + retailCardStartIndex + ":E" + retailCardEndIndex + ")");
			curRowIndex += 2;

			//	4 + bankAccounts.length
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Debt to credit ratio");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(3)).applyFormula("=MAX(D3:D" + bankAccountEndIndex + ",D" + retailCardStartIndex + ":D" + retailCardEndIndex + ")");
			worksheet.rows(curRowIndex).cells(3).cellFormat().formatString("0%");
			self.summaryLineIndex = curRowIndex + 1;

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(8), "Highest Balance Held Ratio");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(9)).applyFormula("=MAX(J3:J" + bankAccountEndIndex + ",J" + retailCardStartIndex + ":J" + retailCardEndIndex + ")");
			worksheet.rows(curRowIndex).cells(9).cellFormat().formatString("0%");

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(12), "Oldest Account");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(13)).applyFormula("=MAX(M3:M" + bankAccountEndIndex + ",M" + retailCardStartIndex + ":M" + retailCardEndIndex + ")");
			curRowIndex += 2;

			//	Aggregate line
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Aggregate ");
			worksheet.rows(curRowIndex).cells(3).cellFormat().formatString("0%");
			fillYellowToCell(worksheet.rows(curRowIndex).cells(3)).applyFormula("=B" + summaryLineIndex + "/C" + summaryLineIndex);
			curRowIndex += 2;

		drawBorderToCells(retailCardStartIndex, 0, retailCardEndIndex - 1, 6);
		drawBorderToCells(retailCardStartIndex, 8, retailCardEndIndex - 1, 9);
		drawBorderToCells(retailCardStartIndex, 11, retailCardEndIndex - 1, 16);

		//	Closed Accounts With Balances and/or Lates line
			bankCardsTitle = worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4);
			setTitleModeToCell(bankCardsTitle, "Closed Accounts With Balances and/or Lates");
			curRowIndex++;

			var closedAccountStartIndex = curRowIndex;

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Account Name");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Account Type");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(3), "Account Number");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(4), "Payment Status");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(5), "Account Status");

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(8), "Date Opened");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(9), "Last Reported");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(10), "30 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(11), "60 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(12), "90 Days Late");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(13), "120 Days Late");
			curRowIndex++;

			for (var i = 0; i < closedAccounts.length; i++) {
				item = closedAccounts[i];
				worksheet.rows(curRowIndex).cells(0).value(item.name);
				worksheet.rows(curRowIndex).cells(1).value(item.type);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(2), item.balance);
				worksheet.rows(curRowIndex).cells(3).value(item.accountNumber);
				worksheet.rows(curRowIndex).cells(4).value(item.payStatus);
				fillGreenToCell(worksheet.rows(curRowIndex).cells(5)).value(getAccountStatus(item.remark));

				worksheet.rows(curRowIndex).cells(8).value(item.opened);
				worksheet.rows(curRowIndex).cells(9).value(item.reported);
				// worksheet.rows(curRowIndex).cells(10).value(item.latePayments[30]);
				// worksheet.rows(curRowIndex).cells(11).value(item.latePayments[60]);
				// worksheet.rows(curRowIndex).cells(12).value(item.latePayments[90]);
				lateDates = getLatePaymentCommonValue(item.latePaymentDates);
				worksheet.rows(curRowIndex).cells(10).value(lateDates['30']);
				worksheet.rows(curRowIndex).cells(11).value(lateDates['60']);
				worksheet.rows(curRowIndex).cells(12).value(lateDates['90']);
				curRowIndex++;
			}

			var closedAccountEndIndex = curRowIndex;

			drawBorderToCells(closedAccountStartIndex, 0, closedAccountEndIndex - 1, 5);
			drawBorderToCells(closedAccountStartIndex, 8, closedAccountEndIndex - 1, 13);

		//	Authorized User Accounts
			bankCardsTitle = worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4);
			setTitleModeToCell(bankCardsTitle, "Authorized User Accounts");
			curRowIndex++;

			authorizedAccountStartIndex = curRowIndex;

			//	Rows 
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Account Name");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Limit");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(3), "Debt to Credit Ratio");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(4), "Amount to Pay");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(5), "New Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(6), "Account Number");
			createInqueryTable(curRowIndex, inquiries);
			curRowIndex++;
			
			for (var i = 0; i < authorized.length; i++) {
				var item = authorized[i];

				worksheet.rows(curRowIndex).cells(0).value(item.name);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(1), item.balance);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(2), item.limit);
				tmpCell = worksheet.getCell('D' + (curRowIndex+1));
				tmpCell.cellFormat().formatString("0%");
				tmpCell.applyFormula("=B" + (curRowIndex+1) + "/C" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(4).applyFormula("=IF(C" + (curRowIndex+1) + "<=1000,B" + (curRowIndex+1) + ",IF(D" + (curRowIndex+1) + "<0.4,0,B" + (curRowIndex+1) + "-(C" + (curRowIndex+1) + "*0.4)))");
				fillGreenToCell(worksheet.rows(curRowIndex).cells(5)).applyFormula("=B" + (curRowIndex+1) + "-E" + (curRowIndex+1));
				worksheet.rows(curRowIndex).cells(6).value(item.accountNumber);
				curRowIndex++;
			}

			drawBorderToCells(authorizedAccountStartIndex, 0, curRowIndex - 1, 6);

			curRowIndex ++;

		//	Installment Accounts
			installmentAccountTitle = worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4);
			setTitleModeToCell(installmentAccountTitle, "Installment Accounts");
			curRowIndex++;

			installmentAccountStartIndex = curRowIndex;

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Account Name");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Type of Loan");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Balance");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(3), "Monthly Payment");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(4), "Date Opened");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(5), "Age");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(6), "Lates");
			curRowIndex++;

			for(var i = 0; i < installmentAccounts.length; i++) {
				item = installmentAccounts[i];
				worksheet.rows(curRowIndex).cells(0).value(item.name);
				worksheet.rows(curRowIndex).cells(1).value(item.type);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(2), item.balance);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(3), item.payment);
				worksheet.rows(curRowIndex).cells(4).value(item.opened);
				worksheet.rows(curRowIndex).cells(5).applyFormula('=DATEDIF(E' + (curRowIndex + 1) + ',TODAY(),"Y")');
				worksheet.rows(curRowIndex).cells(6).value(item.latePayments['30'] + ',' + item.latePayments['60'] + ',' + item.latePayments['90']);
				curRowIndex++;
			}

			installmentAccountEndIndex = curRowIndex;
			drawBorderToCells(installmentAccountStartIndex, 0, installmentAccountEndIndex - 1, 6);
			drawBorderToCells(authorizedAccountStartIndex, 8, installmentAccountEndIndex - 1, 16);

			curRowIndex++;

		//	Last section...

		//	Public records section
			setTitleModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4), "Public Records");
			curRowIndex++;

			publicRecordsStartIndex = curRowIndex;
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Account Name");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Type");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Date");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(3), "Status");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(4), "Amount");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(5), "Account");
			curRowIndex++;

			for(var i = 0; i < publicRecords.length; i ++) {
				item = publicRecords[i];

				worksheet.rows(curRowIndex).cells(0).value(item.name);
				worksheet.rows(curRowIndex).cells(1).value(item.type);
				worksheet.rows(curRowIndex).cells(2).value(item.date);
				worksheet.rows(curRowIndex).cells(3).value(item.status);
				setCurrencyModeToCell(worksheet.rows(curRowIndex).cells(4), item.amount);
				worksheet.rows(curRowIndex).cells(5).value(item.accountNumber);				
				curRowIndex++;
			}

			drawBorderToCells(publicRecordsStartIndex, 0, curRowIndex - 1, 5);

			curRowIndex ++;


		//	Address section
			addressTableStartIndex = curRowIndex;
			setTableHeadModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4), "Current Address");
			curRowIndex++;
			worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4).value(personal.curAddress.split("\n").join(" ").trim());
			curRowIndex++;

			setTableHeadModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4), "Previous Address");
			curRowIndex ++;
			for (var i = 0; i < personal.prevAddress.length; i ++) {
				tempAdrr = personal.prevAddress[i];
				worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 4).value(tempAdrr.split("\n").join(" ").trim());
				curRowIndex ++;
			}
			drawBorderToCells(addressTableStartIndex, 0, curRowIndex - 1, 4);

			curRowIndex++;

		//	Fraud alert section
			setTableHeadModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 6), "Fraud Alert");
			curRowIndex++;
			fraudAlertStartIndex = curRowIndex;
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Bureau");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Date");
			setTableHeadModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 2, curRowIndex, 6), "Statement");
			curRowIndex++;

			worksheet.rows(curRowIndex).cells(0).value("Transunion");
			worksheet.mergedCellsRegions().add(curRowIndex, 2, curRowIndex, 6).value(fraud.TransUnion);
			curRowIndex++;

			worksheet.rows(curRowIndex).cells(0).value("Experian");
			worksheet.mergedCellsRegions().add(curRowIndex, 2, curRowIndex, 6).value(fraud.Experian);
			curRowIndex++;

			worksheet.rows(curRowIndex).cells(0).value("Equifax");
			worksheet.mergedCellsRegions().add(curRowIndex, 2, curRowIndex, 6).value(fraud.Equifax);
			curRowIndex++;

			drawBorderToCells(fraudAlertStartIndex, 0, curRowIndex - 1, 6);

			curRowIndex++;

		// Credit scroes section
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Credit Scores");
			curRowIndex++;

			drawBorderToCells(curRowIndex, 0, curRowIndex + 2, 2);

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Experian");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(1), "Equifax");
			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(2), "Transunion");
			curRowIndex++;

			worksheet.rows(curRowIndex).cells(0).value(parseInt(self.scores.Experian));
			worksheet.rows(curRowIndex).cells(1).value(parseInt(self.scores.Equifax));
			worksheet.rows(curRowIndex).cells(2).value(parseInt(self.scores.Transunion));
			curRowIndex++;

			setTableHeadModeToCell(worksheet.rows(curRowIndex).cells(0), "Age of Client");
			worksheet.rows(curRowIndex).cells(1).value(self.personal.birthday);
			worksheet.rows(curRowIndex).cells(2).applyFormula('=2015-B' + (curRowIndex + 1));
			self.yearBornLineInxex = curRowIndex + 1;
	},

	createVerificationCallWorksheet: function(worksheet) {
		var curRowIndex = 0,
			mergedRegion = worksheet.mergedCellsRegions().add( 0, 0, 0, 9 ),
			setLableModeToCell = function(cell, value, italicFlag) {
				cellFormat = cell.cellFormat();
				cellFormat.font().height(fontSizeMapping['13']);
				cellFormat.font().name("Arial Unicode MS");
				cell.value(value);

				if (italicFlag) {
					cellFormat.font().italic(true);
				}
			},
			setDataFieldModeToCell = function(cell) {
				cellFormat = cell.cellFormat();
				cellFormat.font().height(fontSizeMapping['13']);
				cellFormat.font().name("Arial Unicode MS");
				cellFormat.fill($.ig.excel.CellFill.createSolidFill('#EDE52E'));
				cellFormat.bottomBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
			};

		worksheet.columns(0).setWidth(25, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(1).setWidth(13.71, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(2).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(3).setWidth(13, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(4).setWidth(14.29, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(5).setWidth(10.43, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(6).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(7).setWidth(12, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(8).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(9).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(10).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(11).setWidth(10.14, $.ig.excel.WorksheetColumnWidthUnit.character);

		mergedRegion.value("CORPORATION PROFILE");
		cellFormat = mergedRegion.cellFormat();
		cellFormat.fill($.ig.excel.CellFill.createSolidFill('#000000'));
		cellFormat.font().height(fontSizeMapping['13']);
		cellFormat.alignment($.ig.excel.HorizontalCellAlignment.center);
		cellFormat.font().colorInfo(new $.ig.excel.WorkbookColorInfo('#FFFFFF'));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "Go into Sheet 3 and ask which type of cards they have for Chase, Bank of America, Citi, and Capital One (if any)");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Business Name:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 9));
		
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Mailing Address:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 6));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "Suite #", true);
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 8, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Verify Address on ID and Application");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Mailing Cont.:");
		setLableModeToCell(worksheet.rows(curRowIndex).cells(1), "City");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 2, curRowIndex, 4));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(5), "State", true);
		setDataFieldModeToCell(worksheet.rows(curRowIndex).cells(6));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "ZIP Code", true);
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 8, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Funding Estimate Amounts.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Tax Identification No.:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "# of Employees:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Seek Fee.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Phone Number:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Web Domain:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Multiple applications will be sent for credit cards.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Type of Entity:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "State of Incorporation:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Funding Status Update");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Nature of Business:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Services Provided:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       APR, both introductory and ongoing rates.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Business Incorp Date:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Business Start Date:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Timeline of funding process.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Business Gross Income");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Net Profit");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Do\'s and don\'ts of credit report.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       How to handle bank calls and emails.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "GUARANTOR INFO");
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Industry Experience:");
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Invoicing and Liquidation Instructions.");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Full Name:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Does client understand APR, both introductory and ongoing rates?");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(1), "Last");
		setLableModeToCell(worksheet.rows(curRowIndex).cells(3), "First");
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "Middle Name");
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Timeline of funding process:  ");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Mailing Address:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 6));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "Suite #");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 8, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Do\'s and don\'ts of funding process:  ");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Mailing Cont.:");
		setLableModeToCell(worksheet.rows(curRowIndex).cells(1), "City", true);
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 2, curRowIndex, 4));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(5), "State", true);
		setDataFieldModeToCell(worksheet.rows(curRowIndex).cells(6));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "ZIP Code", true);
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 8, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       How to handle bank calls:  ");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Social Security Number:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(3), "Birth Date:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 4, curRowIndex, 6));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "Age", true);
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 8, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Who their Seek Funding Coordinator is:");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Home Phone Number:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(3), "Cell Number:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 4, curRowIndex, 9));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(11), "       Additional Comments:");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Email Address:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Mother\'s Maiden Name:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Time at Residence:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 1, curRowIndex, 2));
		setLableModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 3, curRowIndex, 4), "Gross Annual Income:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Drivers License:");
		setDataFieldModeToCell(worksheet.rows(curRowIndex).cells(1));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(2), "State:");
		setDataFieldModeToCell(worksheet.rows(curRowIndex).cells(3));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(4), "Issue Date:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 5, curRowIndex, 6));
		setLableModeToCell(worksheet.rows(curRowIndex).cells(7), "Expiration:");
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 8, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Seek Additional Info", true);
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "1. Income used for Personal Or Business?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "2. Business Projection Used?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "3. Business address used on application? (Cannot Be P.O. BOX)");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "4. Time in business?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "5. Business Name Used? Business may have other names such as DBA,");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;
		curRowIndex++;
		curRowIndex++;


		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Business Questions:");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "1. Can they receive mail at business address?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "2. Does client have business checking account? What Bank? How much in deposits?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "3. Are there business Derrogatories/BK?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "4. Are there any existing business accounts?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "5. If Yes, Need name of Bank, Credit Limits, Balances, Average monthly payment being made, current/delinquent on account");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Personal Questions:");
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "1. Can they receive mail at personal address?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "2. Personal BK in the past?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "3. Personal Checking/Savings? What Banks? Current Deposit amounts? (If BOFA/CHASE-also ask last deposit amount, how much, when?)");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "4. Vehicles registered under PG (Year, Model, Color)?");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "5. College Graduated at? Year? Major? Any Special Degrees/License? (Example: real estate License)");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "6. Who else lives in the household? Need First,Middle,Last name for everyone in the household along with Date of Birth");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "7. Do they have personal credit cards with BofA/Chase? Last few purchases made, amount, due dates of each account.");
		curRowIndex++;
		setDataFieldModeToCell(worksheet.mergedCellsRegions().add(curRowIndex, 0, curRowIndex, 9));
		curRowIndex++;

		curRowIndex += 2;
		setLableModeToCell(worksheet.rows(curRowIndex).cells(0), "Go into Sheet 3 and ask which type of cards they have for Chase, Bank of America, Citi, and Capital One (if any)");
	},

	createSummaryWorksheet: function(worksheet) {
		var curRowIndex = 0,
			self = this,
			setBoldToCell = function(cell, value, underlineFlag) {
				cell.cellFormat().font().bold(true);
				cell.value(value);

				if (underlineFlag)
					cell.cellFormat().font().underlineStyle($.ig.excel.FontUnderlineStyle.single);
			},
			fillBlueToCell = function(cell) {
				cell.cellFormat().fill($.ig.excel.CellFill.createSolidFill('#92D050'));
				return cell;
			},
			drawBorderToCells = function(r1, c1, r2, c2) {
				for (var i = r1; i < r2 + 1; i ++) {
					for (var j = c1; j < c2 + 1; j ++) {
						format = worksheet.rows(i).cells(j).cellFormat();

						format.topBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
						format.rightBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
						format.bottomBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
						format.leftBorderColorInfo(new $.ig.excel.WorkbookColorInfo('#000000'));
					}
				}
			};

		//	Column width Config

			worksheet.columns(0).setWidth(38, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(1).setWidth(14.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(2).setWidth(14.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(3).setWidth(14.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(4).setWidth(3.43, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(5).setWidth(38, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(6).setWidth(9.43, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(7).setWidth(8.43, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(8).setWidth(8.43, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(9).setWidth(9.71, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(10).setWidth(6.14, $.ig.excel.WorksheetColumnWidthUnit.character);
			worksheet.columns(11).setWidth(9.14, $.ig.excel.WorksheetColumnWidthUnit.character);

		for (var i = 1; i < 24; i++) {
			if (i !== 17) {
				worksheet.rows(i).cells(6).cellFormat().fill($.ig.excel.CellFill.createSolidFill('#92D050'));
			} else {
				worksheet.rows(i).cells(6).cellFormat().fill($.ig.excel.CellFill.createSolidFill('#D7E4BC'));
			}
		}

		drawBorderToCells(0, 0, 26, 3);
		drawBorderToCells(1, 5, 23, 7);
		drawBorderToCells(2, 9, 3, 11);

		defaultFormat = worksheet.columns(0, 11).cellFormat();
		defaultFormat.font().height(fontSizeMapping['12']);

		setBoldToCell(worksheet.rows(curRowIndex).cells(1), "Tier 1");
		setBoldToCell(worksheet.rows(curRowIndex).cells(2), "Tier 2");
		setBoldToCell(worksheet.rows(curRowIndex).cells(3), "Tier 3");
		setBoldToCell(worksheet.rows(curRowIndex).cells(6), "Inputs");
		setBoldToCell(worksheet.rows(curRowIndex).cells(9), "Credit Score");
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Credit Score");
		worksheet.rows(curRowIndex).cells(1).value("720+");
		worksheet.rows(curRowIndex).cells(2).value("690-719");
		worksheet.rows(curRowIndex).cells(3).value("660-689");
		worksheet.rows(curRowIndex).cells(5).value("Credit Score");
		worksheet.rows(curRowIndex).cells(6).value("Inputs");
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$2>=660,$G$2<=689),"Tier 3",(IF(AND($G$2>=690,$G$2<=719),"Tier 2",(IF(AND($G$2>=720,$G$2<=900),"Tier 1",(IF(AND($G$2>=500,$G$2<=659),"DECLINE",)))))))');
		worksheet.rows(curRowIndex).cells(9).value("Credit Score");
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(9).value("Experian");
		worksheet.rows(curRowIndex).cells(10).value("Equifax");
		worksheet.rows(curRowIndex).cells(11).value("Transunion");
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(9).value(parseInt(self.scores.Experian));
		worksheet.rows(curRowIndex).cells(10).value(parseInt(self.scores.Equifax));
		worksheet.rows(curRowIndex).cells(11).value(parseInt(self.scores.Transunion));
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Debt to Credit Ratio");
		worksheet.rows(curRowIndex).cells(1).value("0-45%");
		worksheet.rows(curRowIndex).cells(2).value("46-50%");
		worksheet.rows(curRowIndex).cells(3).value("51-65%");
		worksheet.rows(curRowIndex).cells(5).value("Highest Utilization");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=Calculator!D" + self.summaryLineIndex);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$5>=0.1,$G$5<=0.45),"Tier 1",(IF(AND($G$5>=0.46,$G$5<=0.5),"Tier 2",(IF(AND($G$5>=0.51,$G$5<=0.65),"Tier 3",(IF(AND($G$5>=0.66,$G$5<=1),"DECLINE",)))))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(5).value("Aggregate Utilization");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=Calculator!D" + (self.summaryLineIndex + 2));
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Minimum # of open Lines");
		worksheet.rows(curRowIndex).cells(1).value(3);
		worksheet.rows(curRowIndex).cells(2).value(2);
		worksheet.rows(curRowIndex).cells(3).value(2);
		worksheet.rows(curRowIndex).cells(5).value("Minimum # of open Lines");
		worksheet.rows(curRowIndex).cells(6).value(0);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$7>=0,$G$7<=1.9),"DECLINE",(IF(AND($G$7>=2,$G$7<=2.9),"Tier 2 Or Tier 3",(IF(AND($G$7>=3,$G$7<=99),"Tier 1",)))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Minimum Age of Accounts (oldest)");
		worksheet.rows(curRowIndex).cells(1).value(4);
		worksheet.rows(curRowIndex).cells(2).value(2);
		worksheet.rows(curRowIndex).cells(3).value(2);
		worksheet.rows(curRowIndex).cells(5).value("Minimum Age of Accounts (oldest)");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=Calculator!N" + self.summaryLineIndex);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$8>=0,$G$8<=1.9),"DECLINE",(IF(AND($G$8>=2,$G$8<=3.9),"Tier 2 Or Tier 3",(IF(AND($G$8>=4,$G$8<=99),"Tier 1",)))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Max # of Inquiries/ per bureau (last 6 months)");
		worksheet.rows(curRowIndex).cells(1).value(2);
		worksheet.rows(curRowIndex).cells(2).value(4);
		worksheet.rows(curRowIndex).cells(3).value(6);
		worksheet.rows(curRowIndex).cells(5).value("Max # of Inquiries/ per bureau (last 6 months)");
		worksheet.rows(curRowIndex).cells(6).value(0);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$9>=0,$G$9<=2),"Tier 1",(IF(AND($G$9>=2.1,$G$9<=4),"Tier 2 ",(IF(AND($G$9>=4.1,$G$9<=6),"Tier 3",(IF(AND($G$9>=6.1,$G$9<=99),"DECLINE")))))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Max # Derogatories (last 2 years)");
		worksheet.rows(curRowIndex).cells(1).value(0);
		worksheet.rows(curRowIndex).cells(2).value(1);
		worksheet.rows(curRowIndex).cells(3).value(3);
		worksheet.rows(curRowIndex).cells(5).value("Max # Deragatories 30 days late (last 2 years)");
		worksheet.rows(curRowIndex).cells(6).value(0);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$10>=0,$G$10<=0.9),"Tier 1",(IF(AND($G$10>=1,$G$10<=1.9),"Tier 2 ",(IF(AND($G$10>=2,$G$10<=3.9),"Tier 3",(IF(AND($G$10>=4,$G$10<=99),"DECLINE")))))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Max # Deragatories 30 days late (last 2 years)");
		worksheet.rows(curRowIndex).cells(1).value(0);
		worksheet.rows(curRowIndex).cells(2).value(1);
		worksheet.rows(curRowIndex).cells(3).value(3);
		worksheet.rows(curRowIndex).cells(5).value("Max # Deragatories 60 days late (last 2 years)");
		worksheet.rows(curRowIndex).cells(6).value(0);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF($G$11=0,"All Tiers","DECLINE")');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Max # Deragatories 60 days late (last 2 years)");
		worksheet.rows(curRowIndex).cells(1).value(0);
		worksheet.rows(curRowIndex).cells(2).value(0);
		worksheet.rows(curRowIndex).cells(3).value(0);
		worksheet.rows(curRowIndex).cells(5).value("Max # Derogatories (last 2 years)");
		worksheet.rows(curRowIndex).cells(6).value(0);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$12>=0,$G$12<=0.9),"Tier 1",(IF(AND($G$12>=1,$G$12<=1.9),"Tier 2 ",(IF(AND($G$12>=2,$G$12<=3.9),"Tier 3",(IF(AND($G$12>=4,$G$12<=99),"DECLINE")))))))');
		curRowIndex++;
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Highest Balance Held Ratio (Highest)");
		worksheet.rows(curRowIndex).cells(1).value("60%+");
		worksheet.rows(curRowIndex).cells(2).value("30-60%");
		worksheet.rows(curRowIndex).cells(3).value("0-29%");
		worksheet.rows(curRowIndex).cells(5).value("Highest Balance Held Ratio (Highest) ");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=Calculator!J" + self.summaryLineIndex);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$14>=0.61,$G$14<=0.99),"Tier 1",(IF(AND($G$14>=0.3,$G$14<=0.6),"Tier 2 ",(IF(AND($G$14>=0,$G$14<=0.29),"Tier 3",)))))');
		curRowIndex++;
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("# of Satisifed Accounts");
		worksheet.rows(curRowIndex).cells(1).value("7+");
		worksheet.rows(curRowIndex).cells(2).value("3--6");
		worksheet.rows(curRowIndex).cells(3).value("1--2");
		worksheet.rows(curRowIndex).cells(5).value("# of Satisifed Accounts");
		worksheet.rows(curRowIndex).cells(6).value(0);
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$16>=7,$G$16<=99),"Tier 1",(IF(AND($G$16>=3,$G$16<=6.9),"Tier 2 ",(IF(AND($G$16>=1,$G$16<=2.9),"Tier 3",(IF(AND($G$16>=0,$G$16<=0.9),"DECLINE")))))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Mortgage Holder (Never Late)");
		worksheet.rows(curRowIndex).cells(1).value("yes");
		worksheet.rows(curRowIndex).cells(2).value("no");
		worksheet.rows(curRowIndex).cells(3).value("no");
		worksheet.rows(curRowIndex).cells(5).value("Mortgage Holder (Never Late)");
		worksheet.rows(curRowIndex).cells(6).value("no");
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF($G$17="yes","Tier 1","Tier 2 Or 3")');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Conservative States ");
		worksheet.rows(curRowIndex).cells(1).value("yes");
		worksheet.rows(curRowIndex).cells(2).value("no");
		worksheet.rows(curRowIndex).cells(3).value("no");
		worksheet.rows(curRowIndex).cells(5).value("Enter State (lower case)");
		worksheet.rows(curRowIndex).cells(6).value("CA");
		worksheet.rows(curRowIndex).cells(7).value('');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(5).value("Conservative States ");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=IF(VLOOKUP($G$18,'State Codes'!$B$1:$C$51,2,FALSE)>=60000,\"no\",\"yes\")");
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF($G$19="yes","Tier 1","Tier 2 Or 3")');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Age of client");
		worksheet.rows(curRowIndex).cells(1).value("25-60");
		worksheet.rows(curRowIndex).cells(2).value("25-60");
		worksheet.rows(curRowIndex).cells(3).value("22-65");
		worksheet.rows(curRowIndex).cells(5).value("Year Born");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=Calculator!B" + self.yearBornLineInxex);
		worksheet.rows(curRowIndex).cells(7).value('');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(5).value("Age of client");
		worksheet.rows(curRowIndex).cells(6).applyFormula("=2015-G20");
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF(AND($G$21>=25,$G$21<=60),"All Tiers",(IF(AND($G$21>=22,$G$21<=24.9),"Tier 3 ",(IF(AND($G$21>=61,$G$21<=65),"Tier 3",(IF(AND($G$21>=66,$G$21<=99),"DECLINE")))))))');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Primary Funding Level");
		worksheet.rows(curRowIndex).cells(1).value("$60,000-$90,000");
		worksheet.rows(curRowIndex).cells(2).value("$30,000- $75,000");
		worksheet.rows(curRowIndex).cells(3).value("$10,000- $40,000");
		setBoldToCell(worksheet.rows(curRowIndex).cells(5), "Funding Holdbacks", true);
		worksheet.rows(curRowIndex).cells(6).value("");
		worksheet.rows(curRowIndex).cells(7).value('');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Secondary Funding Levell");
		worksheet.rows(curRowIndex).cells(1).value("$40,000- $70,000");
		worksheet.rows(curRowIndex).cells(2).value("$10,000- $40,000");
		worksheet.rows(curRowIndex).cells(3).value("$5,000- $30,000");
		worksheet.rows(curRowIndex).cells(5).value("Mortgage Holder (Never Late)");
		worksheet.rows(curRowIndex).cells(6).value("no");
		worksheet.rows(curRowIndex).cells(7).applyFormula('=IF($G$23="yes","All Tiers","DECLINE")');
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(5).value("Bankruptcies, Collections, Judgements ");
		curRowIndex++;

		setBoldToCell(worksheet.rows(curRowIndex).cells(0), "Funding Holdbacks", true);
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Age of client");
		worksheet.rows(curRowIndex).cells(1).value("52-60");
		worksheet.rows(curRowIndex).cells(2).value("22-25, 52-60");
		worksheet.rows(curRowIndex).cells(3).value("22-25, 52-60");
		curRowIndex++;

		worksheet.rows(curRowIndex).cells(0).value("Mortgage Holder (Never Late)");
		worksheet.rows(curRowIndex).cells(1).value("no");
		worksheet.rows(curRowIndex).cells(2).value("no");
		worksheet.rows(curRowIndex).cells(3).value("no");
		curRowIndex++;
	},

	createStateCodesWorksheet: function(worksheet) {
		var self = this;

		worksheet.columns(0).setWidth(18, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(1).setWidth(3.14, $.ig.excel.WorksheetColumnWidthUnit.character);
		worksheet.columns(2).setWidth(8.14, $.ig.excel.WorksheetColumnWidthUnit.character);

		for (var i = 0; i < stateCodes.length; i++) {
			for (var j = 0; j < stateCodes[i].length; j++) {
				if (j === 2)
					worksheet.rows(i).cells(j).value(parseInt(stateCodes[i][j]));
				else
					worksheet.rows(i).cells(j).value(stateCodes[i][j]);
			}
		}
	},

	refine: function(items, flag) {
		var self = this,
			result = "";
		switch(flag) {
			case "acc-num":
			case "type":
			case "name":
			case "employer":
				result = items[0];

				for(var i = 1; i < items.length; i++) {
					if (items[i].length > result.length)
						result = items[i];
				}
				break;

			case "lates":
				result = parseInt(items[0] || 0);
				for (var i = 0; i < items.length; i++) {
					if (parseInt(items[i].substr(1)) > result) {
						result = parseInt(items[i].substr(1));
					}
				}
				break;

			case "balance":
			case "limit":
			case "payment":
				result = parseInt(items[0].substr(1) || 0);
				for (var i = 0; i < items.length; i++) {
					if (parseInt(items[i].substr(1)) > result) {
						result = parseInt(items[i].substr(1));
					}
				}
				break;

			case "prev-addr":
				result = [items[0]];
				for (var i = 1; i < items.length; i++) {
					if (items.indexOf(items[i]) !== -1) {
						result.push(items[i]);
					}
				}
				break;

			default:
				result = items[0];
				for (var i = 1; i < items.length; i++) {
					if (!result) {
						result = items[i];
					} else {
						break;
					}
				}
				break;
		}

		return result;
	},

	getMoreInfo: function() {
		var self = this,
			accounts = self.accounts;

		self.curItem = self.accounts.shift();

		if (self.curItem) {
			self.saveState();
			chrome.tabs.create({url: self.curItem.detailViewLink}, function(tab) {
				console.log(self.curItem);
			});
		} else {
			self.curItem = {};
			self.stop();
		}
	}
}