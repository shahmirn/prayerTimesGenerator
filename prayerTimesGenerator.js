/**
 * Utility to create an excel sheet that contains muslim prayer timings
 */

require('datejs');

var pt = require('prayer-times'),
	excelbuilder = require('msexcel-builder');

var latitude = 29.742278,
	longitude = -95.500213,
	timezoneOffset = -6;

var monthNames = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];

var timeColumnHeadings = ["Day", "Imsak", "Fajr", "Sunrise", "Zuhr", 
	"Sunset", "Maghrib", "Midnight"];

var year = new Date().getFullYear();
var date = new Date(year, 0, 1);

var prayTimes = new pt();
prayTimes.setMethod('Jafari');
prayTimes.adjust({imsak: 18});

var workbook = excelbuilder.createWorkbook('./', 'prayerTimes.xlsx');
var sheet1 = workbook.createSheet('sheet1', 26, 140);

var month = -1,
	rowCount = 2,
	colCountAddl = 0,
	rowMax = 0;

sheet1.set(1, 1, "Salaat(Prayer) Timetable for Houston, TX");
sheet1.merge({col:1,row:1}, {col:26,row:1});
sheet1.align(1, 1, 'center');
sheet1.font(1, 1, {bold: true});

var isLeap = new Date(year, 1, 29).getMonth() == 1;
var daysMax = isLeap ? 366 : 365;

for (var count = 0; count < daysMax; count++) {
	// in the excel sheet, months are displayed horizontall in chunks of three,
	// so determine the column offsets and the maximum number of rows based on the month
	if (date.getMonth() != month) {
		month = date.getMonth();
		if (month%3==0) {
			colCountAddl = 0;
			rowMax = parseInt(33*(month)/3) + 2 + 1*month/3;
		} else if (month%3==1) {
			colCountAddl = 9;
		} else if (month%3==2) {
			colCountAddl = 18;
		}

		rowCount = rowMax;

		// show the month
		sheet1.set(1+colCountAddl, rowCount, monthNames[month]);
		sheet1.merge({col:1+colCountAddl,row:rowCount}, {col:8+colCountAddl,row:rowCount});
		sheet1.align(1+colCountAddl, rowCount, 'center');
		sheet1.font(1+colCountAddl, rowCount, {bold: true});
		rowCount++;

		// Show the column headings (day, imsak, fajr, etc)
		for (var colCount = 0; colCount <= timeColumnHeadings.length; colCount++) {
			sheet1.set(1+colCount+colCountAddl, rowCount, timeColumnHeadings[colCount]);
			sheet1.font(1+colCount+colCountAddl, rowCount, {bold: true});
			sheet1.align(1+colCount+colCountAddl, rowCount, 'center');
			sheet1.width(1+colCount+colCountAddl, 11);
		}

		rowCount++;
	}

	// Calculate the prayer timings
	var times = prayTimes.getTimes(date, [latitude, longitude], timezoneOffset);
	var timesToDisplay = [times.imsak, times.fajr, times.sunrise, times.dhuhr, times.sunset, 
		times.maghrib, times.midnight];
	
	// Display the day (1, 2, 3, ... 31) and the prayer timings for the corresponding day
	sheet1.set(1+colCountAddl, rowCount, date.getDate());
	sheet1.align(1+colCountAddl, rowCount, 'center');
	for (var timesCount = 0; timesCount < timesToDisplay.length; timesCount++) {
		sheet1.set(2+timesCount+colCountAddl, rowCount, timesToDisplay[timesCount]);
		sheet1.align(2+timesCount+colCountAddl, rowCount, 'right');
	}

	rowCount++;

	date = date.add(1).days();
}

workbook.save(function(ok){
	if (!ok) 
		workbook.cancel();
	else
		console.log('congratulations, your workbook created');
});