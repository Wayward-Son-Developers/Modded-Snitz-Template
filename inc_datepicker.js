var strUrl, dtNow, objYear, objMonth, objDay;
var arrMonths, arrDays, intYear, intOldYear, intMonth;
var intDay, intLen, intOldLen, intIndex, intValue;

strUrl = window.location.href.toLowerCase();
dtNow = new Date();
objYear = document.getElementById("year");
objMonth = document.getElementById("month");
objDay = document.getElementById("day");

arrMonths=["January","February","March","April","May","June","July","August","September","October","November","December"];
arrDays=[31,28,31,30,31,30,31,31,30,31,30,31];

function DateSelector(w){
	intYear = objYear.options[objYear.selectedIndex].value;
	if (!objMonth.disabled) {
		intMonth = objMonth.options[objMonth.selectedIndex].value;
	} else {
		intMonth = 0;
	}
	if (!objDay.disabled) {
		intDay=objDay.options[objDay.selectedIndex].value;
	} else {
		intDay=0;
	}
	if (!w) {
		if (intYear != "") {
			if (CheckYear(intYear) || CheckYear(intOldYear)) {
				if (CheckYear(intYear)) {
					intLen = dtNow.getMonth()+1;
				} else {
					intLen = 12;
				}
				Populate(objMonth,"Month",1,intMonth)
			}
			Show(objMonth);
			intOldYear = intYear;
			if (!objDay.disabled) {
				DateSelector(1);
			}
		} else {
			Hide(objDay);
			Hide(objMonth);
		}
	} else {
		if (intYear != "" && intMonth != "") {
			intLen=arrDays[intMonth-1];
			if (intMonth-1 == dtNow.getMonth() && CheckYear(intYear)) {
				intLen=dtNow.getDate();
			}
			if (intMonth == 2 && intLen == 28 && intYear%4 == 0 && !(intYear%100 == 0 && intYear%400 != 0)) {
				intLen++;
			}
			if (intLen != intOldLen) {
				Populate(objDay,"Day",0,intDay);
			}
			Show(objDay);
			intOldLen = intLen;
		} else {
			Hide(objDay);
		}
	}
}

function CheckYear(y) {
	if (y != objYear[1].value) {
 		return false;
	} else {
		return true;
	}
}

function Hide(o){
	if (!o.disabled) {
		o.disabled = true;
		o.selectedIndex = 0;
		o.style.visibility = "hidden";
	}
}

function Populate(o,d,a,v) {
	o.options.length = 0;
	o[0] = new Option(d,"");
	o.selectedIndex = 0;
	for (var x=1; x<=intLen; x++) {
		intIndex = x;
		if (intIndex <= 9) {
			intValue = "0"+intIndex;
		} else {
			intValue = intIndex;
		}
		if (a) {
			intIndex--;
		}
		o[x] = new Option(a?arrMonths[intIndex]:intIndex,intValue);
		if (v == intValue) {
			o.selectedIndex = x;
		}
	}
}

function Show(o){
	if (o.disabled) {
		o.style.visibility = "visible";
		o.disabled = false;
		o.focus();
	}
}

if (strUrl.indexOf("register") != -1 || objYear.selectedIndex==0){
	Hide(objDay);
	Hide(objMonth);
} else {
	intOldYear = objYear.options[objYear.selectedIndex].value;
	intOldLen = parseInt(objDay.options[objDay.length-1].value);
}