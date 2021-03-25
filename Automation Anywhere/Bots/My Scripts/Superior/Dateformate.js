function GetFormattedDate(){
var toDate=new Date();
var month_names =["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
var day = toDate.getDate();
var month_index = toDate.getMonth();
var year = toDate.getFullYear();
return"" + day + "-" + month_names[month_index] + "-" + year;
}
WScript.StdOut.WriteLine(GetFormattedDate());