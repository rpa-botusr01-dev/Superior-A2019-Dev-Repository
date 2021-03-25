var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

var today = new Date();

var args = WScript.Arguments;

var yesterday = new Date(today);
yesterday.setDate(today.getDate() - 1);

var dd,mm,yy,res;

if(args.item(0) == 1){
yy = today.getFullYear().toString();
mm = monthNames[today.getMonth()].toString();
dd = today.getDate().toString();
res = dd+"-"+mm+"-"+yy;
}
else if(args.item(0) == 2){
yy = yesterday.getFullYear().toString();
mm = monthNames[yesterday.getMonth()].toString();
dd = yesterday.getDate().toString();
res = dd+"-"+mm+"-"+yy;
}
else if(args.item(0) == 3){
yy = today.getFullYear();
mm = today.getMonth()+1;
dd = today.getDate();
if(mm <= 9){
mm = "0"+mm;
}
res = yy+"-"+mm+"-"+dd;
}

WScript.StdOut.WriteLine(''+ res);