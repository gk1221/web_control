//------------------------------------------------------------------------------------------
if(window.XMLHttpRequest) {	//FireFox,Mozilla,Opera 8.0+,Safari,...
	http_request = new XMLHttpRequest();
} else	if(window.ActiveXObject) {	//IE
	try{	//·sª©IE
		http_request = new ActiveXObject("Msxml2.XMLHTTP");
	}
	catch(e) {
		try {	//ÂÂª©IE
			http_request = new ActiveXObject("Microsoft.XMLHTTP");
		}
		catch(e) {}
	}

}
//-----------------------------------------------------------------------------------------