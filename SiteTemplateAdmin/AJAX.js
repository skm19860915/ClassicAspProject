function loadXMLDoc(url, asynchFlag, postFlag, callbackFunction) {	
	// branch for native XMLHttpRequest object
	if (window.XMLHttpRequest) {
		xml_req = new XMLHttpRequest();
	} else if (window.ActiveXObject) {
		xml_req = new ActiveXObject("Microsoft.XMLHTTP");
	} else {
		throw new Error("Cannot instantiate XMLHttpRequest");
	}

	if(callbackFunction) {
		xml_req.onreadystatechange = function() {
			if(xml_req.readyState == 4 && xml_req.status == 200) {
				callbackFunction.xml_req = xml_req;
				callbackFunction();
			}
		};
	}

	if(!postFlag) {
		xml_req.open("GET", url, asynchFlag);
		xml_req.send(null);
	} else {	
		var qMarkPos = url.indexOf('?');
		var postContent = url.substr(qMarkPos + 1);
		url = url.substr(0, qMarkPos);

		xml_req.open("POST", url, asynchFlag);
		xml_req.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		xml_req.setRequestHeader("Content-length", postContent.length); 
		xml_req.setRequestHeader("Connection", "close");
		xml_req.send(postContent);
	}
	
	if(!asynchFlag)
		return xml_req.responseText;
}