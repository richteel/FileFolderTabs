function includeFile (elementId, url) {
	var elm = document.getElementById(elementId);
	
	if(elm == null)
		return;
	
	var xhr= new XMLHttpRequest();
	xhr.open('GET', url, true);
	xhr.onreadystatechange= function() {
	    if (this.readyState!==4) return;
	    if (this.status!==200) return; // or whatever error handling you want
	    
	    document.getElementById(elementId).outerHTML= this.responseText;
	};
	xhr.send();
}


function includeSection (elementId) {
	var elm = document.getElementById(elementId);
	
	if(elm == null)
		return;
	
	if(elementId == "_header_") {
		elm.outerHTML = myHeader;
	}
	else if(elementId == "_footer_") {
		elm.outerHTML = myFooter;
	}		
}

function showHideElement (elementId) {
	var elm = document.getElementById(elementId);
	
	if(elm == null)
		return;
	
	if(elm.style == null)
		return;
	
	if(elm.style.display == null)
		return;
	
	if (elm.style.display === "none") {
		elm.style.display = "block";
	} else {
		elm.style.display = "none";
	}
}

var myHeader = `<p id="header">
	<a href="https://github.com/richteel/" target="”_blank">
		<img src="./images/GitHub-Mark-64px.png" alt="GitHub Project Page" title="GitHub Project" width="32" height="32" />
		GitHub Projects
	</a>
</p>`;


var myFooter = `<p id="footer">
	Get latest version at <a href="https://github.com/richteel/FileFolderTabs" target="”_blank">https://github.com/richteel/FileFolderTabs</a>
</p>
</div>`;