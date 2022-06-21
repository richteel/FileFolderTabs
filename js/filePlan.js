/*
    Processes Excel Workbook with:
        SheetName -> Tabs
        Row 1 Columns -> groupOrder	tabColor	textColor	tabGroup	tabSubGroup	tabTitle	print
        With Types    -> int        string      string      string      string      string      bool
*/
// REF Not used: https://www.howtogeek.com/devops/web-apps-can-interact-with-your-filesystem-now/

let fileHandle;
let worksheet;
let previousWidth = 0;
let previousHeight = 0;

// REF: https://developer.mozilla.org/en-US/docs/Web/API/Window/showOpenFilePicker
const pickerOpts = {
    types: [
        {
            description: 'Spreadsheets',
            accept: {
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
            }
        },
    ],
    excludeAcceptAllOption: true,
    multiple: false
};

function createFolderTabs() {
    if(!worksheet) {
        console.log("Worksheet object not set");
        return;
    }

    // Get the output span
    const outSpan = document.getElementById("output");
    const printAllCheckbox = document.getElementById("printAll");
    const tabWidthTextBox = document.getElementById("tabWidth");
    const tabHeightTextBox = document.getElementById("tabHeight");

    if(!outSpan) {
        console.log("Did not find output span");
        return;
    }
    if(!printAllCheckbox) {
        console.log("Did not find print all checkbox");
        return;
    }
    if(!tabWidthTextBox) {
        console.log("Did not find Width textbox");
        return;
    }
    if(!tabHeightTextBox) {
        console.log("Did not find Height textbox");
        return;
    }

    // Clear Span
    outSpan.innerHTML = "";

    const template = `<div class="outerTab" style="width:%width%in; height:%height%in;">
            <div class="tab">
                <div class="tabgroup" style="color:%textColor%; background-color:%tabColor%">%tabGroup%</div>
                <div class="tabtext">%tabText%</div>
            </div>
        </div>`

    // Write tabs
    for(let i=0; i < worksheet.length; i++) {
        if(worksheet[i]["tabGroup"]) {
            // console.log(i, XL_row_object[i]);
            let tabDiv = template.replace("%textColor%", worksheet[i]["textColor"]).replace("%tabColor%", worksheet[i]["tabColor"]).replace("%tabGroup%", worksheet[i]["tabGroup"]);
            let tabText = worksheet[i]["tabTitle"] === undefined ? "" : worksheet[i]["tabTitle"];

            if(!(worksheet[i]["tabSubGroup"] === undefined))
                tabText = worksheet[i]["tabSubGroup"] + ": " + tabText;

            tabDiv = tabDiv.replace("%tabText%", tabText);
            tabDiv = tabDiv.replaceAll("%width%", tabWidthTextBox.value);
            tabDiv = tabDiv.replaceAll("%height%", tabHeightTextBox.value);

            if(worksheet[i]["print"].toLowerCase() == "true" || printAllCheckbox.checked)
                outSpan.innerHTML += tabDiv;
        }
    }
}

function isNumber(evt) {
    evt = (evt) ? evt : window.event;

    const elem = document.getElementById(evt.srcElement.id);
    if(!elem) {
        console.log("Did not find numeric textbox -> ", evt.srcElement.id);
        return;
    }

    if(evt.type == "keypress") {
        if(elem.id == "tabWidth")
            previousWidth = elem.value;
        else if(elem.id == "tabHeight")
            previousHeight = elem.value;

        let charCode = (evt.which) ? evt.which : evt.keyCode;
        if (charCode == 46 || (charCode > 47 && charCode < 58)) {
            return true;
        }
        return false;
    }
    else if(evt.type == "keyup") {
        let floatVal = parseFloat(elem.value);
        let selPos = elem.selectionEnd - 1;
        if(selPos < 0)
            selPos = 0;
        
        if((isNaN(floatVal) || floatVal != elem.value) && !(elem.value.length==0 || elem.value == ".")) {
            if(elem.id == "tabWidth")
                elem.value = previousWidth;
            else if(elem.id == "tabHeight")
                elem.value = previousHeight;

            if(elem.value.length > selPos)
                elem.selectionEnd = selPos;
            else
                elem.selectionEnd = elem.value.length;
        }

        return !isNaN(floatVal);
    }
}

function loadPage() {
    const butProcess = document.getElementById("cmdProcess");

    if (!butProcess) {
        console.log("cmdProcess not loaded yet - come back in 0.1 seconds");
        setTimeout(function () { loadPage(); }, 100);
        return;
    }

    // REF: https://web.dev/file-system-access/
    butProcess.addEventListener('click', async () => {
        [fileHandle] = await window.showOpenFilePicker(pickerOpts);
        const file = await fileHandle.getFile();
        parseExcel(file);
        // console.log("fileHandle -> ", fileHandle);
    });
}

// REF: https://stackoverflow.com/questions/8238407/how-to-parse-excel-xls-file-in-javascript-html5
function parseExcel(file) {
    var reader = new FileReader();

    reader.onload = function(e) {
        const tabWidthTextBox = document.getElementById("tabWidth");
        const tabHeightTextBox = document.getElementById("tabHeight");

        if(!tabWidthTextBox) {
            console.log("Did not find Width textbox");
            return;
        }
        if(!tabHeightTextBox) {
            console.log("Did not find Height textbox");
            return;
        }

        const data = e.target.result;
        const workbook = XLSX.read(data, {
            type: 'binary'
        });

        // console.log("workbook -> ", workbook);
        // REF: https://masteringjs.io/tutorials/fundamentals/foreach-continue
        // console.log("workbook.Workbook.Sheets[\"Tabs\"] -> ", workbook.Workbook.Sheets.filter(v => v.name.toLowerCase() == "tabs"));
        workbook.SheetNames.forEach(function(sheetName) {
            const XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

            // console.log("sheetName -> ", sheetName);
            if(sheetName.toLowerCase() == "tabs") {
                worksheet = XL_row_object.sort(sortTabs);
            }
            else if(sheetName.toLowerCase() == "size") {
                let sizeWorksheet = XL_row_object;

                if(sizeWorksheet.length > 0) {
                    tabWidthTextBox.value = sizeWorksheet[0].width;
                    tabHeightTextBox.value = sizeWorksheet[0].height;
                }
            }
            
            // console.log("XL_row_object.length -> ", XL_row_object.length);
            // console.log("XL_row_object -> ", XL_row_object);
        });

        createFolderTabs();
    };

    reader.onerror = function(ex) {
        console.log(ex);
    };

    reader.readAsBinaryString(file);
}

function sortTabs(a, b) {
    const groupOrderA = a.groupOrder === undefined ? 999 : parseInt(a.groupOrder);
    const groupOrderB = b.groupOrder === undefined ? 999 : parseInt(b.groupOrder);
    const tabGroupA = a.tabGroup === undefined ? "" : a.tabGroup;
    const tabGroupB = b.tabGroup === undefined ? "" : b.tabGroup;
    const tabSubGroupA = a.tabSubGroup === undefined ? "" : a.tabSubGroup;
    const tabSubGroupB = b.tabSubGroup === undefined ? "" : b.tabSubGroup;
    const tabTitleA = a.tabTitle === undefined ? "" : a.tabTitle;
    const tabTitleB = b.tabTitle === undefined ? "" : b.tabTitle;

	if (groupOrderA < groupOrderB)
		return -1;
	else if (groupOrderA > groupOrderB)
		return 1;
	else if (tabGroupA < tabGroupB)
		return -1;
	else if (tabGroupA > tabGroupB)
		return 1;
    else if (tabSubGroupA < tabSubGroupB)
        return -1;
    else if (tabSubGroupA > tabSubGroupB)
        return 1;
    else if (tabTitleA < tabTitleB)
        return -1;
    else if (tabTitleA > tabTitleB)
        return 1;
	else
		return 0;
}


loadPage();