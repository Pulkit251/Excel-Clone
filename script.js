const ps = new PerfectScrollbar('#cells' ,{
    wheelSpeed: 4
});

for (let i = 1; i <= 100; i++) {
    let str = "";
    let n = i;

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = "Z" + str;
            n = Math.floor(n / 26) - 1;
        } else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor(n / 26);
        }
    }
    $("#columns").append(`<div class="column-name">${str}</div>`)
    $("#rows").append(`<div class="row-name">${i}</div>`)
}

let celldata = {
    "Sheet1" : {}
}

let selectedsheet = "Sheet1";
let totalsheets = 1;
let lastlyAddedSheet = 1;

let defaultProperties = {
    "font-family" : "Calibri",
    "font-size" : "14",
    "text" : "",
    "bold" : false,
    "italic" : false,
    "underlined" : false,
    "alignment" : "left",
    "color" : "#444",
    "bgcolor" : "#fff"
};
for(let i = 1;i <= 100;i++){
    let row = $(`<div class="cell-row"></div>`);   
    for(let j = 1;j <= 100;j++){
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);        
    }    
    $("#cells").append(row);
}

$("#cells").scroll(function(e){
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
})

$(".input-cell").dblclick(function(e){    
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    $(this).addClass("selected");
    $(this).attr("contenteditable",true);  
    $(this).focus();
});

$(".input-cell").blur(function(e){
    $(this).attr("contenteditable",false);
    updateCellData("text",$(this).text());
});

function getRowCol(ele){
    let id = $(ele).attr("id");
    let idArray = id.split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId,colId];
}

function getTopLeftBottomRightCell(rowId,colId){
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);
    return [topCell,bottomCell,leftCell,rightCell];
}

$(".input-cell").click(function(e){
    let [rowId,colId] = getRowCol(this);
    let [topCell,bottomCell,leftCell,rightCell] = getTopLeftBottomRightCell(rowId,colId);
    if($(this).hasClass("selected") && e.ctrlKey){
        unselectCell(this,e,topCell,bottomCell,leftCell,rightCell);
    }
    else{
        selectCell(this,e,topCell,bottomCell,leftCell,rightCell);
    }    
});

function unselectCell(ele,e,topCell,bottomCell,leftCell,rightCell){
    if($(ele).attr("contenteditable") == "false"){
        if($(ele).hasClass("top-selected")){
            topCell.removeClass("bottom-selected");
        }
    
        if($(ele).hasClass("bottom-selected")){
            bottomCell.removeClass("top-selected");
        }
    
        if($(ele).hasClass("left-selected")){
            leftCell.removeClass("right-selected");
        }
    
        if($(ele).hasClass("right-selected")){
            rightCell.removeClass("left-selected");
        }
    
        $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
    
}

function selectCell(ele,e,topCell,bottomCell,leftCell,rightCell){
    if(e.ctrlKey){

        let topSelected;
        if(topCell){
            topSelected = topCell.hasClass("selected");
        }

        let bottomSelected;
        if(bottomCell){
            bottomSelected = bottomCell.hasClass("selected");
        }

        let leftSelected;
        if(leftCell){
            leftSelected = leftCell.hasClass("selected");
        }

        let rightSelected;
        if(rightCell){
            rightSelected = rightCell.hasClass("selected");
        }

        if(topSelected){
            $(ele).addClass("top-selected");
            topCell.addClass("bottom-selected");
        }

        if(bottomSelected){
            $(ele).addClass("bottom-selected");
            bottomCell.addClass("top-selected");
        }

        if(leftSelected) {
            $(ele).addClass("left-selected");
            leftCell.addClass("right-selected");
        }

        if(rightSelected) {
            $(ele).addClass("right-selected");
            rightCell.addClass("left-selected");
        }
    }
    else{
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
    $(ele).addClass("selected");
    changeHeader(getRowCol(ele));
}

function changeHeader([rowId,colId]){
    let data;
    if(celldata[selectedsheet][rowId - 1] && celldata[selectedsheet][rowId - 1][colId - 1]){
        data = celldata[selectedsheet][rowId - 1][colId - 1];
    }
    else{
        data = defaultProperties;
    }
    $(".alignemnt.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");    
    addRemoveProperty(data,"bold");
    addRemoveProperty(data,"italic");
    addRemoveProperty(data,"underlined");
    $("#fill-color").css("border-bottom",`4px solid ${data.bgcolor}`);
    $("#text-color").css("border-bottom",`4px solid ${data.color}`);
    $("#font-family").val(data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $("#font-family").css("font-family",data["font-family"]);
}

function addRemoveProperty(data,property){
    if(data[property]){
        $(`#${property}`).addClass("selected");
    }else{
        $(`#${property}`).removeClass("selected");
    }
}

let startCellSelected = false;
let startCell = {};
let endCell = {}; 
let scrollXRstarted = false;
let scrollXLstarted = false;
$(".input-cell").mousemove(function(e){
    e.preventDefault();
    if(e.buttons == 1){  
        if(e.pageX > ($(window).width() - 10) && !scrollXRstarted){
            scrollXR();
        }      
        else if(e.pageX < (10) && !scrollXLstarted){
            scrollXL();
        }
        if(!startCellSelected){            
            let[rowId,colId] = getRowCol(this);
            startCell = {"rowId": rowId, "colId": colId};
            selectBetweenAllCells(startCell,startCell);
            startCellSelected = true;
        }        
    }
    else{
        startCellSelected = false;
    }    
});

$(".input-cell").mouseenter(function(e){
    if(e.buttons == 1){
        if(e.pageX < ($(window).width() - 10) && scrollXRstarted){
            clearInterval(scrollXRinterval);
            scrollXRstarted = false;
        }

        if(e.pageX > 10 && scrollXLstarted){
            clearInterval(scrollXLinterval);
            scrollXLstarted = false;
        }
        let[rowId,colId] = getRowCol(this);
        endCell = {"rowId" : rowId, "colId" : colId};
        selectBetweenAllCells(startCell,endCell);
    }
})

function selectBetweenAllCells(start,end){
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    for(let i = Math.min(start.rowId,end.rowId);i <= Math.max(start.rowId,end.rowId);i++){
        for(let j = Math.min(start.colId,end.colId);j <= Math.max(start.colId,end.colId);j++){
            let [topCell,bottomCell,leftCell,rightCell] = getTopLeftBottomRightCell(i,j);
            selectCell($(`#row-${i}-col-${j}`)[0],{"ctrlKey" : true},topCell,bottomCell,leftCell,rightCell);
        }
    }
}

let scrollXRinterval;
let scrollXLinterval;
function scrollXR(){
    scrollXRstarted = true;
     scrollXRinterval = setInterval(() => {
         $("#cells").scrollLeft($("#cells").scrollLeft() + 100);
     }, 100);
}

function scrollXL(){
    scrollXLstarted = true;
    scrollXLinterval = setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() - 100);
    }, 100);
}

$(".data-container").mousemove(function(e){
    e.preventDefault();
    if(e.buttons == 1){  
        if(e.pageX > ($(window).width() - 10) && !scrollXRstarted){
            scrollXR();
        }      
        else if(e.pageX < (10) && !scrollXLstarted){
            scrollXL();
        }
    }
});

$(".data-container").mouseup(function(e){
    clearInterval(scrollXRinterval);
    clearInterval(scrollXLinterval);
    scrollXRstarted = false;
    scrollXLstarted = false;
});

$(".alignment").click(function(e){
    let alignment = $(this).attr("data-type");
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    $(".input-cell.selected").css("text-align",alignment);
    // $(".input-cell.selected").each(function(index,data){
    //     let [rowId,colId] = getRowCol(data);
    //     celldata[rowId - 1][colId - 1].alignment = alignment;
    // });
    updateCellData("alignment",alignment);
});

$("#bold").click(function(e){
   setStyle(this,"bold","font-weight","bold");
});

$("#italic").click(function(e){
    setStyle(this,"italic","font-style","italic");
});

$("#underlined").click(function(e){
    setStyle(this,"underlined","text-decoration","underline");
});

function setStyle(ele,property,key,value){
    if($(ele).hasClass("selected")){
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key,"");
        // $(".input-cell.selected").each(function(index,data){
        //  let[rowId,colId] = getRowCol(data);
        //  celldata[rowId - 1][colId - 1][property] = false;
        // });
        updateCellData(property,false);
    }
    else{
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key,value);
        // $(".input-cell.selected").each(function(index,data){
        //     let[rowId,colId] = getRowCol(data);
        //     celldata[rowId - 1][colId - 1][property] = true;
        // });
        updateCellData(property,true)
    }
}

$(".pick-color").colorPick({
    'initialColor': "#abcd",
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function(){
        if(this.color != "#ABCD") {
            if($(this.element.children()[1]).attr("id") == "fill-color") {
                $(".input-cell.selected").css("background-color",this.color);
                $("#fill-color").css("border-bottom",`4px solid ${this.color}`);
                // $(".input.cell-selected").each((index,data) =>{
                //     let[rowId,colId] = getRowCol(data);
                //     celldata[rowId - 1][colId - 1].bgcolor = this.color;
                // });
                updateCellData("bgcolor",this.color);
            }
            if($(this.element.children()[1]).attr("id") == "text-color") {
                $(".input-cell.selected").css("color",this.color);
                $("#text-color").css("border-bottom",`4px solid ${this.color}`);
                // $(".input.cell-selected").each((index,data) => {
                //     let[rowId,colId] = getRowCol(data);
                //     celldata[rowId - 1][colId - 1].color = this.color;
                // });
                updateCellData("color",this.color);
            }
        }
    }
});

$("#fill-color").click(function(e){
    setTimeout(() => {
        $(this).parent().click();
    }, 10);    
});

$("#text-color").click(function(e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
}); 

$(".menu-selector").change(function(e){
    let value = $(this).val();
    let key = $(this).attr("id");
    if(key == "font-family"){
        $("#font-family").css(key,value);
    }

    if(!isNaN(value)){
        value = parseInt(value);
    }

    $(".input-cell.selected").css(key,value);
    // $(".input-cell.selected").each((index,data) =>{
    //     let[rowId,colId] = getRowCol(data);
    //     celldata[rowId - 1][colId - 1][key] = value;
    // })
    updateCellData(key,value);
})

function updateCellData(property,value){
    if(value != defaultProperties[property]){
        $(".input-cell.selected").each(function(index,data){
            let[rowId,colId] = getRowCol(data);
            if(celldata[selectedsheet][rowId - 1] == undefined){
                celldata[selectedsheet][rowId - 1] = {};
                celldata[selectedsheet][rowId - 1][colId - 1] = {...defaultProperties};
                celldata[selectedsheet][rowId - 1][colId - 1][property] = value;
            }
            else{
                if(celldata[selectedsheet][rowId - 1][colId - 1] == undefined){
                    celldata[selectedsheet][rowId - 1][colId - 1] = {...defaultProperties};
                    celldata[selectedsheet][rowId - 1][colId - 1][property] = value;
                }
                else{
                    celldata[selectedsheet][rowId - 1][colId - 1][property] = value; 
                }
            }
        })
    }
    else{
        $(".input-cell.selected").each(function(index,data){
            let[rowId,colId] = getRowCol(data);
            if(celldata[selectedsheet][rowId - 1][colId - 1] != undefined){
                celldata[selectedsheet][rowId - 1][colId - 1][property] = value; 
                if(JSON.stringify(celldata[selectedsheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)){
                    delete celldata[selectedsheet][rowId - 1][colId - 1];
                    if(Object.keys(celldata[selectedsheet][rowId - 1]).length == 0){
                        delete celldata[selectedsheet][rowId - 1];
                    }
                }         
                
            }
        });       

    }
}

$(".container").click(function(e){
    $(".sheet-options-modal").remove();
})
function addSheetEvents(){
    $(".sheet-tab.selected").on("contextmenu", function(e){
        e.preventDefault();
        selectSheet(this); 
        $(".sheet-options-modal").remove();
        let modal = $(`<div class="sheet-options-modal">
                                <div class="option sheet-rename">Rename</div>
                                <div class="option sheet-delete">Delete</div>
                        </div>`);
        modal.css({"left": e.pageX});
        $(".container").append(modal);
        $(".sheet-rename").click(function(e){
            let renameModal = $(`<div class="sheet-modal-parent">
                            <div class="sheet-rename-modal">
                                <div class="sheet-modal-title">Rename Sheet</div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">Rename Sheet to:</span>
                                    <input class="sheet-modal-input" type="text">
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button yes-button">OK</div>
                                    <div class="button no-button">Cancel</div>
                                </div>
                            </div>            
                        </div>`);
            $(".container").append(renameModal);
            $(".no-button").click(function(e){
                $(".sheet-modal-parent").remove();
            });
            $(".yes-button").click(function(e){
                renameSheet();
            })
            $(".sheet-modal-input").Keypress(function(e){
                if(e.key == enter){
                    renameSheet();
                }
            });
        })

        $(".sheet-delete").click(function (e) {
            if (totalsheets > 1) {
                let deleteModal = $(`<div class="sheet-modal-parent">
                                    <div class="sheet-delete-modal">
                                        <div class="sheet-modal-title">Sheet Name</div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Are you sure?</span>
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button yes-button">
                                                <div class="material-icons delete-icon">delete</div>
                                                Delete
                                            </div>
                                            <div class="button no-button">Cancel</div>
                                        </div>
                                    </div>
                                </div>`);

                $(".container").append(deleteModal);
                $(".no-button").click(function (e) {
                    $(".sheet-modal-parent").remove();
                });
                $(".yes-button").click(function (e) {
                    deleteSheet();
                });
            } else {
                alert("Not possible");
            }
        })
    });

    $(".sheet-tab.selected").click(function (e) {
        selectSheet(this);
    });
}



addSheetEvents();

$(".add-sheet").click(function(e){
    lastlyAddedSheet++;
    totalsheets++;
    celldata[`Sheet${lastlyAddedSheet}`] = {};
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet${lastlyAddedSheet}</div>`);
    selectSheet();
    addSheetEvents();
});

function selectSheet(ele) {
    if (ele && !$(ele).hasClass("selected")) {
        $(".sheet-tab.selected").removeClass("selected");
        $(ele).addClass("selected");
    }
    emptyPreviousSheet();
    selectedsheet = $(".sheet-tab.selected").text();
    loadCurrentSheet();
    $("#row-1-col-1").click();
}

function emptyPreviousSheet(){
    let data = celldata[selectedsheet];
    let rowkeys = Object.keys(data);
    for(let i of rowkeys){
        let rowId = parseInt(i);
        let colkeys = Object.keys(data[rowId]);
        for(let j of colkeys){
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                "font-family" : "Calibri",
                "font-size" : 14,
                "background-color" : "#fff",
                "color" : "#444",
                "font-weight" : "",
                "font-style" : "",
                "text-decoration" : "",
                "text-align" : "left"
            });
        }
    }
}

function loadCurrentSheet() {
    let data = celldata[selectedsheet];
    let rowkeys = Object.keys(data);
    for(let i of rowkeys) {
        let rowId = parseInt(i);
        let colkeys = Object.keys(data[rowId]);
        for(let j of colkeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId+1}-col-${colId+1}`);
            cell.text(data[rowId][colId].text);
            cell.css({
                "font-family" : data[rowId][colId]["font-family"],
                "font-size" : data[rowId][colId]["font-size"],
                "background-color" : data[rowId][colId]["bgcolor"],
                "color" : data[rowId][colId].color,
                "font-weight" : data[rowId][colId].bold ? "bold" : "",
                "font-style" : data[rowId][colId].italic ? "italic" : "",
                "text-decoration" : data[rowId][colId].underlined ? "underline" : "",
                "text-align" : data[rowId][colId].alignment
            });
        }
    }
} 

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(celldata).includes(newSheetName)) {
        let newCelldata = {};
        for (let i of Object.keys(celldata)) {
            if (i == selectedsheet) {
                newCelldata[newSheetName] = celldata[selectedsheet];
            } else {
                newCelldata[i] = celldata[i];
            }
        }

        celldata = newCelldata;

        selectedsheet = newSheetName;
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
    } else {
        $(".rename-error").remove();
        $(".sheet-modal-input-container").append(`
            <div class="rename-error"> Sheet Name is not valid or Sheet already exists! </div>
        `);
    }
}

function deleteSheet() {
    $(".sheet-modal-parent").remove();
    let sheetIndex = Object.keys(celldata).indexOf(selectedsheet);
    let currselectedsheet = $(".sheet-tab.selected");
    if (sheetIndex == 0) {
        selectSheet(currselectedsheet.next()[0]);
    } else {
        selectSheet(currselectedsheet.prev()[0]);
    }
    delete celldata[currselectedsheet.text()];
    currselectedsheet.remove();
    totalsheets--;
}

$(".left-scroller,.right-scroller").click(function(e){
    let keysArray = Object.keys(celldata);
    let selectedsheetIndex = keysArray.indexOf(selectedsheet);
    if(selectedsheetIndex != 0 && $(this).text() == "arrow_left"){
        selectSheet($(".sheet-tab.selected").prev()[0]);
    }
    else if(selectedsheetIndex != (keysArray.length - 1) && $(this).text() == "arrow_right"){
        selectSheet($(".sheet-tab.selected").next()[0]);
    }

    $(".sheet-tab.selected")[0].scrollIntoView();
})

$("#menu-file").click(function(e){
    let fileModal = $(`<div class="file-modal">
                        <div class="file-options-modal">
                            <div class="close">
                                <div class="material-icons close-icon">arrow_circle_down</div>
                                <div>Close</div>
                            </div>
                            <div class="new">
                                <div class="material-icons new-icon">insert_drive_file</div>
                                <div>New</div>
                            </div>
                            <div class="open">
                                <div class="material-icons open-icon">folder_open</div>
                                <div>Open</div>
                            </div>
                            <div class="save">
                                <div class="material-icons save-icon">save</div>
                                <div>Save</div>
                            </div>
                        </div>
                        <div class="file-recent-modal"></div>
                        <div class="file-transparent"></div>
                    </div>`);
    $(".container").append(fileModal);
    fileModal.animate({
        width: "100vw"
    },300);
    $(".close,.file-transparent,.new,.save,.open").click(function(e){
        fileModal.animate({
            width: "0vw"
        })
        setTimeout(() => {
            fileModal.remove();
        }, 250);
    });
    $(".new").click(function (e) {
        if (save) {
            newFile();
        } else {
            $(".container").append(`<div class="sheet-modal-parent">
                                        <div class="sheet-delete-modal">
                                            <div class="sheet-modal-title">${$(".title").text()}</div>
                                            <div class="sheet-modal-detail-container">
                                                <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                                            </div>
                                            <div class="sheet-modal-confirmation">
                                                <div class="button yes-button">
                                                    Yes
                                                </div>
                                                <div class="button no-button">No</div>
                                            </div>
                                        </div>
                                    </div>`);
            $(".no-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                newFile();
            });
            $(".yes-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                saveFile(true);
            });
        }
    });
    $(".save").click(function (e) {
        if (!save) {
            saveFile();
        }
    });

    $(".open").click(function (e) {
        openFile();
    })
});

function newFile() {
    emptyPreviousSheet();
    celldata = { "Sheet1": {} };
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
    addSheetEvents();
    selectedsheet = "Sheet1";
    totalsheets = 1;
    lastlyAddedSheet = 1;
    $(".title").text("Book1-Excel");
    $("#row-1-col-1").click();
}

function saveFile(newClicked) {
    $(".container").append(`<div class="sheet-modal-parent">
                                <div class="sheet-rename-modal">
                                    <div class="sheet-modal-title">Save File</div>
                                    <div class="sheet-modal-input-container">
                                        <span class="sheet-modal-input-title">File Name:</span>
                                        <input class="sheet-modal-input" value="${$(".title").text()}" type="text" />
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="button yes-button">Save</div>
                                        <div class="button no-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`);
    $(".yes-button").click(function (e) {
        $(".title").text($(".sheet-modal-input").val());
        let a = document.createElement("a");
        a.href = `data:application/json,${encodeURIComponent(JSON.stringify(celldata))}`;
        a.download = $(".title").text() + ".json";
        $(".container").append(a);
        a.click();
        // a.remove();
        save = true;

    });
    $(".no-button,.yes-button").click(function (e) {
        $(".sheet-modal-parent").remove();
        if (newClicked) {
            newFile();
        }
    });
}

function openFile() {
    let inputFile = $(`<input accept="application/json" type="file" />`);
    $(".container").append(inputFile);
    inputFile.click();
    inputFile.change(function (e) {
        console.log(inputFile.val());
        let file = e.target.files[0];
        $(".title").text(file.name.split(".json")[0]);
        let reader = new FileReader();
        reader.readAsText(file);
        reader.onload = () => {
            emptyPreviousSheet();
            $(".sheet-tab").remove();
            celldata = JSON.parse(reader.result);
            let sheets = Object.keys(celldata);
            lastlyAddedSheet = 1;
            for (let i of sheets) {
                if (i.includes("Sheet")) {
                    let splittedSheetArray = i.split("Sheet");
                    if (splittedSheetArray.length == 2 && !isNaN(splittedSheetArray[1])) {
                        lastlyAddedSheet = parseInt(splittedSheetArray[1]);
                    }
                }
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`);
            }
            addSheetEvents();
            $(".sheet-tab").removeClass("selected");
            $($(".sheet-tab")[0]).addClass("selected");
            selectedsheet = sheets[0];
            totalsheets = sheets.length;
            loadCurrentSheet();
            inputFile.remove();
        }
    });
}

let clipboard = { startCell: [], celldata: {} };
let contentCutted = false;
$("#cut,#copy").click(function (e) {
    if ($(this).text() == "content_cut") {
        contentCutted = true;
    }
    clipboard = { startCell: [], celldata: {} };
    clipboard.startCell = getRowCol($(".input-cell.selected")[0]);
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = getRowCol(data);
        if (celldata[selectedsheet][rowId - 1] && celldata[selectedsheet][rowId - 1][colId - 1]) {
            if (!clipboard.celldata[rowId]) {
                clipboard.celldata[rowId] = {};
            }
            clipboard.celldata[rowId][colId] = { ...celldata[selectedsheet][rowId - 1][colId - 1] };
        }
    });
    console.log(clipboard);
});

$("#paste").click(function (e) {
    if (contentCutted) {
        emptyPreviousSheet();
    }
    let startCell = getRowCol($(".input-cell.selected")[0]);
    let rows = Object.keys(clipboard.celldata);
    for (let i of rows) {
        let cols = Object.keys(clipboard.celldata[i]);
        for (let j of cols) {
            if (contentCutted) {
                delete celldata[selectedsheet][i - 1][j - 1];
                if (Object.keys(celldata[selectedsheet][i - 1]).length == 0) {
                    delete celldata[selectedsheet][i - 1];
                }
            }

        }
    }
    for (let i of rows) {
        let cols = Object.keys(clipboard.celldata[i]);
        for (let j of cols) {
            let rowDistance = parseInt(i) - parseInt(clipboard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipboard.startCell[1]);
            if (!celldata[selectedsheet][startCell[0] + rowDistance - 1]) {
                celldata[selectedsheet][startCell[0] + rowDistance - 1] = {};
            }
            celldata[selectedsheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipboard.celldata[i][j] };
        }
    }
    loadCurrentSheet();
    if (contentCutted) {
        contentCutted = false;
        clipboard = { startCell: [], celldata: {} };
    }
});

$("#formula-input").blur(function (e) {
    if ($(".input-cell.selected").length > 0) {
        let formula = $(this).text();
        let tempElements = formula.split(" ");
        let elements = [];
        for (let i of tempElements) {
            if (i.length >= 2) {
                i = i.replace("(", "");
                i = i.replace(")", "");
                if (!elements.includes(i)) {
                    elements.push(i);
                }
            }
        }
        $(".input-cell.selected").each(function (index, data) {
            if (updateStreams(data, elements, false)) {
                let [rowId, colId] = getRowCol(data);
                celldata[selectedsheet][rowId - 1][colId - 1].formula = formula;
                let selfColCode = $(`.column-${colId}`).attr("id");
                evalFormula(selfColCode + rowId);
            } else {
                alert("Formula is not valid");
            }
        })
    } else {
        alert("!Please select a cell First");
    }
});

function updateStreams(ele, elements, update, oldUpstream) {
    let [rowId, colId] = getRowCol(ele);
    let selfColCode = $(`.column-${colId}`).attr("id");
    if (elements.includes(selfColCode + rowId)) {
        return false;
    }
    if (celldata[selectedsheet][rowId - 1] && celldata[selectedsheet][rowId - 1][colId - 1]) {
        let downStream = celldata[selectedsheet][rowId - 1][colId - 1].downStream;
        let upStream = celldata[selectedsheet][rowId - 1][colId - 1].upStream;
        for (let i of downStream) {
            if (elements.includes(i)) {
                return false;
            }
        }
        for (let i of downStream) {
            let [calRowId, calColId] = codeToValue(i);
            console.log(updateStreams($(`#row-${calRowId}-col-${calColId}`)[0], elements, true, upStream));
        }
    }

    if (!celldata[selectedsheet][rowId - 1]) {
        celldata[selectedsheet][rowId - 1] = {};
        celldata[selectedsheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else if (!celldata[selectedsheet][rowId - 1][colId - 1]) {
        celldata[selectedsheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else {

        let upStream = [...celldata[selectedsheet][rowId - 1][colId - 1].upStream];
        if (update) {
            for (let i of oldUpstream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = celldata[selectedsheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                celldata[selectedsheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(celldata[selectedsheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete celldata[selectedsheet][calRowId - 1][calColId - 1];
                    if (Object.keys(celldata[selectedsheet][calRowId - 1]).length == 0) {
                        delete celldata[selectedsheet][calRowId - 1];
                    }
                }
                index = celldata[selectedsheet][rowId - 1][colId - 1].upStream.indexOf(i);
                celldata[selectedsheet][rowId - 1][colId - 1].upStream.splice(index, 1);
            }
            for (let i of elements) {
                celldata[selectedsheet][rowId - 1][colId - 1].upStream.push(i);
            }
        } else {
            for (let i of upStream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = celldata[selectedsheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                celldata[selectedsheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(celldata[selectedsheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete celldata[selectedsheet][calRowId - 1][calColId - 1];
                    if (Object.keys(celldata[selectedsheet][calRowId - 1]).length == 0) {
                        delete celldata[selectedsheet][calRowId - 1];
                    }
                }
            }
            celldata[selectedsheet][rowId - 1][colId - 1].upStream = [...elements];
        }
    }

    for (let i of elements) {
        let [calRowId, calColId] = codeToValue(i);
        if (!celldata[selectedsheet][calRowId - 1]) {
            celldata[selectedsheet][calRowId - 1] = {};
            celldata[selectedsheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else if (!celldata[selectedsheet][calRowId - 1][calColId - 1]) {
            celldata[selectedsheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else {
            celldata[selectedsheet][calRowId - 1][calColId - 1].downStream.push(selfColCode + rowId);
        }
    }
    console.log(celldata);
    return true;

}

function codeToValue(code) {
    let colCode = "";
    let rowCode = "";
    for (let i = 0; i < code.length; i++) {
        if (!isNaN(code.charAt(i))) {
            rowCode += code.charAt(i);
        } else {
            colCode += code.charAt(i);
        }
    }
    let colId = parseInt($(`#${colCode}`).attr("class").split(" ")[1].split("-")[1]);
    let rowId = parseInt(rowCode);
    return [rowId, colId];
}

function evalFormula(cell) {
    let [rowId, colId] = codeToValue(cell);
    let formula = celldata[selectedsheet][rowId - 1][colId - 1].formula;
    console.log(formula);
    if (formula != "") {
        let upStream = celldata[selectedsheet][rowId - 1][colId - 1].upStream;
        let upStreamValue = [];
        for (let i in upStream) {
            let [calRowId, calColId] = codeToValue(upStream[i]);
            let value;
            if (celldata[selectedsheet][calRowId - 1][calColId - 1].text == "") {
                value = "0";
            }
             else {
                value = celldata[selectedsheet][calRowId - 1][calColId - 1].text;
            }
            upStreamValue.push(value);
            console.log(upStreamValue);
            formula = formula.replace(upStream[i], upStreamValue[i]);
        }
        celldata[selectedsheet][rowId - 1][colId - 1].text = eval(formula);
        loadCurrentSheet();
    }
    let downStream = celldata[selectedsheet][rowId - 1][colId - 1].downStream;
    for (let i = downStream.length - 1; i >= 0; i--) {
        evalFormula(downStream[i]);
    }

}