(function () {

let _shadowRoot;
let _id;
let _result;
let div;
let widgetName;
let Ar = [];

let CONFIG = {
    MAX_ROWS: 5000,
    MAX_FILE_SIZE: 10 * 1024 * 1024,
    ALLOWED_TYPES: ["xlsx", "xlsm", "xls"],
    DEBUG: true,
    ENABLE_PREVIEW: true
};

let tmpl = document.createElement("template");

tmpl.innerHTML = `
<style>

:host{
font-family: SAP-icons, Arial;
}

.upload-container{
border:2px dashed #0070f2;
padding:25px;
border-radius:10px;
text-align:center;
transition:0.3s;
background:#fafafa;
}

.upload-container.dragover{
background:#e8f2ff;
border-color:#0040c1;
}

.upload-title{
font-size:18px;
font-weight:bold;
margin-bottom:10px;
}

.progress-bar{
height:8px;
background:#ddd;
border-radius:5px;
margin-top:10px;
overflow:hidden;
}

.progress-inner{
height:100%;
width:0%;
background:#0070f2;
transition:width 0.3s;
}

.preview-table{
margin-top:15px;
max-height:200px;
overflow:auto;
border:1px solid #ddd;
}

.preview-table table{
width:100%;
border-collapse:collapse;
}

.preview-table th{
background:#f5f5f5;
padding:5px;
}

.preview-table td{
padding:5px;
border-bottom:1px solid #eee;
}

.file-info{
font-size:12px;
color:#666;
margin-top:6px;
}

.mode-toggle{
cursor:pointer;
font-size:12px;
float:right;
color:#0070f2;
}

.dark-mode{
background:#1e1e1e;
color:white;
}

.dark-mode .upload-container{
background:#2b2b2b;
border-color:#666;
}

</style>

<div class="upload-container" id="dropZone">

<div class="upload-title">
Excel Upload Widget
<span class="mode-toggle" id="modeToggle">Toggle Theme</span>
</div>

<input type="file" id="fileInput" />

<div class="file-info" id="fileInfo"></div>

<div class="progress-bar">
<div class="progress-inner" id="progressBar"></div>
</div>

<div class="preview-table" id="preview"></div>

</div>
`;

class ExcelUploadWidget extends HTMLElement {

constructor(){

super();

_shadowRoot = this.attachShadow({mode:"open"});
_shadowRoot.appendChild(tmpl.content.cloneNode(true));

_id = createGuid();

this._export_settings = {};
this._firstConnection = 0;

this._initUI();

}

_initUI(){

let dropZone = _shadowRoot.getElementById("dropZone");
let fileInput = _shadowRoot.getElementById("fileInput");

dropZone.addEventListener("dragover",(e)=>{
e.preventDefault();
dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave",(e)=>{
dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop",(e)=>{
e.preventDefault();
dropZone.classList.remove("dragover");
let file = e.dataTransfer.files[0];
this.processFile(file);
});

fileInput.addEventListener("change",(e)=>{
let file = e.target.files[0];
this.processFile(file);
});

_shadowRoot.getElementById("modeToggle").addEventListener("click",()=>{
this.toggleTheme();
});

}

toggleTheme(){

let host = this;

if(host.classList.contains("dark-mode"))
host.classList.remove("dark-mode");
else
host.classList.add("dark-mode");

}

connectedCallback(){

console.log("Widget Connected");

}

processFile(file){

if(!file){
this.showMessage("No file selected");
return;
}

if(file.size > CONFIG.MAX_FILE_SIZE){
this.showMessage("File too large");
return;
}

let ext = file.name.split(".").pop().toLowerCase();

if(!CONFIG.ALLOWED_TYPES.includes(ext)){
this.showMessage("Invalid file type");
return;
}

this.updateFileInfo(file);

this.readExcel(file);

}

updateFileInfo(file){

let info = _shadowRoot.getElementById("fileInfo");

info.innerHTML =
"File : "+file.name+
"<br>Size : "+Math.round(file.size/1024)+" KB";

}

updateProgress(percent){

let bar = _shadowRoot.getElementById("progressBar");
bar.style.width = percent+"%";

}

readExcel(file){

this.updateProgress(10);

let reader = new FileReader();

reader.onload = (e)=>{

let data = e.target.result;

this.updateProgress(30);

let workbook = XLSX.read(data,{type:'binary'});

let sheetNames = workbook.SheetNames;

let rows = [];

sheetNames.forEach(sheet=>{

let csv = XLSX.utils.sheet_to_json(workbook.Sheets[sheet],{header:1});

rows = rows.concat(csv);

});

this.updateProgress(70);

rows = this.cleanRows(rows);

if(rows.length > CONFIG.MAX_ROWS){

this.showMessage("Too many rows");

return;

}

this.previewData(rows);

this.fireUploadEvent(rows);

this.updateProgress(100);

};

reader.readAsBinaryString(file);

}

cleanRows(rows){

let cleaned = [];

rows.forEach(r=>{

if(r.join("").trim() !== "")
cleaned.push(r);

});

return cleaned;

}

previewData(rows){

if(!CONFIG.ENABLE_PREVIEW) return;

let preview = _shadowRoot.getElementById("preview");

let html = "<table>";

rows.slice(0,10).forEach(row=>{

html+="<tr>";

row.forEach(col=>{
html+="<td>"+col+"</td>";
});

html+="</tr>";

});

html+="</table>";

preview.innerHTML = html;

}

fireUploadEvent(data){

_result = JSON.stringify(data);

this.dispatchEvent(new CustomEvent("onUploadSuccess",{
detail:{
data:data
}
}));

}

showMessage(msg){

alert(msg);

}

}

customElements.define("advanced-excel-upload",ExcelUploadWidget);

function createGuid(){

return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g,function(c){

let r = Math.random()*16|0;
let v = c === "x" ? r : (r&0x3|0x8);

return v.toString(16);

});

}

})();
