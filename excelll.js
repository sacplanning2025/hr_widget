(function() {

let _shadowRoot;
let _id;
let _result;
let div;
let widgetName;
var Ar = [];

let CONFIG = {
MAX_ROWS: 2000,
MAX_FILE_SIZE: 10 * 1024 * 1024,
ALLOWED_TYPES: ["xlsx","xlsm","xls"],
DEBUG: true
};

let uploadStats = {
totalRows:0,
validRows:0,
duplicateRows:0,
emptyRows:0
};

let tmpl = document.createElement("template");

tmpl.innerHTML = `
<style>

.upload-wrapper{
padding:10px;
border:1px solid #dcdcdc;
border-radius:6px;
background:#fafafa;
font-family:Arial;
}

.drag-area{
border:2px dashed #0070f2;
padding:20px;
text-align:center;
margin-bottom:10px;
border-radius:6px;
background:#ffffff;
cursor:pointer;
}

.drag-area.dragover{
background:#e8f2ff;
}

.progress{
width:100%;
height:8px;
background:#ddd;
border-radius:5px;
margin-top:8px;
}

.progress-bar{
width:0%;
height:100%;
background:#0070f2;
transition:width 0.4s;
}

.file-info{
font-size:11px;
color:#666;
margin-top:4px;
}

.stats{
font-size:11px;
margin-top:6px;
color:#444;
}

</style>

<div class="upload-wrapper">

<div class="drag-area" id="dragArea">
Drop Excel File Here or Click to Upload
</div>

<div class="file-info" id="fileInfo"></div>

<div class="progress">
<div class="progress-bar" id="progressBar"></div>
</div>

<div class="stats" id="uploadStats"></div>

</div>
`;

class Excel extends HTMLElement {

constructor(){

super();

_shadowRoot = this.attachShadow({mode:"open"});
_shadowRoot.appendChild(tmpl.content.cloneNode(true));

_id = createGuid();

this._export_settings={};
this._firstConnection=0;

this.initializeDragDrop();

}

initializeDragDrop(){

setTimeout(()=>{

let dragArea=_shadowRoot.getElementById("dragArea");

if(!dragArea) return;

dragArea.addEventListener("dragover",(e)=>{
e.preventDefault();
dragArea.classList.add("dragover");
});

dragArea.addEventListener("dragleave",(e)=>{
dragArea.classList.remove("dragover");
});

dragArea.addEventListener("drop",(e)=>{

e.preventDefault();
dragArea.classList.remove("dragover");

let file=e.dataTransfer.files[0];
this.handleFile(file);

});

},500);

}

handleFile(file){

let validation=validateFile(file);

if(!validation.valid){
alert(validation.msg);
return;
}

let info=_shadowRoot.getElementById("fileInfo");

if(info){
info.innerHTML=
"File: "+file.name+
"<br>Size: "+Math.round(file.size/1024)+" KB";
}

this.readExcel(file);

}

readExcel(file){

updateProgress(10);

let reader=new FileReader();

reader.onload=(e)=>{

let data=e.target.result;

updateProgress(30);

let workbook=XLSX.read(data,{type:'binary'});

let rows=[];

workbook.SheetNames.forEach(sheet=>{
let csv=XLSX.utils.sheet_to_json(workbook.Sheets[sheet],{header:1});
rows=rows.concat(csv);
});

updateProgress(60);

rows=this.cleanRows(rows);

uploadStats.totalRows=rows.length;

let seenIDs=new Set();
let finalRows=[];

rows.forEach(r=>{

if(r.join("").trim()===""){
uploadStats.emptyRows++;
return;
}

if(seenIDs.has(r[0])){
uploadStats.duplicateRows++;
return;
}

seenIDs.add(r[0]);
finalRows.push(r);

});

uploadStats.validRows=finalRows.length;

_result=JSON.stringify(finalRows);

updateStats();

updateProgress(100);

this.dispatchEvent(new CustomEvent("onUploadSuccess",{
detail:{
records:finalRows,
stats:uploadStats
}
}));

};

reader.readAsBinaryString(file);

}

cleanRows(rows){

let cleaned=[];

rows.forEach(r=>{
if(r.join("").trim()!=="")
cleaned.push(r);
});

return cleaned;

}

}

customElements.define("com-fd-djaja-sap-sac-excelll",Excel);

function validateFile(file){

if(!file) return {valid:false,msg:"No file selected"};

if(file.size>CONFIG.MAX_FILE_SIZE)
return {valid:false,msg:"File too large"};

let ext=file.name.split(".").pop().toLowerCase();

if(!CONFIG.ALLOWED_TYPES.includes(ext))
return {valid:false,msg:"Invalid file type"};

return {valid:true};

}

function updateProgress(percent){

let bar=_shadowRoot.getElementById("progressBar");

if(bar)
bar.style.width=percent+"%";

}

function updateStats(){

let stats=_shadowRoot.getElementById("uploadStats");

if(stats){

stats.innerHTML=
"Total Rows: "+uploadStats.totalRows+
"<br>Valid Rows: "+uploadStats.validRows+
"<br>Duplicate Rows: "+uploadStats.duplicateRows+
"<br>Empty Rows: "+uploadStats.emptyRows+
"<br>Time: "+new Date().toLocaleTimeString();

}

}

function createGuid(){

return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g,function(c){

let r=Math.random()*16|0;
let v=c==="x"?r:(r&0x3|0x8);

return v.toString(16);

});

}

function loadScript(src,shadowRoot){

return new Promise(function(resolve,reject){

let script=document.createElement("script");
script.src=src;

script.onload=()=>{
console.log("Load: "+src);
resolve(script);
};

script.onerror=()=>reject(new Error("Script load error"));

shadowRoot.appendChild(script);

});

}

let xlsxjs = "https://sacplanning2025.github.io/hr_widget/xlsxxx.js";

async function LoadLibs(){

try{
await loadScript(xlsxjs,_shadowRoot);
}
catch(e){
console.log(e);
}

}

})();
