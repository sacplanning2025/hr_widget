(function() {

    let _shadowRoot;
    let _id;
    let _result ;

    let div;
    let widgetName;
    var Ar = [];

    // ================= ENHANCEMENT CONFIG =================
    var UploadConfig = {
        MAX_FILE_SIZE: 10 * 1024 * 1024,
        MAX_RECORDS: 2000,
        ALLOWED_TYPES: ["xlsm","xlsx","xls"],
        DEBUG:true
    };

    var UploadStats = {
        totalRows:0,
        validRows:0,
        duplicateRows:0,
        emptyRows:0
    };

    let tmpl = document.createElement("template");
    tmpl.innerHTML = `
      <style>

      .drag-zone{
        border:2px dashed #0070f2;
        padding:15px;
        border-radius:6px;
        text-align:center;
        margin-bottom:6px;
        cursor:pointer;
      }

      .drag-zone.dragover{
        background:#e8f2ff;
      }

      .upload-progress{
        width:100%;
        height:6px;
        background:#ddd;
        border-radius:4px;
        margin-top:6px;
      }

      .upload-progress-bar{
        height:100%;
        width:0%;
        background:#0070f2;
        transition:width .4s;
      }

      </style>
    `;

    class Excel extends HTMLElement {

        constructor() {
            super();

            _shadowRoot = this.attachShadow({
                mode: "open"
            });
            _shadowRoot.appendChild(tmpl.content.cloneNode(true));

            _id = createGuid();

            this._export_settings = {};
            this._export_settings.title = "";
            this._export_settings.subtitle = "";
            this._export_settings.icon = "";
            this._export_settings.unit = "";
            this._export_settings.footer = "";

            this._firstConnection = 0;

            this.addEventListener("click", event => {
                console.log('click');
            });

        }

        connectedCallback() {

            // ================= DRAG DROP ENHANCEMENT =================

            setTimeout(()=>{

                let uploader = _shadowRoot.querySelector("input[type=file]");

                if(!uploader) return;

                uploader.addEventListener("dragover",(e)=>{
                    e.preventDefault();
                    uploader.classList.add("dragover");
                });

                uploader.addEventListener("dragleave",(e)=>{
                    uploader.classList.remove("dragover");
                });

                uploader.addEventListener("drop",(e)=>{
                    e.preventDefault();
                    uploader.classList.remove("dragover");
                    uploader.files = e.dataTransfer.files;
                });

            },1000);

            try {
                if (window.commonApp) {
                    let outlineContainer = commonApp.getShell().findElements(true, ele => ele.hasStyleClass && ele.hasStyleClass("sapAppBuildingOutline"))[0];

                    if (outlineContainer && outlineContainer.getReactProps) {

                        let parseReactState = state => {

                            let components = {};
                            let globalState = state.globalState;
                            let instances = globalState.instances;
                            let app = instances.app["[{\"app\":\"MAIN_APPLICATION\"}]"];
                            let names = app.names;

                            for (let key in names) {

                                let name = names[key];

                                let obj = JSON.parse(key).pop();
                                let type = Object.keys(obj)[0];
                                let id = obj[type];

                                components[id] = {
                                    type: type,
                                    name: name
                                };

                            }

                            let metadata = JSON.stringify({
                                components: components,
                                vars: app.globalVars
                            });

                            if (metadata != this.metadata) {

                                this.metadata = metadata;

                                this.dispatchEvent(new CustomEvent("propertiesChanged", {
                                    detail: {
                                        properties: {
                                            metadata: metadata
                                        }
                                    }
                                }));

                            }

                        };

                        let subscribeReactStore = store => {

                            this._subscription = store.subscribe({
                                effect: state => {
                                    parseReactState(state);
                                    return {result:1};
                                }
                            });

                        };

                        let props = outlineContainer.getReactProps();

                        if (props) {
                            subscribeReactStore(props.store);
                        }

                    }
                }
            } catch (e) {}

        }

        disconnectedCallback() {
            if (this._subscription) {
                this._subscription();
                this._subscription = null;
            }
        }

        onCustomWidgetBeforeUpdate(changedProperties) {
            if ("designMode" in changedProperties) {
                this._designMode = changedProperties["designMode"];
            }
        }

        onCustomWidgetAfterUpdate(changedProperties) {

            var that = this;

            let xlsxjs = "https://sacplanning2025.github.io/hr_widget/xlsxxx.js";

            async function LoadLibs() {

                try {

                    await loadScript(xlsxjs, _shadowRoot);

                } catch (e) {
                    console.log(e);
                }

                finally {

                    loadthis(that, changedProperties);

                }

            }

            LoadLibs();

        }

        _renderExportButton() {}

        _firePropertiesChanged() {

            this.unit = "";

            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: {
                        unit: this.unit
                    }
                }
            }));

        }

        get title() {
            return this._export_settings.title;
        }

        set title(value) {
            this._export_settings.title = value;
        }

        get unit() {
            return this._export_settings.unit;
        }

        set unit(value) {
            value = _result;
            this._export_settings.unit = value;
        }

        static get observedAttributes() {
            return ["title","subtitle","icon","unit","footer","link"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue != newValue) {
                this[name] = newValue;
            }
        }

    }

    customElements.define("com-fd-djaja-sap-sac-excelll", Excel);


    // ================= ENHANCEMENT UTILITIES =================

    function debugLog(msg){
        if(UploadConfig.DEBUG){
            console.log("[ExcelWidget]",msg);
        }
    }

    function validateFile(file){

        if(file.size > UploadConfig.MAX_FILE_SIZE){
            sap.m.MessageToast.show("File exceeds 10MB");
            return false;
        }

        var ext = file.name.split(".").pop().toLowerCase();

        if(!UploadConfig.ALLOWED_TYPES.includes(ext)){
            sap.m.MessageToast.show("Invalid file type");
            return false;
        }

        return true;

    }

    function updateUploadStats(){

        console.log("Upload Stats:",UploadStats);

    }

    function updateUploadProgress(percent){

        var bar = document.getElementById("upload-progress-bar");

        if(bar){
            bar.style.width = percent+"%";
        }

    }

    function createGuid() {

        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {

            let r = Math.random() * 16 | 0,
                v = c === "x" ? r : (r & 0x3 | 0x8);

            return v.toString(16);

        });

    }

    function loadScript(src, shadowRoot) {

        return new Promise(function(resolve, reject) {

            let script = document.createElement('script');
            script.src = src;

            script.onload = () => {
                console.log("Load: " + src);
                resolve(script);
            }

            script.onerror = () => reject(new Error(`Script load error for ${src}`));

            shadowRoot.appendChild(script)

        });

    }

})();
