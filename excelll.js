(function() {
    let _shadowRoot;
    let _id;
    let _result ;

    let div;
    let widgetName;
    var Ar = [];

    let tmpl = document.createElement("template");
    tmpl.innerHTML = `
      <style>
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

            //_shadowRoot.querySelector("#oView").id = "oView";

            this._export_settings = {};
            this._export_settings.title = "";
            this._export_settings.subtitle = "";
            this._export_settings.icon = "";
            this._export_settings.unit = "";
            this._export_settings.footer = "";

            this.addEventListener("click", event => {
                console.log('click');

            });

            this._firstConnection = 0;
        }

        connectedCallback() {
            try {
                if (window.commonApp) {
                    let outlineContainer = commonApp.getShell().findElements(true, ele => ele.hasStyleClass && ele.hasStyleClass("sapAppBuildingOutline"))[0]; // sId: "__container0"

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

                            for (let componentId in components) {
                                let component = components[componentId];
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
                                    return {
                                        result: 1
                                    };
                                }
                            });
                        };

                        let props = outlineContainer.getReactProps();
                        if (props) {
                            subscribeReactStore(props.store);
                        } else {
                            let oldRenderReactComponent = outlineContainer.renderReactComponent;
                            outlineContainer.renderReactComponent = e => {
                                let props = outlineContainer.getReactProps();
                                subscribeReactStore(props.store);

                                oldRenderReactComponent.call(outlineContainer, e);
                            }
                        }
                    }
                }
            } catch (e) {}
        }

        disconnectedCallback() {
            if (this._subscription) { // react store subscription
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
                } finally {
                    loadthis(that, changedProperties);
                }
            }
            LoadLibs();
        }

        _renderExportButton() {
            let components = this.metadata ? JSON.parse(this.metadata)["components"] : {};
        }

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

        // SETTINGS
        get title() {
            return this._export_settings.title;
        }
        set title(value) {
            console.log("setTitle:" + value);
            this._export_settings.title = value;
        }

        get subtitle() {
            return this._export_settings.subtitle;
        }
        set subtitle(value) {
            this._export_settings.subtitle = value;
        }

        get icon() {
            return this._export_settings.icon;
        }
        set icon(value) {
            this._export_settings.icon = value;
        }

        get unit() {
            return this._export_settings.unit;
        }
        set unit(value) {
            value = _result;
            console.log("value: " + value);
            this._export_settings.unit = value;
        }

        get footer() {
            return this._export_settings.footer;
        }
        set footer(value) {
            this._export_settings.footer = value;
        }

        static get observedAttributes() {
            return [
                "title",
                "subtitle",
                "icon",
                "unit",
                "footer",
                "link"
            ];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue != newValue) {
                this[name] = newValue;
            }
        }

    }
    customElements.define("com-fd-djaja-sap-sac-excelll", Excel);

    // UTILS
    function loadthis(that, changedProperties) {
        var that_ = that;

        widgetName = changedProperties.widgetName;
        if(typeof widgetName === "undefined") {
            widgetName = that._export_settings.title.split("|")[0];
        }


        div = document.createElement('div');
        div.slot = "content_" + widgetName;

        if(that._firstConnection === 0) {
            let div0 = document.createElement('div');
            div0.innerHTML = '<?xml version="1.0"?><script id="oView_' + widgetName + '" name="oView_' + widgetName + '" type="sapui5/xmlview"><mvc:View height="100%" xmlns="sap.m" xmlns:u="sap.ui.unified" xmlns:f="sap.ui.layout.form" xmlns:core="sap.ui.core" xmlns:mvc="sap.ui.core.mvc" controllerName="myView.Template"><f:SimpleForm editable="true"><f:content><Label text="Upload"></Label><VBox><u:FileUploader id="idfileUploader" width="100%" useMultipart="false" sendXHR="true" sameFilenameAllowed="false" buttonText="" fileType="XLSM" placeholder="Choose a file" style="Emphasized"/><Button text="Upload" press="onValidate" id="__uploadButton" tooltip="Upload a File"/></VBox></f:content></f:SimpleForm></mvc:View></script>';
            _shadowRoot.appendChild(div0);

            let div1 = document.createElement('div');
            div1.innerHTML = '<?xml version="1.0"?><script id="myXMLFragment_' + widgetName + '" type="sapui5/fragment"><core:FragmentDefinition xmlns="sap.m" xmlns:core="sap.ui.core"><SelectDialog title="Partner Number" class="sapUiPopupWithPadding"  items="{' + widgetName + '>/}" search="_handleValueHelpSearch"  confirm="_handleValueHelpClose"  cancel="_handleValueHelpClose"  multiSelect="true" showClearButton="true" rememberSelections="true"><StandardListItem icon="{' + widgetName + '>ProductPicUrl}" iconDensityAware="false" iconInset="false" title="{' + widgetName + '>partner}" description="{' + widgetName + '>partner}" /></SelectDialog></core:FragmentDefinition></script>';
            _shadowRoot.appendChild(div1);

            let div2 = document.createElement('div');
            div2.innerHTML = '<div id="ui5_content_' + widgetName + '" name="ui5_content_' + widgetName + '"><slot name="content_' + widgetName + '"></slot></div>';
            _shadowRoot.appendChild(div2);

            that_.appendChild(div);
            // ================= DRAG & DROP FEATURE =================

            div.addEventListener("dragover", function(e){
                e.preventDefault();
            });
            
            div.addEventListener("drop", function(e){
            
                e.preventDefault();
            
                var file = e.dataTransfer.files[0];
            
                if(!file){
                    sap.m.MessageToast.show("No file detected");
                    return;
                }
            
                // Assign dropped file to uploader
                var uploader = document.querySelector("input[type='file']");
                
                if(uploader){
                    uploader.files = e.dataTransfer.files;
                }
            
                sap.m.MessageToast.show("File dropped successfully");
            
            });
            
            // =======================================================
            

            var mapcanvas_divstr = _shadowRoot.getElementById('oView_' + widgetName);
            var mapcanvas_fragment_divstr = _shadowRoot.getElementById('myXMLFragment_' + widgetName);

            Ar.push({
               'id': widgetName,
               'div': mapcanvas_divstr,
               'divf': mapcanvas_fragment_divstr
            });
        }

        that_._renderExportButton();

        sap.ui.getCore().attachInit(function() {
            "use strict";

            //### Controller ###
            sap.ui.define([
                "jquery.sap.global",
                "sap/ui/core/mvc/Controller",
                "sap/ui/model/json/JSONModel",
                "sap/m/MessageToast",
                "sap/ui/core/library",
                "sap/ui/core/Core",
                'sap/ui/model/Filter',
                'sap/m/library',
                'sap/m/MessageBox',
                'sap/ui/unified/DateRange',
                'sap/ui/core/format/DateFormat',
                'sap/ui/model/BindingMode',
                'sap/ui/core/Fragment',
                'sap/m/Token',
                'sap/ui/model/FilterOperator',
                'sap/ui/model/odata/ODataModel',
                'sap/m/BusyDialog'
            ], function(jQuery, Controller, JSONModel, MessageToast, coreLibrary, Core, Filter, mobileLibrary, MessageBox, DateRange, DateFormat, BindingMode, Fragment, Token, FilterOperator, ODataModel, BusyDialog) {
                "use strict";

                var busyDialog = (busyDialog) ? busyDialog : new BusyDialog({});

                return Controller.extend("myView.Template", {

                    onInit: function() {
                        console.log(that._export_settings.title);
                        console.log("widgetName:" + that.widgetName);

                        if(that._firstConnection === 0) {
                            that._firstConnection = 1;
                        }
                    },

                    /*onValidate: function (e) {
                    var fU = this.getView().byId("idfileUploader");
                    //var domRef = fU.getFocusDomRef();
                    //var domRef = this.getView().byId("__xmlview1--idfileUploader-fu").getFocusDomRef();
                    //var file = domRef.files[0];
                    var file = $("#__xmlview1--idfileUploader-fu")[0].files[0];
                    var this_ = this;

                     this_.wasteTime();*/

                       onValidate: function () {

                        var fU = this.getView().byId("idfileUploader");
                        var file = $("#__xmlview1--idfileUploader-fu")[0].files[0];
                        var this_ = this;
                        
                        // FILE SELECT CHECK
                        if (!file) {
                            sap.m.MessageToast.show("Please select a file");
                            return;
                        }
                        
                        // FILE SIZE VALIDATION
                        var maxSize = 5 * 1024 * 1024;
                        if (file.size > maxSize) {
                            sap.m.MessageToast.show("File must be smaller than 5MB");
                            return;
                        }
                        
                        // FILE TYPE VALIDATION
                        var allowedTypes = ["xls","xlsx","xlsm"];
                        var fileExt = file.name.split(".").pop().toLowerCase();
                        
                        if (!allowedTypes.includes(fileExt)) {
                            sap.m.MessageToast.show("Only Excel files allowed");
                            return;
                        }
                        
                        this_.wasteTime();
                        
                        var reader = new FileReader();
                        
                        reader.onprogress = function (event) {
                        
                            if (event.lengthComputable) {
                        
                                var percent = Math.round((event.loaded / event.total) * 100);
                        
                                busyDialog.setText("Uploading " + percent + "%");
                            }
                        };
                        
                        reader.onload = function (e) {
                        
                            var data = e.target.result;
                        
                            var workbook = XLSX.read(data,{type:'binary'});
                        
                            var result_final = [];
                            var idSet = new Set();
                        
                            workbook.SheetNames.forEach(function(sheetName){
                        
                                if(sheetName === "Sheet1"){
                        
                                    var json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1});
                        
                                    for(var i=1;i<json.length;i++){
                        
                                        var rec = json[i];
                        
                                        if(!rec || rec.length===0) continue;
                        
                                        var id = (rec[0] || "").toString().trim();
                        
                                        // DUPLICATE CHECK
                                        if(idSet.has(id)) continue;
                        
                                        idSet.add(id);
                        
                                        result_final.push({
                        
                                            ID: id,
                                            DESCRIPTION: (rec[1]||"").toString().trim(),
                                            H1: (rec[2]||"").toString().trim(),
                                            Company_Code: (rec[3]||"").toString().trim(),
                                            Costcenter: (rec[4]||"").toString().trim(),
                                            Division: (rec[5]||"").toString().trim(),
                                            Department: (rec[6]||"").toString().trim(),
                                            Position: (rec[7]||"").toString().trim(),
                                            ZZ_PAY_GRADE_LVL: (rec[8]||"").toString().trim(),
                                            Hire_Month: (rec[9]||"").toString().trim(),
                                            Nationality: (rec[10]||"").toString().trim(),
                                            Med_Insu_class: (rec[11]||"").toString().trim(),
                                            No_of_dependents: (rec[12]||"").toString().trim(),
                                            ACCOM: (rec[13]||"").toString().trim(),
                                            TRANSPORT: (rec[14]||"").toString().trim(),
                                            EMP_CLASS: (rec[15]||"").toString().trim(),
                                            OT: (rec[16]||"").toString().trim()
                        
                                        });
                                    }
                                }
                        
                            });
                        
                            if(result_final.length===0){
                        
                                sap.m.MessageToast.show("No valid records found");
                                this_.runNext();
                                return;
                        
                            }
                        
                            if(result_final.length>2000){
                        
                                sap.m.MessageToast.show("Maximum records allowed: 2000");
                                this_.runNext();
                                return;
                        
                            }
                        
                            // SAVE DATA FOR SAC SCRIPT
                            _result = JSON.stringify(result_final);
                        
                            that._firePropertiesChanged();
                        
                            that.dispatchEvent(new CustomEvent("onStart",{
                                detail:{settings:{}}
                            }));
                        
                            // STORE HISTORY
                            localStorage.setItem("lastExcelUpload",new Date());
                        
                            // RECORD COUNT
                            console.log("Uploaded Records:",result_final.length);
                        
                            // PREVIEW TABLE
                            console.table(result_final.slice(0,10));
                        
                            this_.runNext();
                        
                            fU.setValue("");
                        
                        };
                        
                        reader.readAsBinaryString(file);
                        
                        }
                    },

                    wasteTime: function() {
                        busyDialog.open();
                    },

                    runNext: function() {
                        busyDialog.close();
                    },

                });
            });

            console.log("widgetName Final:" + widgetName);
            var foundIndex = Ar.findIndex(x => x.id == widgetName);
            var divfinal = Ar[foundIndex].div;
            console.log(divfinal);

            //### THE APP: place the XMLView somewhere into DOM ###
            var oView = sap.ui.xmlview({
                viewContent: jQuery(divfinal).html(),
            });

            oView.placeAt(div);
            if (that_._designMode) {
                oView.byId("idfileUploader").setEnabled(false);
            }
        });
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
