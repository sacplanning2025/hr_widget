(function() {
    let _shadowRoot;
    let _id;
    let _result ;

    let div;
    let widgetName;
    var Ar = [];

    let tmpl = document.createElement("template");
    tmpl.innerHTML = `<style></style>`;

    class Excel extends HTMLElement {

        constructor() {
            super();

            _shadowRoot = this.attachShadow({ mode: "open" });
            _shadowRoot.appendChild(tmpl.content.cloneNode(true));

            _id = createGuid();

            this._export_settings = {
                title: "",
                subtitle: "",
                icon: "",
                unit: "",
                footer: ""
            };

            this.addEventListener("click", () => console.log('click'));
            this._firstConnection = 0;
        }

        connectedCallback() {
            try {
                if (window.commonApp) {
                    let outlineContainer = commonApp.getShell()
                        .findElements(true, ele => ele.hasStyleClass && ele.hasStyleClass("sapAppBuildingOutline"))[0];

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

                                components[id] = { type, name };
                            }

                            let metadata = JSON.stringify({
                                components,
                                vars: app.globalVars
                            });

                            if (metadata != this.metadata) {
                                this.metadata = metadata;

                                this.dispatchEvent(new CustomEvent("propertiesChanged", {
                                    detail: { properties: { metadata } }
                                }));
                            }
                        };

                        let subscribeReactStore = store => {
                            this._subscription = store.subscribe({
                                effect: state => {
                                    parseReactState(state);
                                    return { result: 1 };
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

        onCustomWidgetAfterUpdate(changedProperties) {
            let that = this;
            let xlsxjs = "https://sacplanning2025.github.io/hr_widget/xlsxxx.js";

            async function LoadLibs() {
                try {
                    await loadScript(xlsxjs, _shadowRoot);
                } finally {
                    loadthis(that, changedProperties);
                }
            }
            LoadLibs();
        }

        _firePropertiesChanged() {
            this.unit = "";
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: { properties: { unit: this.unit } }
            }));
        }

        get unit() { return this._export_settings.unit; }
        set unit(value) {
            value = _result;
            this._export_settings.unit = value;
        }

        static get observedAttributes() {
            return ["title","subtitle","icon","unit","footer","link"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue != newValue) this[name] = newValue;
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelll", Excel);

    function loadthis(that, changedProperties) {

        widgetName = changedProperties.widgetName;
        if(typeof widgetName === "undefined") {
            widgetName = that._export_settings.title.split("|")[0];
        }

        div = document.createElement('div');
        div.slot = "content_" + widgetName;

        if(that._firstConnection === 0) {

            let div0 = document.createElement('div');
            div0.innerHTML =
            '<?xml version="1.0"?><script id="oView_' + widgetName +
            '" type="sapui5/xmlview"><mvc:View height="100%" xmlns="sap.m" xmlns:u="sap.ui.unified" xmlns:f="sap.ui.layout.form" xmlns:core="sap.ui.core" xmlns:mvc="sap.ui.core.mvc" controllerName="myView.Template"><f:SimpleForm editable="true"><f:content><Label text="Upload"></Label><VBox><u:FileUploader id="idfileUploader" width="100%" useMultipart="false" sendXHR="true" sameFilenameAllowed="false" buttonText="" fileType="XLS,XLSX,XLSM" placeholder="Choose a file"/><Button text="Upload" press="onValidate"/></VBox></f:content></f:SimpleForm></mvc:View></script>';

            _shadowRoot.appendChild(div0);
            that.appendChild(div);
            that._firstConnection = 1;
        }

        sap.ui.getCore().attachInit(function() {

            sap.ui.define([
                "sap/ui/core/mvc/Controller",
                "sap/ui/model/json/JSONModel",
                "sap/m/MessageToast",
                "sap/m/BusyDialog"
            ], function(Controller, JSONModel, MessageToast, BusyDialog) {

                var busyDialog = new BusyDialog({});

                return Controller.extend("myView.Template", {

                    onValidate: function () {

                        var fU = this.getView().byId("idfileUploader");
                        var file = fU.oFileUpload.files[0];
                        var this_ = this;

                        busyDialog.open();

                        var reader = new FileReader();

                        reader.onload = function (e) {

                            var workbook = XLSX.read(e.target.result, { type: 'binary' });

                            // Take first sheet automatically
                            var sheetName = workbook.SheetNames[0];
                            var sheet = workbook.Sheets[sheetName];

                            // JSON conversion (robust)
                            var jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

                            var result_final = [];

                            jsonData.forEach(function(row) {

                                if (
                                    row.ID || row.Description || row.H1 ||
                                    row.Company_Code || row.Costcenter ||
                                    row.Position || row.Grade ||
                                    row.Hire_Month || row.Nationality
                                ) {

                                    result_final.push({
                                        ID: (row.ID || "").toString().trim(),
                                        DESCRIPTION: (row.Description || "").trim(),
                                        H1: (row.H1 || "").trim(),
                                        Company_Code: (row.Company_Code || "").trim(),
                                        Costcenter: (row.Costcenter || "").trim(),
                                        Position: (row.Position || "").trim(),
                                        Grade: (row.Grade || "").trim(),
                                        Hire_Month: (row.Hire_Month || "").trim(),
                                        Nationality: (row.Nationality || "").trim()
                                    });
                                }
                            });

                            if (result_final.length === 0) {
                                MessageToast.show("No valid records found.");
                                busyDialog.close();
                                return;
                            }

                            if (result_final.length > 2000) {
                                MessageToast.show("Maximum 2000 records allowed.");
                                busyDialog.close();
                                return;
                            }

                            _result = JSON.stringify(result_final);
                            that._firePropertiesChanged();

                            that.dispatchEvent(new CustomEvent("onStart", {
                                detail: { settings: {} }
                            }));

                            busyDialog.close();
                            fU.setValue("");
                        };

                        if (file) {
                            reader.readAsBinaryString(file);
                        } else {
                            busyDialog.close();
                        }
                    }
                });
            });

            var oView = sap.ui.xmlview({
                viewContent: jQuery(_shadowRoot.getElementById('oView_' + widgetName)).html(),
            });

            oView.placeAt(div);
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
            script.onload = () => resolve(script);
            script.onerror = () => reject(new Error(`Script load error`));
            shadowRoot.appendChild(script);
        });
    }
})();
