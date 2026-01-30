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

                  onValidate: function (e) {
  var fU = this.getView().byId("idfileUploader");
  var file = $("#__xmlview1--idfileUploader-fu")[0]?.files?.[0];
  var this_ = this;
  this_.wasteTime();

  if (!file) {
    this_.runNext();
    sap.m.MessageToast.show("Please choose a file");
    return;
  }

  var reader = new FileReader();

  reader.onload = function (ev) {
    try {
      // Read workbook (xls, xlsx, xlsm)
      var data = ev.target.result;
      var workbook = XLSX.read(data, { type: "binary" });

      // Prefer "Sheet1", else first sheet
      var ws = workbook.Sheets["Sheet1"] || workbook.Sheets[workbook.SheetNames[0]];
      if (!ws) {
        sap.m.MessageToast.show("Please upload the correct file (missing Sheet1)");
        this_.runNext(); fU.setValue(""); return;
      }

      // Read as array-of-arrays with header row
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });

      if (!rows || rows.length < 2) {
        sap.m.MessageToast.show("There is no record to be uploaded");
        this_.runNext(); fU.setValue(""); return;
      }

      // ---- Header mapping (case/underscore-insensitive) ----
      var hdr = rows[0].map(h => String(h).trim());
      var expected = ["ID","Description","H1","Company_Code","Costcenter","Position","Grade","Hire_Month","Nationality"];
      var norm = s => String(s).toLowerCase().replace(/\s+/g, "").replace(/[_-]/g, "");
      var hmap = {};
      hdr.forEach((h, i) => { hmap[norm(h)] = i; });

      var missing = expected.map(norm).filter(k => !(k in hmap));
      if (missing.length) {
        sap.m.MessageToast.show("Template columns missing: " + missing.join(", "));
        this_.runNext(); fU.setValue(""); return;
      }
      var idx = name => hmap[norm(name)];

      // ---- Helpers: Description clean + Proper Case ----
      var cleanDesc = s => String(s).replace(/[^A-Za-z ]+/g, " ").replace(/\s+/g, " ").trim();
      var properCase = s => String(s).toLowerCase().replace(/\b\w/g, c => c.toUpperCase());

      // ---- Build result ----
      var result_final = [];
      for (var r = 1; r < rows.length; r++) {
        var row = rows[r];
        var empty = !row || row.every(v => String(v).trim() === "");
        if (empty) continue;

        var rec = {
          "ID"          : String(row[idx("ID")] ?? "").trim(),
          "DESCRIPTION" : properCase(cleanDesc(String(row[idx("Description")] ?? ""))),
          "H1"          : String(row[idx("H1")] ?? "").trim(),
          "Company_Code": String(row[idx("Company_Code")] ?? "").trim(),
          "Costcenter"  : String(row[idx("Costcenter")] ?? "").trim(),
          "Position"    : String(row[idx("Position")] ?? "").trim(),
          "Grade"       : String(row[idx("Grade")] ?? "").trim(),
          "Hire_Month"  : String(row[idx("Hire_Month")] ?? "").trim(),
          "Nationality" : String(row[idx("Nationality")] ?? "").replace(/_x000D_/gi, " ").trim()
        };

        if (Object.values(rec).join("").trim().length > 0) {
          result_final.push(rec);
        }
      }

      if (result_final.length === 0) {
        fU.setValue("");
        sap.m.MessageToast.show("There is no record to be uploaded");
        this_.runNext(); return;
      }
      if (result_final.length >= 2001) {
        fU.setValue("");
        sap.m.MessageToast.show("Maximum records are 2000.");
        this_.runNext(); return;
      }

      // ---- Keep your original eventing ----
      var oModel = new sap.ui.model.json.JSONModel();
      oModel.setSizeLimit("5000");
      oModel.setData({ result_final: result_final });

      var oModel1 = new sap.ui.model.json.JSONModel();
      oModel1.setData({ fname: file.name });

      _result = JSON.stringify(result_final);
      that._firePropertiesChanged();
      this.settings = {};
      this.settings.result = "";
      that.dispatchEvent(new CustomEvent("onStart", { detail: { settings: this.settings } }));

      this_.runNext();
      fU.setValue("");
    } catch (err) {
      console.log(err);
      this_.runNext();
      fU.setValue("");
      sap.m.MessageToast.show("Please upload the correct file");
    }
  };

  // Works for xlsx/xlsm
  reader.readAsBinaryString(file);
},
                            } else {
                                this_.runNext();
                                console.log("Error: wrong Excel File template");
                                MessageToast.show("Please upload the correct file");
                            }
                        };

                        if (typeof file !== 'undefined') {
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
