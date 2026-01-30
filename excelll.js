(function () {
    let _shadowRoot;
    let _id;
    let _result;

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

            this._firstConnection = 0;
        }

        connectedCallback() {}

        disconnectedCallback() {}

        onCustomWidgetBeforeUpdate(changedProperties) {
            if ("designMode" in changedProperties) {
                this._designMode = changedProperties["designMode"];
            }
        }

        onCustomWidgetAfterUpdate(changedProperties) {
            var that = this;

            let xlsxjs =
                "https://sacplanning2025.github.io/hr_widget/xlsxxx.js";

            loadScript(xlsxjs, _shadowRoot).finally(() => {
                loadthis(that, changedProperties);
            });
        }

        _firePropertiesChanged() {
            this.unit = "";
            this.dispatchEvent(
                new CustomEvent("propertiesChanged", {
                    detail: { properties: { unit: this.unit } }
                })
            );
        }

        get unit() {
            return this._export_settings.unit;
        }
        set unit(value) {
            value = _result;
            this._export_settings.unit = value;
        }

        static get observedAttributes() {
            return ["title", "subtitle", "icon", "unit", "footer"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelll", Excel);

    function loadthis(that, changedProperties) {
        widgetName =
            changedProperties.widgetName ||
            that._export_settings.title.split("|")[0];

        div = document.createElement("div");
        div.slot = "content_" + widgetName;

        if (that._firstConnection === 0) {
            let div0 = document.createElement("div");

            div0.innerHTML =
                '<?xml version="1.0"?><script id="oView_' +
                widgetName +
                '" type="sapui5/xmlview"><mvc:View height="100%" xmlns="sap.m" xmlns:u="sap.ui.unified" xmlns:f="sap.ui.layout.form" xmlns:core="sap.ui.core" xmlns:mvc="sap.ui.core.mvc" controllerName="myView.Template"><f:SimpleForm editable="true"><f:content><Label text="Upload"></Label><VBox><u:FileUploader id="idfileUploader" width="100%" useMultipart="false" sendXHR="true" buttonText="" fileType="XLS,XLSX,XLSM" placeholder="Choose a file"/><Button text="Upload" press="onValidate"/></VBox></f:content></f:SimpleForm></mvc:View></script>';

            _shadowRoot.appendChild(div0);
            that.appendChild(div);

            Ar.push({
                id: widgetName,
                div: _shadowRoot.getElementById("oView_" + widgetName)
            });

            that._firstConnection = 1;
        }

        sap.ui.getCore().attachInit(function () {
            sap.ui.define(
                [
                    "sap/ui/core/mvc/Controller",
                    "sap/ui/model/json/JSONModel",
                    "sap/m/MessageToast",
                    "sap/m/BusyDialog"
                ],
                function (Controller, JSONModel, MessageToast, BusyDialog) {
                    var busyDialog = new BusyDialog();

                    return Controller.extend("myView.Template", {
                        wasteTime: function () {
                            busyDialog.open();
                        },

                        runNext: function () {
                            busyDialog.close();
                        },

                        onValidate: function () {
                            var fU =
                                this.getView().byId("idfileUploader");
                            var file =
                                fU.oFileUpload &&
                                fU.oFileUpload.files[0];

                            var this_ = this;
                            this_.wasteTime();

                            if (!file) {
                                this_.runNext();
                                MessageToast.show(
                                    "Please choose a file"
                                );
                                return;
                            }

                            var reader = new FileReader();

                            reader.onload = function (ev) {
                                try {
                                    var workbook = XLSX.read(
                                        ev.target.result,
                                        { type: "binary" }
                                    );

                                    var ws =
                                        workbook.Sheets["Sheet1"] ||
                                        workbook.Sheets[
                                            workbook.SheetNames[0]
                                        ];

                                    if (!ws) throw "Sheet missing";

                                    var rows =
                                        XLSX.utils.sheet_to_json(ws, {
                                            header: 1,
                                            defval: ""
                                        });

                                    if (rows.length < 2)
                                        throw "No data";

                                    var hdr = rows[0];
                                    var hmap = {};
                                    hdr.forEach((h, i) => {
                                        hmap[
                                            h
                                                .toLowerCase()
                                                .replace(/[_ ]/g, "")
                                        ] = i;
                                    });

                                    var idx = name =>
                                        hmap[
                                            name
                                                .toLowerCase()
                                                .replace(/[_ ]/g, "")
                                        ];

                                    var result_final = [];

                                    for (
                                        var r = 1;
                                        r < rows.length;
                                        r++
                                    ) {
                                        var row = rows[r];
                                        if (!row) continue;

                                        result_final.push({
                                            ID:
                                                row[idx("ID")] || "",
                                            DESCRIPTION:
                                                row[
                                                    idx("Description")
                                                ] || ""
                                        });
                                    }

                                    _result = JSON.stringify(result_final);

                                    that._firePropertiesChanged();

                                    that.dispatchEvent(
                                        new CustomEvent("onStart", {
                                            detail: { result: _result }
                                        })
                                    );

                                    MessageToast.show(
                                        "File uploaded successfully"
                                    );

                                    this_.runNext();
                                    fU.setValue("");
                                } catch (err) {
                                    this_.runNext();
                                    fU.setValue("");
                                    MessageToast.show(
                                        "Invalid file format"
                                    );
                                }
                            };

                            reader.readAsBinaryString(file);
                        }
                    });
                }
            );

            var foundIndex = Ar.findIndex(
                x => x.id == widgetName
            );

            var oView = sap.ui.xmlview({
                viewContent: $(Ar[foundIndex].div).html()
            });

            oView.placeAt(div);
        });
    }

    function createGuid() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(
            /[xy]/g,
            function (c) {
                var r = (Math.random() * 16) | 0;
                var v = c == "x" ? r : (r & 0x3) | 0x8;
                return v.toString(16);
            }
        );
    }

    function loadScript(src, shadowRoot) {
        return new Promise(function (resolve, reject) {
            let script = document.createElement("script");
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            shadowRoot.appendChild(script);
        });
    }
})();           
