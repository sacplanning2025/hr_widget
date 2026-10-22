(function () {
  class AssetEntryWidget extends HTMLElement {
    constructor() {
      super();
      this._shadowRoot = this.attachShadow({ mode: "open" });

      this._rows = [];
      this._validationErrors = [];
      this._validationResult = "true";
      this._lastEvent = "";
      this._savePayload = [];
      this._widgetStatus = "READY";
      this._rowSequence = 1;

      this._companyOptions = [];
      this._costCenterOptions = [];
      this._assetClassOptions = [];

      this._columns = [
        { key: "selected", label: "Sel", type: "checkbox", width: "70px" },
        { key: "ID", label: "ID", type: "input", width: "200px" },
        { key: "DESCRIPTION", label: "Description", type: "input", width: "240px" },
        { key: "COMPANY", label: "Company", type: "select", width: "160px" },
        { key: "COSTCENTER", label: "Cost Center", type: "select", width: "180px" },
        { key: "ASSET_CLASS", label: "Asset Class", type: "select", width: "180px" },
        { key: "CAPITALIZED", label: "Capitalized", type: "month", width: "170px" }
      ];

      this._render();
    }

    connectedCallback() {
      if (!this._rows || this._rows.length === 0) {
        this._rows = [this._createEmptyRow()];
      }

      this._normalizeAllRows();
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onReady", { status: "ready" });
    }

    static get observedAttributes() {
      return [
        "rows",
        "companyOptions",
        "costCenterOptions",
        "assetClassOptions"
      ];
    }

    attributeChangedCallback(name, oldValue, newValue) {
      if (oldValue === newValue) return;

      if (name === "rows") {
        this.setRows(newValue || "[]");
        return;
      }

      if (name === "companyOptions") {
        this._companyOptions = this._parseOptions(newValue);
        this._refreshTable();
        return;
      }

      if (name === "costCenterOptions") {
        this._costCenterOptions = this._parseOptions(newValue);
        this._refreshTable();
        return;
      }

      if (name === "assetClassOptions") {
        this._assetClassOptions = this._parseOptions(newValue);
        this._refreshTable();
        return;
      }
    }

    _createEmptyRow() {
      return {
        rowId: "ROW_" + String(this._rowSequence++),
        selected: false,
        isModified: false,
        rowStatus: "NEW",
        ID: "",
        DESCRIPTION: "",
        COMPANY: "",
        COSTCENTER: "",
        ASSET_CLASS: "",
        CAPITALIZED: ""
      };
    }

    _normalizeRow(row) {
      if (!row.rowId) row.rowId = "ROW_" + String(this._rowSequence++);
      if (row.selected !== true) row.selected = false;
      if (row.isModified !== true) row.isModified = false;
      if (!row.rowStatus) row.rowStatus = "LOADED";

      row.ID = this._safeString(row.ID);
      row.DESCRIPTION = this._safeString(row.DESCRIPTION);
      row.COMPANY = this._safeString(row.COMPANY);
      row.COSTCENTER = this._safeString(row.COSTCENTER);
      row.ASSET_CLASS = this._safeString(row.ASSET_CLASS);
      row.CAPITALIZED = this._normalizeCapitalized(this._safeString(row.CAPITALIZED));
    }

    _normalizeAllRows() {
      for (var i = 0; i < this._rows.length; i++) {
        this._normalizeRow(this._rows[i]);
      }
    }

    _safeString(value) {
      if (value === undefined || value === null) return "";
      return String(value).trim();
    }

    _escape(value) {
      if (value === undefined || value === null) return "";
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }

    _parseOptions(json) {
      try {
        var arr = JSON.parse(json || "[]");
        if (!Array.isArray(arr)) return [];

        var normalized = [];
        for (var i = 0; i < arr.length; i++) {
          var item = arr[i];
          if (item && typeof item === "object") {
            var key = item.key !== undefined && item.key !== null ? String(item.key) : "";
            var text = item.text !== undefined && item.text !== null ? String(item.text) : key;
            normalized.push({ key: key, text: text });
          } else {
            var val = item !== undefined && item !== null ? String(item) : "";
            normalized.push({ key: val, text: val });
          }
        }
        return normalized;
      } catch (e) {
        return [];
      }
    }

    _getOptionsForField(fieldName) {
      if (fieldName === "COMPANY") return this._companyOptions || [];
      if (fieldName === "COSTCENTER") return this._costCenterOptions || [];
      if (fieldName === "ASSET_CLASS") return this._assetClassOptions || [];
      return [];
    }

    _normalizeCapitalized(value) {
      var v = this._safeString(value);
      if (/^\d{6}$/.test(v)) {
        return v;
      }
      if (/^\d{4}-\d{2}$/.test(v)) {
        return v.replace("-", "");
      }
      return "";
    }

    _capitalizedToMonthInput(value) {
      var v = this._normalizeCapitalized(value);
      if (/^\d{6}$/.test(v)) {
        return v.substring(0, 4) + "-" + v.substring(4, 6);
      }
      return "";
    }

    _monthInputToCapitalized(value) {
      var v = this._safeString(value);
      if (/^\d{4}-\d{2}$/.test(v)) {
        return v.replace("-", "");
      }
      return "";
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          :host {
            display: block;
            font-family: "72", Arial, sans-serif;
            color: #223548;
          }

          .wrap {
            border: 1px solid #d9e2ef;
            border-radius: 12px;
            background: #ffffff;
            overflow: hidden;
          }

          .toolbar {
            display: flex;
            justify-content: flex-end;
            gap: 8px;
            padding: 12px;
            border-bottom: 1px solid #e5edf7;
            background: #f8fbff;
            flex-wrap: wrap;
          }

          .btn {
            border: 1px solid #c7d7ea;
            background: #ffffff;
            color: #0a6ed1;
            border-radius: 8px;
            padding: 8px 14px;
            cursor: pointer;
            font-weight: 600;
            font-size: 13px;
          }

          .btn:hover {
            background: #f3f8fd;
          }

          .btn.primary {
            background: #0a6ed1;
            color: #ffffff;
            border-color: #0a6ed1;
          }

          .btn.danger {
            color: #bb1e1e;
            border-color: #efb4b4;
            background: #fff7f7;
          }

          .gridWrap {
            overflow: auto;
            max-height: 520px;
            background: #ffffff;
          }

          table {
            border-collapse: separate;
            border-spacing: 0;
            width: max-content;
            min-width: 100%;
          }

          th, td {
            border-bottom: 1px solid #edf2f7;
            padding: 8px;
            vertical-align: top;
            white-space: nowrap;
          }

          th {
            position: sticky;
            top: 0;
            background: #eef4fb;
            z-index: 1;
            text-align: left;
            font-size: 12px;
            color: #223548;
            font-weight: 700;
          }

          tr:hover td {
            background: #fafcff;
          }

          tr.errorRow td {
            background: #fff7f7;
          }

          tr.modifiedRow td {
            background: #fffbeb;
          }

          .cell, .select {
            width: 100%;
            min-height: 34px;
            height: 34px;
            border: 1px solid #c9d6e5;
            border-radius: 6px;
            padding: 6px 10px;
            font-size: 13px;
            background: #fff;
            color: #223548;
            box-sizing: border-box;
            outline: none;
          }

          .cell.error, .select.error {
            border-color: #e25555;
            background: #fff5f5;
          }

          .rowErr {
            margin-top: 4px;
            font-size: 11px;
            color: #c53030;
            white-space: normal;
            max-width: 240px;
            line-height: 1.3;
          }

          .summary {
            padding: 10px 12px;
            border-top: 1px solid #e5edf7;
            display: flex;
            gap: 18px;
            font-size: 12px;
            background: #fafcff;
            flex-wrap: wrap;
          }

          .row-checkbox {
            width: 22px;
            height: 22px;
            cursor: pointer;
            margin-top: 8px;
          }

          .select-all-wrap {
            display: flex;
            align-items: center;
            gap: 6px;
          }

          .select-all-checkbox {
            width: 16px;
            height: 16px;
            cursor: pointer;
          }
        </style>
        <div class="wrap" id="widgetWrap"></div>
      `;
    }

    _hasSelectedRows() {
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) return true;
      }
      return false;
    }

    _areAllRowsSelected() {
      if (!this._rows || this._rows.length === 0) return false;
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected !== true) return false;
      }
      return true;
    }

    _toggleSelectAll(checked) {
      for (var i = 0; i < this._rows.length; i++) {
        this._rows[i].selected = checked;
      }

      this._validationErrors = [];
      this._validationResult = "true";
      this._lastEvent = JSON.stringify({ type: "selectAll", selected: checked });
      this._widgetStatus = "CHANGED";
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    _refreshStatusBarHtml() {
      return `
        <div class="summary">
          <div>Total Rows: ${this.getRowCount()}</div>
          <div>Selected Rows: ${this.getSelectedRowCount()}</div>
          <div>Validation: ${this._validationResult}</div>
          <div>Status: ${this._widgetStatus}</div>
        </div>
      `;
    }

    _refreshTable() {
      var container = this._shadowRoot.getElementById("widgetWrap");
      var hasSelection = this._hasSelectedRows();
      var allSelected = this._areAllRowsSelected();
      var rowErrorMap = this._getRowErrorMap();

      var html = '';
      html += '<div class="toolbar">';
      html += '<button class="btn" id="btnAdd">Add Row</button>';

      if (hasSelection) {
        html += '<button class="btn" id="btnCopy">Copy</button>';
        html += '<button class="btn danger" id="btnDelete">Delete Selected</button>';
      }

      html += '<button class="btn" id="btnValidate">Validate</button>';
      html += '<button class="btn primary" id="btnSave">Save</button>';
      html += '<button class="btn" id="btnClear">Clear</button>';
      html += '</div>';

      html += '<div class="gridWrap">';
      html += '<table>';
      html += '<thead><tr>';

      for (var h = 0; h < this._columns.length; h++) {
        var col = this._columns[h];
        if (col.key === "selected") {
          html += '<th style="width:' + col.width + '"><div class="select-all-wrap"><span>Sel</span><input class="select-all-checkbox" type="checkbox" id="selectAll" ' + (allSelected ? 'checked' : '') + ' /></div></th>';
        } else {
          html += '<th style="width:' + col.width + '">' + col.label + '</th>';
        }
      }

      html += '</tr></thead><tbody>';

      for (var i = 0; i < this._rows.length; i++) {
        var row = this._rows[i];
        var rowErrors = rowErrorMap[i] || [];
        var rowClass = "";

        if (rowErrors.length) rowClass = "errorRow";
        else if (row.isModified === true) rowClass = "modifiedRow";

        html += '<tr class="' + rowClass + '">';
        for (var c = 0; c < this._columns.length; c++) {
          html += '<td style="width:' + this._columns[c].width + '">' + this._renderCell(row, i, this._columns[c], rowErrors) + '</td>';
        }
        html += '</tr>';
      }

      html += '</tbody></table></div>';
      html += this._refreshStatusBarHtml();

      container.innerHTML = html;
      this._bindEvents();
    }

    _renderCell(row, rowIndex, column, rowErrors) {
      var value = row[column.key] !== undefined && row[column.key] !== null ? row[column.key] : "";
      var hasError = this._hasFieldError(column.key, rowErrors);
      var errorCss = hasError ? "error" : "";

      if (column.type === "checkbox") {
        return '<input class="row-checkbox" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="checkbox" type="checkbox" ' + (value === true ? 'checked' : '') + ' />';
      }

      if (column.type === "input") {
        return ''
          + '<input class="cell ' + errorCss + '"'
          + ' data-row="' + rowIndex + '"'
          + ' data-field="' + column.key + '"'
          + ' data-type="input"'
          + ' type="text"'
          + ' value="' + this._escape(String(value)) + '" />'
          + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "month") {
        return ''
          + '<input class="cell ' + errorCss + '"'
          + ' data-row="' + rowIndex + '"'
          + ' data-field="' + column.key + '"'
          + ' data-type="month"'
          + ' type="month"'
          + ' value="' + this._escape(this._capitalizedToMonthInput(String(value))) + '" />'
          + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "select") {
        var options = this._getOptionsForField(column.key);
        var html = '<select class="select ' + errorCss + '" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="select">';
        html += '<option value="">Select</option>';
        for (var i = 0; i < options.length; i++) {
          var selected = String(options[i].key) === String(value) ? 'selected' : '';
          html += '<option value="' + this._escape(options[i].key) + '" ' + selected + '>' + this._escape(options[i].text) + '</option>';
        }
        html += '</select>';
        html += this._renderFieldErrors(column.key, rowErrors);
        return html;
      }

      return "";
    }

    _renderFieldErrors(fieldName, rowErrors) {
      var messages = [];
      for (var i = 0; i < rowErrors.length; i++) {
        var err = rowErrors[i];
        if (err.field === fieldName) messages.push(err.message);
      }
      if (!messages.length) return "";
      return '<div class="rowErr">' + messages.join("<br>") + '</div>';
    }

    _hasFieldError(fieldName, rowErrors) {
      for (var i = 0; i < rowErrors.length; i++) {
        if (rowErrors[i].field === fieldName) return true;
      }
      return false;
    }

    _getRowErrorMap() {
      var map = {};
      for (var i = 0; i < this._validationErrors.length; i++) {
        var err = this._validationErrors[i];
        var rowIndex = Number(err.rowIndex) - 1;
        if (!map[rowIndex]) map[rowIndex] = [];
        map[rowIndex].push(err);
      }
      return map;
    }

    _bindEvents() {
      var that = this;

      var btnAdd = this._shadowRoot.getElementById("btnAdd");
      if (btnAdd) btnAdd.addEventListener("click", function () { that.addRow(); });

      var btnCopy = this._shadowRoot.getElementById("btnCopy");
      if (btnCopy) btnCopy.addEventListener("click", function () { that.copySelectedRows(); });

      var btnDelete = this._shadowRoot.getElementById("btnDelete");
      if (btnDelete) btnDelete.addEventListener("click", function () { that.deleteSelectedRows(); });

      var btnValidate = this._shadowRoot.getElementById("btnValidate");
      if (btnValidate) btnValidate.addEventListener("click", function () { that.validate(); });

      var btnSave = this._shadowRoot.getElementById("btnSave");
      if (btnSave) btnSave.addEventListener("click", function () { that.save(); });

      var btnClear = this._shadowRoot.getElementById("btnClear");
      if (btnClear) btnClear.addEventListener("click", function () { that.clear(); });

      var selectAll = this._shadowRoot.getElementById("selectAll");
      if (selectAll) {
        selectAll.addEventListener("change", function () {
          that._toggleSelectAll(selectAll.checked);
        });
      }

      var allElements = this._shadowRoot.querySelectorAll("[data-row][data-field]");
      Array.prototype.forEach.call(allElements, function (el) {
        var type = el.getAttribute("data-type");

        if (type === "checkbox") {
          el.addEventListener("change", function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.checked;
            that._updateField(rowIndex, fieldName, value);
          });
          return;
        }

        if (type === "input" || type === "select") {
          el.addEventListener("change", function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.value;
            that._updateField(rowIndex, fieldName, value);
          });
          return;
        }

        if (type === "month") {
          el.addEventListener("change", function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = that._monthInputToCapitalized(this.value);
            that._updateField(rowIndex, fieldName, value);
          });
        }
      });
    }

    _updateField(rowIndex, fieldName, value) {
      this._rows[rowIndex][fieldName] = value;
      this._rows[rowIndex].isModified = true;
      this._rows[rowIndex].rowStatus = "CHANGED";
      this._validationErrors = [];
      this._validationResult = "true";
      this._widgetStatus = "CHANGED";
      this._lastEvent = JSON.stringify({
        type: "fieldChange",
        rowIndex: rowIndex,
        field: fieldName,
        value: value
      });

      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onFieldChange", { rowIndex: rowIndex, field: fieldName, value: value });
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    _fireSimpleEvent(name, detail) {
      this.dispatchEvent(new CustomEvent(name, { detail: detail }));
    }

    _firePropertiesChanged() {
      this.dispatchEvent(new CustomEvent("propertiesChanged", {
        detail: {
          properties: {
            rows: JSON.stringify(this._rows),
            lastEvent: this._lastEvent,
            validationResult: this._validationResult,
            validationErrors: JSON.stringify(this._validationErrors || []),
            savePayload: JSON.stringify(this._savePayload || []),
            rowCount: this.getRowCount(),
            selectedRowCount: this.getSelectedRowCount(),
            widgetStatus: this._widgetStatus
          }
        }
      }));
    }

    _syncRows() {
      this._firePropertiesChanged();
    }

    addRow() {
      this._rows.push(this._createEmptyRow());
      this._widgetStatus = "CHANGED";
      this._lastEvent = JSON.stringify({ type: "addRow" });
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    copySelectedRows() {
      var copiedRows = [];
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          copiedRows.push({
            rowId: "ROW_" + String(this._rowSequence++),
            selected: false,
            isModified: true,
            rowStatus: "NEW",
            ID: this._rows[i].ID || "",
            DESCRIPTION: this._rows[i].DESCRIPTION || "",
            COMPANY: this._rows[i].COMPANY || "",
            COSTCENTER: this._rows[i].COSTCENTER || "",
            ASSET_CLASS: this._rows[i].ASSET_CLASS || "",
            CAPITALIZED: this._rows[i].CAPITALIZED || ""
          });
        }
      }

      for (var j = 0; j < copiedRows.length; j++) {
        this._rows.push(copiedRows[j]);
      }

      this._validationErrors = [];
      this._validationResult = "true";
      this._widgetStatus = "CHANGED";
      this._lastEvent = JSON.stringify({ type: "copySelectedRows", copiedCount: copiedRows.length });
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    deleteSelectedRows() {
      var remainingRows = [];
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected !== true) {
          remainingRows.push(this._rows[i]);
        }
      }

      if (!remainingRows.length) {
        remainingRows = [this._createEmptyRow()];
      }

      this._rows = remainingRows;
      this._normalizeAllRows();
      this._validationErrors = [];
      this._validationResult = "true";
      this._widgetStatus = "CHANGED";
      this._lastEvent = JSON.stringify({ type: "deleteSelectedRows" });
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    clear() {
      this._rows = [this._createEmptyRow()];
      this._validationErrors = [];
      this._validationResult = "true";
      this._savePayload = [];
      this._lastEvent = JSON.stringify({ type: "clear" });
      this._widgetStatus = "READY";
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    validate() {
      var errors = [];
      var assetMap = {};

      for (var i = 0; i < this._rows.length; i++) {
        var row = this._rows[i];
        var rowIndex = i + 1;

        if (!row.ID) errors.push({ rowIndex: rowIndex, field: "ID", message: "ID is mandatory" });
        if (!row.DESCRIPTION) errors.push({ rowIndex: rowIndex, field: "DESCRIPTION", message: "Description is mandatory" });
        if (!row.COMPANY) errors.push({ rowIndex: rowIndex, field: "COMPANY", message: "Company is mandatory" });
        if (!row.COSTCENTER) errors.push({ rowIndex: rowIndex, field: "COSTCENTER", message: "Cost Center is mandatory" });
        if (!row.ASSET_CLASS) errors.push({ rowIndex: rowIndex, field: "ASSET_CLASS", message: "Asset Class is mandatory" });
        if (!row.CAPITALIZED) {
          errors.push({ rowIndex: rowIndex, field: "CAPITALIZED", message: "Capitalized is mandatory" });
        } else if (!/^\d{6}$/.test(String(row.CAPITALIZED))) {
          errors.push({ rowIndex: rowIndex, field: "CAPITALIZED", message: "Capitalized must be in YYYYMM format" });
        }

        var assetKey = this._safeString(row.ID);
        if (assetKey) {
          if (assetMap[assetKey]) {
            errors.push({ rowIndex: rowIndex, field: "ID", message: "Duplicate ID found" });
          } else {
            assetMap[assetKey] = true;
          }
        }
      }

      this._validationErrors = errors;
      this._validationResult = errors.length === 0 ? "true" : "false";
      this._lastEvent = JSON.stringify({
        type: "validate",
        validationResult: this._validationResult,
        errorCount: errors.length
      });
      this._widgetStatus = errors.length === 0 ? "VALID" : "ERROR";

      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onValidate", {
        validationResult: this._validationResult,
        validationErrors: errors
      });

      return this._validationResult;
    }

    save() {
      var validationResult = this.validate();

      if (validationResult !== "true") {
        this._savePayload = [];
        this._widgetStatus = "ERROR";
        this._syncRows();
        this._fireSimpleEvent("onDataChange", {
          rows: this._rows,
          savePayload: [],
          validationErrors: this._validationErrors
        });
        return;
      }

      var payload = [];
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          payload.push({
            ID: this._rows[i].ID,
            DESCRIPTION: this._rows[i].DESCRIPTION,
            COMPANY: this._rows[i].COMPANY,
            COSTCENTER: this._rows[i].COSTCENTER,
            ASSET_CLASS: this._rows[i].ASSET_CLASS,
            CAPITALIZED: this._rows[i].CAPITALIZED
          });
        }
      }

      if (payload.length === 0) {
        this._savePayload = [];
        this._validationErrors = [{
          rowIndex: 1,
          field: "selected",
          message: "Please select at least one row to save"
        }];
        this._validationResult = "false";
        this._widgetStatus = "ERROR";
        this._syncRows();
        this._refreshTable();
        this._fireSimpleEvent("onDataChange", {
          rows: this._rows,
          savePayload: [],
          validationErrors: this._validationErrors
        });
        return;
      }

      this._validationErrors = [];
      this._validationResult = "true";
      this._savePayload = payload;
      this._lastEvent = JSON.stringify({
        type: "save",
        status: "READY",
        payloadCount: payload.length
      });
      this._widgetStatus = "SAVE_READY";
      this._syncRows();
      this._fireSimpleEvent("onDataChange", {
        rows: this._rows,
        savePayload: payload
      });
    }

    getRows() {
      return JSON.stringify(this._rows || []);
    }

    setRows(rowsJson) {
      try {
        this._rows = JSON.parse(rowsJson || "[]");
        if (!Array.isArray(this._rows) || this._rows.length === 0) {
          this._rows = [this._createEmptyRow()];
        }
      } catch (e) {
        this._rows = [this._createEmptyRow()];
      }

      this._normalizeAllRows();
      this._widgetStatus = "LOADED";
      this._syncRows();
      this._refreshTable();
    }

    getRowCount() {
      return this._rows.length;
    }

    getSelectedRowCount() {
      var count = 0;
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) count++;
      }
      return count;
    }

    getRowValue(rowIndex, fieldName) {
      if (rowIndex < 0 || rowIndex >= this._rows.length) return "";
      var row = this._rows[rowIndex];
      if (!row || row[fieldName] === undefined || row[fieldName] === null) return "";
      return String(row[fieldName]);
    }

    getSelectedRowValue(selectedIndex, fieldName) {
      var selectedRows = [];
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) selectedRows.push(this._rows[i]);
      }

      if (selectedIndex < 0 || selectedIndex >= selectedRows.length) return "";
      var row = selectedRows[selectedIndex];
      if (!row || row[fieldName] === undefined || row[fieldName] === null) return "";
      return String(row[fieldName]);
    }

    getValidationErrors() {
      return JSON.stringify(this._validationErrors || []);
    }

    getValidationResult() {
      return this._validationResult || "false";
    }

    getLastEvent() {
      return this._lastEvent || "";
    }

    setCompanyOptions(json) {
      this._companyOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setCostCenterOptions(json) {
      this._costCenterOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setAssetClassOptions(json) {
      this._assetClassOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setVisible(flag) {
      this.style.display = flag ? "block" : "none";
    }
  }

  if (!customElements.get("com-company-assetentrywidget")) {
    customElements.define("com-company-assetentrywidget", AssetEntryWidget);
  }
})(); update js also remove hierachy in this also in capitalized add calender to select capetalized month
