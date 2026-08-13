(function () {
  class ProjectEntryWidget extends HTMLElement {
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
      this._suspendAttributeSync = false;

      this._companyCodeOptions = [];
      this._customerOptions = [];

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
        "lastEvent",
        "validationResult",
        "validationErrors",
        "savePayload",
        "rowCount",
        "selectedRowCount",
        "widgetStatus",
        "companyCodeOptions",
        "customerOptions"
      ];
    }

    attributeChangedCallback(name, oldValue, newValue) {
      if (oldValue === newValue || this._suspendAttributeSync) {
        return;
      }

      if (name === "rows") {
        this.setRows(newValue || "[]");
        return;
      }

      if (name === "companyCodeOptions") {
        this._companyCodeOptions = this._parseOptions(newValue);
        this._refreshTable();
        return;
      }

      if (name === "customerOptions") {
        this._customerOptions = this._parseOptions(newValue);
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
        CompanyCode: "",
        ProjectID: "",
        Description: "",
        CustomerID: "",
        ProjectStartDate: "",
        ProjectEndDate: "",
        ChanceOfWinning: ""
      };
    }

    _normalizeRow(row) {
      if (!row.rowId) {
        row.rowId = "ROW_" + String(this._rowSequence++);
      }

      if (row.selected !== true) {
        row.selected = false;
      }

      if (row.isModified !== true) {
        row.isModified = false;
      }

      if (!row.rowStatus) {
        row.rowStatus = "LOADED";
      }

      row.CompanyCode = this._safeString(row.CompanyCode);
      row.ProjectID = this._safeString(row.ProjectID);
      row.Description = this._safeString(row.Description);
      row.CustomerID = this._safeString(row.CustomerID);
      row.ProjectStartDate = this._safeString(row.ProjectStartDate);
      row.ProjectEndDate = this._safeString(row.ProjectEndDate);
      row.ChanceOfWinning = this._safeString(row.ChanceOfWinning);
    }

    _normalizeAllRows() {
      for (var i = 0; i < this._rows.length; i++) {
        this._normalizeRow(this._rows[i]);
      }
    }

    _safeString(value) {
      if (value === undefined || value === null) {
        return "";
      }
      return String(value).trim();
    }

    _escape(value) {
      if (value === undefined || value === null) {
        return "";
      }
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }

    _parseOptions(json) {
      try {
        var arr = JSON.parse(json || "[]");
        return Array.isArray(arr) ? arr : [];
      } catch (e) {
        return [];
      }
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          :host { display:block; font-family:"72", Arial, sans-serif; color:#223548; }
          .wrap { border:1px solid #d9e2ef; border-radius:12px; background:#ffffff; overflow:hidden; }
          .toolbar { display:flex; justify-content:flex-end; gap:8px; padding:12px; border-bottom:1px solid #e5edf7; background:#f8fbff; flex-wrap:wrap; }
          .btn { border:1px solid #c7d7ea; background:#ffffff; color:#0a6ed1; border-radius:8px; padding:8px 14px; cursor:pointer; font-weight:600; font-size:13px; }
          .btn:hover { background:#f3f8fd; }
          .btn.primary { background:#0a6ed1; color:#ffffff; border-color:#0a6ed1; }
          .btn.danger { color:#bb1e1e; border-color:#efb4b4; background:#fff7f7; }
          .gridWrap { overflow:auto; max-height:520px; background:#ffffff; }
          table { border-collapse:separate; border-spacing:0; width:max-content; min-width:100%; }
          th, td { border-bottom:1px solid #edf2f7; padding:8px; vertical-align:top; white-space:nowrap; }
          th { position:sticky; top:0; background:#eef4fb; z-index:1; text-align:left; font-size:12px; color:#223548; font-weight:700; }
          tr:hover td { background:#fafcff; }
          tr.errorRow td { background:#fff7f7; }
          tr.modifiedRow td { background:#fffbeb; }
          .cell {
            width:100%;
            box-sizing:border-box;
            min-height:34px;
            height:34px;
            border:1px solid #c9d6e5;
            border-radius:6px;
            padding:6px 10px;
            font-size:13px;
            background:#fff;
            color:#223548;
            outline:none;
          }
          .rowErr { margin-top:4px; font-size:11px; color:#c53030; white-space:normal; max-width:220px; line-height:1.3; }
          .summary { padding:10px 12px; border-top:1px solid #e5edf7; display:flex; gap:18px; font-size:12px; background:#fafcff; flex-wrap:wrap; }
          .row-checkbox { width:22px; height:22px; cursor:pointer; margin-top:6px; }
          .select-all-wrap { display:flex; align-items:center; gap:6px; }
          .select-all-checkbox { width:16px; height:16px; cursor:pointer; }
        </style>
        <div class="wrap" id="widgetWrap"></div>
      `;
    }

    _hasSelectedRows() {
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          return true;
        }
      }
      return false;
    }

    _areAllRowsSelected() {
      if (!this._rows || this._rows.length === 0) {
        return false;
      }

      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected !== true) {
          return false;
        }
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
      html += '<th style="width:70px"><div class="select-all-wrap"><span>Sel</span><input class="select-all-checkbox" type="checkbox" id="selectAll" ' + (allSelected ? 'checked' : '') + ' /></div></th>';
      html += '<th style="width:180px">Company Code</th>';
      html += '<th style="width:180px">Project ID</th>';
      html += '<th style="width:260px">Description</th>';
      html += '<th style="width:180px">Customer ID</th>';
      html += '<th style="width:170px">Project Start Date</th>';
      html += '<th style="width:170px">Project End Date</th>';
      html += '<th style="width:180px">% Chance of Winning</th>';
      html += '</tr></thead><tbody>';

      for (var i = 0; i < this._rows.length; i++) {
        var row = this._rows[i];
        var rowErrors = rowErrorMap[i] || [];
        var rowClass = "";

        if (rowErrors.length) {
          rowClass = "errorRow";
        } else if (row.isModified === true) {
          rowClass = "modifiedRow";
        }

        html += '<tr class="' + rowClass + '">';
        html += '<td>' + this._renderCheckboxCell(i, row.selected) + '</td>';
        html += '<td>' + this._renderSelectCell(i, "CompanyCode", row.CompanyCode, this._companyCodeOptions, rowErrors) + '</td>';
        html += '<td>' + this._renderInputCell(i, "ProjectID", row.ProjectID, "text", rowErrors) + '</td>';
        html += '<td>' + this._renderInputCell(i, "Description", row.Description, "text", rowErrors) + '</td>';
        html += '<td>' + this._renderSelectCell(i, "CustomerID", row.CustomerID, this._customerOptions, rowErrors) + '</td>';
        html += '<td>' + this._renderInputCell(i, "ProjectStartDate", row.ProjectStartDate, "date", rowErrors) + '</td>';
        html += '<td>' + this._renderInputCell(i, "ProjectEndDate", row.ProjectEndDate, "date", rowErrors) + '</td>';
        html += '<td>' + this._renderInputCell(i, "ChanceOfWinning", row.ChanceOfWinning, "number", rowErrors) + '</td>';
        html += '</tr>';
      }

      html += '</tbody></table></div>';
      html += this._refreshStatusBarHtml();

      container.innerHTML = html;
      this._bindEvents();
    }

    _renderCheckboxCell(rowIndex, checked) {
      return '<input class="row-checkbox" data-row="' + rowIndex + '" data-field="selected" data-type="checkbox" type="checkbox" ' + (checked ? 'checked' : '') + ' />';
    }

    _renderInputCell(rowIndex, fieldName, value, inputType, rowErrors) {
      return ''
        + '<input class="cell"'
        + ' data-row="' + rowIndex + '"'
        + ' data-field="' + fieldName + '"'
        + ' data-type="input"'
        + ' type="' + inputType + '"'
        + ' value="' + this._escape(value) + '" />'
        + this._renderFieldErrors(fieldName, rowErrors);
    }

    _renderSelectCell(rowIndex, fieldName, value, options, rowErrors) {
      var html = '';
      html += '<select class="cell" data-row="' + rowIndex + '" data-field="' + fieldName + '" data-type="select">';
      html += '<option value="">Select</option>';

      for (var i = 0; i < options.length; i++) {
        var opt = options[i];
        var selected = String(opt.key) === String(value) ? 'selected' : '';
        html += '<option value="' + this._escape(opt.key) + '" ' + selected + '>' + this._escape(opt.text) + '</option>';
      }

      html += '</select>';
      html += this._renderFieldErrors(fieldName, rowErrors);
      return html;
    }

    _renderFieldErrors(fieldName, rowErrors) {
      var messages = [];

      for (var i = 0; i < rowErrors.length; i++) {
        var err = rowErrors[i];
        if (err.field === fieldName) {
          messages.push(err.message);
        }
      }

      if (!messages.length) {
        return "";
      }

      return '<div class="rowErr">' + messages.join("<br>") + '</div>';
    }

    _getRowErrorMap() {
      var map = {};
      for (var i = 0; i < this._validationErrors.length; i++) {
        var err = this._validationErrors[i];
        var rowIndex = Number(err.rowIndex) - 1;
        if (!map[rowIndex]) {
          map[rowIndex] = [];
        }
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

            that._rows[rowIndex][fieldName] = value;
            that._rows[rowIndex].isModified = true;
            that._rows[rowIndex].rowStatus = "CHANGED";
            that._validationErrors = [];
            that._validationResult = "true";
            that._widgetStatus = "CHANGED";
            that._lastEvent = JSON.stringify({
              type: "fieldChange",
              rowIndex: rowIndex,
              field: fieldName,
              value: value
            });

            that._syncRows();
            that._refreshTable();
            that._fireSimpleEvent("onFieldChange", { rowIndex: rowIndex, field: fieldName, value: value });
            that._fireSimpleEvent("onDataChange", { rows: that._rows });
          });
          return;
        }

        el.addEventListener("change", function () {
          var rowIndex = parseInt(this.getAttribute("data-row"), 10);
          var fieldName = this.getAttribute("data-field");
          var value = this.value;

          that._rows[rowIndex][fieldName] = value;
          that._rows[rowIndex].isModified = true;
          that._rows[rowIndex].rowStatus = "CHANGED";
          that._validationErrors = [];
          that._validationResult = "true";
          that._widgetStatus = "CHANGED";
          that._lastEvent = JSON.stringify({
            type: "fieldChange",
            rowIndex: rowIndex,
            field: fieldName,
            value: value
          });

          that._syncRows();
          that._refreshTable();
          that._fireSimpleEvent("onFieldChange", { rowIndex: rowIndex, field: fieldName, value: value });
          that._fireSimpleEvent("onDataChange", { rows: that._rows });
        });
      });
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
            CompanyCode: this._rows[i].CompanyCode || "",
            ProjectID: this._rows[i].ProjectID || "",
            Description: this._rows[i].Description || "",
            CustomerID: this._rows[i].CustomerID || "",
            ProjectStartDate: this._rows[i].ProjectStartDate || "",
            ProjectEndDate: this._rows[i].ProjectEndDate || "",
            ChanceOfWinning: this._rows[i].ChanceOfWinning || ""
          });
        }
      }

      if (!copiedRows.length) {
        return;
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

    _isValidDate(value) {
      if (!value) {
        return false;
      }
      return /^\d{4}-\d{2}-\d{2}$/.test(String(value));
    }

    validate() {
      var errors = [];
      var projectMap = {};

      for (var i = 0; i < this._rows.length; i++) {
        var row = this._rows[i];
        var rowIndex = i + 1;

        if (!row.CompanyCode) {
          errors.push({ rowIndex: rowIndex, field: "CompanyCode", message: "Company Code is mandatory" });
        }
        if (!row.ProjectID) {
          errors.push({ rowIndex: rowIndex, field: "ProjectID", message: "Project ID is mandatory" });
        }
        if (!row.Description) {
          errors.push({ rowIndex: rowIndex, field: "Description", message: "Description is mandatory" });
        }
        if (!row.CustomerID) {
          errors.push({ rowIndex: rowIndex, field: "CustomerID", message: "Customer ID is mandatory" });
        }
        if (!row.ProjectStartDate) {
          errors.push({ rowIndex: rowIndex, field: "ProjectStartDate", message: "Project Start Date is mandatory" });
        }
        if (row.ProjectStartDate && !this._isValidDate(row.ProjectStartDate)) {
          errors.push({ rowIndex: rowIndex, field: "ProjectStartDate", message: "Project Start Date must be in YYYY-MM-DD format" });
        }
        if (!row.ProjectEndDate) {
          errors.push({ rowIndex: rowIndex, field: "ProjectEndDate", message: "Project End Date is mandatory" });
        }
        if (row.ProjectEndDate && !this._isValidDate(row.ProjectEndDate)) {
          errors.push({ rowIndex: rowIndex, field: "ProjectEndDate", message: "Project End Date must be in YYYY-MM-DD format" });
        }
        if (row.ProjectStartDate && row.ProjectEndDate && row.ProjectEndDate < row.ProjectStartDate) {
          errors.push({ rowIndex: rowIndex, field: "ProjectEndDate", message: "Project End Date must be greater than or equal to Start Date" });
        }
        if (row.ChanceOfWinning === "") {
          errors.push({ rowIndex: rowIndex, field: "ChanceOfWinning", message: "Chance of Winning is mandatory" });
        }
        if (row.ChanceOfWinning !== "" && (Number(row.ChanceOfWinning) < 0 || Number(row.ChanceOfWinning) > 100)) {
          errors.push({ rowIndex: rowIndex, field: "ChanceOfWinning", message: "Chance of Winning must be between 0 and 100" });
        }

        var projectKey = [
          this._safeString(row.CompanyCode),
          this._safeString(row.ProjectID)
        ].join("|");

        if (projectKey !== "|") {
          if (projectMap[projectKey]) {
            errors.push({ rowIndex: rowIndex, field: "ProjectID", message: "Duplicate Company Code + Project ID found" });
          } else {
            projectMap[projectKey] = true;
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
        return;
      }

      var payload = [];

      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          payload.push({
            CompanyCode: this._rows[i].CompanyCode,
            ProjectID: this._rows[i].ProjectID,
            Description: this._rows[i].Description,
            CustomerID: this._rows[i].CustomerID,
            ProjectStartDate: this._rows[i].ProjectStartDate,
            ProjectEndDate: this._rows[i].ProjectEndDate,
            ChanceOfWinning: this._rows[i].ChanceOfWinning
          });
        }
      }

      if (payload.length === 0) {
        this._validationErrors = [{
          rowIndex: 0,
          field: "selected",
          message: "Please select at least one row to save"
        }];
        this._validationResult = "false";
        this._widgetStatus = "ERROR";
        this._lastEvent = JSON.stringify({
          type: "save",
          status: "NO_SELECTION"
        });
        this._syncRows();
        this._refreshTable();
        this._fireSimpleEvent("onValidate", {
          validationResult: this._validationResult,
          validationErrors: this._validationErrors
        });
        return;
      }

      this._savePayload = payload;
      this._lastEvent = JSON.stringify({
        type: "save",
        status: "READY",
        payloadCount: payload.length
      });
      this._widgetStatus = "SAVE_READY";
      this._syncRows();
      this._fireSimpleEvent("onDataChange", { rows: this._rows, savePayload: payload });
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
        if (this._rows[i].selected === true) {
          count++;
        }
      }
      return count;
    }

    getRowValue(rowIndex, fieldName) {
      if (rowIndex < 0 || rowIndex >= this._rows.length) {
        return "";
      }
      var row = this._rows[rowIndex];
      if (!row || row[fieldName] === undefined || row[fieldName] === null) {
        return "";
      }
      return String(row[fieldName]);
    }

    getSelectedRowValue(selectedIndex, fieldName) {
      var selectedRows = [];

      for (var i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          selectedRows.push(this._rows[i]);
        }
      }

      if (selectedIndex < 0 || selectedIndex >= selectedRows.length) {
        return "";
      }

      var row = selectedRows[selectedIndex];
      if (!row || row[fieldName] === undefined || row[fieldName] === null) {
        return "";
      }

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

    setCompanyCodeOptions(json) {
      this._companyCodeOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setCustomerOptions(json) {
      this._customerOptions = this._parseOptions(json);
      this._refreshTable();
    }
  }

  if (!customElements.get("com-company-projectentrywidget")) {
    customElements.define("com-company-projectentrywidget", ProjectEntryWidget);
  }
})();
