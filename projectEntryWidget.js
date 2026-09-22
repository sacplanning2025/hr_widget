/*(function () {
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
*/
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

      this._companyCodeOptions = [];
      this._customerOptions = [];

      this._rowOptions = {};

      this._dropdownPanel = null;
      this._dropdownSearch = null;
      this._dropdownList = null;
      this._dropdownOpen = false;
      this._activeDropdownTrigger = null;
      this._activeDropdownRow = -1;
      this._activeDropdownField = "";
      this._activeDropdownOptions = [];
      this._activeDropdownSelectedKey = "";

      this._filteredDropdownOptions = [];
      this._dropdownBatchSize = 100;
      this._dropdownRenderedCount = 0;
      this._dropdownSearchTimer = null;
      this._dropdownListMoreEl = null;

      this._columns = [
        { key: "selected", label: "Sel", type: "checkbox", width: "70px" },
        { key: "CompanyCode", label: "Company Code", type: "select", width: "220px" },
        { key: "ProjectID", label: "Project ID", type: "input", width: "180px" },
        { key: "Description", label: "Description", type: "input", width: "260px" },
        { key: "WBS_Element", label: "WBS Element", type: "input", width: "220px" },
        { key: "CustomerID", label: "Customer ID", type: "select", width: "230px" },
        { key: "ProjectStartDate", label: "Project Start Date", type: "date", width: "180px" },
        { key: "ProjectEndDate", label: "Project End Date", type: "date", width: "180px" },
        { key: "ChanceOfWinning", label: "% Chance of Winning", type: "number", width: "180px" }
      ];

      this._render();
    }

    connectedCallback() {
      if (!this._rows || this._rows.length === 0) {
        this._rows = [this._createEmptyRow()];
      }

      this._createDropdownPanel();
      this._normalizeAllRows();
      this._cleanupRowOptions();
      this._syncRows();
      this._refreshTable();
      this._fireSimpleEvent("onReady", { status: "ready" });
    }

    disconnectedCallback() {
      this._closeDropdown();

      if (this._dropdownPanel && this._dropdownPanel.parentNode) {
        this._dropdownPanel.parentNode.removeChild(this._dropdownPanel);
      }

      if (this._documentClickHandler) {
        document.removeEventListener("click", this._documentClickHandler);
      }

      if (this._windowResizeHandler) {
        window.removeEventListener("resize", this._windowResizeHandler);
      }

      if (this._windowScrollHandler) {
        window.removeEventListener("scroll", this._windowScrollHandler, true);
      }

      this._dropdownPanel = null;
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
      if (oldValue === newValue) {
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
        WBS_Element: "",
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
      row.WBS_Element = this._safeString(row.WBS_Element);
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
        if (!Array.isArray(arr)) {
          return [];
        }

        var normalized = [];
        for (var i = 0; i < arr.length; i++) {
          var item = arr[i];

          if (item && typeof item === "object") {
            var key = item.key !== undefined && item.key !== null ? String(item.key) : "";
            var text = item.text !== undefined && item.text !== null ? String(item.text) : key;
            normalized.push({
              key: key,
              text: text,
              searchText: (key + " " + text).toLowerCase()
            });
          } else {
            var val = item !== undefined && item !== null ? String(item) : "";
            normalized.push({
              key: val,
              text: val,
              searchText: val.toLowerCase()
            });
          }
        }

        return normalized;
      } catch (e) {
        return [];
      }
    }

    _cloneOptions(options) {
      var cloned = [];
      if (!options || !options.length) {
        return cloned;
      }

      for (var i = 0; i < options.length; i++) {
        cloned.push({
          key: options[i].key,
          text: options[i].text,
          searchText: options[i].searchText
        });
      }
      return cloned;
    }

    _ensureRowOptionBucket(rowIndex) {
      if (!this._rowOptions[rowIndex]) {
        this._rowOptions[rowIndex] = {};
      }
    }

    _getRowOptionsForField(rowIndex, fieldName) {
      if (
        this._rowOptions &&
        this._rowOptions[rowIndex] &&
        this._rowOptions[rowIndex][fieldName] &&
        Array.isArray(this._rowOptions[rowIndex][fieldName])
      ) {
        return this._rowOptions[rowIndex][fieldName];
      }
      return null;
    }

    _setRowOptionsForField(rowIndex, fieldName, options) {
      this._ensureRowOptionBucket(rowIndex);
      this._rowOptions[rowIndex][fieldName] = this._cloneOptions(options || []);
    }

    _cloneRowFieldOptions(rowFieldOptions) {
      var result = {};
      if (!rowFieldOptions) {
        return result;
      }

      for (var fieldName in rowFieldOptions) {
        result[fieldName] = this._cloneOptions(rowFieldOptions[fieldName] || []);
      }

      return result;
    }

    _rebuildRowOptionsAfterRowChange(oldRows, newRows) {
      var oldMap = {};
      var newMap = {};
      var rebuilt = {};
      var i = 0;

      for (i = 0; i < oldRows.length; i++) {
        if (oldRows[i] && oldRows[i].rowId) {
          oldMap[oldRows[i].rowId] = i;
        }
      }

      for (i = 0; i < newRows.length; i++) {
        if (newRows[i] && newRows[i].rowId) {
          newMap[newRows[i].rowId] = i;
        }
      }

      for (var rowId in newMap) {
        if (oldMap[rowId] !== undefined && this._rowOptions[oldMap[rowId]]) {
          rebuilt[newMap[rowId]] = this._cloneRowFieldOptions(this._rowOptions[oldMap[rowId]]);
        }
      }

      this._rowOptions = rebuilt;
    }

    _cleanupRowOptions() {
      var rebuilt = {};
      for (var i = 0; i < this._rows.length; i++) {
        if (this._rowOptions[i]) {
          rebuilt[i] = this._cloneRowFieldOptions(this._rowOptions[i]);
        }
      }
      this._rowOptions = rebuilt;
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          :host {
            display:block;
            font-family:"72", Arial, sans-serif;
            color:#223548;
          }

          .wrap {
            border:1px solid #d9e2ef;
            border-radius:12px;
            background:#ffffff;
            overflow:hidden;
          }

          .toolbar {
            display:flex;
            justify-content:flex-end;
            gap:8px;
            padding:12px;
            border-bottom:1px solid #e5edf7;
            background:#f8fbff;
            flex-wrap:wrap;
          }

          .btn {
            border:1px solid #c7d7ea;
            background:#ffffff;
            color:#0a6ed1;
            border-radius:8px;
            padding:8px 14px;
            cursor:pointer;
            font-weight:600;
            font-size:13px;
          }

          .btn:hover {
            background:#f3f8fd;
          }

          .btn.primary {
            background:#0a6ed1;
            color:#ffffff;
            border-color:#0a6ed1;
          }

          .btn.danger {
            color:#bb1e1e;
            border-color:#efb4b4;
            background:#fff7f7;
          }

          .gridWrap {
            overflow:auto;
            max-height:520px;
            background:#ffffff;
          }

          table {
            border-collapse:separate;
            border-spacing:0;
            width:max-content;
            min-width:100%;
          }

          th, td {
            border-bottom:1px solid #edf2f7;
            padding:8px;
            vertical-align:top;
            white-space:nowrap;
          }

          th {
            position:sticky;
            top:0;
            background:#eef4fb;
            z-index:1;
            text-align:left;
            font-size:12px;
            color:#223548;
            font-weight:700;
          }

          tr:hover td {
            background:#fafcff;
          }

          tr.errorRow td {
            background:#fff7f7;
          }

          tr.modifiedRow td {
            background:#fffbeb;
          }

          .cell {
            width:100%;
            min-height:34px;
            height:34px;
            border:1px solid #c9d6e5;
            border-radius:6px;
            padding:6px 10px;
            font-size:13px;
            background:#fff;
            color:#223548;
            box-sizing:border-box;
            outline:none;
          }

          .cell.error {
            border-color:#e25555;
            background:#fff5f5;
          }

          .dropdown-trigger {
            width:100%;
            min-height:34px;
            height:34px;
            border:1px solid #c9d6e5;
            border-radius:6px;
            background:#fff;
            display:flex;
            align-items:center;
            justify-content:space-between;
            box-sizing:border-box;
            padding:0 10px;
            cursor:pointer;
            font-size:13px;
            color:#223548;
            user-select:none;
          }

          .dropdown-trigger .label {
            overflow:hidden;
            text-overflow:ellipsis;
            white-space:nowrap;
            padding-right:8px;
          }

          .dropdown-trigger .arrow {
            color:#6a7f94;
            font-size:11px;
            flex:0 0 auto;
          }

          .dropdown-trigger.error {
            border-color:#e25555;
            background:#fff5f5;
          }

          .rowErr {
            margin-top:4px;
            font-size:11px;
            color:#c53030;
            white-space:normal;
            max-width:240px;
            line-height:1.3;
          }

          .summary {
            padding:10px 12px;
            border-top:1px solid #e5edf7;
            display:flex;
            gap:18px;
            font-size:12px;
            background:#fafcff;
            flex-wrap:wrap;
          }

          .row-checkbox {
            width:22px;
            height:22px;
            cursor:pointer;
            margin-top:8px;
          }

          .select-all-wrap {
            display:flex;
            align-items:center;
            gap:6px;
          }

          .select-all-checkbox {
            width:16px;
            height:16px;
            cursor:pointer;
          }
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

      html += '<div class="gridWrap" id="gridWrap">';
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

        if (rowErrors.length) {
          rowClass = "errorRow";
        } else if (row.isModified === true) {
          rowClass = "modifiedRow";
        }

        html += '<tr class="' + rowClass + '">';

        for (var c = 0; c < this._columns.length; c++) {
          html += '<td style="width:' + this._columns[c].width + '">' + this._renderCell(row, i, this._columns[c], rowErrors) + '</td>';
        }

        html += '</tr>';
      }

      html += '</tbody></table></div>';
      html += this._refreshStatusBarHtml();

      var oldGrid = this._shadowRoot.getElementById("gridWrap");
      var oldLeft = oldGrid ? oldGrid.scrollLeft : 0;
      var oldTop = oldGrid ? oldGrid.scrollTop : 0;

      container.innerHTML = html;
      this._bindEvents();

      var newGrid = this._shadowRoot.getElementById("gridWrap");
      if (newGrid) {
        newGrid.scrollLeft = oldLeft;
        newGrid.scrollTop = oldTop;
      }

      if (this._dropdownOpen && this._activeDropdownTrigger) {
        this._repositionDropdown();
      }
    }

    _renderCell(row, rowIndex, column, rowErrors) {
      var value = row[column.key] !== undefined && row[column.key] !== null ? row[column.key] : "";
      var hasError = this._hasFieldError(column.key, rowErrors);
      var errorCss = hasError ? "error" : "";

      if (column.type === "checkbox") {
        return '<input class="row-checkbox ' + errorCss + '" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="checkbox" type="checkbox" ' + (value === true ? 'checked' : '') + ' />';
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

      if (column.type === "date") {
        return ''
          + '<input class="cell ' + errorCss + '"'
          + ' data-row="' + rowIndex + '"'
          + ' data-field="' + column.key + '"'
          + ' data-type="input"'
          + ' type="date"'
          + ' value="' + this._escape(String(value)) + '" />'
          + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "number") {
        return ''
          + '<input class="cell ' + errorCss + '"'
          + ' data-row="' + rowIndex + '"'
          + ' data-field="' + column.key + '"'
          + ' data-type="input"'
          + ' type="number"'
          + ' min="0" max="100"'
          + ' value="' + this._escape(String(value)) + '" />'
          + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "select") {
        var displayText = this._getOptionText(column.key, value, rowIndex);
        if (!displayText) {
          displayText = "Select";
        }

        return ''
          + '<div class="dropdown-trigger ' + errorCss + '" tabindex="0" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="select">'
          + '<span class="label">' + this._escape(String(displayText)) + '</span>'
          + '<span class="arrow">▼</span>'
          + '</div>'
          + this._renderFieldErrors(column.key, rowErrors);
      }

      return "";
    }

    _getGlobalOptionsForField(fieldName) {
      if (fieldName === "CompanyCode") return this._companyCodeOptions || [];
      if (fieldName === "CustomerID") return this._customerOptions || [];
      return [];
    }

    _getOptionsForField(fieldName, rowIndex) {
      var rowOptions = this._getRowOptionsForField(rowIndex, fieldName);
      if (rowOptions !== null) {
        return rowOptions;
      }
      return this._getGlobalOptionsForField(fieldName);
    }

    _getOptionText(fieldName, value, rowIndex) {
      var options = this._getOptionsForField(fieldName, rowIndex);
      var valueStr = String(value || "");

      for (var i = 0; i < options.length; i++) {
        if (String(options[i].key) === valueStr) {
          return options[i].text;
        }
      }

      var globalOptions = this._getGlobalOptionsForField(fieldName);
      for (var j = 0; j < globalOptions.length; j++) {
        if (String(globalOptions[j].key) === valueStr) {
          return globalOptions[j].text;
        }
      }

      return "";
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

    _hasFieldError(fieldName, rowErrors) {
      for (var i = 0; i < rowErrors.length; i++) {
        if (rowErrors[i].field === fieldName) {
          return true;
        }
      }
      return false;
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

    _createDropdownPanel() {
      if (this._dropdownPanel) {
        return;
      }

      var dropdownPanel = document.createElement("div");
      dropdownPanel.className = "project-widget-dropdown-panel";
      dropdownPanel.style.display = "none";

      dropdownPanel.innerHTML =
        '<div class="dropdown-search-wrap">' +
          '<input type="text" class="dropdown-search-input" placeholder="Search..." />' +
        '</div>' +
        '<div class="dropdown-list"></div>';

      document.body.appendChild(dropdownPanel);

      this._dropdownPanel = dropdownPanel;
      this._dropdownSearch = dropdownPanel.querySelector(".dropdown-search-input");
      this._dropdownList = dropdownPanel.querySelector(".dropdown-list");

      var that = this;

      this._dropdownSearch.addEventListener("input", function () {
        if (that._dropdownSearchTimer) {
          clearTimeout(that._dropdownSearchTimer);
        }

        that._dropdownSearchTimer = setTimeout(function () {
          that._applyDropdownSearch(that._dropdownSearch.value);
        }, 120);
      });

      this._dropdownList.addEventListener("scroll", function () {
        var nearBottom = that._dropdownList.scrollTop + that._dropdownList.clientHeight >= that._dropdownList.scrollHeight - 30;
        if (nearBottom) {
          that._appendNextDropdownBatch();
        }
      });

      this._documentClickHandler = function (e) {
        if (!that._dropdownOpen) {
          return;
        }

        var insidePanel = that._dropdownPanel && that._dropdownPanel.contains(e.target);
        var insideTrigger = that._activeDropdownTrigger && that._activeDropdownTrigger.contains(e.target);

        if (!insidePanel && !insideTrigger) {
          that._closeDropdown();
        }
      };

      this._windowResizeHandler = function () {
        if (that._dropdownOpen) {
          that._repositionDropdown();
        }
      };

      this._windowScrollHandler = function () {
        if (that._dropdownOpen) {
          that._repositionDropdown();
        }
      };

      document.addEventListener("click", this._documentClickHandler);
      window.addEventListener("resize", this._windowResizeHandler);
      window.addEventListener("scroll", this._windowScrollHandler, true);
    }

    _openDropdown(triggerEl, rowIndex, fieldName) {
      this._createDropdownPanel();

      if (
        this._dropdownOpen &&
        this._activeDropdownTrigger === triggerEl &&
        this._activeDropdownRow === rowIndex &&
        this._activeDropdownField === fieldName
      ) {
        this._closeDropdown();
        return;
      }

      this._activeDropdownTrigger = triggerEl;
      this._activeDropdownRow = rowIndex;
      this._activeDropdownField = fieldName;
      this._activeDropdownOptions = this._getOptionsForField(fieldName, rowIndex) || [];
      this._activeDropdownSelectedKey = this._rows[rowIndex] ? this._rows[rowIndex][fieldName] : "";

      this._dropdownSearch.value = "";
      this._filteredDropdownOptions = this._activeDropdownOptions.slice(0);
      this._dropdownRenderedCount = 0;
      this._dropdownList.innerHTML = "";
      this._dropdownOpen = true;

      this._dropdownPanel.style.display = "block";
      this._dropdownPanel.style.position = "fixed";
      this._dropdownPanel.style.zIndex = "999999";

      this._repositionDropdown();
      this._appendNextDropdownBatch();

      var that = this;
      setTimeout(function () {
        if (that._dropdownSearch) {
          that._dropdownSearch.focus();
        }
      }, 0);
    }

    _repositionDropdown() {
      if (!this._dropdownOpen || !this._activeDropdownTrigger || !this._dropdownPanel) {
        return;
      }

      var triggerRect = this._activeDropdownTrigger.getBoundingClientRect();
      var dropdownWidth = Math.max(triggerRect.width, 320);
      var viewportWidth = window.innerWidth;
      var viewportHeight = window.innerHeight;
      var panelMaxHeight = 360;
      var gap = 4;

      var left = triggerRect.left;
      var top = triggerRect.bottom + gap;

      if (left + dropdownWidth > viewportWidth - 10) {
        left = viewportWidth - dropdownWidth - 10;
      }

      if (left < 10) {
        left = 10;
      }

      var spaceBelow = viewportHeight - triggerRect.bottom - 10;
      var spaceAbove = triggerRect.top - 10;

      if (spaceBelow < 220 && spaceAbove > spaceBelow) {
        var computedHeight = Math.min(panelMaxHeight, Math.max(180, spaceAbove));
        this._dropdownPanel.style.maxHeight = computedHeight + "px";
        top = Math.max(10, triggerRect.top - computedHeight - gap);
      } else {
        var computedHeightBelow = Math.min(panelMaxHeight, Math.max(180, spaceBelow));
        this._dropdownPanel.style.maxHeight = computedHeightBelow + "px";
      }

      this._dropdownPanel.style.left = left + "px";
      this._dropdownPanel.style.top = top + "px";
      this._dropdownPanel.style.width = dropdownWidth + "px";
    }

    _closeDropdown() {
      if (this._dropdownPanel) {
        this._dropdownPanel.style.display = "none";
      }

      this._dropdownOpen = false;
      this._activeDropdownTrigger = null;
      this._activeDropdownRow = -1;
      this._activeDropdownField = "";
      this._activeDropdownOptions = [];
      this._activeDropdownSelectedKey = "";
      this._filteredDropdownOptions = [];
      this._dropdownRenderedCount = 0;

      if (this._dropdownList) {
        this._dropdownList.innerHTML = "";
        this._dropdownList.scrollTop = 0;
      }
    }

    _applyDropdownSearch(searchText) {
      var source = this._activeDropdownOptions || [];
      var searchValue = String(searchText || "").toLowerCase().trim();
      var filtered = [];
      var i = 0;

      if (searchValue === "") {
        this._filteredDropdownOptions = source.slice(0);
      } else {
        for (i = 0; i < source.length; i++) {
          if (source[i].searchText.indexOf(searchValue) > -1) {
            filtered.push(source[i]);
          }
        }
        this._filteredDropdownOptions = filtered;
      }

      this._dropdownRenderedCount = 0;
      this._dropdownList.innerHTML = "";
      this._appendNextDropdownBatch();
    }

    _appendNextDropdownBatch() {
      if (!this._dropdownList || !this._filteredDropdownOptions) {
        return;
      }

      var start = this._dropdownRenderedCount;
      var end = start + this._dropdownBatchSize;

      if (start >= this._filteredDropdownOptions.length) {
        if (this._filteredDropdownOptions.length === 0) {
          this._dropdownList.innerHTML = '<div class="dropdown-empty">No results found</div>';
        }
        return;
      }

      var fragment = document.createDocumentFragment();
      var that = this;

      for (var i = start; i < end && i < this._filteredDropdownOptions.length; i++) {
        var opt = this._filteredDropdownOptions[i];
        var item = document.createElement("div");
        item.className = "dropdown-item";

        if (String(opt.key) === String(this._activeDropdownSelectedKey)) {
          item.className += " selected";
        }

        item.textContent = opt.text;
        item.setAttribute("data-key", opt.key);

        item.addEventListener("mousedown", function (e) {
          e.preventDefault();

          var selectedKeyValue = this.getAttribute("data-key");
          var rowIndex = that._activeDropdownRow;
          var fieldName = that._activeDropdownField;

          if (!that._rows[rowIndex]) {
            that._closeDropdown();
            return;
          }

          that._rows[rowIndex][fieldName] = selectedKeyValue;
          that._rows[rowIndex].isModified = true;
          that._rows[rowIndex].rowStatus = "CHANGED";
          that._validationErrors = [];
          that._validationResult = "true";
          that._widgetStatus = "CHANGED";
          that._lastEvent = JSON.stringify({
            type: "fieldChange",
            rowIndex: rowIndex,
            field: fieldName,
            value: selectedKeyValue
          });

          that._syncRows();
          that._refreshTable();
          that._fireSimpleEvent("onFieldChange", { rowIndex: rowIndex, field: fieldName, value: selectedKeyValue });
          that._fireSimpleEvent("onDataChange", { rows: that._rows });
          that._closeDropdown();
        });

        fragment.appendChild(item);
      }

      if (this._dropdownListMoreEl && this._dropdownListMoreEl.parentNode) {
        this._dropdownListMoreEl.parentNode.removeChild(this._dropdownListMoreEl);
      }

      this._dropdownList.appendChild(fragment);
      this._dropdownRenderedCount = Math.min(end, this._filteredDropdownOptions.length);

      if (this._dropdownRenderedCount < this._filteredDropdownOptions.length) {
        if (!this._dropdownListMoreEl) {
          this._dropdownListMoreEl = document.createElement("div");
          this._dropdownListMoreEl.className = "dropdown-more";
        }

        this._dropdownListMoreEl.textContent =
          "Showing " + String(this._dropdownRenderedCount) + " of " + String(this._filteredDropdownOptions.length);

        this._dropdownList.appendChild(this._dropdownListMoreEl);
      }
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

        if (type === "select") {
          el.addEventListener("click", function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            that._openDropdown(this, rowIndex, fieldName);
          });

          el.addEventListener("keydown", function (e) {
            if (e.key === "Enter" || e.key === " ") {
              e.preventDefault();
              var rowIndex = parseInt(this.getAttribute("data-row"), 10);
              var fieldName = this.getAttribute("data-field");
              that._openDropdown(this, rowIndex, fieldName);
            }
          });

          return;
        }

        if (type === "input") {
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
        }
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
      var copiedOptions = [];
      var i = 0;

      for (i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          var newRow = {
            rowId: "ROW_" + String(this._rowSequence++),
            selected: false,
            isModified: true,
            rowStatus: "NEW",
            CompanyCode: this._rows[i].CompanyCode || "",
            ProjectID: this._rows[i].ProjectID || "",
            Description: this._rows[i].Description || "",
            WBS_Element: this._rows[i].WBS_Element || "",
            CustomerID: this._rows[i].CustomerID || "",
            ProjectStartDate: this._rows[i].ProjectStartDate || "",
            ProjectEndDate: this._rows[i].ProjectEndDate || "",
            ChanceOfWinning: this._rows[i].ChanceOfWinning || ""
          };

          copiedRows.push(newRow);

          if (this._rowOptions[i]) {
            copiedOptions.push(this._cloneRowFieldOptions(this._rowOptions[i]));
          } else {
            copiedOptions.push(null);
          }
        }
      }

      if (!copiedRows.length) {
        return;
      }

      for (var j = 0; j < copiedRows.length; j++) {
        this._rows.push(copiedRows[j]);

        var newIndex = this._rows.length - 1;
        if (copiedOptions[j]) {
          this._rowOptions[newIndex] = copiedOptions[j];
        }
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
      var oldRows = this._rows.slice(0);
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
      this._rebuildRowOptionsAfterRowChange(oldRows, this._rows);
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
      this._rowOptions = {};
      this._validationErrors = [];
      this._validationResult = "true";
      this._savePayload = [];
      this._lastEvent = JSON.stringify({ type: "clear" });
      this._widgetStatus = "READY";
      this._closeDropdown();
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
        if (!row.WBS_Element) {
          errors.push({ rowIndex: rowIndex, field: "WBS_Element", message: "WBS Element is mandatory" });
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
        this._savePayload = [];
        this._lastEvent = JSON.stringify({
          type: "save",
          status: "VALIDATION_FAILED",
          validationResult: this._validationResult,
          errorCount: this._validationErrors.length
        });
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
            CompanyCode: this._rows[i].CompanyCode,
            ProjectID: this._rows[i].ProjectID,
            Description: this._rows[i].Description,
            WBS_Element: this._rows[i].WBS_Element,
            CustomerID: this._rows[i].CustomerID,
            ProjectStartDate: this._rows[i].ProjectStartDate,
            ProjectEndDate: this._rows[i].ProjectEndDate,
            ChanceOfWinning: this._rows[i].ChanceOfWinning
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
        this._lastEvent = JSON.stringify({
          type: "save",
          status: "NO_SELECTION",
          payloadCount: 0
        });
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
      var oldRows = this._rows.slice(0);

      try {
        this._rows = JSON.parse(rowsJson || "[]");
        if (!Array.isArray(this._rows) || this._rows.length === 0) {
          this._rows = [this._createEmptyRow()];
        }
      } catch (e) {
        this._rows = [this._createEmptyRow()];
      }

      this._normalizeAllRows();
      this._rebuildRowOptionsAfterRowChange(oldRows, this._rows);
      this._cleanupRowOptions();
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

    setRowFieldOptions(rowIndex, fieldName, json) {
      var index = parseInt(rowIndex, 10);
      if (isNaN(index) || index < 0) {
        return;
      }

      var field = this._safeString(fieldName);
      if (!field) {
        return;
      }

      var parsed = this._parseOptions(json);
      this._setRowOptionsForField(index, field, parsed);

      if (this._rows[index]) {
        var currentValue = this._safeString(this._rows[index][field]);
        var keepValue = false;

        for (var i = 0; i < parsed.length; i++) {
          if (String(parsed[i].key) === currentValue) {
            keepValue = true;
            break;
          }
        }

        if (currentValue !== "" && keepValue === false) {
          this._rows[index][field] = "";
          this._rows[index].isModified = true;
          this._rows[index].rowStatus = "CHANGED";
        }
      }

      this._syncRows();
      this._refreshTable();
    }

    setVisible(flag) {
      this.style.display = flag ? "block" : "none";
    }
  }

  if (!customElements.get("com-company-projectentrywidget")) {
    customElements.define("com-company-projectentrywidget", ProjectEntryWidget);
  }

  (function () {
    if (document.getElementById("project-widget-dropdown-global-style")) {
      return;
    }

    var globalStyleEl = document.createElement("style");
    globalStyleEl.id = "project-widget-dropdown-global-style";
    globalStyleEl.textContent =
      '.project-widget-dropdown-panel {' +
        'background:#ffffff;' +
        'border:1px solid #cfd9e3;' +
        'border-radius:8px;' +
        'box-shadow:0 8px 24px rgba(34,53,72,0.18);' +
        'overflow:hidden;' +
        'min-width:220px;' +
        'max-width:460px;' +
        'max-height:360px;' +
        'z-index:999999;' +
        'font-family:"72", Arial, sans-serif;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-search-wrap {' +
        'padding:8px;' +
        'border-bottom:1px solid #e8eef5;' +
        'background:#ffffff;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-search-input {' +
        'width:100%;' +
        'height:32px;' +
        'border:1px solid #b9cae0;' +
        'border-radius:6px;' +
        'padding:0 10px;' +
        'box-sizing:border-box;' +
        'font-size:13px;' +
        'outline:none;' +
        'color:#223548;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-search-input:focus {' +
        'border-color:#0a6ed1;' +
        'box-shadow:0 0 0 2px rgba(10,110,209,0.12);' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-list {' +
        'max-height:300px;' +
        'overflow:auto;' +
        'background:#ffffff;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-item {' +
        'padding:9px 10px;' +
        'font-size:13px;' +
        'color:#223548;' +
        'cursor:pointer;' +
        'border-bottom:1px solid #f3f6f9;' +
        'line-height:1.35;' +
        'word-break:break-word;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-item:hover {' +
        'background:#edf5ff;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-item.selected {' +
        'background:#e8f2ff;' +
        'color:#0a6ed1;' +
        'font-weight:600;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-empty {' +
        'padding:12px;' +
        'color:#7b8a9a;' +
        'text-align:center;' +
        'font-size:13px;' +
      '}' +
      '.project-widget-dropdown-panel .dropdown-more {' +
        'padding:8px 10px;' +
        'font-size:12px;' +
        'color:#6b7c8f;' +
        'text-align:center;' +
        'background:#fafcff;' +
        'border-top:1px solid #eef3f8;' +
        'position:sticky;' +
        'bottom:0;' +
      '}';
    document.head.appendChild(globalStyleEl);
  })();
})();


