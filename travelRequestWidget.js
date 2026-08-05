(function () {
  class TravelRequestWidget extends HTMLElement {
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

      this._costCenterOptions = [];
      this._employeeOptions = [];
      this._positionOptions = [];
      this._routeOptions = [];
      this._initiativeOptions = [];
      this._purposeOptions = [];
      this._yesNoOptions = [
        { key: "1", text: "Yes" },
        { key: "0", text: "No" }
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
        "lastEvent",
        "validationResult",
        "validationErrors",
        "savePayload",
        "rowCount",
        "selectedRowCount",
        "widgetStatus",
        "costCenterOptions",
        "employeeOptions",
        "positionOptions",
        "routeOptions",
        "initiativeOptions",
        "purposeOptions",
        "yesNoOptions"
      ];
    }

    attributeChangedCallback(name, oldValue, newValue) {
      if (oldValue === newValue) {
        return;
      }
    }

    _createEmptyRow() {
      return {
        rowId: "ROW_" + String(this._rowSequence++),
        selected: false,
        isModified: false,
        rowStatus: "NEW",
        saveStatus: "",
        saveMessage: "",

        CostCenter: "",
        Employee: "",
        Position: "",
        Route: "",
        ExpectedTiming: "",
        TripDays: "",
        Purpose: "",
        Initiative: "",
        Accommodation: "",
        AirTicket: "",
        Others: "",

        CombinationKey: "",
        AuditUser: "",
        AuditTimestamp: ""
      };
    }

    _normalizeRow(row) {
      if (!row.rowId) {
        row.rowId = "ROW_" + String(this._rowSequence++);
      }

      if (row.selected === undefined) {
        row.selected = false;
      }

      if (row.isModified === undefined) {
        row.isModified = false;
      }

      if (!row.rowStatus) {
        row.rowStatus = "LOADED";
      }

      if (!row.saveStatus) {
        row.saveStatus = "";
      }

      if (!row.saveMessage) {
        row.saveMessage = "";
      }

      row.CostCenter = this._safeString(row.CostCenter);
      row.Employee = this._safeString(row.Employee);
      row.Position = this._safeString(row.Position);
      row.Route = this._safeString(row.Route);
      row.ExpectedTiming = this._safeString(row.ExpectedTiming);
      row.TripDays = this._safeString(row.TripDays);
      row.Purpose = this._safeString(row.Purpose);
      row.Initiative = this._safeString(row.Initiative);
      row.Accommodation = this._safeString(row.Accommodation);
      row.AirTicket = this._safeString(row.AirTicket);
      row.Others = this._safeString(row.Others);

      row.CombinationKey = this._buildCombinationKey(row);
      row.AuditUser = this._safeString(row.AuditUser);
      row.AuditTimestamp = this._safeString(row.AuditTimestamp);
    }

    _normalizeAllRows() {
      var i = 0;
      for (i = 0; i < this._rows.length; i++) {
        this._normalizeRow(this._rows[i]);
      }
    }

    _buildCombinationKey(row) {
      return [
        this._safeString(row.CostCenter),
        this._safeString(row.Employee),
        this._safeString(row.Position),
        this._safeString(row.Route),
        this._safeString(row.ExpectedTiming),
        this._safeString(row.Purpose),
        this._safeString(row.Initiative)
      ].join("|");
    }

    _buildTransactionPayloadRow(row) {
      return {
        rowId: row.rowId,
        combinationKey: this._buildCombinationKey(row),

        dimensions: {
          AHC_COSTCENTER: this._safeString(row.CostCenter),
          AHC_EMPLOYEE: this._safeString(row.Employee),
          AHC_POSITION: this._safeString(row.Position),
          AHC_ROUTE: this._safeString(row.Route),
          Date: this._safeString(row.ExpectedTiming),
          AHC_AUDIT: this._safeString(row.Purpose),
          AHC_SBI: this._safeString(row.Initiative)
        },

        transactions: {
          NO_OF_DAYS: this._safeString(row.TripDays),
          ACCOMODATION: this._safeString(row.Accommodation),
          AIRTICKET: this._safeString(row.AirTicket),
          OTHERS: this._safeString(row.Others)
        },

        display: {
          Purpose: this._safeString(row.Purpose)
        },

        status: {
          rowStatus: this._safeString(row.rowStatus),
          saveStatus: this._safeString(row.saveStatus),
          saveMessage: this._safeString(row.saveMessage)
        }
      };
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          * { box-sizing: border-box; font-family: Arial, sans-serif; }
          .toolbar {
            margin-bottom: 10px;
            display: flex;
            justify-content: flex-end;
            gap: 8px;
            width: 100%;
            flex-wrap: wrap;
          }
          button {
            background: #0a6ed1;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 8px 12px;
            cursor: pointer;
            font-size: 12px;
            font-weight: 600;
          }
          button.secondary { background: #6c757d; }
          button.danger { background: #d9534f; }
          table {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
          }
          th, td {
            border: 1px solid #d9d9d9;
            padding: 6px;
            text-align: left;
            vertical-align: top;
          }
          th {
            background: #eaf3ff;
            font-size: 12px;
          }
          td input, td select {
            width: 100%;
            padding: 5px;
            border: 1px solid #cfcfcf;
            border-radius: 3px;
            font-size: 12px;
            height: 28px;
          }
          .wrap {
            width: 100%;
          }
          .statusbar {
            display: flex;
            gap: 16px;
            margin-bottom: 8px;
            font-size: 12px;
            color: #444;
            flex-wrap: wrap;
          }
          .row-error {
            background: #fff2f2;
          }
          .row-modified {
            background: #fffceb;
          }
          .small {
            font-size: 11px;
            color: #666;
          }
          .save-ok {
            color: #0a7d32;
            font-weight: 600;
          }
          .save-error {
            color: #c62828;
            font-weight: 600;
          }
        </style>

        <div class="wrap">
          <div class="statusbar">
            <div>Total Rows: <span id="totalRows">0</span></div>
            <div>Selected Rows: <span id="selectedRows">0</span></div>
            <div>Validation: <span id="validationState">true</span></div>
            <div>Status: <span id="widgetState">READY</span></div>
          </div>

          <div class="toolbar">
            <button id="btnAdd">Add Row</button>
            <button id="btnDelete" class="danger">Delete Selected</button>
            <button id="btnValidate">Validate</button>
            <button id="btnPrepare">Prepare Payload</button>
            <button id="btnClear" class="secondary">Clear</button>
          </div>
          <div id="tableContainer"></div>
        </div>
      `;

      this._shadowRoot.getElementById("btnAdd").addEventListener("click", () => this.addRow());
      this._shadowRoot.getElementById("btnDelete").addEventListener("click", () => this.deleteSelectedRows());
      this._shadowRoot.getElementById("btnValidate").addEventListener("click", () => this.validate());
      this._shadowRoot.getElementById("btnPrepare").addEventListener("click", () => this.prepareSavePayload());
      this._shadowRoot.getElementById("btnClear").addEventListener("click", () => this.clear());
    }

    _refreshStatusBar() {
      var totalRowsEl = this._shadowRoot.getElementById("totalRows");
      var selectedRowsEl = this._shadowRoot.getElementById("selectedRows");
      var validationStateEl = this._shadowRoot.getElementById("validationState");
      var widgetStateEl = this._shadowRoot.getElementById("widgetState");

      if (totalRowsEl) {
        totalRowsEl.textContent = String(this.getRowCount());
      }
      if (selectedRowsEl) {
        selectedRowsEl.textContent = String(this.getSelectedRowCount());
      }
      if (validationStateEl) {
        validationStateEl.textContent = this._validationResult;
      }
      if (widgetStateEl) {
        widgetStateEl.textContent = this._widgetStatus;
      }
    }

    _refreshTable() {
      const container = this._shadowRoot.getElementById("tableContainer");
      let html = `
        <table>
          <thead>
            <tr>
              <th style="width:40px">Sel</th>
              <th>Cost Center</th>
              <th>Employee</th>
              <th>Position</th>
              <th>Route</th>
              <th>Expected Timing</th>
              <th>Duration of Trip in Days</th>
              <th>Purpose</th>
              <th>Initiative</th>
              <th>Accommodation</th>
              <th>Air Ticket</th>
              <th>Others</th>
              <th>Save Status</th>
            </tr>
          </thead>
          <tbody>
      `;

      this._rows.forEach((row, index) => {
        const rowHasError = this._hasRowError(index + 1);
        const rowClass = rowHasError ? "row-error" : (row.isModified ? "row-modified" : "");

        html += `
          <tr class="${rowClass}">
            <td><input type="checkbox" data-row="${index}" data-field="selected" ${row.selected ? "checked" : ""}></td>
            <td>${this._buildSelect(index, "CostCenter", this._costCenterOptions, row.CostCenter)}</td>
            <td>${this._buildSelect(index, "Employee", this._employeeOptions, row.Employee)}</td>
            <td>${this._buildSelect(index, "Position", this._positionOptions, row.Position)}</td>
            <td>${this._buildSelect(index, "Route", this._routeOptions, row.Route)}</td>
            <td><input type="text" data-row="${index}" data-field="ExpectedTiming" value="${this._escape(row.ExpectedTiming)}" placeholder="YYYYMM"></td>
            <td><input type="number" data-row="${index}" data-field="TripDays" value="${this._escape(row.TripDays)}"></td>
            <td>${this._buildSelect(index, "Purpose", this._purposeOptions, row.Purpose)}</td>
            <td>${this._buildSelect(index, "Initiative", this._initiativeOptions, row.Initiative)}</td>
            <td>${this._buildSelect(index, "Accommodation", this._yesNoOptions, row.Accommodation)}</td>
            <td>${this._buildSelect(index, "AirTicket", this._yesNoOptions, row.AirTicket)}</td>
            <td>${this._buildSelect(index, "Others", this._yesNoOptions, row.Others)}</td>
            <td>${this._renderSaveStatus(row)}</td>
          </tr>
        `;
      });

      html += `</tbody></table>`;
      container.innerHTML = html;

      Array.prototype.forEach.call(container.querySelectorAll("input, select"), (el) => {
        el.addEventListener("change", (e) => this._handleFieldChange(e));
      });

      this._refreshStatusBar();
    }

    _renderSaveStatus(row) {
      if (row.saveStatus === "SUCCESS") {
        return `<span class="save-ok">SUCCESS</span><div class="small">${this._escape(row.saveMessage || "")}</div>`;
      }
      if (row.saveStatus === "ERROR") {
        return `<span class="save-error">ERROR</span><div class="small">${this._escape(row.saveMessage || "")}</div>`;
      }
      return `<span class="small">${this._escape(row.rowStatus || "")}</span>`;
    }

    _buildSelect(rowIndex, fieldName, options, selectedValue) {
      let html = `<select data-row="${rowIndex}" data-field="${fieldName}">`;
      html += `<option value=""></option>`;
      options.forEach((opt) => {
        const selected = String(opt.key) === String(selectedValue) ? "selected" : "";
        html += `<option value="${this._escape(opt.key)}" ${selected}>${this._escape(opt.text)}</option>`;
      });
      html += `</select>`;
      return html;
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

    _handleFieldChange(e) {
      const rowIndex = Number(e.target.getAttribute("data-row"));
      const field = e.target.getAttribute("data-field");
      let value = "";

      if (field === "selected") {
        value = e.target.checked;
      } else {
        value = e.target.value;
      }

      this._rows[rowIndex][field] = value;
      this._rows[rowIndex].isModified = true;
      this._rows[rowIndex].rowStatus = "CHANGED";
      this._rows[rowIndex].saveStatus = "";
      this._rows[rowIndex].saveMessage = "";
      this._rows[rowIndex].CombinationKey = this._buildCombinationKey(this._rows[rowIndex]);

      const eventObj = {
        type: "fieldChange",
        rowIndex: rowIndex,
        field: field,
        value: value
      };

      this._lastEvent = JSON.stringify(eventObj);
      this._widgetStatus = "CHANGED";
      this._validationErrors = [];
      this._validationResult = "true";

      this._syncRows();

      this._fireSimpleEvent("onFieldChange", eventObj);
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
      this._refreshTable();
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
      this._syncRows();
      this._refreshTable();
    }

    deleteSelectedRows() {
      this._rows = this._rows.filter((r) => !r.selected);
      if (this._rows.length === 0) {
        this._rows.push(this._createEmptyRow());
      }
      this._normalizeAllRows();
      this._widgetStatus = "CHANGED";
      this._syncRows();
      this._refreshTable();
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
    }

    _hasRowError(displayRowIndex) {
      var i = 0;
      for (i = 0; i < this._validationErrors.length; i++) {
        if (Number(this._validationErrors[i].rowIndex) === Number(displayRowIndex)) {
          return true;
        }
      }
      return false;
    }

    _isValidExpectedTiming(value) {
      if (!value) {
        return false;
      }
      return /^[0-9]{6}$/.test(String(value));
    }

    validate() {
      const errors = [];
      const combinationMap = {};
      let i = 0;

      for (i = 0; i < this._rows.length; i++) {
        const row = this._rows[i];
        const rowIndex = i + 1;

        if (!row.CostCenter) { errors.push({ rowIndex: rowIndex, field: "CostCenter", message: "Cost Center is mandatory" }); }
        if (!row.Employee) { errors.push({ rowIndex: rowIndex, field: "Employee", message: "Employee is mandatory" }); }
        if (!row.Position) { errors.push({ rowIndex: rowIndex, field: "Position", message: "Position is mandatory" }); }
        if (!row.Route) { errors.push({ rowIndex: rowIndex, field: "Route", message: "Route is mandatory" }); }
        if (!row.ExpectedTiming) { errors.push({ rowIndex: rowIndex, field: "ExpectedTiming", message: "Expected Timing is mandatory" }); }
        if (row.ExpectedTiming && !this._isValidExpectedTiming(row.ExpectedTiming)) {
          errors.push({ rowIndex: rowIndex, field: "ExpectedTiming", message: "Expected Timing must be in YYYYMM format" });
        }
        if (!row.TripDays) { errors.push({ rowIndex: rowIndex, field: "TripDays", message: "Trip Days is mandatory" }); }
        if (row.TripDays && Number(row.TripDays) <= 0) {
          errors.push({ rowIndex: rowIndex, field: "TripDays", message: "Trip Days must be greater than 0" });
        }
        if (!row.Purpose) { errors.push({ rowIndex: rowIndex, field: "Purpose", message: "Purpose is mandatory" }); }
        if (!row.Initiative) { errors.push({ rowIndex: rowIndex, field: "Initiative", message: "Initiative is mandatory" }); }
        if (!row.Accommodation) { errors.push({ rowIndex: rowIndex, field: "Accommodation", message: "Accommodation is mandatory" }); }
        if (!row.AirTicket) { errors.push({ rowIndex: rowIndex, field: "AirTicket", message: "Air Ticket is mandatory" }); }
        if (!row.Others) { errors.push({ rowIndex: rowIndex, field: "Others", message: "Others is mandatory" }); }

        const combinationKey = this._buildCombinationKey(row);
        if (combinationMap[combinationKey]) {
          errors.push({ rowIndex: rowIndex, field: "CombinationKey", message: "Duplicate dimension combination found" });
        } else {
          combinationMap[combinationKey] = true;
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

      this._firePropertiesChanged();

      this._fireSimpleEvent("onValidate", {
        validationResult: this._validationResult,
        validationErrors: errors
      });

      this._refreshTable();
      return this._validationResult;
    }

    prepareSavePayload() {
      this.validate();

      if (this._validationResult !== "true") {
        this._savePayload = [];
        this._firePropertiesChanged();
        return "[]";
      }

      const payload = [];
      let i = 0;
      for (i = 0; i < this._rows.length; i++) {
        payload.push(this._buildTransactionPayloadRow(this._rows[i]));
      }

      this._savePayload = payload;
      this._lastEvent = JSON.stringify({
        type: "prepareSavePayload",
        payloadCount: payload.length
      });
      this._widgetStatus = "PAYLOAD_READY";
      this._firePropertiesChanged();
      return JSON.stringify(payload);
    }

    getSavePayload() {
      return JSON.stringify(this._savePayload || []);
    }

    getRows() {
      return JSON.stringify(this._rows || []);
    }

    getRowCount() {
      return this._rows.length;
    }

    getSelectedRowCount() {
      let count = 0;
      let i = 0;
      for (i = 0; i < this._rows.length; i++) {
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
      const row = this._rows[rowIndex];
      if (!row || row[fieldName] === undefined || row[fieldName] === null) {
        return "";
      }
      return String(row[fieldName]);
    }

    getSelectedRowValue(selectedIndex, fieldName) {
      const selectedRows = [];
      let i = 0;

      for (i = 0; i < this._rows.length; i++) {
        if (this._rows[i].selected === true) {
          selectedRows.push(this._rows[i]);
        }
      }

      if (selectedIndex < 0 || selectedIndex >= selectedRows.length) {
        return "";
      }

      const row = selectedRows[selectedIndex];
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

    setCostCenterOptions(json) {
      this._costCenterOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setEmployeeOptions(json) {
      this._employeeOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setPositionOptions(json) {
      this._positionOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setRouteOptions(json) {
      this._routeOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setInitiativeOptions(json) {
      this._initiativeOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setPurposeOptions(json) {
      this._purposeOptions = this._parseOptions(json);
      this._refreshTable();
    }

    setYesNoOptions(json) {
      this._yesNoOptions = this._parseOptions(json);
      this._refreshTable();
    }

    _parseOptions(json) {
      try {
        const arr = JSON.parse(json || "[]");
        return Array.isArray(arr) ? arr : [];
      } catch (e) {
        return [];
      }
    }
  }

  if (!customElements.get("com-company-travelrequestwidget")) {
    customElements.define("com-company-travelrequestwidget", TravelRequestWidget);
  }
})();
