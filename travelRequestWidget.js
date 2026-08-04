(function () {
  class TravelRequestWidget extends HTMLElement {
    constructor() {
      super();
      this._shadowRoot = this.attachShadow({ mode: "open" });
      this._rows = [];
      this._validationErrors = [];
      this._lastEvent = "";

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
      if (this._rows.length === 0) {
        this._rows.push(this._createEmptyRow());
      }
      this._refreshTable();
      this._fireSimpleEvent("onReady", { status: "ready" });
    }

    _createEmptyRow() {
      return {
        selected: false,
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
        Others: ""
      };
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          * { box-sizing: border-box; font-family: Arial, sans-serif; }
          .toolbar { margin-bottom: 10px; display: flex; gap: 8px; }
          button {
            background: #0a6ed1; color: white; border: none; border-radius: 4px;
            padding: 8px 12px; cursor: pointer;
          }
          button.secondary { background: #6c757d; }
          button.danger { background: #d9534f; }
          table { width: 100%; border-collapse: collapse; table-layout: fixed; }
          th, td { border: 1px solid #d9d9d9; padding: 6px; text-align: left; vertical-align: top; }
          th { background: #eaf3ff; font-size: 12px; }
          td input, td select {
            width: 100%; padding: 5px; border: 1px solid #cfcfcf; border-radius: 3px; font-size: 12px;
          }
        </style>

        <div class="toolbar">
          <button id="btnAdd">Add Row</button>
          <button id="btnDelete" class="danger">Delete Selected</button>
          <button id="btnClear" class="secondary">Clear</button>
        </div>

        <div id="tableContainer"></div>
      `;

      this._shadowRoot.getElementById("btnAdd").addEventListener("click", () => this.addRow());
      this._shadowRoot.getElementById("btnDelete").addEventListener("click", () => this.deleteSelectedRows());
      this._shadowRoot.getElementById("btnClear").addEventListener("click", () => this.clear());
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
            </tr>
          </thead>
          <tbody>
      `;

      this._rows.forEach((row, index) => {
        html += `
          <tr>
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
          </tr>
        `;
      });

      html += `</tbody></table>`;
      container.innerHTML = html;

      Array.from(container.querySelectorAll("input, select")).forEach((el) => {
        el.addEventListener("change", (e) => this._handleFieldChange(e));
      });
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

    _escape(value) {
      if (value === undefined || value === null) return "";
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

      const eventObj = {
        type: "fieldChange",
        rowIndex: rowIndex,
        field: field,
        value: value,
        rowData: this._rows[rowIndex]
      };

      this._lastEvent = JSON.stringify(eventObj);

      this.dispatchEvent(new CustomEvent("propertiesChanged", {
        detail: {
          properties: {
            rows: JSON.stringify(this._rows),
            lastEvent: this._lastEvent
          }
        }
      }));

      this._fireSimpleEvent("onFieldChange", eventObj);
      this._fireSimpleEvent("onDataChange", { rows: this._rows });
    }

    _fireSimpleEvent(name, detail) {
      this.dispatchEvent(new CustomEvent(name, { detail: detail }));
    }

    addRow() {
      this._rows.push(this._createEmptyRow());
      this._syncRows();
      this._refreshTable();
    }

    deleteSelectedRows() {
      this._rows = this._rows.filter((r) => !r.selected);
      if (this._rows.length === 0) {
        this._rows.push(this._createEmptyRow());
      }
      this._syncRows();
      this._refreshTable();
    }

    clear() {
      this._rows = [this._createEmptyRow()];
      this._validationErrors = [];
      this._syncRows();
      this._refreshTable();
    }

    validate() {
      const errors = [];

      this._rows.forEach((row, i) => {
        const rowIndex = i + 1;

        if (!row.CostCenter) errors.push({ rowIndex: rowIndex, field: "CostCenter", message: "Cost Center is mandatory" });
        if (!row.Employee) errors.push({ rowIndex: rowIndex, field: "Employee", message: "Employee is mandatory" });
        if (!row.Position) errors.push({ rowIndex: rowIndex, field: "Position", message: "Position is mandatory" });
        if (!row.Route) errors.push({ rowIndex: rowIndex, field: "Route", message: "Route is mandatory" });
        if (!row.ExpectedTiming) errors.push({ rowIndex: rowIndex, field: "ExpectedTiming", message: "Expected Timing is mandatory" });
        if (!row.TripDays) errors.push({ rowIndex: rowIndex, field: "TripDays", message: "Trip Days is mandatory" });
        if (row.TripDays && Number(row.TripDays) <= 0) errors.push({ rowIndex: rowIndex, field: "TripDays", message: "Trip Days must be greater than 0" });
        if (!row.Purpose) errors.push({ rowIndex: rowIndex, field: "Purpose", message: "Purpose is mandatory" });
        if (!row.Initiative) errors.push({ rowIndex: rowIndex, field: "Initiative", message: "Initiative is mandatory" });
      });

      this._validationErrors = errors;

      this.dispatchEvent(new CustomEvent("propertiesChanged", {
        detail: {
          properties: {
            validationResult: errors.length === 0 ? "true" : "false",
            validationErrors: JSON.stringify(errors)
          }
        }
      }));

      this._fireSimpleEvent("onValidate", {
        validationResult: errors.length === 0 ? "true" : "false",
        validationErrors: errors
      });

      return errors.length === 0 ? "true" : "false";
    }

    getRows() { return JSON.stringify(this._rows); }
    getValidationErrors() { return JSON.stringify(this._validationErrors || []); }
    getLastEvent() { return this._lastEvent || ""; }

    setRows(rowsJson) {
      try {
        this._rows = JSON.parse(rowsJson || "[]");
        if (!Array.isArray(this._rows) || this._rows.length === 0) {
          this._rows = [this._createEmptyRow()];
        }
      } catch (e) {
        this._rows = [this._createEmptyRow()];
      }
      this._syncRows();
      this._refreshTable();
    }

    setCostCenterOptions(json) { this._costCenterOptions = this._parseOptions(json); this._refreshTable(); }
    setEmployeeOptions(json) { this._employeeOptions = this._parseOptions(json); this._refreshTable(); }
    setPositionOptions(json) { this._positionOptions = this._parseOptions(json); this._refreshTable(); }
    setRouteOptions(json) { this._routeOptions = this._parseOptions(json); this._refreshTable(); }
    setInitiativeOptions(json) { this._initiativeOptions = this._parseOptions(json); this._refreshTable(); }
    setPurposeOptions(json) { this._purposeOptions = this._parseOptions(json); this._refreshTable(); }
    setYesNoOptions(json) { this._yesNoOptions = this._parseOptions(json); this._refreshTable(); }

    _parseOptions(json) {
      try {
        const arr = JSON.parse(json || "[]");
        return Array.isArray(arr) ? arr : [];
      } catch (e) {
        return [];
      }
    }

    _syncRows() {
      this.dispatchEvent(new CustomEvent("propertiesChanged", {
        detail: {
          properties: {
            rows: JSON.stringify(this._rows)
          }
        }
      }));
    }
  }

  customElements.define("com-company-travelrequestwidget", TravelRequestWidget);
})();
