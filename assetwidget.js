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

      this._openDropdown = null;
      this._dropdownSearchState = {};
      this._openMonthPicker = null;
      this._floatingPanelState = null;

      this._dropdownBatchSize = 100;
      this._dropdownSearchTimer = null;

      this._columns = [
        { key: "selected", label: "Sel", type: "checkbox", width: "70px" },
        { key: "COMPANY", label: "Company", type: "select", width: "190px" },
        { key: "COSTCENTER", label: "Cost Center", type: "select", width: "210px" },
        { key: "ASSET_CLASS", label: "Asset Class", type: "select", width: "210px" },
        { key: "CAPITALIZED", label: "Capitalized", type: "month", width: "210px" },
        { key: "ID", label: "ID", type: "readonly", width: "200px" },
        { key: "DESCRIPTION", label: "Description", type: "input", width: "240px" }
      ];

      this._render();
      this._bindOutsideClick();
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
      return ["rows", "companyOptions", "costCenterOptions", "assetClassOptions"];
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
        COMPANY: "",
        COSTCENTER: "",
        ASSET_CLASS: "",
        CAPITALIZED: "",
        ID: "",
        DESCRIPTION: ""
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

    _getOptionsForField(fieldName) {
      if (fieldName === "COMPANY") return this._companyOptions || [];
      if (fieldName === "COSTCENTER") return this._costCenterOptions || [];
      if (fieldName === "ASSET_CLASS") return this._assetClassOptions || [];
      return [];
    }

    _normalizeCapitalized(value) {
      var v = this._safeString(value);
      if (/^\d{6}$/.test(v)) return v;
      if (/^\d{4}-\d{2}$/.test(v)) return v.replace("-", "");
      return "";
    }

    _capitalizedToDisplay(value) {
      var v = this._normalizeCapitalized(value);
      if (/^\d{6}$/.test(v)) {
        var year = v.substring(0, 4);
        var month = v.substring(4, 6);
        return year + " - " + this._getMonthShortName(month);
      }
      return "";
    }

    _getMonthShortName(month) {
      var map = {
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec"
      };
      return map[month] || "";
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          :host { display: block; font-family: "72", Arial, sans-serif; color: #223548; }
          .wrap { border: 1px solid #d9e2ef; border-radius: 12px; background: #ffffff; overflow: visible; position: relative; }
          .titleBar { padding: 14px 16px 8px 16px; border-bottom: 1px solid #e5edf7; background: #ffffff; }
          .titleText { font-size: 24px; font-weight: 700; color: #0a6ed1; line-height: 1.2; }
          .toolbar { display: flex; justify-content: flex-end; gap: 8px; padding: 12px; border-bottom: 1px solid #e5edf7; background: #f8fbff; flex-wrap: wrap; }
          .btn { border: 1px solid #c7d7ea; background: #ffffff; color: #0a6ed1; border-radius: 8px; padding: 8px 14px; cursor: pointer; font-weight: 600; font-size: 13px; }
          .btn:hover { background: #f3f8fd; }
          .btn.primary { background: #0a6ed1; color: #ffffff; border-color: #0a6ed1; }
          .btn.danger { color: #bb1e1e; border-color: #efb4b4; background: #fff7f7; }
          .gridWrap { overflow-x: auto; overflow-y: visible; max-height: 520px; background: #ffffff; position: relative; }
          table { border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; }
          thead, tbody, tr { overflow: visible; }
          th, td { border-bottom: 1px solid #edf2f7; padding: 8px; vertical-align: top; white-space: nowrap; overflow: visible; position: relative; }
          th { position: sticky; top: 0; background: #eef4fb; z-index: 1; text-align: left; font-size: 12px; color: #223548; font-weight: 700; }
          tr:hover td { background: #fafcff; }
          tr.errorRow td { background: #fff7f7; }
          tr.modifiedRow td { background: #fffbeb; }
          .cell { width: 100%; min-height: 34px; height: 34px; border: 1px solid #c9d6e5; border-radius: 6px; padding: 6px 10px; font-size: 13px; background: #fff; color: #223548; box-sizing: border-box; outline: none; }
          .cell.error { border-color: #e25555; background: #fff5f5; }
          .readonlyCell { width: 100%; min-height: 34px; height: 34px; border: 1px solid #d5dee8; border-radius: 6px; padding: 6px 10px; font-size: 13px; background: #f7f9fb; color: #223548; box-sizing: border-box; display: flex; align-items: center; }
          .rowErr { margin-top: 4px; font-size: 11px; color: #c53030; white-space: normal; max-width: 240px; line-height: 1.3; }
          .summary { padding: 10px 12px; border-top: 1px solid #e5edf7; display: flex; gap: 18px; font-size: 12px; background: #fafcff; flex-wrap: wrap; }
          .row-checkbox { width: 22px; height: 22px; cursor: pointer; margin-top: 8px; }
          .select-all-wrap { display: flex; align-items: center; gap: 6px; }
          .select-all-checkbox { width: 16px; height: 16px; cursor: pointer; }

          .dropdownWrap, .monthWrap { position: relative; width: 100%; overflow: visible; }
          .dropdownDisplay, .monthDisplay {
            width: 100%;
            min-height: 34px;
            height: 34px;
            border: 1px solid #c9d6e5;
            border-radius: 6px;
            background: #fff;
            box-sizing: border-box;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            padding: 0 8px 0 10px;
            cursor: pointer;
            font-size: 13px;
          }
          .dropdownDisplay.error, .monthDisplay.error { border-color: #e25555; background: #fff5f5; }
          .dropdownValue, .monthValue { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; color: #223548; }
          .dropdownPlaceholder, .monthPlaceholder { color: #6b7c93; }
          .dropdownIcons { display: flex; align-items: center; gap: 6px; flex-shrink: 0; }
          .dropdownArrow, .closeIcon {
            border: none;
            background: transparent;
            cursor: pointer;
            color: #5f6b7a;
            font-size: 14px;
            line-height: 1;
            padding: 0;
            width: 18px;
            height: 18px;
          }

          .dropdownHeader, .monthHeader { display: flex; align-items: center; gap: 8px; padding: 8px; border-bottom: 1px solid #edf2f7; background: #f8fbff; }
          .dropdownSearch { flex: 1; height: 32px; border: 1px solid #c9d6e5; border-radius: 6px; padding: 6px 10px; font-size: 13px; box-sizing: border-box; outline: none; }
          .dropdownList { max-height: 220px; overflow: auto; background: #fff; }
          .dropdownItem { padding: 8px 10px; cursor: pointer; font-size: 13px; border-bottom: 1px solid #f4f7fb; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
          .dropdownItem:hover { background: #f3f8fd; }
          .dropdownItem.selected { background: #eaf3fc; color: #0a6ed1; font-weight: 600; }
          .dropdownEmpty { padding: 10px; font-size: 12px; color: #6b7c93; }
          .dropdownMore { padding: 8px 10px; font-size: 12px; color: #6b7c93; text-align: center; border-top: 1px solid #edf2f7; background: #fafcff; }

          .dropdownFooter, .monthFooter { display: flex; justify-content: space-between; align-items: center; padding: 8px; border-top: 1px solid #edf2f7; background: #f8fbff; }
          .linkBtn { border: none; background: transparent; color: #0a6ed1; font-size: 12px; cursor: pointer; padding: 0; }
          .yearSelect { flex: 1; height: 32px; border: 1px solid #c9d6e5; border-radius: 6px; padding: 4px 8px; font-size: 13px; outline: none; background: #fff; }
          .monthGrid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; padding: 10px; }
          .monthBtn { border: 1px solid #c9d6e5; background: #fff; border-radius: 6px; padding: 8px 0; cursor: pointer; font-size: 13px; color: #223548; }
          .monthBtn:hover { background: #f3f8fd; }
          .monthBtn.selected { background: #5d6673; border-color: #5d6673; color: #fff; }

          .floatingLayer { position: absolute; inset: 0; overflow: visible; pointer-events: none; z-index: 999998; }
          .floatingPanel { position: absolute; background: #fff; border: 1px solid #c9d6e5; border-radius: 8px; box-shadow: 0 8px 24px rgba(34,53,72,0.16); z-index: 999999; overflow: hidden; pointer-events: auto; }
          .floatingDropdownPanel { min-width: 220px; }
          .floatingMonthPanel { width: 250px; }
        </style>
        <div class="wrap">
          <div class="titleBar">
            <div class="titleText">New Asset Creation</div>
          </div>
          <div id="widgetWrap"></div>
          <div class="floatingLayer" id="floatingLayer"></div>
        </div>
      `;
    }

    _bindOutsideClick() {
      var that = this;

      this._shadowRoot.addEventListener("click", function (e) {
        e.stopPropagation();
      });

      document.addEventListener("click", function () {
        if (that._openDropdown !== null || that._openMonthPicker !== null) {
          that._openDropdown = null;
          that._openMonthPicker = null;
          that._floatingPanelState = null;
          that._renderFloatingPanel();
        }
      });
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
      var html = "";

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
          html += '<th style="width:' + col.width + '"><div class="select-all-wrap"><span>Sel</span><input class="select-all-checkbox" type="checkbox" id="selectAll" ' + (allSelected ? "checked" : "") + " /></div></th>";
        } else {
          html += '<th style="width:' + col.width + '">' + col.label + "</th>";
        }
      }

      html += "</tr></thead><tbody>";

      for (var i = 0; i < this._rows.length; i++) {
        var row = this._rows[i];
        var rowErrors = rowErrorMap[i] || [];
        var rowClass = "";

        if (rowErrors.length) rowClass = "errorRow";
        else if (row.isModified === true) rowClass = "modifiedRow";

        html += '<tr class="' + rowClass + '">';
        for (var c = 0; c < this._columns.length; c++) {
          html += '<td style="width:' + this._columns[c].width + '">' + this._renderCell(row, i, this._columns[c], rowErrors) + "</td>";
        }
        html += "</tr>";
      }

      html += "</tbody></table></div>";
      html += this._refreshStatusBarHtml();

      container.innerHTML = html;
      this._renderFloatingPanel();
      this._bindEvents();
    }

    _renderFloatingPanel() {
      var layer = this._shadowRoot.getElementById("floatingLayer");
      if (!layer) return;

      if (!this._floatingPanelState) {
        layer.innerHTML = "";
        return;
      }

      var p = this._floatingPanelState;
      var html = "";

      if (p.type === "select") {
        html += '<div class="floatingPanel floatingDropdownPanel" style="top:' + p.top + "px; left:" + p.left + "px; width:" + p.width + 'px;">';
        html += '<div class="dropdownHeader">';
        html += '<input class="dropdownSearch" type="text" placeholder="Search..." data-action="searchDropdown" data-dropdown-id="' + p.dropdownId + '" value="' + this._escape(p.searchValue || "") + '" />';
        html += "</div>";
        html += '<div class="dropdownList" data-dropdown-list="' + p.dropdownId + '"></div>';
        html += '<div class="dropdownFooter">';
        html += '<button type="button" class="linkBtn" data-action="clearSelect" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '">Clear selection</button>';
        html += '<button type="button" class="linkBtn" data-action="closeDropdown" data-dropdown-id="' + p.dropdownId + '">Close</button>';
        html += "</div>";
        html += "</div>";
      }

      if (p.type === "month") {
        html += '<div class="floatingPanel floatingMonthPanel" style="top:' + p.top + "px; left:" + p.left + 'px;">';
        html += '<div class="monthHeader">';
        html += '<select class="yearSelect" data-action="changeMonthYear" data-picker-id="' + p.pickerId + '" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '">';
        for (var year = 2026; year <= 3200; year++) {
          var yearStr = String(year);
          var selected = yearStr === p.selectedYear ? "selected" : "";
          html += '<option value="' + yearStr + '" ' + selected + ">" + yearStr + "</option>";
        }
        html += "</select>";
        html += '<button type="button" class="closeIcon" data-action="closeMonth" data-picker-id="' + p.pickerId + '" title="Close">×</button>';
        html += "</div>";

        var monthNames = [
          { key: "01", text: "Jan" }, { key: "02", text: "Feb" }, { key: "03", text: "Mar" }, { key: "04", text: "Apr" },
          { key: "05", text: "May" }, { key: "06", text: "Jun" }, { key: "07", text: "Jul" }, { key: "08", text: "Aug" },
          { key: "09", text: "Sep" }, { key: "10", text: "Oct" }, { key: "11", text: "Nov" }, { key: "12", text: "Dec" }
        ];

        html += '<div class="monthGrid">';
        for (var m = 0; m < monthNames.length; m++) {
          var selectedMonthCss = monthNames[m].key === p.selectedMonth ? "selected" : "";
          html += '<button type="button" class="monthBtn ' + selectedMonthCss + '" data-action="pickMonth" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '" data-picker-id="' + p.pickerId + '" data-year="' + p.selectedYear + '" data-month="' + monthNames[m].key + '">' + monthNames[m].text + "</button>";
        }
        html += "</div>";

        html += '<div class="monthFooter">';
        html += '<button type="button" class="linkBtn" data-action="clearMonth" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '">Clear selection</button>';
        html += '<button type="button" class="linkBtn" data-action="closeMonth" data-picker-id="' + p.pickerId + '">Close</button>';
        html += "</div>";
        html += "</div>";
      }

      layer.innerHTML = html;

      if (p.type === "select") {
        this._renderDropdownBatch(true);
      }
    }

    _renderDropdownBatch(reset) {
      if (!this._floatingPanelState || this._floatingPanelState.type !== "select") return;

      var p = this._floatingPanelState;
      var listEl = this._shadowRoot.querySelector('[data-dropdown-list="' + p.dropdownId + '"]');
      if (!listEl) return;

      if (reset) {
        p.renderedCount = 0;
        listEl.innerHTML = "";
      }

      var start = p.renderedCount || 0;
      var end = Math.min(start + this._dropdownBatchSize, p.filteredOptions.length);

      if (start >= p.filteredOptions.length) {
        if (p.filteredOptions.length === 0 && reset) {
          listEl.innerHTML = '<div class="dropdownEmpty">No results found</div>';
        }
        return;
      }

      var html = "";
      for (var i = start; i < end; i++) {
        var opt = p.filteredOptions[i];
        var selectedCss = String(opt.key) === String(p.value) ? "selected" : "";
        html += '<div class="dropdownItem ' + selectedCss + '" data-action="chooseOption" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '" data-value="' + this._escape(opt.key) + '">' + this._escape(opt.text) + "</div>";
      }

      var temp = document.createElement("div");
      temp.innerHTML = html;
      while (temp.firstChild) {
        listEl.appendChild(temp.firstChild);
      }

      p.renderedCount = end;

      var oldMore = listEl.querySelector(".dropdownMore");
      if (oldMore) {
        oldMore.parentNode.removeChild(oldMore);
      }

      if (p.renderedCount < p.filteredOptions.length) {
        var more = document.createElement("div");
        more.className = "dropdownMore";
        more.textContent = "Showing " + p.renderedCount + " of " + p.filteredOptions.length;
        listEl.appendChild(more);
      }
    }

    _filterDropdownOptions(fieldName, searchValue) {
      var options = this._getOptionsForField(fieldName);
      var q = this._safeString(searchValue).toLowerCase();
      if (q === "") return options.slice(0);

      var filtered = [];
      for (var i = 0; i < options.length; i++) {
        if ((options[i].searchText || "").indexOf(q) !== -1) {
          filtered.push(options[i]);
        }
      }
      return filtered;
    }

    _getSelectedText(fieldName, value) {
      var options = this._getOptionsForField(fieldName);
      for (var i = 0; i < options.length; i++) {
        if (String(options[i].key) === String(value)) {
          return options[i].text;
        }
      }
      return "";
    }

    _renderCell(row, rowIndex, column, rowErrors) {
      var value = row[column.key] !== undefined && row[column.key] !== null ? row[column.key] : "";
      var hasError = this._hasFieldError(column.key, rowErrors);
      var errorCss = hasError ? "error" : "";

      if (column.type === "checkbox") {
        return '<input class="row-checkbox" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="checkbox" type="checkbox" ' + (value === true ? "checked" : "") + " />";
      }

      if (column.type === "readonly") {
        return '<div class="readonlyCell">' + this._escape(String(value)) + "</div>" + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "input") {
        return '<input class="cell ' + errorCss + '" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="input" type="text" value="' + this._escape(String(value)) + '" />' + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "month") {
        return this._renderMonthCell(row, rowIndex, column, rowErrors, errorCss);
      }

      if (column.type === "select") {
        return this._renderSelectCell(row, rowIndex, column, rowErrors, errorCss);
      }

      return "";
    }

    _renderSelectCell(row, rowIndex, column, rowErrors, errorCss) {
      var value = this._safeString(row[column.key]);
      var displayText = this._getSelectedText(column.key, value);
      var placeholder = "Select";
      var dropdownId = "dropdown_" + rowIndex + "_" + column.key;
      var html = "";

      html += '<div class="dropdownWrap">';
      html += '<div class="dropdownDisplay ' + errorCss + '" data-action="toggleDropdown" data-dropdown-id="' + dropdownId + '" data-row="' + rowIndex + '" data-field="' + column.key + '">';
      if (displayText) {
        html += '<div class="dropdownValue">' + this._escape(displayText) + "</div>";
      } else {
        html += '<div class="dropdownValue dropdownPlaceholder">' + this._escape(placeholder) + "</div>";
      }
      html += '<div class="dropdownIcons">';
      html += '<button type="button" class="dropdownArrow" data-action="toggleDropdown" data-dropdown-id="' + dropdownId + '" data-row="' + rowIndex + '" data-field="' + column.key + '" title="Open">▼</button>';
      html += "</div>";
      html += "</div>";
      html += "</div>";
      html += this._renderFieldErrors(column.key, rowErrors);

      return html;
    }

    _renderMonthCell(row, rowIndex, column, rowErrors, errorCss) {
      var value = this._safeString(row[column.key]);
      var displayValue = this._capitalizedToDisplay(value);
      var pickerId = "month_" + rowIndex + "_" + column.key;
      var html = "";

      html += '<div class="monthWrap">';
      html += '<div class="monthDisplay ' + errorCss + '" data-action="toggleMonth" data-picker-id="' + pickerId + '" data-row="' + rowIndex + '" data-field="' + column.key + '">';
      if (displayValue) {
        html += '<div class="monthValue">' + this._escape(displayValue) + "</div>";
      } else {
        html += '<div class="monthValue monthPlaceholder">Select month</div>';
      }
      html += '<div class="dropdownIcons">';
      html += '<button type="button" class="dropdownArrow" data-action="toggleMonth" data-picker-id="' + pickerId + '" data-row="' + rowIndex + '" data-field="' + column.key + '" title="Open">📅</button>';
      html += "</div>";
      html += "</div>";
      html += "</div>";
      html += this._renderFieldErrors(column.key, rowErrors);

      return html;
    }

    _openSelectPanel(rowIndex, fieldName, dropdownId, anchorEl) {
      var searchValue = this._dropdownSearchState[dropdownId] || "";
      var filteredOptions = this._filterDropdownOptions(fieldName, searchValue);
      var wrapRect = this._shadowRoot.querySelector(".wrap").getBoundingClientRect();
      var rect = anchorEl.getBoundingClientRect();

      this._floatingPanelState = {
        type: "select",
        dropdownId: dropdownId,
        rowIndex: rowIndex,
        fieldName: fieldName,
        value: this._safeString(this._rows[rowIndex][fieldName]),
        searchValue: searchValue,
        filteredOptions: filteredOptions,
        renderedCount: 0,
        top: rect.bottom - wrapRect.top + 4,
        left: rect.left - wrapRect.left,
        width: Math.max(rect.width, 220)
      };

      this._openMonthPicker = null;
      this._openDropdown = dropdownId;
      this._renderFloatingPanel();
      this._bindEvents();

      var listEl = this._shadowRoot.querySelector('[data-dropdown-list="' + dropdownId + '"]');
      if (listEl) {
        var that = this;
        listEl.onscroll = function () {
          if (!that._floatingPanelState || that._floatingPanelState.dropdownId !== dropdownId) return;
          var nearBottom = listEl.scrollTop + listEl.clientHeight >= listEl.scrollHeight - 20;
          if (nearBottom) {
            that._renderDropdownBatch(false);
            that._bindEvents();
          }
        };
      }
    }

    _openMonthPanel(rowIndex, fieldName, pickerId, anchorEl) {
      var value = this._safeString(this._rows[rowIndex][fieldName]);
      var selectedYear = "2026";
      var selectedMonth = "";

      if (/^\d{6}$/.test(value)) {
        selectedYear = value.substring(0, 4);
        selectedMonth = value.substring(4, 6);
      }

      var wrapRect = this._shadowRoot.querySelector(".wrap").getBoundingClientRect();
      var rect = anchorEl.getBoundingClientRect();

      this._floatingPanelState = {
        type: "month",
        pickerId: pickerId,
        rowIndex: rowIndex,
        fieldName: fieldName,
        selectedYear: selectedYear,
        selectedMonth: selectedMonth,
        top: rect.bottom - wrapRect.top + 4,
        left: rect.left - wrapRect.left
      };

      this._openDropdown = null;
      this._openMonthPicker = pickerId;
      this._renderFloatingPanel();
      this._bindEvents();
    }

    _renderFieldErrors(fieldName, rowErrors) {
      var messages = [];
      for (var i = 0; i < rowErrors.length; i++) {
        var err = rowErrors[i];
        if (err.field === fieldName) messages.push(err.message);
      }
      if (!messages.length) return "";
      return '<div class="rowErr">' + messages.join("<br>") + "</div>";
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
      if (btnAdd) btnAdd.onclick = function () { that.addRow(); };

      var btnCopy = this._shadowRoot.getElementById("btnCopy");
      if (btnCopy) btnCopy.onclick = function () { that.copySelectedRows(); };

      var btnDelete = this._shadowRoot.getElementById("btnDelete");
      if (btnDelete) btnDelete.onclick = function () { that.deleteSelectedRows(); };

      var btnValidate = this._shadowRoot.getElementById("btnValidate");
      if (btnValidate) btnValidate.onclick = function () { that.validate(); };

      var btnSave = this._shadowRoot.getElementById("btnSave");
      if (btnSave) btnSave.onclick = function () { that.save(); };

      var btnClear = this._shadowRoot.getElementById("btnClear");
      if (btnClear) btnClear.onclick = function () { that.clear(); };

      var selectAll = this._shadowRoot.getElementById("selectAll");
      if (selectAll) {
        selectAll.onchange = function () {
          that._toggleSelectAll(selectAll.checked);
        };
      }

      var allElements = this._shadowRoot.querySelectorAll("[data-row][data-field]");
      Array.prototype.forEach.call(allElements, function (el) {
        var type = el.getAttribute("data-type");

        if (type === "checkbox") {
          el.onchange = function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.checked;
            that._updateField(rowIndex, fieldName, value);
          };
          return;
        }

        if (type === "input") {
          el.onchange = function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.value;
            that._updateField(rowIndex, fieldName, value);
          };
          return;
        }
      });

      var actionElements = this._shadowRoot.querySelectorAll("[data-action]");
      Array.prototype.forEach.call(actionElements, function (el) {
        var action = el.getAttribute("data-action");

        if (action === "toggleDropdown") {
          el.onclick = function (e) {
            e.stopPropagation();
            var dropdownId = this.getAttribute("data-dropdown-id");
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");

            if (that._openDropdown === dropdownId) {
              that._openDropdown = null;
              that._floatingPanelState = null;
              that._renderFloatingPanel();
            } else {
              that._openSelectPanel(rowIndex, fieldName, dropdownId, this.closest(".dropdownDisplay"));
              setTimeout(function () {
                var input = that._shadowRoot.querySelector('.dropdownSearch[data-dropdown-id="' + dropdownId + '"]');
                if (input) input.focus();
              }, 0);
            }
          };
        }

        if (action === "closeDropdown") {
          el.onclick = function (e) {
            e.stopPropagation();
            that._openDropdown = null;
            that._floatingPanelState = null;
            that._renderFloatingPanel();
          };
        }

        if (action === "searchDropdown") {
          el.onclick = function (e) {
            e.stopPropagation();
          };

          el.oninput = function (e) {
            e.stopPropagation();

            var inputEl = this;
            var dropdownId = inputEl.getAttribute("data-dropdown-id");

            if (that._dropdownSearchTimer) {
              clearTimeout(that._dropdownSearchTimer);
            }

            that._dropdownSearchTimer = setTimeout(function () {
              that._dropdownSearchState[dropdownId] = inputEl.value;

              if (that._floatingPanelState && that._floatingPanelState.type === "select") {
                var fieldName = that._floatingPanelState.fieldName;
                that._floatingPanelState.searchValue = inputEl.value;
                that._floatingPanelState.filteredOptions = that._filterDropdownOptions(fieldName, inputEl.value);
                that._floatingPanelState.renderedCount = 0;
                that._renderFloatingPanel();
                that._bindEvents();

                setTimeout(function () {
                  var input = that._shadowRoot.querySelector('.dropdownSearch[data-dropdown-id="' + dropdownId + '"]');
                  if (input) {
                    input.focus();
                    var len = input.value.length;
                    input.setSelectionRange(len, len);
                  }
                }, 0);
              }
            }, 120);
          };
        }

        if (action === "chooseOption") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.getAttribute("data-value");
            that._openDropdown = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, value);
          };
        }

        if (action === "clearSelect") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            that._openDropdown = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, "");
          };
        }

        if (action === "toggleMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            var pickerId = this.getAttribute("data-picker-id");
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");

            if (that._openMonthPicker === pickerId) {
              that._openMonthPicker = null;
              that._floatingPanelState = null;
              that._renderFloatingPanel();
            } else {
              that._openMonthPanel(rowIndex, fieldName, pickerId, this.closest(".monthDisplay"));
            }
          };
        }

        if (action === "closeMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            that._openMonthPicker = null;
            that._floatingPanelState = null;
            that._renderFloatingPanel();
          };
        }

        if (action === "changeMonthYear") {
          el.onclick = function (e) {
            e.stopPropagation();
          };

          el.onchange = function (e) {
            e.stopPropagation();
            var pickerId = this.getAttribute("data-picker-id");
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var currentValue = that._safeString(that._rows[rowIndex][fieldName]);
            var month = "01";

            if (/^\d{6}$/.test(currentValue)) {
              month = currentValue.substring(4, 6);
            }

            that._rows[rowIndex][fieldName] = this.value + month;
            that._rows[rowIndex].isModified = true;
            that._rows[rowIndex].rowStatus = "CHANGED";
            that._widgetStatus = "CHANGED";
            that._validationErrors = [];
            that._validationResult = "true";
            that._syncRows();

            var anchor = that._shadowRoot.querySelector('.monthDisplay[data-picker-id="' + pickerId + '"]');
            if (anchor) {
              that._openMonthPanel(rowIndex, fieldName, pickerId, anchor);
            } else {
              that._refreshTable();
            }
          };
        }

        if (action === "pickMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var year = this.getAttribute("data-year");
            var month = this.getAttribute("data-month");
            that._openMonthPicker = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, year + month);
          };
        }

        if (action === "clearMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            that._openMonthPicker = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, "");
          };
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

    setCellValue(rowIndex, fieldName, value) {
      if (rowIndex < 0 || rowIndex >= this._rows.length) {
        return;
      }

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
            COMPANY: this._rows[i].COMPANY || "",
            COSTCENTER: this._rows[i].COSTCENTER || "",
            ASSET_CLASS: this._rows[i].ASSET_CLASS || "",
            CAPITALIZED: this._rows[i].CAPITALIZED || "",
            ID: "",
            DESCRIPTION: this._rows[i].DESCRIPTION || ""
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

      for (var k = this._rows.length - copiedRows.length; k < this._rows.length; k++) {
        if (this._rows[k].COMPANY && this._rows[k].CAPITALIZED) {
          this._lastEvent = JSON.stringify({
            type: "fieldChange",
            rowIndex: k,
            field: "COMPANY",
            value: this._rows[k].COMPANY
          });
          this._syncRows();
          this._fireSimpleEvent("onFieldChange", {
            rowIndex: k,
            field: "COMPANY",
            value: this._rows[k].COMPANY
          });
        }
      }
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
      this._openDropdown = null;
      this._openMonthPicker = null;
      this._dropdownSearchState = {};
      this._floatingPanelState = null;
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
        this._validationErrors = [{ rowIndex: 1, field: "selected", message: "Please select at least one row to save" }];
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
})();
this is the current widget code and var evt = AssetEntryWidget_1.getLastEvent();

// Ignore non-field change events
if (evt.indexOf("\"type\":\"fieldChange\"") === -1) {
	return;
}

//================= PARSE EVENT STRING =================//

var rowIndex = -1;
var fieldName = "";
var fieldValue = "";

var rowKey = "\"rowIndex\":";
var fieldKey = "\"field\":\"";
var valueKey = "\"value\":\"";

var rowPos = evt.indexOf(rowKey);
if (rowPos > -1) {
	var rowStart = rowPos + rowKey.length;
	var rowEnd = evt.indexOf(",", rowStart);

	if (rowEnd === -1) {
		rowEnd = evt.indexOf("}", rowStart);
	}

	var rowText = evt.substring(rowStart, rowEnd);
	rowIndex = ConvertUtils.stringToInteger(rowText);
}

var fieldPos = evt.indexOf(fieldKey);
if (fieldPos > -1) {
	var fieldStart = fieldPos + fieldKey.length;
	var fieldEnd = evt.indexOf("\"", fieldStart);
	fieldName = evt.substring(fieldStart, fieldEnd);
}

var valuePos = evt.indexOf(valueKey);
if (valuePos > -1) {
	var valueStart = valuePos + valueKey.length;
	var valueEnd = evt.indexOf("\"", valueStart);
	fieldValue = evt.substring(valueStart, valueEnd);
}

if (fieldName !== "COMPANY" && fieldName !== "CAPITALIZED") {
	return;
}

if (rowIndex < 0) {
	return;
}

var companyVal = AssetEntryWidget_1.getRowValue(rowIndex, "COMPANY");
var capitalizedVal = AssetEntryWidget_1.getRowValue(rowIndex, "CAPITALIZED");

companyVal = companyVal.trim();
capitalizedVal = capitalizedVal.trim();

if (companyVal === "" || capitalizedVal === "") {
	return;
}

if (capitalizedVal.length !== 6) {
	return;
}

Application.showBusyIndicator("Generating Asset ID...");

//================= LOAD EXISTING ASSET IDS FROM MODEL =================//

var assetMembers = AHCP_CAPEX.getMembers("AHC_ASSET", {limit: 10000});
var existingIds = ArrayUtils.create(Type.string);

var i = 0;
for (i = 0; i < assetMembers.length; i++) {
	var modelId = "";

	if (assetMembers[i].id !== undefined) {
		modelId = assetMembers[i].id;
	}

	modelId = modelId.trim();

	if (modelId !== "") {
		existingIds.push(modelId);
	}
}

//================= INCLUDE IDS ALREADY PRESENT IN WIDGET =================//

var totalRows = AssetEntryWidget_1.getRowCount();

for (i = 0; i < totalRows; i++) {
	if (i !== rowIndex) {
		var rowIdVal = AssetEntryWidget_1.getRowValue(i, "ID");
		rowIdVal = rowIdVal.trim();

		if (rowIdVal !== "") {
			existingIds.push(rowIdVal);
		}
	}
}

//================= GENERATE NEXT ID =================//

var yearPart = capitalizedVal.substr(0, 4);
var prefix = companyVal + "-" + yearPart + "-";
var maxSeq = 0;

for (i = 0; i < existingIds.length; i++) {
	var testId = existingIds[i];

	if (testId.indexOf(prefix) === 0) {
		var suffix = testId.substr(prefix.length);
		var seq = ConvertUtils.stringToInteger(suffix);

		if (seq > maxSeq) {
			maxSeq = seq;
		}
	}
}

maxSeq = maxSeq + 1;

var seqText = "";
if (maxSeq < 10) {
	seqText = "0" + ConvertUtils.numberToString(maxSeq);
} else {
	seqText = ConvertUtils.numberToString(maxSeq);
}

var generatedId = prefix + seqText;

//================= WRITE BACK TO WIDGET =================//

AssetEntryWidget_1.setCellValue(rowIndex, "ID", generatedId);

Application.hideBusyIndicator(); 

/*(function () {
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

      this._openDropdown = null;
      this._dropdownSearchState = {};
      this._openMonthPicker = null;
      this._floatingPanelState = null;

      this._dropdownBatchSize = 100;
      this._dropdownSearchTimer = null;

      this._columns = [
        { key: "selected", label: "Sel", type: "checkbox", width: "70px" },
         { key: "COMPANY", label: "Company", type: "select", width: "190px" },
        { key: "COSTCENTER", label: "Cost Center", type: "select", width: "210px" },
        { key: "ASSET_CLASS", label: "Asset Class", type: "select", width: "210px" },
        { key: "CAPITALIZED", label: "Capitalized", type: "month", width: "210px" },
        { key: "ID", label: "ID", type: "input", width: "200px" },
        { key: "DESCRIPTION", label: "Description", type: "input", width: "240px" }
      ];

      this._render();
      this._bindOutsideClick();
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
      return ["rows", "companyOptions", "costCenterOptions", "assetClassOptions"];
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
        COMPANY: "",
        COSTCENTER: "",
        ASSET_CLASS: "",
        CAPITALIZED: "",
         ID: "",
         DESCRIPTION: ""
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

    _getOptionsForField(fieldName) {
      if (fieldName === "COMPANY") return this._companyOptions || [];
      if (fieldName === "COSTCENTER") return this._costCenterOptions || [];
      if (fieldName === "ASSET_CLASS") return this._assetClassOptions || [];
      return [];
    }

    _normalizeCapitalized(value) {
      var v = this._safeString(value);
      if (/^\d{6}$/.test(v)) return v;
      if (/^\d{4}-\d{2}$/.test(v)) return v.replace("-", "");
      return "";
    }

    _capitalizedToDisplay(value) {
      var v = this._normalizeCapitalized(value);
      if (/^\d{6}$/.test(v)) {
        var year = v.substring(0, 4);
        var month = v.substring(4, 6);
        return year + " - " + this._getMonthShortName(month);
      }
      return "";
    }

    _getMonthShortName(month) {
      var map = {
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec"
      };
      return map[month] || "";
    }

    _render() {
      this._shadowRoot.innerHTML = `
        <style>
          :host { display: block; font-family: "72", Arial, sans-serif; color: #223548; }
          .wrap { border: 1px solid #d9e2ef; border-radius: 12px; background: #ffffff; overflow: visible; position: relative; }
          .titleBar { padding: 14px 16px 8px 16px; border-bottom: 1px solid #e5edf7; background: #ffffff; }
          .titleText { font-size: 24px; font-weight: 700; color: #0a6ed1; line-height: 1.2; }
          .toolbar { display: flex; justify-content: flex-end; gap: 8px; padding: 12px; border-bottom: 1px solid #e5edf7; background: #f8fbff; flex-wrap: wrap; }
          .btn { border: 1px solid #c7d7ea; background: #ffffff; color: #0a6ed1; border-radius: 8px; padding: 8px 14px; cursor: pointer; font-weight: 600; font-size: 13px; }
          .btn:hover { background: #f3f8fd; }
          .btn.primary { background: #0a6ed1; color: #ffffff; border-color: #0a6ed1; }
          .btn.danger { color: #bb1e1e; border-color: #efb4b4; background: #fff7f7; }
          .gridWrap { overflow-x: auto; overflow-y: visible; max-height: 520px; background: #ffffff; position: relative; }
          table { border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; }
          thead, tbody, tr { overflow: visible; }
          th, td { border-bottom: 1px solid #edf2f7; padding: 8px; vertical-align: top; white-space: nowrap; overflow: visible; position: relative; }
          th { position: sticky; top: 0; background: #eef4fb; z-index: 1; text-align: left; font-size: 12px; color: #223548; font-weight: 700; }
          tr:hover td { background: #fafcff; }
          tr.errorRow td { background: #fff7f7; }
          tr.modifiedRow td { background: #fffbeb; }
          .cell { width: 100%; min-height: 34px; height: 34px; border: 1px solid #c9d6e5; border-radius: 6px; padding: 6px 10px; font-size: 13px; background: #fff; color: #223548; box-sizing: border-box; outline: none; }
          .cell.error { border-color: #e25555; background: #fff5f5; }
          .rowErr { margin-top: 4px; font-size: 11px; color: #c53030; white-space: normal; max-width: 240px; line-height: 1.3; }
          .summary { padding: 10px 12px; border-top: 1px solid #e5edf7; display: flex; gap: 18px; font-size: 12px; background: #fafcff; flex-wrap: wrap; }
          .row-checkbox { width: 22px; height: 22px; cursor: pointer; margin-top: 8px; }
          .select-all-wrap { display: flex; align-items: center; gap: 6px; }
          .select-all-checkbox { width: 16px; height: 16px; cursor: pointer; }

          .dropdownWrap, .monthWrap { position: relative; width: 100%; overflow: visible; }
          .dropdownDisplay, .monthDisplay {
            width: 100%;
            min-height: 34px;
            height: 34px;
            border: 1px solid #c9d6e5;
            border-radius: 6px;
            background: #fff;
            box-sizing: border-box;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            padding: 0 8px 0 10px;
            cursor: pointer;
            font-size: 13px;
          }
          .dropdownDisplay.error, .monthDisplay.error { border-color: #e25555; background: #fff5f5; }
          .dropdownValue, .monthValue { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; color: #223548; }
          .dropdownPlaceholder, .monthPlaceholder { color: #6b7c93; }
          .dropdownIcons { display: flex; align-items: center; gap: 6px; flex-shrink: 0; }
          .dropdownArrow, .closeIcon {
            border: none;
            background: transparent;
            cursor: pointer;
            color: #5f6b7a;
            font-size: 14px;
            line-height: 1;
            padding: 0;
            width: 18px;
            height: 18px;
          }

          .dropdownHeader, .monthHeader { display: flex; align-items: center; gap: 8px; padding: 8px; border-bottom: 1px solid #edf2f7; background: #f8fbff; }
          .dropdownSearch { flex: 1; height: 32px; border: 1px solid #c9d6e5; border-radius: 6px; padding: 6px 10px; font-size: 13px; box-sizing: border-box; outline: none; }
          .dropdownList { max-height: 220px; overflow: auto; background: #fff; }
          .dropdownItem { padding: 8px 10px; cursor: pointer; font-size: 13px; border-bottom: 1px solid #f4f7fb; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
          .dropdownItem:hover { background: #f3f8fd; }
          .dropdownItem.selected { background: #eaf3fc; color: #0a6ed1; font-weight: 600; }
          .dropdownEmpty { padding: 10px; font-size: 12px; color: #6b7c93; }
          .dropdownMore { padding: 8px 10px; font-size: 12px; color: #6b7c93; text-align: center; border-top: 1px solid #edf2f7; background: #fafcff; }

          .dropdownFooter, .monthFooter { display: flex; justify-content: space-between; align-items: center; padding: 8px; border-top: 1px solid #edf2f7; background: #f8fbff; }
          .linkBtn { border: none; background: transparent; color: #0a6ed1; font-size: 12px; cursor: pointer; padding: 0; }
          .yearSelect { flex: 1; height: 32px; border: 1px solid #c9d6e5; border-radius: 6px; padding: 4px 8px; font-size: 13px; outline: none; background: #fff; }
          .monthGrid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; padding: 10px; }
          .monthBtn { border: 1px solid #c9d6e5; background: #fff; border-radius: 6px; padding: 8px 0; cursor: pointer; font-size: 13px; color: #223548; }
          .monthBtn:hover { background: #f3f8fd; }
          .monthBtn.selected { background: #5d6673; border-color: #5d6673; color: #fff; }

          .floatingLayer { position: absolute; inset: 0; overflow: visible; pointer-events: none; z-index: 999998; }
          .floatingPanel { position: absolute; background: #fff; border: 1px solid #c9d6e5; border-radius: 8px; box-shadow: 0 8px 24px rgba(34,53,72,0.16); z-index: 999999; overflow: hidden; pointer-events: auto; }
          .floatingDropdownPanel { min-width: 220px; }
          .floatingMonthPanel { width: 250px; }
        </style>
        <div class="wrap">
          <div class="titleBar">
            <div class="titleText">New Asset Creation</div>
          </div>
          <div id="widgetWrap"></div>
          <div class="floatingLayer" id="floatingLayer"></div>
        </div>
      `;
    }

    _bindOutsideClick() {
      var that = this;

      this._shadowRoot.addEventListener("click", function (e) {
        e.stopPropagation();
      });

      document.addEventListener("click", function () {
        if (that._openDropdown !== null || that._openMonthPicker !== null) {
          that._openDropdown = null;
          that._openMonthPicker = null;
          that._floatingPanelState = null;
          that._renderFloatingPanel();
        }
      });
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
      var html = "";

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
          html += '<th style="width:' + col.width + '"><div class="select-all-wrap"><span>Sel</span><input class="select-all-checkbox" type="checkbox" id="selectAll" ' + (allSelected ? "checked" : "") + " /></div></th>";
        } else {
          html += '<th style="width:' + col.width + '">' + col.label + "</th>";
        }
      }

      html += "</tr></thead><tbody>";

      for (var i = 0; i < this._rows.length; i++) {
        var row = this._rows[i];
        var rowErrors = rowErrorMap[i] || [];
        var rowClass = "";

        if (rowErrors.length) rowClass = "errorRow";
        else if (row.isModified === true) rowClass = "modifiedRow";

        html += '<tr class="' + rowClass + '">';
        for (var c = 0; c < this._columns.length; c++) {
          html += '<td style="width:' + this._columns[c].width + '">' + this._renderCell(row, i, this._columns[c], rowErrors) + "</td>";
        }
        html += "</tr>";
      }

      html += "</tbody></table></div>";
      html += this._refreshStatusBarHtml();

      container.innerHTML = html;
      this._renderFloatingPanel();
      this._bindEvents();
    }

    _renderFloatingPanel() {
      var layer = this._shadowRoot.getElementById("floatingLayer");
      if (!layer) return;

      if (!this._floatingPanelState) {
        layer.innerHTML = "";
        return;
      }

      var p = this._floatingPanelState;
      var html = "";

      if (p.type === "select") {
        html += '<div class="floatingPanel floatingDropdownPanel" style="top:' + p.top + "px; left:" + p.left + "px; width:" + p.width + 'px;">';
        html += '<div class="dropdownHeader">';
        html += '<input class="dropdownSearch" type="text" placeholder="Search..." data-action="searchDropdown" data-dropdown-id="' + p.dropdownId + '" value="' + this._escape(p.searchValue || "") + '" />';
        html += "</div>";
        html += '<div class="dropdownList" data-dropdown-list="' + p.dropdownId + '"></div>';
        html += '<div class="dropdownFooter">';
        html += '<button type="button" class="linkBtn" data-action="clearSelect" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '">Clear selection</button>';
        html += '<button type="button" class="linkBtn" data-action="closeDropdown" data-dropdown-id="' + p.dropdownId + '">Close</button>';
        html += "</div>";
        html += "</div>";
      }

      if (p.type === "month") {
        html += '<div class="floatingPanel floatingMonthPanel" style="top:' + p.top + "px; left:" + p.left + 'px;">';
        html += '<div class="monthHeader">';
        html += '<select class="yearSelect" data-action="changeMonthYear" data-picker-id="' + p.pickerId + '" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '">';
        for (var year = 2026; year <= 3200; year++) {
          var yearStr = String(year);
          var selected = yearStr === p.selectedYear ? "selected" : "";
          html += '<option value="' + yearStr + '" ' + selected + ">" + yearStr + "</option>";
        }
        html += "</select>";
        html += '<button type="button" class="closeIcon" data-action="closeMonth" data-picker-id="' + p.pickerId + '" title="Close">×</button>';
        html += "</div>";

        var monthNames = [
          { key: "01", text: "Jan" }, { key: "02", text: "Feb" }, { key: "03", text: "Mar" }, { key: "04", text: "Apr" },
          { key: "05", text: "May" }, { key: "06", text: "Jun" }, { key: "07", text: "Jul" }, { key: "08", text: "Aug" },
          { key: "09", text: "Sep" }, { key: "10", text: "Oct" }, { key: "11", text: "Nov" }, { key: "12", text: "Dec" }
        ];

        html += '<div class="monthGrid">';
        for (var m = 0; m < monthNames.length; m++) {
          var selectedMonthCss = monthNames[m].key === p.selectedMonth ? "selected" : "";
          html += '<button type="button" class="monthBtn ' + selectedMonthCss + '" data-action="pickMonth" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '" data-picker-id="' + p.pickerId + '" data-year="' + p.selectedYear + '" data-month="' + monthNames[m].key + '">' + monthNames[m].text + "</button>";
        }
        html += "</div>";

        html += '<div class="monthFooter">';
        html += '<button type="button" class="linkBtn" data-action="clearMonth" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '">Clear selection</button>';
        html += '<button type="button" class="linkBtn" data-action="closeMonth" data-picker-id="' + p.pickerId + '">Close</button>';
        html += "</div>";
        html += "</div>";
      }

      layer.innerHTML = html;

      if (p.type === "select") {
        this._renderDropdownBatch(true);
      }
    }

    _renderDropdownBatch(reset) {
      if (!this._floatingPanelState || this._floatingPanelState.type !== "select") return;

      var p = this._floatingPanelState;
      var listEl = this._shadowRoot.querySelector('[data-dropdown-list="' + p.dropdownId + '"]');
      if (!listEl) return;

      if (reset) {
        p.renderedCount = 0;
        listEl.innerHTML = "";
      }

      var start = p.renderedCount || 0;
      var end = Math.min(start + this._dropdownBatchSize, p.filteredOptions.length);

      if (start >= p.filteredOptions.length) {
        if (p.filteredOptions.length === 0 && reset) {
          listEl.innerHTML = '<div class="dropdownEmpty">No results found</div>';
        }
        return;
      }

      var html = "";
      for (var i = start; i < end; i++) {
        var opt = p.filteredOptions[i];
        var selectedCss = String(opt.key) === String(p.value) ? "selected" : "";
        html += '<div class="dropdownItem ' + selectedCss + '" data-action="chooseOption" data-row="' + p.rowIndex + '" data-field="' + p.fieldName + '" data-value="' + this._escape(opt.key) + '">' + this._escape(opt.text) + "</div>";
      }

      var temp = document.createElement("div");
      temp.innerHTML = html;
      while (temp.firstChild) {
        listEl.appendChild(temp.firstChild);
      }

      p.renderedCount = end;

      var oldMore = listEl.querySelector(".dropdownMore");
      if (oldMore) {
        oldMore.parentNode.removeChild(oldMore);
      }

      if (p.renderedCount < p.filteredOptions.length) {
        var more = document.createElement("div");
        more.className = "dropdownMore";
        more.textContent = "Showing " + p.renderedCount + " of " + p.filteredOptions.length;
        listEl.appendChild(more);
      }
    }

    _filterDropdownOptions(fieldName, searchValue) {
      var options = this._getOptionsForField(fieldName);
      var q = this._safeString(searchValue).toLowerCase();
      if (q === "") return options.slice(0);

      var filtered = [];
      for (var i = 0; i < options.length; i++) {
        if ((options[i].searchText || "").indexOf(q) !== -1) {
          filtered.push(options[i]);
        }
      }
      return filtered;
    }

    _getSelectedText(fieldName, value) {
      var options = this._getOptionsForField(fieldName);
      for (var i = 0; i < options.length; i++) {
        if (String(options[i].key) === String(value)) {
          return options[i].text;
        }
      }
      return "";
    }

    _renderCell(row, rowIndex, column, rowErrors) {
      var value = row[column.key] !== undefined && row[column.key] !== null ? row[column.key] : "";
      var hasError = this._hasFieldError(column.key, rowErrors);
      var errorCss = hasError ? "error" : "";

      if (column.type === "checkbox") {
        return '<input class="row-checkbox" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="checkbox" type="checkbox" ' + (value === true ? "checked" : "") + " />";
      }

      if (column.type === "input") {
        return '<input class="cell ' + errorCss + '" data-row="' + rowIndex + '" data-field="' + column.key + '" data-type="input" type="text" value="' + this._escape(String(value)) + '" />' + this._renderFieldErrors(column.key, rowErrors);
      }

      if (column.type === "month") {
        return this._renderMonthCell(row, rowIndex, column, rowErrors, errorCss);
      }

      if (column.type === "select") {
        return this._renderSelectCell(row, rowIndex, column, rowErrors, errorCss);
      }

      return "";
    }

    _renderSelectCell(row, rowIndex, column, rowErrors, errorCss) {
      var value = this._safeString(row[column.key]);
      var displayText = this._getSelectedText(column.key, value);
      var placeholder = "Select";
      var dropdownId = "dropdown_" + rowIndex + "_" + column.key;
      var html = "";

      html += '<div class="dropdownWrap">';
      html += '<div class="dropdownDisplay ' + errorCss + '" data-action="toggleDropdown" data-dropdown-id="' + dropdownId + '" data-row="' + rowIndex + '" data-field="' + column.key + '">';
      if (displayText) {
        html += '<div class="dropdownValue">' + this._escape(displayText) + "</div>";
      } else {
        html += '<div class="dropdownValue dropdownPlaceholder">' + this._escape(placeholder) + "</div>";
      }
      html += '<div class="dropdownIcons">';
      html += '<button type="button" class="dropdownArrow" data-action="toggleDropdown" data-dropdown-id="' + dropdownId + '" data-row="' + rowIndex + '" data-field="' + column.key + '" title="Open">▼</button>';
      html += "</div>";
      html += "</div>";
      html += "</div>";
      html += this._renderFieldErrors(column.key, rowErrors);

      return html;
    }

    _renderMonthCell(row, rowIndex, column, rowErrors, errorCss) {
      var value = this._safeString(row[column.key]);
      var displayValue = this._capitalizedToDisplay(value);
      var pickerId = "month_" + rowIndex + "_" + column.key;
      var html = "";

      html += '<div class="monthWrap">';
      html += '<div class="monthDisplay ' + errorCss + '" data-action="toggleMonth" data-picker-id="' + pickerId + '" data-row="' + rowIndex + '" data-field="' + column.key + '">';
      if (displayValue) {
        html += '<div class="monthValue">' + this._escape(displayValue) + "</div>";
      } else {
        html += '<div class="monthValue monthPlaceholder">Select month</div>';
      }
      html += '<div class="dropdownIcons">';
      html += '<button type="button" class="dropdownArrow" data-action="toggleMonth" data-picker-id="' + pickerId + '" data-row="' + rowIndex + '" data-field="' + column.key + '" title="Open">📅</button>';
      html += "</div>";
      html += "</div>";
      html += "</div>";
      html += this._renderFieldErrors(column.key, rowErrors);

      return html;
    }

    _openSelectPanel(rowIndex, fieldName, dropdownId, anchorEl) {
      var searchValue = this._dropdownSearchState[dropdownId] || "";
      var filteredOptions = this._filterDropdownOptions(fieldName, searchValue);
      var wrapRect = this._shadowRoot.querySelector(".wrap").getBoundingClientRect();
      var rect = anchorEl.getBoundingClientRect();

      this._floatingPanelState = {
        type: "select",
        dropdownId: dropdownId,
        rowIndex: rowIndex,
        fieldName: fieldName,
        value: this._safeString(this._rows[rowIndex][fieldName]),
        searchValue: searchValue,
        filteredOptions: filteredOptions,
        renderedCount: 0,
        top: rect.bottom - wrapRect.top + 4,
        left: rect.left - wrapRect.left,
        width: Math.max(rect.width, 220)
      };

      this._openMonthPicker = null;
      this._openDropdown = dropdownId;
      this._renderFloatingPanel();
      this._bindEvents();

      var listEl = this._shadowRoot.querySelector('[data-dropdown-list="' + dropdownId + '"]');
      if (listEl) {
        var that = this;
        listEl.addEventListener("scroll", function () {
          if (!that._floatingPanelState || that._floatingPanelState.dropdownId !== dropdownId) return;
          var nearBottom = listEl.scrollTop + listEl.clientHeight >= listEl.scrollHeight - 20;
          if (nearBottom) {
            that._renderDropdownBatch(false);
            that._bindEvents();
          }
        });
      }
    }

    _openMonthPanel(rowIndex, fieldName, pickerId, anchorEl) {
      var value = this._safeString(this._rows[rowIndex][fieldName]);
      var selectedYear = "2026";
      var selectedMonth = "";

      if (/^\d{6}$/.test(value)) {
        selectedYear = value.substring(0, 4);
        selectedMonth = value.substring(4, 6);
      }

      var wrapRect = this._shadowRoot.querySelector(".wrap").getBoundingClientRect();
      var rect = anchorEl.getBoundingClientRect();

      this._floatingPanelState = {
        type: "month",
        pickerId: pickerId,
        rowIndex: rowIndex,
        fieldName: fieldName,
        selectedYear: selectedYear,
        selectedMonth: selectedMonth,
        top: rect.bottom - wrapRect.top + 4,
        left: rect.left - wrapRect.left
      };

      this._openDropdown = null;
      this._openMonthPicker = pickerId;
      this._renderFloatingPanel();
      this._bindEvents();
    }

    _renderFieldErrors(fieldName, rowErrors) {
      var messages = [];
      for (var i = 0; i < rowErrors.length; i++) {
        var err = rowErrors[i];
        if (err.field === fieldName) messages.push(err.message);
      }
      if (!messages.length) return "";
      return '<div class="rowErr">' + messages.join("<br>") + "</div>";
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
      if (btnAdd) btnAdd.onclick = function () { that.addRow(); };

      var btnCopy = this._shadowRoot.getElementById("btnCopy");
      if (btnCopy) btnCopy.onclick = function () { that.copySelectedRows(); };

      var btnDelete = this._shadowRoot.getElementById("btnDelete");
      if (btnDelete) btnDelete.onclick = function () { that.deleteSelectedRows(); };

      var btnValidate = this._shadowRoot.getElementById("btnValidate");
      if (btnValidate) btnValidate.onclick = function () { that.validate(); };

      var btnSave = this._shadowRoot.getElementById("btnSave");
      if (btnSave) btnSave.onclick = function () { that.save(); };

      var btnClear = this._shadowRoot.getElementById("btnClear");
      if (btnClear) btnClear.onclick = function () { that.clear(); };

      var selectAll = this._shadowRoot.getElementById("selectAll");
      if (selectAll) {
        selectAll.onchange = function () {
          that._toggleSelectAll(selectAll.checked);
        };
      }

      var allElements = this._shadowRoot.querySelectorAll("[data-row][data-field]");
      Array.prototype.forEach.call(allElements, function (el) {
        var type = el.getAttribute("data-type");

        if (type === "checkbox") {
          el.onchange = function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.checked;
            that._updateField(rowIndex, fieldName, value);
          };
          return;
        }

        if (type === "input") {
          el.onchange = function () {
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.value;
            that._updateField(rowIndex, fieldName, value);
          };
          return;
        }
      });

      var actionElements = this._shadowRoot.querySelectorAll("[data-action]");
      Array.prototype.forEach.call(actionElements, function (el) {
        var action = el.getAttribute("data-action");

        if (action === "toggleDropdown") {
          el.onclick = function (e) {
            e.stopPropagation();
            var dropdownId = this.getAttribute("data-dropdown-id");
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");

            if (that._openDropdown === dropdownId) {
              that._openDropdown = null;
              that._floatingPanelState = null;
              that._renderFloatingPanel();
            } else {
              that._openSelectPanel(rowIndex, fieldName, dropdownId, this.closest(".dropdownDisplay"));
              setTimeout(function () {
                var input = that._shadowRoot.querySelector('.dropdownSearch[data-dropdown-id="' + dropdownId + '"]');
                if (input) input.focus();
              }, 0);
            }
          };
        }

        if (action === "closeDropdown") {
          el.onclick = function (e) {
            e.stopPropagation();
            that._openDropdown = null;
            that._floatingPanelState = null;
            that._renderFloatingPanel();
          };
        }

        if (action === "searchDropdown") {
          el.onclick = function (e) {
            e.stopPropagation();
          };

          el.oninput = function (e) {
            e.stopPropagation();

            var inputEl = this;
            var dropdownId = inputEl.getAttribute("data-dropdown-id");

            if (that._dropdownSearchTimer) {
              clearTimeout(that._dropdownSearchTimer);
            }

            that._dropdownSearchTimer = setTimeout(function () {
              that._dropdownSearchState[dropdownId] = inputEl.value;

              if (that._floatingPanelState && that._floatingPanelState.type === "select") {
                var fieldName = that._floatingPanelState.fieldName;
                that._floatingPanelState.searchValue = inputEl.value;
                that._floatingPanelState.filteredOptions = that._filterDropdownOptions(fieldName, inputEl.value);
                that._floatingPanelState.renderedCount = 0;
                that._renderFloatingPanel();
                that._bindEvents();

                setTimeout(function () {
                  var input = that._shadowRoot.querySelector('.dropdownSearch[data-dropdown-id="' + dropdownId + '"]');
                  if (input) {
                    input.focus();
                    var len = input.value.length;
                    input.setSelectionRange(len, len);
                  }
                }, 0);
              }
            }, 120);
          };
        }

        if (action === "chooseOption") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var value = this.getAttribute("data-value");
            that._openDropdown = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, value);
          };
        }

        if (action === "clearSelect") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            that._openDropdown = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, "");
          };
        }

        if (action === "toggleMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            var pickerId = this.getAttribute("data-picker-id");
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");

            if (that._openMonthPicker === pickerId) {
              that._openMonthPicker = null;
              that._floatingPanelState = null;
              that._renderFloatingPanel();
            } else {
              that._openMonthPanel(rowIndex, fieldName, pickerId, this.closest(".monthDisplay"));
            }
          };
        }

        if (action === "closeMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            that._openMonthPicker = null;
            that._floatingPanelState = null;
            that._renderFloatingPanel();
          };
        }

        if (action === "changeMonthYear") {
          el.onclick = function (e) {
            e.stopPropagation();
          };

          el.onchange = function (e) {
            e.stopPropagation();
            var pickerId = this.getAttribute("data-picker-id");
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var currentValue = that._safeString(that._rows[rowIndex][fieldName]);
            var month = "01";

            if (/^\d{6}$/.test(currentValue)) {
              month = currentValue.substring(4, 6);
            }

            that._rows[rowIndex][fieldName] = this.value + month;
            that._rows[rowIndex].isModified = true;
            that._rows[rowIndex].rowStatus = "CHANGED";
            that._widgetStatus = "CHANGED";
            that._validationErrors = [];
            that._validationResult = "true";
            that._syncRows();

            var anchor = that._shadowRoot.querySelector('.monthDisplay[data-picker-id="' + pickerId + '"]');
            if (anchor) {
              that._openMonthPanel(rowIndex, fieldName, pickerId, anchor);
            } else {
              that._refreshTable();
            }
          };
        }

        if (action === "pickMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            var year = this.getAttribute("data-year");
            var month = this.getAttribute("data-month");
            that._openMonthPicker = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, year + month);
          };
        }

        if (action === "clearMonth") {
          el.onclick = function (e) {
            e.stopPropagation();
            var rowIndex = parseInt(this.getAttribute("data-row"), 10);
            var fieldName = this.getAttribute("data-field");
            that._openMonthPicker = null;
            that._floatingPanelState = null;
            that._updateField(rowIndex, fieldName, "");
          };
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
           
            COMPANY: this._rows[i].COMPANY || "",
            COSTCENTER: this._rows[i].COSTCENTER || "",
            ASSET_CLASS: this._rows[i].ASSET_CLASS || "",
            CAPITALIZED: this._rows[i].CAPITALIZED || "",
             ID: this._rows[i].ID || "",
            DESCRIPTION: this._rows[i].DESCRIPTION || ""
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
      this._openDropdown = null;
      this._openMonthPicker = null;
      this._dropdownSearchState = {};
      this._floatingPanelState = null;
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
        this._validationErrors = [{ rowIndex: 1, field: "selected", message: "Please select at least one row to save" }];
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
})(); */
