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

      this._dropdownPanel = null;
      this._dropdownSearch = null;
      this._dropdownList = null;
      this._dropdownCloseBtn = null;
      this._dropdownClearBtn = null;
      this._dropdownOpen = false;
      this._activeDropdownTrigger = null;
      this._activeDropdownRow = -1;
      this._activeDropdownField = "";
      this._activeDropdownOptions = [];
      this._activeDropdownSelectedKey = "";
      this._filteredDropdownOptions = [];
      this._dropdownSearchTimer = null;

      this._monthPanel = null;
      this._monthYearSelect = null;
      this._monthGrid = null;
      this._monthCloseBtn = null;
      this._monthClearBtn = null;
      this._monthOpen = false;
      this._activeMonthTrigger = null;
      this._activeMonthRow = -1;
      this._activeMonthField = "";
      this._activeMonthYear = "2026";
      this._activeMonthSelected = "";

      this._columns = [
        { key: "selected", label: "Sel", type: "checkbox", width: "70px" },
        { key: "ID", label: "ID", type: "input", width: "200px" },
        { key: "DESCRIPTION", label: "Description", type: "input", width: "240px" },
        { key: "COMPANY", label: "Company", type: "select", width: "190px" },
        { key: "COSTCENTER", label: "Cost Center", type: "select", width: "210px" },
        { key: "ASSET_CLASS", label: "Asset Class", type: "select", width: "210px" },
        { key: "CAPITALIZED", label: "Capitalized", type: "month", width: "210px" }
      ];

      this._render();
      this._createDropdownPanel();
      this._createMonthPanel();
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

    disconnectedCallback() {
      this._closeDropdown();
      this._closeMonthPanel();

      if (this._dropdownPanel && this._dropdownPanel.parentNode) {
        this._dropdownPanel.parentNode.removeChild(this._dropdownPanel);
      }

      if (this._monthPanel && this._monthPanel.parentNode) {
        this._monthPanel.parentNode.removeChild(this._monthPanel);
      }

      if (this._documentClickHandler) {
        document.removeEventListener("click", this._documentClickHandler);
      }
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
        return;
      }

      if (name === "costCenterOptions") {
        this._costCenterOptions = this._parseOptions(newValue);
        return;
      }

      if (name === "assetClassOptions") {
        this._assetClassOptions = this._parseOptions(newValue);
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
          :host { display:block; font-family:"72", Arial, sans-serif; color:#223548; }
          .wrap { border:1px solid #d9e2ef; border-radius:12px; background:#ffffff; overflow:visible; position:relative; }
          .titleBar { padding:14px 16px 8px 16px; border-bottom:1px solid #e5edf7; background:#ffffff; }
          .titleText { font-size:24px; font-weight:700; color:#0a6ed1; line-height:1.2; }
          .toolbar { display:flex; justify-content:flex-end; gap:8px; padding:12px; border-bottom:1px solid #e5edf7; background:#f8fbff; flex-wrap:wrap; }
          .btn { border:1px solid #c7d7ea; background:#ffffff; color:#0a6ed1; border-radius:8px; padding:8px 14px; cursor:pointer; font-weight:600; font-size:13px; }
          .btn:hover { background:#f3f8fd; }
          .btn.primary { background:#0a6ed1; color:#ffffff; border-color:#0a6ed1; }
          .btn.danger { color:#bb1e1e; border-color:#efb4b4; background:#fff7f7; }
          .gridWrap { overflow-x:auto; overflow-y:visible; max-height:520px; background:#ffffff; position:relative; }
          table { border-collapse:separate; border-spacing:0; width:max-content; min-width:100%; }
          thead, tbody, tr { overflow:visible; }
          th, td { border-bottom:1px solid #edf2f7; padding:8px; vertical-align:top; white-space:nowrap; overflow:visible; position:relative; }
          th { position:sticky; top:0; background:#eef4fb; z-index:1; text-align:left; font-size:12px; color:#223548; font-weight:700; }
          tr:hover td { background:#fafcff; }
          tr.errorRow td { background:#fff7f7; }
          tr.modifiedRow td { background:#fffbeb; }
          .cell { width:100%; min-height:34px; height:34px; border:1px solid #c9d6e5; border-radius:6px; padding:6px 10px; font-size:13px; background:#fff; color:#223548; box-sizing:border-box; outline:none; }
          .cell.error { border-color:#e25555; background:#fff5f5; }
          .rowErr { margin-top:4px; font-size:11px; color:#c53030; white-space:normal; max-width:240px; line-height:1.3; }
          .summary { padding:10px 12px; border-top:1px solid #e5edf7; display:flex; gap:18px; font-size:12px; background:#fafcff; flex-wrap:wrap; }
          .row-checkbox { width:22px; height:22px; cursor:pointer; margin-top:8px; }
          .select-all-wrap { display:flex; align-items:center; gap:6px; }
          .select-all-checkbox { width:16px; height:16px; cursor:pointer; }
          .dropdownWrap,.monthWrap { position:relative; width:100%; overflow:visible; }
          .dropdownDisplay,.monthDisplay {
            width:100%; min-height:34px; height:34px; border:1px solid #c9d6e5; border-radius:6px; background:#fff;
            box-sizing:border-box; display:flex; align-items:center; justify-content:space-between; gap:8px; padding:0 8px 0 10px; cursor:pointer; font-size:13px;
          }
          .dropdownDisplay.error,.monthDisplay.error { border-color:#e25555; background:#fff5f5; }
          .dropdownValue,.monthValue { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; flex:1; color:#223548; }
          .dropdownPlaceholder,.monthPlaceholder { color:#6b7c93; }
          .dropdownIcons { display:flex; align-items:center; gap:6px; flex-shrink:0; }
          .dropdownArrow { color:#5f6b7a; font-size:14px; line-height:1; width:18px; height:18px; display:flex; align-items:center; justify-content:center; }
        </style>
        <div class="wrap">
          <div class="titleBar">
            <div class="titleText">New Asset Creation</div>
          </div>
          <div id="widgetWrap"></div>
        </div>
      `;
    }

    _createDropdownPanel() {
      if (this._dropdownPanel) return;

      var dropdownPanel = document.createElement("div");
      dropdownPanel.className = "asset-widget-dropdown-panel";
      dropdownPanel.style.display = "none";
      dropdownPanel.innerHTML =
        '<div class="dropdown-topbar">' +
          '<div class="dropdown-search-wrap">' +
            '<input type="text" class="dropdown-search-input" placeholder="Search..." />' +
          '</div>' +
          '<button type="button" class="dropdown-close-btn" title="Close">×</button>' +
        '</div>' +
        '<div class="dropdown-actions">' +
          '<button type="button" class="dropdown-clear-btn">Clear selection</button>' +
        '</div>' +
        '<div class="dropdown-list"></div>';

      document.body.appendChild(dropdownPanel);

      this._dropdownPanel = dropdownPanel;
      this._dropdownSearch = dropdownPanel.querySelector(".dropdown-search-input");
      this._dropdownList = dropdownPanel.querySelector(".dropdown-list");
      this._dropdownCloseBtn = dropdownPanel.querySelector(".dropdown-close-btn");
      this._dropdownClearBtn = dropdownPanel.querySelector(".dropdown-clear-btn");

      var that = this;

      this._dropdownCloseBtn.onclick = function (e) {
        e.preventDefault();
        e.stopPropagation();
        that._closeDropdown();
      };

      this._dropdownClearBtn.onclick = function (e) {
        e.preventDefault();
        e.stopPropagation();

        var rowIndex = that._activeDropdownRow;
        var fieldName = that._activeDropdownField;
        if (rowIndex < 0 || !that._rows[rowIndex]) {
          that._closeDropdown();
          return;
        }

        that._rows[rowIndex][fieldName] = "";
        that._rows[rowIndex].isModified = true;
        that._rows[rowIndex].rowStatus = "CHANGED";
        that._validationErrors = [];
        that._validationResult = "true";
        that._widgetStatus = "CHANGED";
        that._lastEvent = JSON.stringify({
          type: "fieldChange",
          rowIndex: rowIndex,
          field: fieldName,
          value: ""
        });

        that._syncRows();
        that._refreshTable();
        that._fireSimpleEvent("onFieldChange", { rowIndex: rowIndex, field: fieldName, value: "" });
        that._fireSimpleEvent("onDataChange", { rows: that._rows });
        that._closeDropdown();
      };

      this._dropdownSearch.oninput = function () {
        if (that._dropdownSearchTimer) {
          clearTimeout(that._dropdownSearchTimer);
        }
        that._dropdownSearchTimer = setTimeout(function () {
          that._applyDropdownSearch(that._dropdownSearch.value);
        }, 100);
      };

      this._documentClickHandler = function (e) {
        if (!that._dropdownOpen && !that._monthOpen) return;

        var insideDropdownPanel = that._dropdownPanel && that._dropdownPanel.contains(e.target);
        var insideDropdownTrigger = that._activeDropdownTrigger && that._activeDropdownTrigger.contains(e.target);
        var insideMonthPanel = that._monthPanel && that._monthPanel.contains(e.target);
        var insideMonthTrigger = that._activeMonthTrigger && that._activeMonthTrigger.contains(e.target);

        if (!insideDropdownPanel && !insideDropdownTrigger && !insideMonthPanel && !insideMonthTrigger) {
          that._closeDropdown();
          that._closeMonthPanel();
        }
      };

      document.addEventListener("click", this._documentClickHandler);
    }

    _createMonthPanel() {
      if (this._monthPanel) return;

      var monthPanel = document.createElement("div");
      monthPanel.className = "asset-widget-month-panel";
      monthPanel.style.display = "none";
      monthPanel.innerHTML =
        '<div class="month-topbar">' +
          '<select class="month-year-select"></select>' +
          '<button type="button" class="month-close-btn" title="Close">×</button>' +
        '</div>' +
        '<div class="month-grid"></div>' +
        '<div class="month-actions">' +
          '<button type="button" class="month-clear-btn">Clear selection</button>' +
        '</div>';

      document.body.appendChild(monthPanel);

      this._monthPanel = monthPanel;
      this._monthYearSelect = monthPanel.querySelector(".month-year-select");
      this._monthGrid = monthPanel.querySelector(".month-grid");
      this._monthCloseBtn = monthPanel.querySelector(".month-close-btn");
      this._monthClearBtn = monthPanel.querySelector(".month-clear-btn");

      var yearHtml = "";
      for (var year = 2026; year <= 3200; year++) {
        yearHtml += '<option value="' + year + '">' + year + '</option>';
      }
      this._monthYearSelect.innerHTML = yearHtml;

      var that = this;

      this._monthCloseBtn.onclick = function (e) {
        e.preventDefault();
        e.stopPropagation();
        that._closeMonthPanel();
      };

      this._monthClearBtn.onclick = function (e) {
        e.preventDefault();
        e.stopPropagation();
        var rowIndex = that._activeMonthRow;
        var fieldName = that._activeMonthField;

        if (rowIndex < 0 || !that._rows[rowIndex]) {
          that._closeMonthPanel();
          return;
        }

        that._rows[rowIndex][fieldName] = "";
        that._rows[rowIndex].isModified = true;
        that._rows[rowIndex].rowStatus = "CHANGED";
        that._validationErrors = [];
        that._validationResult = "true";
        that._widgetStatus = "CHANGED";
        that._lastEvent = JSON.stringify({
          type: "fieldChange",
          rowIndex: rowIndex,
          field: fieldName,
          value: ""
        });

        that._syncRows();
        that._refreshTable();
        that._fireSimpleEvent("onFieldChange", { rowIndex: rowIndex, field: fieldName, value: "" });
        that._fireSimpleEvent("onDataChange", { rows: that._rows });
        that._closeMonthPanel();
      };

      this._monthYearSelect.onchange = function (e) {
        e.stopPropagation();
        that._activeMonthYear = this.value;
        that._renderMonthGrid();
      };
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
        return this._renderMonthCell(row, rowIndex, column, rowErrors, errorCss);
      }

      if (column.type === "select") {
        return this._renderSelectCell(row, rowIndex, column, rowErrors, errorCss);
      }

      return "";
    }

    _renderSelectCell(row, rowIndex, column, rowErrors, errorCss) {
      var value = this._safeString(row[column.key]);
      var displayText = this._getSelectedText(column.key, value) || "Select";
      var dropdownId = "dropdown_" + rowIndex + "_" + column.key;

      return ''
        + '<div class="dropdownWrap">'
        +   '<div class="dropdownDisplay ' + errorCss + '" data-dropdown-id="' + dropdownId + '" data-row="' + rowIndex + '" data-field="' + column.key + '">'
        +     '<div class="dropdownValue' + (displayText === "Select" ? ' dropdownPlaceholder' : '') + '">' + this._escape(displayText) + '</div>'
        +     '<div class="dropdownIcons"><span class="dropdownArrow">▼</span></div>'
        +   '</div>'
        + '</div>'
        + this._renderFieldErrors(column.key, rowErrors);
    }

    _renderMonthCell(row, rowIndex, column, rowErrors, errorCss) {
      var value = this._safeString(row[column.key]);
      var displayValue = this._capitalizedToDisplay(value) || "Select month";
      var pickerId = "month_" + rowIndex + "_" + column.key;

      return ''
        + '<div class="monthWrap">'
        +   '<div class="monthDisplay ' + errorCss + '" data-picker-id="' + pickerId + '" data-row="' + rowIndex + '" data-field="' + column.key + '">'
        +     '<div class="monthValue' + (displayValue === "Select month" ? ' monthPlaceholder' : '') + '">' + this._escape(displayValue) + '</div>'
        +     '<div class="dropdownIcons"><span class="dropdownArrow">📅</span></div>'
        +   '</div>'
        + '</div>'
        + this._renderFieldErrors(column.key, rowErrors);
    }

    _bindOutsideClick() {
      var that = this;
      this._shadowRoot.addEventListener("click", function (e) {
        e.stopPropagation();
      });
    }

    _openDropdownPanel(triggerEl, rowIndex, fieldName) {
      this._createDropdownPanel();

      this._activeDropdownTrigger = triggerEl;
      this._activeDropdownRow = rowIndex;
      this._activeDropdownField = fieldName;
      this._activeDropdownOptions = this._getOptionsForField(fieldName) || [];
      this._activeDropdownSelectedKey = this._rows[rowIndex] ? this._rows[rowIndex][fieldName] : "";

      this._dropdownClearBtn.style.display = this._activeDropdownSelectedKey ? "block" : "none";

      var triggerRect = triggerEl.getBoundingClientRect();
      var dropdownWidth = Math.max(triggerRect.width, 320);
      var dropdownHeight = 360;
      var gap = 4;
      var top = triggerRect.bottom + gap;
      var left = triggerRect.left;

      if (left + dropdownWidth > window.innerWidth - 10) {
        left = window.innerWidth - dropdownWidth - 10;
      }
      if (left < 10) left = 10;

      if (top + dropdownHeight > window.innerHeight - 10) {
        top = Math.max(10, triggerRect.top - dropdownHeight - gap);
      }

      this._dropdownPanel.style.display = "block";
      this._dropdownPanel.style.position = "fixed";
      this._dropdownPanel.style.left = left + "px";
      this._dropdownPanel.style.top = top + "px";
      this._dropdownPanel.style.width = dropdownWidth + "px";
      this._dropdownPanel.style.maxHeight = dropdownHeight + "px";
      this._dropdownPanel.style.zIndex = "999999";

      this._dropdownSearch.value = this._dropdownSearchState[fieldName + "_" + rowIndex] || "";
      this._applyDropdownSearch(this._dropdownSearch.value);
      this._dropdownOpen = true;

      var that = this;
      setTimeout(function () {
        if (that._dropdownSearch) that._dropdownSearch.focus();
      }, 0);
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
    }

    _applyDropdownSearch(searchText) {
      var source = this._activeDropdownOptions || [];
      var searchValue = String(searchText || "").toLowerCase().trim();
      var filtered = [];

      if (searchValue === "") {
        this._filteredDropdownOptions = source.slice(0);
      } else {
        for (var i = 0; i < source.length; i++) {
          if ((source[i].searchText || "").indexOf(searchValue) > -1) {
            filtered.push(source[i]);
          }
        }
        this._filteredDropdownOptions = filtered;
      }

      this._renderDropdownListOnly();
    }

    _renderDropdownListOnly() {
      if (!this._dropdownList) return;

      var html = "";
      if (this._filteredDropdownOptions.length > 0) {
        for (var i = 0; i < this._filteredDropdownOptions.length; i++) {
          var opt = this._filteredDropdownOptions[i];
          var selectedCss = String(opt.key) === String(this._activeDropdownSelectedKey) ? " selected" : "";
          html += '<div class="dropdown-item' + selectedCss + '" data-key="' + this._escape(opt.key) + '">' + this._escape(opt.text) + '</div>';
        }
      } else {
        html = '<div class="dropdown-empty">No results found</div>';
      }

      this._dropdownList.innerHTML = html;

      var that = this;
      var items = this._dropdownList.querySelectorAll(".dropdown-item");
      Array.prototype.forEach.call(items, function (item) {
        item.onmousedown = function (e) {
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
        };
      });
    }

    _openMonthPanelUI(triggerEl, rowIndex, fieldName) {
      this._createMonthPanel();

      this._activeMonthTrigger = triggerEl;
      this._activeMonthRow = rowIndex;
      this._activeMonthField = fieldName;

      var value = this._safeString(this._rows[rowIndex][fieldName]);
      this._activeMonthYear = "2026";
      this._activeMonthSelected = "";

      if (/^\d{6}$/.test(value)) {
        this._activeMonthYear = value.substring(0, 4);
        this._activeMonthSelected = value.substring(4, 6);
      }

      this._monthYearSelect.value = this._activeMonthYear;

      var triggerRect = triggerEl.getBoundingClientRect();
      var panelWidth = 250;
      var panelHeight = 300;
      var gap = 4;
      var top = triggerRect.bottom + gap;
      var left = triggerRect.left;

      if (left + panelWidth > window.innerWidth - 10) {
        left = window.innerWidth - panelWidth - 10;
      }
      if (left < 10) left = 10;

      if (top + panelHeight > window.innerHeight - 10) {
        top = Math.max(10, triggerRect.top - panelHeight - gap);
      }

      this._monthPanel.style.display = "block";
      this._monthPanel.style.position = "fixed";
      this._monthPanel.style.left = left + "px";
      this._monthPanel.style.top = top + "px";
      this._monthPanel.style.width = panelWidth + "px";
      this._monthPanel.style.zIndex = "999999";

      this._renderMonthGrid();
      this._monthOpen = true;
    }

    _renderMonthGrid() {
      if (!this._monthGrid) return;

      var monthNames = [
        { key: "01", text: "Jan" }, { key: "02", text: "Feb" }, { key: "03", text: "Mar" }, { key: "04", text: "Apr" },
        { key: "05", text: "May" }, { key: "06", text: "Jun" }, { key: "07", text: "Jul" }, { key: "08", text: "Aug" },
        { key: "09", text: "Sep" }, { key: "10", text: "Oct" }, { key: "11", text: "Nov" }, { key: "12", text: "Dec" }
      ];

      var html = "";
      for (var i = 0; i < monthNames.length; i++) {
        var selectedCss = monthNames[i].key === this._activeMonthSelected ? " selected" : "";
        html += '<button type="button" class="month-btn' + selectedCss + '" data-month="' + monthNames[i].key + '">' + monthNames[i].text + '</button>';
      }

      this._monthGrid.innerHTML = html;

      var that = this;
      var btns = this._monthGrid.querySelectorAll(".month-btn");
      Array.prototype.forEach.call(btns, function (btn) {
        btn.onclick = function (e) {
          e.stopPropagation();
          var month = this.getAttribute("data-month");
          var rowIndex = that._activeMonthRow;
          var fieldName = that._activeMonthField;
          var value = that._activeMonthYear + month;

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
          that._closeMonthPanel();
        };
      });
    }

    _closeMonthPanel() {
      if (this._monthPanel) {
        this._monthPanel.style.display = "none";
      }
      this._monthOpen = false;
      this._activeMonthTrigger = null;
      this._activeMonthRow = -1;
      this._activeMonthField = "";
      this._activeMonthYear = "2026";
      this._activeMonthSelected = "";
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

      this._bindRowEvents();
    }

    _bindRowEvents() {
      var that = this;

      var rowCheckboxes = this._shadowRoot.querySelectorAll(".row-checkbox");
      Array.prototype.forEach.call(rowCheckboxes, function (el) {
        el.onchange = function () {
          var rowIndex = parseInt(this.getAttribute("data-row"), 10);
          var fieldName = this.getAttribute("data-field");
          that._updateField(rowIndex, fieldName, this.checked);
        };
      });

      var rowInputs = this._shadowRoot.querySelectorAll(".cell[data-type='input']");
      Array.prototype.forEach.call(rowInputs, function (el) {
        el.onchange = function () {
          var rowIndex = parseInt(this.getAttribute("data-row"), 10);
          var fieldName = this.getAttribute("data-field");
          that._updateField(rowIndex, fieldName, this.value);
        };
      });

      var dropdownDisplays = this._shadowRoot.querySelectorAll(".dropdownDisplay");
      Array.prototype.forEach.call(dropdownDisplays, function (el) {
        el.onclick = function (e) {
          e.stopPropagation();
          var dropdownId = this.getAttribute("data-dropdown-id");
          var rowIndex = parseInt(this.getAttribute("data-row"), 10);
          var fieldName = this.getAttribute("data-field");

          if (that._dropdownOpen && that._activeDropdownTrigger === this) {
            that._closeDropdown();
          } else {
            that._closeMonthPanel();
            that._openDropdownPanel(this, rowIndex, fieldName);
          }
        };
      });

      var monthDisplays = this._shadowRoot.querySelectorAll(".monthDisplay");
      Array.prototype.forEach.call(monthDisplays, function (el) {
        el.onclick = function (e) {
          e.stopPropagation();
          var rowIndex = parseInt(this.getAttribute("data-row"), 10);
          var fieldName = this.getAttribute("data-field");

          if (that._monthOpen && that._activeMonthTrigger === this) {
            that._closeMonthPanel();
          } else {
            that._closeDropdown();
            that._openMonthPanelUI(this, rowIndex, fieldName);
          }
        };
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
      this._closeDropdown();
      this._closeMonthPanel();
      this._dropdownSearchState = {};
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

  (function () {
    if (document.getElementById("asset-widget-global-style")) {
      return;
    }

    var globalStyleEl = document.createElement("style");
    globalStyleEl.id = "asset-widget-global-style";
    globalStyleEl.textContent =
      '.asset-widget-dropdown-panel {' +
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
      '.asset-widget-dropdown-panel .dropdown-topbar {' +
        'display:flex;' +
        'align-items:center;' +
        'gap:8px;' +
        'padding:8px 8px 0 8px;' +
        'background:#ffffff;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-search-wrap {' +
        'flex:1 1 auto;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-search-input {' +
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
      '.asset-widget-dropdown-panel .dropdown-close-btn {' +
        'width:32px;' +
        'height:32px;' +
        'border:1px solid #b9cae0;' +
        'border-radius:6px;' +
        'background:#ffffff;' +
        'color:#5d7288;' +
        'font-size:14px;' +
        'line-height:1;' +
        'cursor:pointer;' +
        'flex:0 0 auto;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-actions {' +
        'padding:6px 8px;' +
        'border-bottom:1px solid #e8eef5;' +
        'background:#ffffff;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-clear-btn {' +
        'width:100%;' +
        'height:28px;' +
        'border:1px solid #c7d7ea;' +
        'border-radius:6px;' +
        'background:#ffffff;' +
        'color:#0a6ed1;' +
        'font-size:12px;' +
        'font-weight:600;' +
        'cursor:pointer;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-list {' +
        'max-height:240px;' +
        'overflow:auto;' +
        'background:#ffffff;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-item {' +
        'padding:9px 10px;' +
        'font-size:13px;' +
        'color:#223548;' +
        'cursor:pointer;' +
        'border-bottom:1px solid #f3f6f9;' +
        'line-height:1.35;' +
        'word-break:break-word;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-item:hover {' +
        'background:#edf5ff;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-item.selected {' +
        'background:#e8f2ff;' +
        'color:#0a6ed1;' +
        'font-weight:600;' +
      '}' +
      '.asset-widget-dropdown-panel .dropdown-empty {' +
        'padding:12px;' +
        'color:#7b8a9a;' +
        'text-align:center;' +
        'font-size:13px;' +
      '}' +
      '.asset-widget-month-panel {' +
        'background:#ffffff;' +
        'border:1px solid #cfd9e3;' +
        'border-radius:8px;' +
        'box-shadow:0 8px 24px rgba(34,53,72,0.18);' +
        'overflow:hidden;' +
        'width:250px;' +
        'z-index:999999;' +
        'font-family:"72", Arial, sans-serif;' +
      '}' +
      '.asset-widget-month-panel .month-topbar {' +
        'display:flex;' +
        'align-items:center;' +
        'gap:8px;' +
        'padding:8px;' +
        'border-bottom:1px solid #e8eef5;' +
        'background:#ffffff;' +
      '}' +
      '.asset-widget-month-panel .month-year-select {' +
        'flex:1 1 auto;' +
        'height:32px;' +
        'border:1px solid #b9cae0;' +
        'border-radius:6px;' +
        'padding:0 10px;' +
        'box-sizing:border-box;' +
        'font-size:13px;' +
        'outline:none;' +
        'color:#223548;' +
        'background:#ffffff;' +
      '}' +
      '.asset-widget-month-panel .month-close-btn {' +
        'width:32px;' +
        'height:32px;' +
        'border:1px solid #b9cae0;' +
        'border-radius:6px;' +
        'background:#ffffff;' +
        'color:#5d7288;' +
        'font-size:14px;' +
        'line-height:1;' +
        'cursor:pointer;' +
        'flex:0 0 auto;' +
      '}' +
      '.asset-widget-month-panel .month-grid {' +
        'display:grid;' +
        'grid-template-columns:repeat(4, 1fr);' +
        'gap:8px;' +
        'padding:10px;' +
      '}' +
      '.asset-widget-month-panel .month-btn {' +
        'border:1px solid #c9d6e5;' +
        'background:#ffffff;' +
        'border-radius:6px;' +
        'padding:8px 0;' +
        'cursor:pointer;' +
        'font-size:13px;' +
        'color:#223548;' +
      '}' +
      '.asset-widget-month-panel .month-btn:hover {' +
        'background:#f3f8fd;' +
      '}' +
      '.asset-widget-month-panel .month-btn.selected {' +
        'background:#5d6673;' +
        'border-color:#5d6673;' +
        'color:#ffffff;' +
      '}' +
      '.asset-widget-month-panel .month-actions {' +
        'padding:8px;' +
        'border-top:1px solid #e8eef5;' +
        'background:#ffffff;' +
      '}' +
      '.asset-widget-month-panel .month-clear-btn {' +
        'width:100%;' +
        'height:28px;' +
        'border:1px solid #c7d7ea;' +
        'border-radius:6px;' +
        'background:#ffffff;' +
        'color:#0a6ed1;' +
        'font-size:12px;' +
        'font-weight:600;' +
        'cursor:pointer;' +
      '}';
    document.head.appendChild(globalStyleEl);
  })();
})();

