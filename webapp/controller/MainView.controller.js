sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/m/Column",
  "sap/m/ColumnListItem",
  "sap/m/Text",
  "sap/m/Label",
  "sap/m/MessageToast",
  "sap/m/MessageBox"
], function (
  Controller,
  Column,
  ColumnListItem,
  Text,
  Label,
  MessageToast,
  MessageBox
) {
  "use strict";

  return Controller.extend("com.excel.excelupload.controller.MainView", {

    // ─────────────────────────────────────────
    // INIT
    // ─────────────────────────────────────────
    onInit: function () {
      this._oRawData      = [];
      this._oColumns      = [];
      this._oSelectedFile = null;
    },

    // ─────────────────────────────────────────
    // FILE CHANGE
    // ─────────────────────────────────────────
    onFileChange: function (oEvent) {
      var oFileUploader = this.byId("fileUploader");
      var oFocusDom     = oFileUploader.getFocusDomRef();

      if (!oFocusDom || !oFocusDom.files || oFocusDom.files.length === 0) {
        MessageToast.show("No file selected.");
        return;
      }

      var oFile = oFocusDom.files[0];
      this._validateFile(oFile);
    },

    // ─────────────────────────────────────────
    // VALIDATE FILE
    // ─────────────────────────────────────────
    _validateFile: function (oFile) {
      var oErrorStrip = this.byId("errorStrip");
      var oInfoStrip  = this.byId("fileInfoStrip");
      var oUploadBtn  = this.byId("uploadBtn");

      oErrorStrip.setVisible(false);
      oInfoStrip.setVisible(false);
      oUploadBtn.setEnabled(false);
      this._oSelectedFile = null;

      // Validate type
      var sName  = oFile.name.toLowerCase();
      var bValid = sName.endsWith(".xlsx") ||
                   sName.endsWith(".xls")  ||
                   sName.endsWith(".csv");

      if (!bValid) {
        oErrorStrip.setText(
          "Invalid file: " + oFile.name +
          ". Only .xlsx, .xls or .csv allowed."
        );
        oErrorStrip.setVisible(true);
        return;
      }

      // Validate size (10MB)
      if (oFile.size > 10 * 1024 * 1024) {
        oErrorStrip.setText(
          "File too large (" + this._formatSize(oFile.size) +
          "). Max allowed is 10MB."
        );
        oErrorStrip.setVisible(true);
        return;
      }

      // All good
      this._oSelectedFile = oFile;
      oInfoStrip.setText(
        "Ready: " + oFile.name + " (" + this._formatSize(oFile.size) + ")"
      );
      oInfoStrip.setVisible(true);
      oUploadBtn.setEnabled(true);
    },

    // ─────────────────────────────────────────
    // FORMAT FILE SIZE
    // ─────────────────────────────────────────
    _formatSize: function (n) {
      if (n < 1024)        return n + " B";
      if (n < 1048576)     return (n / 1024).toFixed(1) + " KB";
      return (n / 1048576).toFixed(1) + " MB";
    },

    // ─────────────────────────────────────────
    // UPLOAD PRESS
    // ─────────────────────────────────────────
    onUploadPress: function () {
      if (!this._oSelectedFile) {
        MessageBox.warning("Please select a file first.");
        return;
      }
      this._readFile(this._oSelectedFile);
    },

    // ─────────────────────────────────────────
    // READ FILE
    // ─────────────────────────────────────────
    _readFile: function (oFile) {
      var oProgress  = this.byId("uploadProgress");
      var oUploadBtn = this.byId("uploadBtn");
      var that       = this;

      oProgress.setVisible(true);
      oProgress.setPercentValue(10);
      oProgress.setDisplayValue("Reading file...");
      oUploadBtn.setEnabled(false);

      var oReader = new FileReader();

      oReader.onload = function (e) {
        try {
          oProgress.setPercentValue(50);
          oProgress.setDisplayValue("Parsing...");

          var sName   = oFile.name.toLowerCase();
          var oResult = sName.endsWith(".csv")
            ? that._parseCSV(e.target.result)
            : that._parseExcel(e.target.result);

          oProgress.setPercentValue(80);
          oProgress.setDisplayValue("Building table...");

          that._buildTable(
            oResult.columns,
            oResult.data,
            oResult.sheetName,
            oFile.name
          );

          oProgress.setPercentValue(100);
          oProgress.setDisplayValue("Done!");

          setTimeout(function () {
            oProgress.setVisible(false);
            oUploadBtn.setEnabled(true);
          }, 500);

          MessageToast.show(
            "Loaded " + oResult.data.length + " rows successfully!"
          );

        } catch (err) {
          oProgress.setVisible(false);
          oUploadBtn.setEnabled(true);
          MessageBox.error("Parse error: " + err.message);
        }
      };

      oReader.onerror = function () {
        oProgress.setVisible(false);
        oUploadBtn.setEnabled(true);
        MessageBox.error("Could not read file. Please try again.");
      };

      if (oFile.name.toLowerCase().endsWith(".csv")) {
        oReader.readAsText(oFile);
      } else {
        oReader.readAsArrayBuffer(oFile);
      }
    },

    // ─────────────────────────────────────────
    // PARSE EXCEL
    // ─────────────────────────────────────────
    _parseExcel: function (oBuffer) {
      if (typeof XLSX === "undefined") {
        throw new Error("XLSX library not loaded. Check index.html.");
      }

      var oWB = XLSX.read(new Uint8Array(oBuffer), { type: "array" });

      if (!oWB.SheetNames.length) {
        throw new Error("No sheets found in Excel file.");
      }

      var sSheet = oWB.SheetNames[0];
      var oWS    = oWB.Sheets[sSheet];
      var aRaw   = XLSX.utils.sheet_to_json(oWS, {
        header   : 1,
        defval   : "",
        blankrows: false
      });

      if (aRaw.length < 2) {
        throw new Error("File has no data rows.");
      }

      return this._arrayToData(aRaw, sSheet);
    },

    // ─────────────────────────────────────────
    // PARSE CSV
    // ─────────────────────────────────────────
    _parseCSV: function (sText) {
      if (!sText || !sText.trim()) {
        throw new Error("CSV file is empty.");
      }

      var aLines = sText
        .split("\n")
        .map(function (l) { return l.replace(/\r/g, ""); })
        .filter(function (l) { return l.trim() !== ""; });

      if (aLines.length < 2) {
        throw new Error("CSV needs at least 1 header + 1 data row.");
      }

      var fnSplit = function (sRow) {
        var aRes = [], bQ = false, s = "";
        for (var i = 0; i < sRow.length; i++) {
          var c = sRow[i];
          if (c === '"')      { bQ = !bQ; }
          else if (c === "," && !bQ) { aRes.push(s.trim()); s = ""; }
          else                { s += c; }
        }
        aRes.push(s.trim());
        return aRes;
      };

      return this._arrayToData(
        aLines.map(function (l) { return fnSplit(l); }),
        "CSV"
      );
    },

    // ─────────────────────────────────────────
    // ARRAY TO DATA OBJECT
    // ─────────────────────────────────────────
    _arrayToData: function (aRaw, sSheet) {
      var aHeaders = aRaw[0].map(function (h, i) {
        return String(h).trim() || "Column_" + (i + 1);
      });

      var aData = [];
      for (var i = 1; i < aRaw.length; i++) {
        var oRow   = {};
        var bEmpty = true;
        aHeaders.forEach(function (h, j) {
          var v = aRaw[i][j] !== undefined ? String(aRaw[i][j]) : "";
          oRow["col_" + j] = v;
          if (v.trim()) bEmpty = false;
        });
        if (!bEmpty) aData.push(oRow);
      }

      return { columns: aHeaders, data: aData, sheetName: sSheet };
    },

    // ─────────────────────────────────────────
    // BUILD TABLE
    // ─────────────────────────────────────────
    _buildTable: function (aColumns, aData, sSheet, sFileName) {
      var oTable = this.byId("excelTable");

      this._oColumns = aColumns;
      this._oRawData = aData;

      oTable.removeAllColumns();
      oTable.removeAllItems();

      // Add columns
      aColumns.forEach(function (sCol) {
        oTable.addColumn(new Column({
          header        : new Label({ text: sCol }),
          width         : "auto",
          minScreenWidth: "Tablet",
          demandPopin   : true,
          popinDisplay  : "Inline"
        }));
      });

      // Add rows
      this._fillRows(aData);

      // Update stats
      this.byId("totalRowsText").setText(String(aData.length));
      this.byId("totalColsText").setText(String(aColumns.length));
      this.byId("sheetNameText").setText(sSheet || "Sheet1");
      this.byId("tableTitle").setText(
        sFileName + "  (" + aData.length + " rows)"
      );

      // Show panels
      this.byId("statsPanel").setVisible(true);
      this.byId("tablePanel").setVisible(true);
    },

    // ─────────────────────────────────────────
    // FILL TABLE ROWS
    // ─────────────────────────────────────────
    _fillRows: function (aData) {
      var oTable   = this.byId("excelTable");
      var aColumns = this._oColumns;

      oTable.removeAllItems();

      aData.forEach(function (oRow) {
        var aCells = aColumns.map(function (h, j) {
          return new Text({ text: oRow["col_" + j] || "", wrapping: false });
        });
        oTable.addItem(new ColumnListItem({ cells: aCells }));
      });
    },

    // ─────────────────────────────────────────
    // SEARCH
    // ─────────────────────────────────────────
    onSearch: function (oEvent) {
      var sQ = oEvent.getParameter("newValue").toLowerCase().trim();

      if (!sQ) {
        this._fillRows(this._oRawData);
        return;
      }

      var aFiltered = this._oRawData.filter(function (oRow) {
        return Object.values(oRow).some(function (v) {
          return v.toLowerCase().includes(sQ);
        });
      });

      this._fillRows(aFiltered);
    },

    // ─────────────────────────────────────────
    // DELETE ROWS
    // ─────────────────────────────────────────
    onDeleteRows: function () {
      var oTable = this.byId("excelTable");
      var aSelected = oTable.getSelectedItems();

      if (!aSelected.length) {
        MessageToast.show("Select at least one row to delete.");
        return;
      }

      MessageBox.confirm(
        "Delete " + aSelected.length + " row(s)?",
        {
          title: "Confirm Delete",
          onClose: function (sAction) {
            if (sAction === MessageBox.Action.OK) {
              var that = this;

              aSelected
                .map(function (o) { return oTable.indexOfItem(o); })
                .sort(function (a, b) { return b - a; })
                .forEach(function (idx) {
                  that._oRawData.splice(idx, 1);
                  oTable.removeItem(oTable.getItems()[idx]);
                });

              that.byId("totalRowsText").setText(
                String(that._oRawData.length)
              );
              MessageToast.show(aSelected.length + " row(s) deleted.");
            }
          }.bind(this)
        }
      );
    },

    // ─────────────────────────────────────────
    // EXPORT CSV
    // ─────────────────────────────────────────
    onExportCSV: function () {
      if (!this._oRawData.length) {
        MessageToast.show("No data to export.");
        return;
      }

      var aCol = this._oColumns;
      var sCSV = aCol.join(",") + "\n";

      this._oRawData.forEach(function (oRow) {
        sCSV += aCol.map(function (h, j) {
          var v = oRow["col_" + j] || "";
          if (v.includes(",") || v.includes('"')) {
            v = '"' + v.replace(/"/g, '""') + '"';
          }
          return v;
        }).join(",") + "\n";
      });

      var oBlob = new Blob([sCSV], { type: "text/csv;charset=utf-8;" });
      var sUrl  = URL.createObjectURL(oBlob);
      var oA    = document.createElement("a");
      oA.href     = sUrl;
      oA.download = "export_" + Date.now() + ".csv";
      document.body.appendChild(oA);
      oA.click();
      document.body.removeChild(oA);
      URL.revokeObjectURL(sUrl);

      MessageToast.show("CSV exported!");
    },

    // ─────────────────────────────────────────
    // RESET
    // ─────────────────────────────────────────
    onReset: function () {
      MessageBox.confirm("Clear all data and reset?", {
        title: "Confirm Reset",
        onClose: function (sAction) {
          if (sAction === MessageBox.Action.OK) {
            this._oRawData      = [];
            this._oColumns      = [];
            this._oSelectedFile = null;

            var oTable = this.byId("excelTable");
            oTable.removeAllColumns();
            oTable.removeAllItems();

            this.byId("statsPanel").setVisible(false);
            this.byId("tablePanel").setVisible(false);
            this.byId("fileUploader").clear();
            this.byId("uploadBtn").setEnabled(false);
            this.byId("fileInfoStrip").setVisible(false);
            this.byId("errorStrip").setVisible(false);
            this.byId("uploadProgress").setVisible(false);
            this.byId("totalRowsText").setText("0");
            this.byId("totalColsText").setText("0");
            this.byId("sheetNameText").setText("-");

            MessageToast.show("Reset done.");
          }
        }.bind(this)
      });
    }

  });
});