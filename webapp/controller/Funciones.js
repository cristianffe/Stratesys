sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "com/co/stratesys/zpscrearproyectos/utils/xlsx.full.min",
    "sap/ui/core/routing/History",
    "sap/ui/model/json/JSONModel",
    "sap/m/SearchField",
    "sap/ui/model/type/String",
    "sap/ui/table/Column",
    "sap/m/Column",
    "sap/m/Label",
    "sap/ui/core/Fragment",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    'sap/ui/model/Sorter',
    'sap/m/Dialog',
    "sap/m/HBox",
    "sap/m/VBox",
    "sap/m/Button",
    "sap/m/Switch",
    "sap/ui/export/Spreadsheet",
    "sap/ui/export/library",
    "sap/m/MessageBox"
], function (Controller, History, JSONModel, SearchField, TypeString, UIColumn, MColumn, Label, Fragment, Filter, FilterOperator, Sorter, Dialog, HBox, VBox, Button, Switch, Spreadsheet, exportLibrary, MessageBox) {
    "use strict";

    return Controller.extend("com.co.stratesys.zpscrearproyectos.controller.Funciones", {

        onOpenDialog: function () {

            var modelApp = this.getOwnerComponent().getModel("modelApp");

            if (!this._oDialog) {
                this._oDialog = sap.ui.xmlfragment("com.co.stratesys.zpscrearproyectos.view.FilterDialog", this);
                this.getView().addDependent(this._oDialog);
            }

            this._oDialog.open();
        },

        onFileChange: function (oEvent) {
            this._oExcelFile = oEvent.getParameter("files")[0];
        },

        closeDialog: function () {
            this._oExcelFile = null;
            this._oDialog.close();
        },

        onProcesarExcel: function (oEvent) {

            var oFile = this._oExcelFile;

            if (!oFile) {
                sap.m.MessageBox.warning("Seleccione un archivo primero");
                return;
            }

            if (!oFile.name.endsWith(".xlsx")) {
                sap.m.MessageBox.error("El archivo debe ser un Excel (.xlsx)");
                return;
            }

            // Limpiar FileUploader antes de leer
            var oDialog = oEvent.getSource().getParent();
            var oFileUploader = oDialog.findAggregatedObjects(true, function (oObj) {
                return oObj.isA("sap.ui.unified.FileUploader");
            })[0];
            if (oFileUploader) { oFileUploader.clear(); }

            sap.ui.core.BusyIndicator.show(0);

            var oReader = new FileReader();

            oReader.onload = async function (e) {

                try {

                    // ── 1. Leer workbook ──────────────────────────────────
                    var aData     = new Uint8Array(e.target.result);
                    var oWorkbook = XLSX.read(aData, { type: "array" });

                    console.log("Hojas disponibles:", oWorkbook.SheetNames);

                    var oWS = oWorkbook.Sheets["Datos Proyecto"]
                           || oWorkbook.Sheets[oWorkbook.SheetNames[1]];

                    if (!oWS) {
                        sap.m.MessageBox.error("No se encontró la hoja 'Datos Proyecto' en el Excel");
                        return;
                    }

                    // ── 2. Extraer columna A y construir mapa campo→valor ─
                    var aMatrix = XLSX.utils.sheet_to_json(oWS, { header: 1, defval: "" });

                    var oFieldMap = {};

                    aMatrix.forEach(function (aRow) {
                        var sCellA = String(aRow[0] || "").trim();
                        if (!sCellA) return;

                        // Patrón: "Campo": "Valor",  o  "Campo": null,
                        var oMatch = sCellA.match(/^"([^"]+)"\s*:\s*(.+)$/);
                        if (!oMatch) return;

                        var sKey = oMatch[1].trim();
                        var sRaw = oMatch[2].trim().replace(/,$/, ""); // quitar coma final

                        var vValue;
                        if (sRaw === "null") {
                            vValue = null;
                        } else if (/^".*"$/.test(sRaw)) {
                            vValue = sRaw.slice(1, -1); // quitar comillas
                        } else {
                            vValue = sRaw;
                        }

                        // Solo primera ocurrencia de cada campo
                        if (!(sKey in oFieldMap)) {
                            oFieldMap[sKey] = vValue;
                        }
                    });

                    console.log("FieldMap extraído:", oFieldMap);

                    // ── 3. Modelo PROYECTO (Master + Header) ──────────────
                    var oProjectData = {
                        ProjectName             : oFieldMap["ProjectName"]            || "",
                        ProjectStage            : oFieldMap["ProjectStage"]           || "",
                        OrgID                   : oFieldMap["OrgID"]                  || "",
                        ProjectCategory         : oFieldMap["ProjectCategory"]        || "",
                        Currency                : oFieldMap["Currency"]               || "",
                        StartDate               : oFieldMap["StartDate"]              || null,
                        EndDate                 : oFieldMap["EndDate"]                || null,
                        ProjManagerExtId        : oFieldMap["ProjManagerExtId"]       || "",
                        ProjManagerCompCode     : oFieldMap["ProjManagerCompCode"]    || "",
                        Customer                : oFieldMap["Customer"]               || "",
                        CostCenter              : oFieldMap["CostCenter"]             || "",
                        ProfitCenter            : oFieldMap["ProfitCenter"]           || "",
                        ProjAccountantExtId     : oFieldMap["ProjAccountantExtId"]    || "",
                        ProjAccountantCompCode  : oFieldMap["ProjAccountantCompCode"] || "",
                        ProjControllerExtId     : oFieldMap["ProjControllerExtId"]    || "",
                        ProjControllerCompCode  : oFieldMap["ProjControllerCompCode"] || "",
                        ProjPartnerExtId        : oFieldMap["ProjPartnerExtId"]       || "",
                        ProjPartnerCompCode     : oFieldMap["ProjPartnerCompCode"]    || "",
                        ProjectDesc             : oFieldMap["ProjectDesc"]            || "",
                        Confidential            : oFieldMap["Confidential"]           || "",
                        UseProjectBilling       : oFieldMap["UseProjectBilling"]      || "",
                        RestrictTimePosting     : oFieldMap["RestrictTimePosting"]    || "",
                        YY1_ACTIVE_Cpr          : oFieldMap["YY1_ACTIVE_Cpr"]         || "",
                        YY1_Fechadeventa_Cpr    : oFieldMap["YY1_Fechadeventa_Cpr"]   || null,
                        YY1_Geografia_Cpr       : oFieldMap["YY1_Geografia_Cpr"]      || "",
                        YY1_Producto_Cpr        : oFieldMap["YY1_Producto_Cpr"]       || "",
                        YY1_Tipodeproyecto_Cpr  : oFieldMap["YY1_Tipodeproyecto_Cpr"] || "",
                        // Campos display del header expandido
                        PlannedEffort           : "",
                        PlannedCost             : "",
                        PlannedRevenue          : "",
                        PlannedMargin           : "",
                        SalesRevenue            : "",
                        SalesMargin             : ""
                    };

                    // ── 4. Modelo WBS ─────────────────────────────────────
                    var aWBSData = [{
                        WorkPackageName           : oFieldMap["WorkPackageName"]            || "",
                        Description               : oFieldMap["Description"]               || "",
                        WPStartDate               : oFieldMap["WPStartDate"]               || null,
                        WPEndDate                 : oFieldMap["WPEndDate"]                 || null,
                        WorkPackageType           : oFieldMap["WorkPackageType"]           || "",
                        UnitQuantity              : oFieldMap["UnitQuantity"]              || "",
                        UnitId                    : oFieldMap["UnitId"]                    || "",
                        YY1_TipodeproyectoSub_cpd : oFieldMap["YY1_TipodeproyectoSub_cpd"] || ""
                    }];

                    // ── 5. Modelo Resource ────────────────────────────────
                    var aResourceData = [{
                        Version                       : oFieldMap["Version"]                       || "",
                        EngagementProject             : oFieldMap["EngagementProject"]             || "",
                        WorkItem                      : oFieldMap["WorkItem"]                      || "",
                        BillingControlCategory        : oFieldMap["BillingControlCategory"]        || "",
                        DeliveryOrganization          : oFieldMap["DeliveryOrganization"]          || "",
                        EngagementProjectResourceType : oFieldMap["EngagementProjectResourceType"] || "",
                        EngagementProjectResource     : oFieldMap["EngagementProjectResource"]     || "",
                        WorkforcePersonUserID         : oFieldMap["WorkforcePersonUserID"]         || "",
                        PersonWorkAgreement           : oFieldMap["PersonWorkAgreement"]           || "",
                        ResourceDemandStatus          : oFieldMap["ResourceDemandStatus"]          || "",
                        UnitOfMeasure                 : oFieldMap["UnitOfMeasure"]                 || "",
                        Quantity                      : oFieldMap["Quantity"]                      || "",
                        Currency                      : oFieldMap["Currency"]                      || ""
                    }];

                    // ── 6. Modelo Billing ─────────────────────────────────
                    var aBillingData = [{
                        SalesOrderItem                : oFieldMap["SalesOrderItem"]                || "",
                        BillingPlanItem               : oFieldMap["BillingPlanItem"]               || "",
                        SalesOrder                    : oFieldMap["SalesOrder"]                    || "",
                        BillingPlanBillingDate        : oFieldMap["BillingPlanBillingDate"]        || null,
                        BillingPlanRelatedBillgStatus : oFieldMap["BillingPlanRelatedBillgStatus"] || "",
                        BillingPlanAmount             : oFieldMap["BillingPlanAmount"]             || "",
                        TransactionCurrency           : oFieldMap["TransactionCurrency"]           || "",
                        BillingPlanItemUsage          : oFieldMap["BillingPlanItemUsage"]          || "",
                        BillingPlanItemDescription    : oFieldMap["BillingPlanItemDescription"]    || "",
                        BillingPlanServiceStartDate   : oFieldMap["BillingPlanServiceStartDate"]   || null,
                        BillingPlanServiceEndDate     : oFieldMap["BillingPlanServiceEndDate"]     || null
                    }];

                    // ── 7. Setear modelos en la vista ─────────────────────
                    var oView = this.getView();

                    oView.setModel(new sap.ui.model.json.JSONModel(oProjectData),  "ProjectSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(aWBSData),      "WBSSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(aResourceData), "ResourceSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(aBillingData),  "BillingSet");

                    // ── 8. Bindear master y detail al raíz de ProjectSet ──
                    oView.byId("masterPage").bindElement("ProjectSet>/");
                    oView.byId("detailPage").bindElement("ProjectSet>/");

                    console.log("Proyecto:", oProjectData);
                    console.log("WBS:",      aWBSData);
                    console.log("Resource:", aResourceData);
                    console.log("Billing:",  aBillingData);

                    this.closeDialog();
                    sap.m.MessageToast.show("Archivo cargado correctamente");

                } catch (ex) {
                    console.error("Error procesando Excel:", ex);
                    sap.m.MessageBox.error("Error al cargar el archivo Excel: " + ex.message);
                } finally {
                    sap.ui.core.BusyIndicator.hide();
                }

            }.bind(this);

            oReader.readAsArrayBuffer(oFile);
        },

        onUploadFile: function () {

           this._oDialog.open();

            //return this._oDialog;
        },

    });
});