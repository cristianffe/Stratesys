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
], function (Controller, xls, History, JSONModel, SearchField, TypeString, UIColumn, MColumn, Label, Fragment, Filter, FilterOperator, Sorter, Dialog, HBox, VBox, Button, Switch, Spreadsheet, exportLibrary, MessageBox) {
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
                    var aData = new Uint8Array(e.target.result);
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
                        ProjectID: oFieldMap["ProjectID"] || "",
                        ProjectName: oFieldMap["ProjectName"] || "",
                        ProjectStage: oFieldMap["ProjectStage"] || "",
                        OrgID: oFieldMap["OrgID"] || "",
                        ProjectCategory: oFieldMap["ProjectCategory"] || "",
                        Currency: oFieldMap["Currency"] || "",
                        StartDate: oFieldMap["StartDate"] || null,
                        EndDate: oFieldMap["EndDate"] || null,
                        ProjManagerExtId: oFieldMap["ProjManagerExtId"] || "",
                        ProjManagerCompCode: oFieldMap["ProjManagerCompCode"] || "",
                        Customer: oFieldMap["Customer"] || "",
                        CostCenter: oFieldMap["CostCenter"] || "",
                        ProfitCenter: oFieldMap["ProfitCenter"] || "",
                        ProjAccountantExtId: oFieldMap["ProjAccountantExtId"] || "",
                        ProjAccountantCompCode: oFieldMap["ProjAccountantCompCode"] || "",
                        ProjControllerExtId: oFieldMap["ProjControllerExtId"] || "",
                        ProjControllerCompCode: oFieldMap["ProjControllerCompCode"] || "",
                        ProjPartnerExtId: oFieldMap["ProjPartnerExtId"] || "",
                        ProjPartnerCompCode: oFieldMap["ProjPartnerCompCode"] || "",
                        ProjectDesc: oFieldMap["ProjectDesc"] || "",
                        Confidential: oFieldMap["Confidential"] || "",
                        UseProjectBilling: oFieldMap["UseProjectBilling"] || "",
                        RestrictTimePosting: oFieldMap["RestrictTimePosting"] || "",
                        YY1_ACTIVE_Cpr: oFieldMap["YY1_ACTIVE_Cpr"] || "",
                        YY1_Fechadeventa_Cpr: oFieldMap["YY1_Fechadeventa_Cpr"] || null,
                        YY1_Geografia_Cpr: oFieldMap["YY1_Geografia_Cpr"] || "",
                        YY1_Producto_Cpr: oFieldMap["YY1_Producto_Cpr"] || "",
                        YY1_Tipodeproyecto_Cpr: oFieldMap["YY1_Tipodeproyecto_Cpr"] || "",
                        // Campos display del header expandido
                        PlannedEffort: "",
                        PlannedCost: "",
                        PlannedRevenue: "",
                        PlannedMargin: "",
                        SalesRevenue: "",
                        SalesMargin: ""
                    };

                    // ── 4. Modelo WBS ─────────────────────────────────────
                    var aWBSData = [{
                        WorkPackageName: oFieldMap["WorkPackageName"] || "",
                        Description: oFieldMap["Description"] || "",
                        WPStartDate: oFieldMap["WPStartDate"] || null,
                        WPEndDate: oFieldMap["WPEndDate"] || null,
                        WorkPackageType: oFieldMap["WorkPackageType"] || "",
                        UnitQuantity: oFieldMap["UnitQuantity"] || "",
                        UnitId: oFieldMap["UnitId"] || "",
                        YY1_TipodeproyectoSub_cpd: oFieldMap["YY1_TipodeproyectoSub_cpd"] || ""
                    }];

                    // ── 5. Modelo Resource ────────────────────────────────
                    var aResourceData = [{
                        Version: oFieldMap["Version"] || "",
                        EngagementProject: oFieldMap["EngagementProject"] || "",
                        WorkItem: oFieldMap["WorkItem"] || "",
                        BillingControlCategory: oFieldMap["BillingControlCategory"] || "",
                        DeliveryOrganization: oFieldMap["DeliveryOrganization"] || "",
                        EngagementProjectResourceType: oFieldMap["EngagementProjectResourceType"] || "",
                        EngagementProjectResource: oFieldMap["EngagementProjectResource"] || "",
                        WorkforcePersonUserID: oFieldMap["WorkforcePersonUserID"] || "",
                        PersonWorkAgreement: oFieldMap["PersonWorkAgreement"] || "",
                        ResourceDemandStatus: oFieldMap["ResourceDemandStatus"] || "",
                        UnitOfMeasure: oFieldMap["UnitOfMeasure"] || "",
                        Quantity: oFieldMap["Quantity"] || "",
                        Currency: oFieldMap["Currency"] || ""
                    }];

                    // ── 6. Modelo Billing ─────────────────────────────────
                    var aBillingData = [{
                        SalesOrderItem: oFieldMap["SalesOrderItem"] || "",
                        BillingPlanItem: oFieldMap["BillingPlanItem"] || "",
                        SalesOrder: oFieldMap["SalesOrder"] || "",
                        BillingPlanBillingDate: oFieldMap["BillingPlanBillingDate"] || null,
                        BillingPlanRelatedBillgStatus: oFieldMap["BillingPlanRelatedBillgStatus"] || "",
                        BillingPlanAmount: oFieldMap["BillingPlanAmount"] || "",
                        TransactionCurrency: oFieldMap["TransactionCurrency"] || "",
                        BillingPlanItemUsage: oFieldMap["BillingPlanItemUsage"] || "",
                        BillingPlanItemDescription: oFieldMap["BillingPlanItemDescription"] || "",
                        BillingPlanServiceStartDate: oFieldMap["BillingPlanServiceStartDate"] || null,
                        BillingPlanServiceEndDate: oFieldMap["BillingPlanServiceEndDate"] || null
                    }];

                    // ── 7. Setear modelos en la vista ─────────────────────
                    var oView = this.getView();

                    oView.setModel(new sap.ui.model.json.JSONModel(oProjectData), "ProjectSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(aWBSData), "WBSSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(aResourceData), "ResourceSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(aBillingData), "BillingSet");

                    // ── 8. Bindear master y detail al raíz de ProjectSet ──
                    oView.byId("masterPage").bindElement("ProjectSet>/");
                    oView.byId("detailPage").bindElement("ProjectSet>/");

                    console.log("Proyecto:", oProjectData);
                    console.log("WBS:", aWBSData);
                    console.log("Resource:", aResourceData);
                    console.log("Billing:", aBillingData);

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

        onMasterSearch: function (oEvent) {
            var sQuery = (oEvent.getParameter("query")
                || oEvent.getParameter("newValue")
                || "").trim().toLowerCase();

            var oContainer = this.byId("masterFieldsBox");
            if (!oContainer) return;

            oContainer.getItems().forEach(function (oHBox) {
                if (!oHBox.isA("sap.m.HBox")) {
                    oHBox.setVisible(true);
                    return;
                }

                var aChildren = oHBox.getItems();
                var sLabel = "";
                var sValue = "";

                aChildren.forEach(function (oChild) {
                    if (oChild.isA("sap.m.Label")) {
                        sLabel = (oChild.getText() || "").toLowerCase();
                    } else if (oChild.isA("sap.m.Text")) {
                        sValue = (oChild.getText() || "").toLowerCase();
                    }
                });

                // Sin query → mostrar todo
                // Con query → mostrar solo si el nombre del campo O el valor coincide
                oHBox.setVisible(
                    !sQuery || sLabel.includes(sQuery) || sValue.includes(sQuery)
                );
            });
        },

        onOpenFilterDialog: function () {
            var oDialog = this.byId("filterDialog");
            if (oDialog) {
                oDialog.open();
            }
        },

        onCloseFilterDialog: function () {
            var oDialog = this.byId("filterDialog");
            if (oDialog) {
                oDialog.close();
            }
        },

        onClearFilter: function () {
            // Limpiar todos los inputs del diálogo
            var aInputIds = [
                "fProjectName", "fProjectStage", "fOrgID", "fProjectCategory",
                "fCurrency", "fProjectDesc", "fProjManagerExtId", "fProjAccountantExtId",
                "fProjControllerExtId", "fProjPartnerExtId", "fCustomer", "fCostCenter",
                "fProfitCenter", "fGeografia", "fProducto", "fTipoproyecto"
            ];
            aInputIds.forEach(function (sId) {
                var oInput = this.byId(sId);
                if (oInput) oInput.setValue("");
            }.bind(this));

            // Resetear los Select a "(Todos)"
            ["fActive", "fConfidential", "fUseProjectBilling", "fRestrictTimePosting"]
                .forEach(function (sId) {
                    var oSelect = this.byId(sId);
                    if (oSelect) oSelect.setSelectedKey("");
                }.bind(this));

            // Limpiar DatePickers
            ["fStartDateFrom", "fStartDateTo", "fEndDateFrom", "fEndDateTo"]
                .forEach(function (sId) {
                    var oDp = this.byId(sId);
                    if (oDp) oDp.setValue("");
                }.bind(this));

            // Mostrar todos los campos del master de nuevo
            this.byId("masterFieldsBox").getItems().forEach(function (oHBox) {
                oHBox.setVisible(true);
            });
        },

        onApplyFilter: function () {
            // Leer valores del diálogo
            var oData = this.getView().getModel("ProjectSet").getData();

            var checks = [
                { id: "fProjectName", field: "ProjectName" },
                { id: "fProjectStage", field: "ProjectStage" },
                { id: "fOrgID", field: "OrgID" },
                { id: "fProjectCategory", field: "ProjectCategory" },
                { id: "fCurrency", field: "Currency" },
                { id: "fProjectDesc", field: "ProjectDesc" },
                { id: "fProjManagerExtId", field: "ProjManagerExtId" },
                { id: "fProjAccountantExtId", field: "ProjAccountantExtId" },
                { id: "fProjControllerExtId", field: "ProjControllerExtId" },
                { id: "fProjPartnerExtId", field: "ProjPartnerExtId" },
                { id: "fCustomer", field: "Customer" },
                { id: "fCostCenter", field: "CostCenter" },
                { id: "fProfitCenter", field: "ProfitCenter" },
                { id: "fGeografia", field: "YY1_Geografia_Cpr" },
                { id: "fProducto", field: "YY1_Producto_Cpr" },
                { id: "fTipoproyecto", field: "YY1_Tipodeproyecto_Cpr" }
            ];

            // Para cada HBox del master, mostrar/ocultar según si su campo pasa el filtro
            var oContainer = this.byId("masterFieldsBox");
            oContainer.getItems().forEach(function (oHBox) {
                if (!oHBox.isA("sap.m.HBox")) { oHBox.setVisible(true); return; }

                var sLabelText = oHBox.getItems()[0].getText(); // el Label

                // Buscar si este campo tiene un filtro activo
                var oCheck = checks.find(function (c) { return c.field === sLabelText; });
                if (!oCheck) { oHBox.setVisible(true); return; }

                var sFilter = (this.byId(oCheck.id).getValue() || "").trim().toLowerCase();
                if (!sFilter) { oHBox.setVisible(true); return; }

                var sValue = (oData[oCheck.field] || "").toString().toLowerCase();
                oHBox.setVisible(sValue.includes(sFilter));
            }.bind(this));

            this.onCloseFilterDialog();
        },

        onCreate: async function () {

            var oView = this.getView();
            var header = oView.getModel("ProjectSet").getData();
            var workPackage = oView.getModel("WBSSet").getData();
            var resourceDemand = oView.getModel("ResourceSet").getData();

            var payloadHeader = this.setHeaderProyecto(header);
            var payloadWorkPackage = this.setWorkPackageProyecto(workPackage);
            var payloadResourceDemand = this.setResourceDemandProyecto(resourceDemand);

            var payloadApi = this.parsearPayloadApi(payloadHeader, payloadWorkPackage, payloadResourceDemand);

            const base64 = btoa(payloadApi);
            var oPayloadBase64 = {
                id: "0123456789",
                body: base64
            };

            var sEndpoint = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Project";

            // ✅ Confirmación antes de ejecutar
            MessageBox.confirm("¿Esta Seguro De crear el proyecto?", {
                title: "Confirmar creación",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                onClose: async function (sAction) {

                    if (sAction !== MessageBox.Action.YES) return;

                    sap.ui.core.BusyIndicator.show(0);

                    try {
                        var rpta = await this._postODataV4(sEndpoint, oPayloadBase64);

                        if (rpta.numericSeverity === 4) {
                            MessageBox.error(rpta.mensaje);
                        } else {
                            MessageBox.success(rpta.mensaje || `Proyecto ${rpta.id} creado exitosamente`);
                        }

                    } catch (oError) {
                        MessageBox.error("Error al crear el proyecto");

                    } finally {
                        sap.ui.core.BusyIndicator.hide();
                    }

                }.bind(this)
            });
        },

        setHeaderProyecto: function (oData) {

            var oPayload = {
                ProjectID: oData.ProjectID || "",
                Confidential: oData.Confidential || "N",
                CostCenter: oData.CostCenter || "",
                Currency: oData.Currency || "",
                Customer: oData.Customer || "",
                EndDate: oData.EndDate || null,
                OrgID: oData.OrgID || "",
                ProfitCenter: oData.ProfitCenter || "",
                ProjAccountantCompCode: oData.ProjAccountantCompCode || "",
                ProjAccountantExtId: oData.ProjAccountantExtId || "",
                ProjControllerCompCode: oData.ProjControllerCompCode || "",
                ProjControllerExtId: oData.ProjControllerExtId || "",
                ProjManagerCompCode: oData.ProjManagerCompCode || "",
                ProjManagerExtId: oData.ProjManagerExtId || "",
                ProjPartnerCompCode: oData.ProjPartnerCompCode || "",
                ProjPartnerExtId: oData.ProjPartnerExtId || "",
                ProjectCategory: oData.ProjectCategory || "",
                ProjectDesc: oData.ProjectDesc || "",
                ProjectName: oData.ProjectName || "",
                ProjectStage: oData.ProjectStage || "",
                RestrictTimePosting: oData.RestrictTimePosting || "N",
                StartDate: oData.StartDate || null,
                UseProjectBilling: oData.UseProjectBilling || "",
                YY1_ACTIVE_Cpr: oData.YY1_ACTIVE_Cpr || "",
                YY1_Fechadeventa_Cpr: oData.YY1_Fechadeventa_Cpr || null,
                YY1_Geografia_Cpr: oData.YY1_Geografia_Cpr || "",
                YY1_Producto_Cpr: oData.YY1_Producto_Cpr || "",
                YY1_Tipodeproyecto_Cpr: oData.YY1_Tipodeproyecto_Cpr || ""
            };

            return oPayload;

        },

        setWorkPackageProyecto: function (aWBSData) {

            var aWBSPayloads = [];

            for (var i = 0; i < aWBSData.length; i++) {
                var oItem = aWBSData[i];

                var oPayload = {
                    // EngagementProject: sProjectID,
                    WorkPackageName: oItem.WorkPackageName || "",
                    Description: oItem.Description || "",
                    WPStartDate: oItem.WPStartDate || null,
                    WPEndDate: oItem.WPEndDate || null,
                    WorkPackageType: oItem.WorkPackageType || "",
                    UnitQuantity: oItem.UnitQuantity || "0.000",
                    UnitId: oItem.UnitId || "",
                    YY1_TipodeproyectoSub_cpd: oItem.YY1_TipodeproyectoSub_cpd || ""
                };

                aWBSPayloads.push(oPayload);
            }

            return aWBSPayloads;

        },

        setResourceDemandProyecto: function (aResourceData) {

            var aResourcePayloads = [];

            for (var i = 0; i < aResourceData.length; i++) {
                var oItem = aResourceData[i];

                var oPayload = {
                    BillingControlCategory: oItem.BillingControlCategory || "",
                    Currency: oItem.Currency || "",
                    DeliveryOrganization: oItem.DeliveryOrganization || "",
                    //  EngagementProject: sProjectID,
                    Country2DigitISOCode: "ES",
                    EngagementProjectResource: oItem.EngagementProjectResource || "",
                    EngagementProjectResourceType: oItem.EngagementProjectResourceType || "",
                    PersonWorkAgreement: oItem.PersonWorkAgreement || "",
                    Quantity: oItem.Quantity || "0.000",
                    ResourceDemandStatus: oItem.ResourceDemandStatus || "",
                    UnitOfMeasure: oItem.UnitOfMeasure || "",
                    Version: oItem.Version || "",
                    WorkItem: oItem.WorkItem || "",
                    WorkforcePersonUserID: oItem.WorkforcePersonUserID || ""
                };

                aResourcePayloads.push(oPayload);

            }

            return aResourcePayloads;

        },

        parsearPayloadApi: function (payloadHeader, payloadWorkPackage, payloadResourceDemand) {

            var oPayload = Object.assign({}, payloadHeader);

            if (payloadResourceDemand) {
                // Anidar resources dentro de cada WorkPackage
                for (var i = 0; i < payloadWorkPackage.length; i++) {
                    payloadWorkPackage[i].to_ResourceDemand = payloadResourceDemand;
                }
            }

            oPayload.WorkPackageSet = payloadWorkPackage;

            var payloadApi = JSON.stringify(oPayload, null, 2);
            return payloadApi;

        },

        postODataV4: function (sEndpoint, oBody) {
            return new Promise(function (resolve, reject) {

                $.ajax({
                    url: sEndpoint,
                    method: "POST",
                    contentType: "application/json",
                    headers: {
                        "Accept": "application/json",
                        "X-Requested-With": "XMLHttpRequest"
                    },
                    data: JSON.stringify(oBody),
                    success: function (oResponse) {
                        console.log("POST exitoso:", oResponse);
                        resolve(oResponse);
                    },
                    error: function (oError) {
                        console.error("Error en POST:", oError);
                        reject(oError);
                    }
                });

            });
        },

        _postODataV4: async function (sEndpoint, oBody) {
            return new Promise(function (resolve, reject) {

                // ── 1. Fetch CSRF Token ────────────────────────────────────
                $.ajax({
                    url: "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/$metadata",
                    method: "GET",
                    headers: {
                        "X-CSRF-Token": "Fetch"
                    },
                    success: function (data, status, oXHR) {
                        var sToken = oXHR.getResponseHeader("X-CSRF-Token");

                        console.log("CSRF Token obtenido:", sToken);

                        // ── 2. POST con el token ───────────────────────────
                        $.ajax({
                            url: sEndpoint,
                            method: "POST",
                            contentType: "application/json",
                            headers: {
                                "X-CSRF-Token": sToken,
                                "Accept": "application/json",
                                "X-Requested-With": "XMLHttpRequest"
                            },
                            data: JSON.stringify(oBody),
                            success: function (oResponse, sStatus, oXHR) {
                                let sMensaje = "Proyecto creado exitosamente";
                                let nSeverity = 1;

                                try {
                                    const sSapMessages = oXHR.getResponseHeader("sap-messages");

                                    if (sSapMessages) {
                                        const aMessages = JSON.parse(sSapMessages).reverse();
                                        sMensaje = aMessages.map(function (o) { return o.message; }).join("");
                                        nSeverity = aMessages[0]?.numericSeverity;
                                    }

                                } catch (e) {
                                    console.warn("No se pudo parsear sap-messages:", e);
                                }

                                resolve({
                                    mensaje: sMensaje,
                                    numericSeverity: nSeverity,
                                    id: oResponse?.id   // ✅ retornar el id del proyecto creado
                                });
                            },
                            error: function (oError) {
                                console.error("Error en POST:", oError);
                                reject(oError);
                            }
                        });
                    },
                    error: function (oError) {
                        console.error("Error obteniendo CSRF Token:", oError);
                        reject(oError);
                    }
                });

            });


        },

    });
});