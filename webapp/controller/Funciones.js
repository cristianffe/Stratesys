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

        onAfterRendering: function () {
            alert.show("Bienvenido a la aplicación de creación de proyectos. Para comenzar, haz clic en el botón 'Cargar Excel' y selecciona tu archivo .xlsm con los datos del proyecto. Asegúrate de que el formato del Excel sea correcto para que los datos se mapeen correctamente. ¡Gracias por usar la aplicación!");
        },

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

        onProcesarExcel: async function (oEvent) {

            var oFile = this._oExcelFile;

            if (!oFile) {
                sap.m.MessageBox.warning("Seleccione un archivo primero");
                return;
            }

            if (!oFile.name.endsWith(".xlsm")) {
                sap.m.MessageBox.error("El archivo debe ser un Excel (.xlsm)");
                return;
            }

            var oDialog = oEvent.getSource().getParent();
            var oFileUploader = oDialog.findAggregatedObjects(true, function (oObj) {
                return oObj.isA("sap.ui.unified.FileUploader");
            })[0];
            if (oFileUploader) { oFileUploader.clear(); }

            this.getOwnerComponent().getModel("AppModel").setProperty("/OrgId", "");
            this.getOwnerComponent().getModel("AppModel").setProperty("/CostCenter", "");
            this.getOwnerComponent().getModel("AppModel").setProperty("/Company", "");

            sap.ui.core.BusyIndicator.show(0);
            await this.obtenerProjectID();
            // this.getView().getModel("AppModel").refresh(true);

            var oReader = new FileReader();

            oReader.onload = async function (e) {
                try {
                    // ── 1. Leer workbook ──────────────────────────────────
                    var aData = new Uint8Array(e.target.result);
                    var oWorkbook = XLSX.read(aData, {
                        type: "array",
                        password: "managemente",
                        cellDates: true   // <-- fechas como Date objects
                    });

                    console.log("Hojas disponibles:", oWorkbook.SheetNames);

                    var oWS = oWorkbook.Sheets["Datos Generales"]
                        || oWorkbook.Sheets[oWorkbook.SheetNames[2]];

                    if (!oWS) {
                        sap.m.MessageBox.error("No se encontró la hoja 'Datos Generales'");
                        sap.ui.core.BusyIndicator.hide();
                        return;
                    }

                    var oResource = oWorkbook.Sheets["Planificacion"]
                        || oWorkbook.Sheets[oWorkbook.SheetNames[3]];

                    var oBilling = oWorkbook.Sheets["Plan de Facturacion"]
                        || oWorkbook.Sheets[oWorkbook.SheetNames[7]];

                    // ── 2. Helper para leer celda de forma segura ─────────
                    var fnCell = function (sAddr) {
                        var oCell = oWS[sAddr];
                        if (!oCell) return "";
                        // Si es fecha, formatear a YYYY-MM-DD
                        if (oCell.t === "d" || oCell.v instanceof Date) {
                            var d = oCell.v instanceof Date ? oCell.v : new Date(oCell.v);
                            return d.toISOString().split("T")[0];
                        }
                        return oCell.v !== undefined ? String(oCell.v).trim() : "";
                    };

                    var fnCellResource = function (sAddr) {
                        var oCell = oResource[sAddr];
                        if (!oCell) return "";
                        // Si es fecha, formatear a YYYY-MM-DD
                        if (oCell.t === "d" || oCell.v instanceof Date) {
                            var d = oCell.v instanceof Date ? oCell.v : new Date(oCell.v);
                            return d.toISOString().split("T")[0];
                        }
                        return oCell.v !== undefined ? String(oCell.v).trim() : "";
                    };

                    var fnCellBilling = function (sAddr) {
                        var oCell = oBilling[sAddr];
                        if (!oCell) return "";
                        // Si es fecha, formatear a YYYY-MM-DD
                        if (oCell.t === "d" || oCell.v instanceof Date) {
                            var d = oCell.v instanceof Date ? oCell.v : new Date(oCell.v);
                            return d.toISOString().split("T")[0];
                        }
                        return oCell.v !== undefined ? String(oCell.v).trim() : "";
                    };



                    // ── 3. MAPEO DE CELDAS (ajusta direcciones según tu Excel) ──
                    //    Basado en la imagen:
                    //    Fila 6:  Moneda de HPP          → B6
                    //    Fila 7:  Nombre Proyecto         → H7  (celda después del label)
                    //    Fila 7:  Industria               → P7
                    //    Fila 8:  Centro de Competencia   → K8  (aprox)
                    //    Fila 9:  Tipo de Tecnología      → P9
                    //    Fila 10: Inicio                  → E10
                    //    Fila 11: Fin                     → E11
                    //    Fila 10: Gerente                 → H10 (aprox)
                    //    Fila 10: Geografia               → K10 (aprox)
                    //    Fila 10: Empresa que factura      → P10
                    //    etc.
                    //
                    //  ⚠️ IMPORTANTE: Abre el Excel, haz clic en cada dato visible
                    //     y ajusta las referencias de celda abajo con lo que muestre
                    //     el cuadro de nombre (esquina superior izquierda, ej: "H7")

                    var valores = {
                        tpProyectoInf: { valor: fnCell("F5"), codigo: "1" },
                        tpProyectoSup: { valor: fnCell("F6"), codigo: "2" },
                        tpProyectoCer: { valor: fnCell("F7"), codigo: "4" },
                        tpProyectoLin: { valor: fnCell("F8"), codigo: "3" }
                    };

                    var ProjectStage = Object.values(valores).find(k => k.valor.trim() == "X")?.codigo || "";



                    var oProjectData = {
                        // Datos generales
                        Currency: fnCell("B6").split("-")[0].trim(),   // "EUR - Euro"
                        ProjectName: fnCell("H7"),   // "(Cod. - CLiente ) IA Expansión - JP"
                        YY1_Geografia_Cpr: fnCell("L11"),   // Industria label row
                        Customer: fnCell("H5"),   // Cliente → ajustar

                        // Responsables
                        ProjManagerExtId: fnCell("H9"),   // Socio/Director Responsable → ajustar
                        CostCenter: fnCell("L9"),   // Centro de Competencia → ajustar
                        YY1_Tipodeproyecto_Cpr: fnCell("F5") == 'X' ? '1' : fnCell("F6") == 'X' ? '2' : fnCell("F7") == 'X' ? '4' : fnCell("F8") == 'X' ? '3' : '',   // Tipo de Tecnología → ajustar
                        ProjControllerExtId: fnCell("H11"),  // Gerente → ajustar
                        ProfitCenter: fnCell("P7"),
                        ProjPartnerExtId: fnCell("H9"),

                        // Fechas
                        StartDate: fnCell("E10"),  // "01.04.2026"
                        EndDate: fnCell("E11"),  // "31.10.2026"

                        // Empresa
                        OrgID: fnCell("P11"),  // Empresa que factura → ajustar

                        // Contacto cliente
                        Customer_Contact: fnCell("H13"),  // Persona de Contacto
                        Customer_RazonSocial: fnCell("H14"),  // Razón Social → "SERCOTEL"
                        Customer_Direccion: fnCell("H15"),  // Dirección
                        Customer_Telefono: fnCell("K15"),  // Teléfono → ajustar
                        Customer_Email: fnCell("P15"),  // email → ajustar

                        // CONSTANTES
                        ProjectStage: "P001",
                        ProjectCategory: "C",
                        Confidential: "N",
                        UseProjectBilling: "X",
                        RestrictTimePosting: "N",
                        YY1_ACTIVE_Cpr: "Y",
                        YY1_Producto_Cpr: "SW070",
                    };

                    var workPackage = [{
                        WorkPackageName: "Prueba",
                        StartDate: fnCell("E10"),  // "01.04.2026"
                        EndDate: fnCell("E11"),  // "31.10.2026"                       
                        YY1_TipodeproyectoSub_cpd: ProjectStage,
                    }];

                    var resourceDemand = [{
                        Version: "1",
                        WorkPackage: this.getOwnerComponent().getModel("AppModel").getProperty("/ProjectID") + ".1.1",  // Asignar al WBS creado (ej: "12345.1.1")
                        EngagementProjectResource: fnCellResource("H9"),   // Socio/Director Responsable → ajustar
                        EngagementProjectResourceType: "1",  // Recurso Interno
                        UnitOfMeasure: "H",
                        Quantity: fnCellResource("Q8"),
                        Currency: fnCellResource("B6"),
                    }];

                    let IPlanBilling = 7;
                    let BillingPlan = [];
                    debugger
                    while(fnCellBilling("D" + IPlanBilling) != '') {
                        const str = fnCellBilling("E" + IPlanBilling);
                        const [day, month, year] = str.split("/").map(Number);
                        const date = new Date(year, month - 1, day);

                        BillingPlan.push({
                        SalesOrderItem: IPlanBilling - 6,
                        BillingPlanItem: IPlanBilling - 6,
                        SalesOrder: "",  // Empresa que factura → ajustar
                        BillingPlanBillingDate: date,  // "01.04.2026"
                        BillingPlanRelatedBillgStatus: "A",  // No tengo este dato en el Excel de ejemplo
                        BillingPlanAmount: fnCellBilling("H" + IPlanBilling),
                        TransactionCurrency: fnCellBilling("B6"),
                        BillingPlanItemDescription: fnCellBilling("F" + IPlanBilling)
                    });

                    IPlanBilling++;
                    }

                    // ── 4. Log para verificar mapeo ───────────────────────
                    console.log("ProjectData extraído:", oProjectData);

                    // ── 5. DEBUG: Imprimir TODAS las celdas con valor ─────
                    //    (Útil para encontrar las direcciones exactas)
                    //    Descomenta esto la primera vez que pruebes:
                    /*
                    Object.keys(oWS).forEach(function(k) {
                        if (!k.startsWith("!")) {
                            console.log(k, "→", oWS[k].v);
                        }
                    });
                    */



                    // ── 6. Setear modelos ─────────────────────────────────
                    var oView = this.getView();
                    oView.setModel(new sap.ui.model.json.JSONModel(oProjectData), "ProjectSet"); 
                    oView.setModel(new sap.ui.model.json.JSONModel(workPackage), "WBSSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(resourceDemand), "ResourceSet");
                    oView.setModel(new sap.ui.model.json.JSONModel(BillingPlan), "BillingSet");

                    // Fire change para customer
                    setTimeout(() => {
                        let customerComboBox = oView.byId("customerComboBox");
                        customerComboBox.fireChange({ value: customerComboBox.getValue() });

                        if (customerComboBox.getSelectedKey()) {
                            let oldProjectName = this.getView().getModel("ProjectSet").getProperty("/ProjectName");
                            this.getView().getModel("ProjectSet").setProperty("/ProjectName", customerComboBox.getSelectedKey() + " - " + oldProjectName);
                            this.getView().getModel("ProjectSet").setProperty("/ProjectName", customerComboBox.getSelectedKey() + " - " + oldProjectName);
                            this.getView().getModel("ProjectSet").setProperty("/ProjectName", customerComboBox.getSelectedKey() + " - " + oldProjectName);
                        }
                    }, 500);

                    
                    await this.fixTextToCode(oView);
                    

                    oView.byId("masterPage").bindElement("ProjectSet>/");
                    oView.byId("detailPage").bindElement("ProjectSet>/");

                    var filter = "?$filter=NombreEmpresaFactura eq '" + oProjectData.OrgID + "' and ServiceOrgDefaultCostCenter eq '" + oProjectData.CostCenter + "'";
                    await this.obtenerDatosSociedad(filter);



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

        async fixTextToCode (oView) {
            // Convertir Geografia a code
            var sGeografia = oView.getModel("ProjectSet").getProperty("/YY1_Geografia_Cpr");
            // Buscar en odata /Oficinas
            var url = `/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Oficinas?$filter=SAP_Description eq '${sGeografia}'`;
            var response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json'
                }
            })
            debugger
            var data = await response.json();
            if (data.value && data.value.length > 0) {
                var sCodigoGeografia = data.value[0].IDOficina;
                oView.getModel("ProjectSet").setProperty("/YY1_Geografia_Cpr", sCodigoGeografia);
            }
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
            var billing = oView.getModel("BillingSet").getData();

            var projectId = this.getOwnerComponent().getModel("AppModel").getProperty("/ProjectID");
            var payloadHeader = this.setHeaderProyecto(header, projectId);
            var payloadWorkPackage = this.setWorkPackageProyecto(workPackage, projectId);
            var payloadResourceDemand = this.setResourceDemandProyecto(resourceDemand, projectId);

            var payloadApi = this.parsearPayloadApi(payloadHeader, payloadWorkPackage, payloadResourceDemand);
            debugger;
            var base64 = btoa(payloadApi);
            var oPayloadBase64 = {
                projectId: this.getOwnerComponent().getModel("AppModel").getProperty("/ProjectID"),
                body: base64,
                urlApi: '/sap/opu/odata/CPD/SC_PROJ_ENGMT_CREATE_UPD_SRV/ProjectSet',
                urlMet: '/sap/opu/odata/CPD/SC_EXTERNAL_SERVICES_SRV/$metadata',
                entidad: 'Project'
            };

            var sEndpoint = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Project";

            // ✅ Confirmación antes de ejecutar
            MessageBox.confirm("¿Esta Seguro De crear el proyecto?", {
                title: "Confirmar creación",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                onClose: async function (sAction) {

                    if (sAction !== MessageBox.Action.YES) return;

                    sap.ui.core.BusyIndicator.show(0);
                    var aMessages = [];

                    try {
                        var rpta = await this._postODataV4(sEndpoint, oPayloadBase64);

                        if (rpta.numericSeverity === 4) {
                            MessageBox.error(rpta.mensaje);
                        } else {

                            aMessages.push(new sap.m.MessageItem({
                                type: "Success",
                                title: "Proyecto creado correctamente",
                                description: "Proyecto " + payloadHeader.ProjectID + " creado correctamente"
                            }));

                            debugger;

                            var payloadBilling = this.planFacturacion(billing, projectId);
                            base64 = btoa(payloadBilling);
                            oPayloadBase64 = {
                                projectId: this.getOwnerComponent().getModel("AppModel").getProperty("/ProjectID"),
                                body: base64,
                                urlApi: '/sap/opu/odata/CPD/SC_PROJ_ENGMT_CREATE_UPD_SRV/A_CustProjSlsOrd',
                                urlMet: '/sap/opu/odata/CPD/SC_EXTERNAL_SERVICES_SRV/$metadata',
                                entidad: 'A_CustProjSlsOrd'
                            };

                            rpta = await this._postODataV4(sEndpoint, oPayloadBase64);
                            debugger;

                            if (rpta.numericSeverity === 4) {
                                aMessages.push(new sap.m.MessageItem({
                                    type: "Error",
                                    title: "Error en la creación del Pedido",
                                    description: rpta.mensaje
                                }));

                            } else {

                                aMessages.push(new sap.m.MessageItem({
                                    type: "Success",
                                    title: "Pedido creado correctamente",
                                    description: "Pedido " + rpta.mensaje + " creado correctamente"
                                }));

                            }

                            this.visualizarLog(aMessages, payloadHeader.ProjectID, rpta.mensaje);


                            //MessageBox.success(`Proyecto ${payloadHeader.ProjectID} creado exitosamente`);
                        }

                    } catch (oError) {
                        MessageBox.error("Error al crear el proyecto");

                    } finally {
                        sap.ui.core.BusyIndicator.hide();
                    }

                }.bind(this)
            });
        },

        setHeaderProyecto: function (oData, projectId) {

            var oProjectModel = this.getView().getModel("ProjectSet");

            var oPayload = {
                ProjectID: projectId,
                Confidential: oData.Confidential || "N",
                //CostCenter: oData.CostCenter || "",
                CostCenter: this.getOwnerComponent().getModel("AppModel").getProperty("/CostCenter"),
                Currency: oData.Currency.split("-")[0].trim(),
                Customer: oProjectModel.getProperty("/Customer").trim(),
                // Customer: oData.Customer || ""
                EndDate: oData.EndDate ? oData.EndDate + "T00:00:00" : null,
                //OrgID: oData.OrgID || "",
                OrgID: this.getOwnerComponent().getModel("AppModel").getProperty("/OrgId"),
                ProfitCenter: oData.ProfitCenter.split("(")[0].trim(),
                ProjAccountantCompCode: oData.ProjAccountantCompCode || "",
                ProjAccountantExtId: oData.ProjAccountantExtId || "",
                ProjControllerCompCode: oData.ProjControllerCompCode || "",
                //ProjControllerExtId: oData.ProjControllerExtId || "",
                ProjControllerExtId: "IKER.LARRANAGA",
                ProjManagerCompCode: oData.ProjManagerCompCode || "",
                //ProjManagerExtId: oData.ProjManagerExtId || "",
                ProjManagerExtId: "0005",
                ProjPartnerCompCode: oData.ProjPartnerCompCode || "",
                // ProjPartnerExtId: oData.ProjPartnerExtId || "",
                ProjPartnerExtId: "0004",
                ProjectCategory: oData.ProjectCategory || "",
                ProjectDesc: oData.ProjectName || "",
                ProjectName: oData.ProjectName || "",
                // ProjectStage: oData.ProjectStage || "",
                ProjectStage: "P001",
                RestrictTimePosting: oData.RestrictTimePosting || "N",
                StartDate: oData.StartDate ? oData.StartDate + "T00:00:00" : null,
                UseProjectBilling: oData.UseProjectBilling || "",
                YY1_ACTIVE_Cpr: oData.YY1_ACTIVE_Cpr || "",
                YY1_Fechadeventa_Cpr: oData.YY1_Fechadeventa_Cpr || null,
                YY1_Geografia_Cpr: oData.YY1_Geografia_Cpr || "",
                YY1_Producto_Cpr: oData.YY1_Producto_Cpr || "",
                YY1_Tipodeproyecto_Cpr: oData.YY1_Tipodeproyecto_Cpr || ""
            };

            return oPayload;

        },

        setWorkPackageProyecto: function (aWBSData, projectId) {

            var aWBSPayloads = [];



            for (var i = 0; i < aWBSData.length; i++) {
                var oItem = aWBSData[i];

                var oPayload = {
                    // EngagementProject: projectId,
                    WorkPackageName: oItem.WorkPackageName || "",
                    Description: oItem.Description || "",
                    WPStartDate: oItem.StartDate + "T00:00:00" || null,
                    WPEndDate: oItem.EndDate + "T00:00:00" || null,
                    WorkPackageType: oItem.WorkPackageType || "",
                    UnitQuantity: oItem.UnitQuantity || "0.000",
                    UnitId: oItem.UnitId || "",
                    YY1_TipodeproyectoSub_cpd: oItem.YY1_TipodeproyectoSub_cpd || ""
                };

                aWBSPayloads.push(oPayload);
            }

            return aWBSPayloads;

        },

        setResourceDemandProyecto: function (aResourceData, projectId) {

            var aResourcePayloads = [];
            var WorkPackage = projectId + ".1.1"

            for (var i = 0; i < aResourceData.length; i++) {
                var oItem = aResourceData[i];

                var oPayload = {
                    WorkPackage: WorkPackage,
                    BillingControlCategory: oItem.BillingControlCategory || "",
                    Currency: oItem.Currency.split("-")[0].trim(),
                    //DeliveryOrganization: oItem.DeliveryOrganization || "",
                    DeliveryOrganization: "EC05",
                    //  EngagementProject: sProjectID,
                    Country2DigitISOCode: "ES",
                    //EngagementProjectResource: oItem.EngagementProjectResource || "",     
                    EngagementProjectResource: "T001",
                    //EngagementProjectResourceType: oItem.EngagementProjectResourceType || "",
                    EngagementProjectResourceType: "0ACT",
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

        obtenerProjectID: async function () {

            var url = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/NewProject";

            // Para que sea "síncrono", lo metemos en una función async con await
            try {
                var response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json'
                    }
                });
                var data = await response.json();
                //console.log("Resultado del GET:", data[0].value);
            } catch (error) {
                // console.error("Fallo en la petición:", error);
            }

            var projectId = data?.value[0]?.ProjectID || "1";
            projectId = (BigInt(projectId.replace(/[^0-9]/g, "")) + 1n).toString();

            this.getOwnerComponent().getModel("AppModel").setProperty("/ProjectID", projectId);


            // return sProjectID;

        },

        obtenerOdata: async function (url) {
            // Para que sea "síncrono", lo metemos en una función async con await
            try {
                var response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json'
                    }
                });

                var data = await response.json();
                return data.value;
                //console.log("Resultado del GET:", data[0].value);
            } catch (error) {
                // console.error("Fallo en la petición:", error);
            }

        },


        obtenerDatosSociedad: async function (filter) {

            var url = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/DataHpp" + filter;

            // Para que sea "síncrono", lo metemos en una función async con await
            try {
                var response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json'
                    }
                });
                var data = await response.json();

                var empresa = data.value.find(i => i.campo === "EMPRESA")?.valores;
                var orgId = data.value.find(i => i.campo === "SOCIEDAD")?.valores;
                var ceco = data.value.find(i => i.campo === "CECO")?.valores;

                this.getOwnerComponent().getModel("AppModel").setProperty("/OrgId", orgId);
                this.getOwnerComponent().getModel("AppModel").setProperty("/CostCenter", ceco);
                this.getOwnerComponent().getModel("AppModel").setProperty("/Company", empresa);
                //console.log("Resultado del GET:", data[0].value);
            } catch (error) {
                // console.error("Fallo en la petición:", error);
            }
            // return sProjectID;

        },

        planFacturacion: function (billing, projectId) {
            debugger
            var oItemBilling = billing[0];

            var oBodyBilling = {
                CustomerProject: projectId,  // ProjectID del proyecto
                //PurchaseOrderByCustomer:    "",  // Referencia orden de compra cliente
                // CustomerPurchaseOrderDate:  "",  // Fecha OC formato "2019-01-28T00:00:00"
                PaymentMethod: "",  // Método de pago ej: "F"

                to_CustProjSlsOrdItem: [{
                    SalesOrderItemCategory: "PS01",  // Número item ej: "10"
                    ExpectedNetAmount: oItemBilling.BillingPlanAmount,  // Monto neto esperado
                    TransactionCurrency: oItemBilling.TransactionCurrency.split("-")[0].trim(),  // Moneda ej: "EUR"
                    SalesOrderItem: "1",
                    SalesOrderItemText: "",

                    to_CustProjSlsOrdItemWorkPckg: [{
                        WorkPackage: projectId + ".1.1"       // ProjectID.1.1
                    }],

                    to_CustProjSOIBillgPlnItm: [{
                        BillingPlanBillingDate: oItemBilling.BillingPlanBillingDate.split("/").reverse().join("-") + "T00:00:00",  // Fecha facturación "2019-01-28T00:00:00"
                        BillingPlanAmount: oItemBilling.BillingPlanAmount,  // Monto plan facturación
                        TransactionCurrency: oItemBilling.TransactionCurrency.split("-")[0].trim(), // Moneda ej: "EUR"
                        BillingPlanItemDescription: "",  // Descripción
                        BillingPlanItemUsage: ""   // Uso
                    }]
                }]
            };

            oBodyBilling = JSON.stringify(oBodyBilling, null, 2);
            return oBodyBilling;

        },

        visualizarLog: function (aMessages, project, salesOrder) {

            var oMessageView = new sap.m.MessageView({
                showDetailsPageHeader: true,
                items: aMessages
            });

            // ── Botón Back para cuando se navega al detalle del mensaje ──
            var oBackButton = new sap.m.Button({
                icon: "sap-icon://nav-back",
                visible: false,
                press: function () {
                    oMessageView.navigateBack();
                    oBackButton.setVisible(false);
                }
            });

            // ── Mostrar back solo cuando se navega al detalle ──
            oMessageView.attachItemSelect(function () {
                oBackButton.setVisible(true);
            });

            var oDialog = new sap.m.Dialog({
                resizable: true,
                content: oMessageView,
                title: "Resultado de la operación",
                contentHeight: "300px",
                contentWidth: "450px",
                verticalScrolling: false,
                state: "None",
                customHeader: new sap.m.Bar({
                    contentLeft: [oBackButton],
                    contentMiddle: [
                        new sap.m.Title({ text: "Resultado de la operación" })
                    ]
                }),
                buttons: [
                    new sap.m.Button({
                        text: "Visualizar Proyecto",
                        type: "Emphasized",
                        icon: "sap-icon://project-definition-triangle",
                        press: function () {
                            //oDialog.close();
                            sap.m.URLHelper.redirect(
                                "sap/bc/ui2/flp#CustomerProject-maintainCustomerProject&/Display/ProjEngagementsSet/" + project,
                                true
                            );
                        }
                    }),
                    new sap.m.Button({
                        text: "Visualizar Pedido",
                        type: "Default",
                        icon: "sap-icon://sales-order",
                        press: function () {
                            //oDialog.close();
                            sap.m.URLHelper.redirect(
                                "/sap/bc/ui2/flp#SalesOrder-displayFactSheet?SalesOrder=" + salesOrder,
                                true
                            );
                        }
                    }),
                    new sap.m.Button({
                        text: "Cerrar",
                        type: "Transparent",
                        press: function () { oDialog.close(); }
                    })
                ]
            });

            oDialog.open();
        },

        onCustomerValueHelp: async function () {
            if (!this._oCustomerDialog) {
                this._oCustomerDialog = sap.ui.xmlfragment(
                    "com.co.stratesys.zpscrearproyectos.view.HelpCustomer",
                    this
                );
                this.getView().addDependent(this._oCustomerDialog);
            }

            var aCustomers = await this.obtenerOdata("/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Clientes");

            const oModel = new sap.ui.model.json.JSONModel({
                customers: aCustomers
            });

            // ✅ Solo cambia esta línea
            this.getView().setModel(oModel, "CustomerSet");

            this._oCustomerDialog.open();

            var oBinding = this._oCustomerDialog.getBinding("items");
            if (oBinding) {
                oBinding.filter([]);
            }
        },

        onCustomerSearch: function (oEvent) {
            var sValue = oEvent.getParameter("value");
            var oFilter = new sap.ui.model.Filter({
                filters: [
                    new sap.ui.model.Filter("CustomerID", sap.ui.model.FilterOperator.Contains, sValue),
                    new sap.ui.model.Filter("CustomerName", sap.ui.model.FilterOperator.Contains, sValue)
                ],
                and: false  // OR entre ID y Nombre
            });
            oEvent.getSource().getBinding("items").filter([oFilter]);
        },

        onCustomerCancel: function () {
            this._oCustomerDialog.close();
        },

        onCustomerConfirm: function (oEvent) {
            const oSelected = oEvent.getParameter("selectedItem");

            if (oSelected) {
                const oContext = oSelected.getBindingContext("CustomerSet");
                const sPath = oContext.getPath(); // "/customers/3"
                const oModel = oContext.getModel();

                // Leer el objeto directamente por path
                const oCustomer = oModel.getProperty(sPath);
                console.log("Cliente:", oCustomer);

                var oProjectModel = this.getView().getModel("ProjectSet");
                if (oProjectModel && oCustomer) {
                    oProjectModel.setProperty("/Customer", oCustomer.Customer);
                }
            }
        },

        onCustomerLiveChangeCustomer: function (oEvent) {
            var sQuery = oEvent.getParameter("value");
            var oTableSelectDialog = oEvent.getSource();

            // 1. Si el usuario sigue escribiendo, reiniciamos el temporizador de 2 segundos
            if (this._searchTimeout) {
                clearTimeout(this._searchTimeout);
            }

            // 2. Creamos la espera de 2 segundos (2000ms) antes de disparar el fetch
            this._searchTimeout = setTimeout(async function () {

                // URL base de tu servicio (ajústala a tu necesidad, ej: "/sap/opu/odata/sap/ZSERVICIO_SRV/CustomerSet")
                var sBaseUrl = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Clientes";
                var sUrlConFiltro = sBaseUrl;

                // 3. Si hay texto, construimos el parámetro $filter de OData usando 'substringof' (OData V2)
                // Nota: Si usas OData V4, se usa: contains(Customer,'texto')


                if (sQuery && sQuery.length > 0) {
                    var sODataFilter =
                        //                  "contains(Customer, '" + sQuery + "') or " +
                        //                "contains(CustomerName, '" + sQuery + "') or " +
                        "contains(CustomerFullName, '" + sQuery + "')";

                    // Agregamos el ? $filter = de manera limpia y codificamos SOLO las funciones contains
                    sUrlConFiltro += "?$filter=" + encodeURIComponent(sODataFilter);
                }

                // 4. Encendemos el indicador de carga en el TableSelectDialog
                oTableSelectDialog.setBusy(true);

                try {
                    // 5. Llamamos a tu función actual para obtener los datos del backend
                    var aClientesFiltrados = await this.obtenerOdata(sUrlConFiltro);

                    // 6. Actualizamos el JSONModel que está amarrado al fragmento
                    // Nota: Asumo que tu modelo se llama 'CustomerSet'. Ajusta el nombre si es necesario.
                    var oModel = this.getView().getModel("CustomerSet");
                    if (!oModel) {
                        oModel = new sap.ui.model.json.JSONModel();
                        this.getView().setModel(oModel, "CustomerSet");
                    }

                    // Seteamos los nuevos datos filtrados que trajo el fetch
                    oModel.setProperty("/customers", aClientesFiltrados || []);

                } catch (oError) {
                    console.error("Error al filtrar OData: ", oError);
                } finally {
                    // 7. Apagamos el indicador de carga pase lo que pase
                    oTableSelectDialog.setBusy(false);
                }

            }.bind(this), 1000);

        }

    });
});