sap.ui.define([
    "com/co/stratesys/zpscrearproyectos/controller/Funciones",
    'sap/ui/core/library'
], (Funciones, coreLibrary) => {
    "use strict";

    return Funciones.extend("com.co.stratesys.zpscrearproyectos.controller.Principal", {
        _searchTimeout: null,
        async onInit() {

            // Configurar tamaño inicial del splitter
            var oSplitter = this.byId("mainSplitter");
            var oMasterPage = this.byId("masterPage");

            oMasterPage.addStyleClass("sapUiSizeCompact");

            // Asignar SplitterLayoutData por código
            var oLayoutData = new sap.ui.layout.SplitterLayoutData({
                size: "380px",
                minSize: 200,
                resizable: true
            });
            oMasterPage.setLayoutData(oLayoutData);



            this.onOpenDialog();
        },

        onAfterRendering: async function () {
            await this.loadODataMasterData();
        },

        async loadODataMasterData() {
            // Load clientes
            await this.loadClientes();

            // Load Proj Managers and Partners
            await this.loadProjManagersAndPartner();

            // Load profit centers
            await this.loadProfitCenters();
        },
        async loadProfitCenters() {
            var url = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Industria";
            var response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json'
                }
            })

            var data = await response.json();

            var oProfitCenterModel = new sap.ui.model.json.JSONModel({
                ProfitCenterSet: data.value
            });

            let profitCenterComboBox = this.byId("profitCenterComboBox");
            profitCenterComboBox.setModel(oProfitCenterModel);
        },
        async loadClientes() {
            var url = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Clientes";
            var response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json'
                }
            })

            var data = await response.json();

            var oClientsModel = new sap.ui.model.json.JSONModel({
                ClientsSet: data.value
            });

            let customerComboBox = this.byId("customerComboBox");
            customerComboBox.setModel(oClientsModel);
        },
        async loadProjManagersAndPartner() {
            let url = "/sap/opu/odata4/sap/zsrv_project_entry/srvd/sap/zsrv_project_entry/0001/Empleados";
            let response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json'
                }
            })

            let data = await response.json();

            let oDataModelManager = new sap.ui.model.json.JSONModel({
                ProjectManagerSet: data.value
            });

            let oDataModelPartner = new sap.ui.model.json.JSONModel({
                ProjectPartnerSet: data.value
            });

            let oDataModelController = new sap.ui.model.json.JSONModel({
                ProjectControllerSet: data.value
            });

            let projectManagerComboBox = this.byId("projectManagerComboBox");
            projectManagerComboBox.setModel(oDataModelManager);

            let projectPartnerComboBox = this.byId("projectPartnerComboBox");
            projectPartnerComboBox.setModel(oDataModelPartner);

            let projectControllerComboBox = this.byId("projectControllerComboBox");
            projectControllerComboBox.setModel(oDataModelController);
        },

        onCustomerChange: function (oEvent) {
            var ValueState = coreLibrary.ValueState;

            var oValidatedComboBox = oEvent.getSource(),
                sSelectedKey = oValidatedComboBox.getSelectedKey(),
                oSelectedItem = oValidatedComboBox.getSelectedItem(),
                sValue = oValidatedComboBox.getValue();

            if (!sSelectedKey && sValue) {
                oValidatedComboBox.setValueState(ValueState.Error);
                oValidatedComboBox.setValueStateText("Por favor indique un cliente válido!");
            } else {
                oValidatedComboBox.setValueState(ValueState.None);
                this.getView().getModel("ProjectSet").setProperty("/CustomerID", sSelectedKey);
                this.getView().getModel("WBSSet").setProperty("/0/WorkPackageName", this.getView().getModel("ProjectSet").getProperty("/CustomerID") + " - " + this.getView().getModel("ProjectSet").getProperty("/ProjectName"));

                if (oSelectedItem.getModel().getProperty(oSelectedItem.getBindingContext().getPath()).Currency !== this.getView().getModel("ProjectSet").getProperty("/Currency")) {
                    oValidatedComboBox.setValueState(ValueState.Error);
                    oValidatedComboBox.setValueStateText("La moneda del cliente no coincide con la del proyecto!");
                } else {
                    oValidatedComboBox.setValueState(ValueState.None);
                }
            }


        },

        onProjManagerChange(oEvent) {
            var ValueState = coreLibrary.ValueState;

            var oValidatedComboBox = oEvent.getSource(),
                sSelectedKey = oValidatedComboBox.getSelectedKey(),
                sValue = oValidatedComboBox.getValue();

            if (!sSelectedKey && sValue) {
                oValidatedComboBox.setValueState(ValueState.Error);
                oValidatedComboBox.setValueStateText("Por favor indique un proj manager válido!");
            } else {
                oValidatedComboBox.setValueState(ValueState.None);
                this.getView().getModel("ProjectSet").setProperty("/ProjManagerExtId", sSelectedKey);
            }
        },

        onProjPartnerChange(oEvent) {
            var ValueState = coreLibrary.ValueState;

            var oValidatedComboBox = oEvent.getSource(),
                sSelectedKey = oValidatedComboBox.getSelectedKey(),
                sValue = oValidatedComboBox.getValue();

            if (!sSelectedKey && sValue) {
                oValidatedComboBox.setValueState(ValueState.Error);
                oValidatedComboBox.setValueStateText("Por favor indique un proj partner válido!");
            } else {
                oValidatedComboBox.setValueState(ValueState.None);
                this.getView().getModel("ProjectSet").setProperty("/ProjPartnerExtId", sSelectedKey);
            }
        },

        onProjectControllerChange(oEvent) {
            var ValueState = coreLibrary.ValueState;

            var oValidatedComboBox = oEvent.getSource(),
                sSelectedKey = oValidatedComboBox.getSelectedKey(),
                sValue = oValidatedComboBox.getValue();

            if (!sSelectedKey && sValue) {
                oValidatedComboBox.setValueState(ValueState.Error);
                oValidatedComboBox.setValueStateText("Por favor indique un proj controller válido!");
            } else {
                oValidatedComboBox.setValueState(ValueState.None);
                this.getView().getModel("ProjectSet").setProperty("/ProjControllerExtId", sSelectedKey);
            }
        },
        onProfitCenterChange(oEvent) {
            var ValueState = coreLibrary.ValueState;

            var oValidatedComboBox = oEvent.getSource(),
                sSelectedKey = oValidatedComboBox.getSelectedKey(),
                sValue = oValidatedComboBox.getValue();

            if (!sSelectedKey && sValue) {
                oValidatedComboBox.setValueState(ValueState.Error);
                oValidatedComboBox.setValueStateText("Por favor indique un profit center válido!");
            } else {
                oValidatedComboBox.setValueState(ValueState.None);
                this.getView().getModel("ProjectSet").setProperty("/ProfitCenter", sSelectedKey);
            }
        },
        formatCurrency: function (value, currency, locale) {
            if (value === null || value === undefined || isNaN(value)) {
                return "";
            }
            var sCurrency = currency || "COP";
            var sLocale = locale || "es-CO";
            return new Intl.NumberFormat(sLocale, {
                style: "currency",
                currency: sCurrency,
                minimumFractionDigits: 0,
                maximumFractionDigits: 2
            }).format(Number(value));
        },
    });
});