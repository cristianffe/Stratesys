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
            // Fill Clientes
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

         onCustomerChange: function (oEvent) {
            var ValueState = coreLibrary.ValueState;

            var oValidatedComboBox = oEvent.getSource(),
                sSelectedKey = oValidatedComboBox.getSelectedKey(),
                sValue = oValidatedComboBox.getValue();

            if (!sSelectedKey && sValue) {
                oValidatedComboBox.setValueState(ValueState.Error);
                oValidatedComboBox.setValueStateText("Por favor indique un cliente válido!");
            } else {
                oValidatedComboBox.setValueState(ValueState.None);
                let oldProjectName = this.getView().getModel("ProjectSet").getProperty("/ProjectName");
                this.getView().getModel("ProjectSet").setProperty("/ProjectName", sSelectedKey + " - " + oldProjectName);
            }
        }
    });
});