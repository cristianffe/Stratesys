sap.ui.define([
    "sap/ui/core/UIComponent",
    "com/co/stratesys/zpscrearproyectos/model/models"
], (UIComponent, models) => {
    "use strict";

    return UIComponent.extend("com.co.stratesys.zpscrearproyectos.Component", {
        metadata: {
            manifest: "json",
            interfaces: [
                "sap.ui.core.IAsyncContentCreation"
            ]
        },

        init() {
            // call the base component's init function
            UIComponent.prototype.init.apply(this, arguments);

            // Modelo para guardar ProjectID
            var oProjectIDModel = new sap.ui.model.json.JSONModel({
                ProjectID: "",
                OrgId: "",
                CostCenter: "",
                Company:"",
                CountryIso: "",
                enableBtnCreate: false,
                enableBtnPlantilla: false,
                ClientsSet: []
            });
            
            // set the device model
            this.setModel(models.createDeviceModel(), "device");

            // enable routing
            this.getRouter().initialize();

            this.setModel(oProjectIDModel, "AppModel");
        }
    });
});