sap.ui.define([
    "com/co/stratesys/zpscrearproyectos/controller/Funciones"
], (Funciones) => {
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



    });
});