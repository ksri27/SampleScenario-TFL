// JavaScript source code for case ribbon
caseRibbon = {
    //check user access to show/hide resolve button
    checkUserAccess: function (context) {
        debugger;
        var hasAccess = false;
        //Get users role
        var userSettings = Xrm.Utility.getGlobalContext().userSettings;
        var userRoles = userSettings.roles;
        if (Object.keys(userRoles._collection).length > 0) {
            userRoles.forEach(function (item) {
                if (item.name.toLowerCase() === "customer service manager") {
                    hasAccess = true;
                }
            });
        }
        return hasAccess; 
    },
    escalateAction: function (context) {
        if (context.getAttribute("tfl_teamid").getValue() == null) {
            var textContent = "";
            if (context.getAttribute("tfl_isconfidential").getValue()) {
                textContent = "Please select Confidential Case Team in Team lookup";
            }
            else {
                textContent = "Please select Escalation Team in Team lookup";
            }
            var alertStrings = { confirmButtonLabel: "Yes", text: textContent, title: "Notification" };
            var alertOptions = { height: 120, width: 260 };
            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                function (success) {
                    console.log("Alert dialog closed");
                    context.getControl("tfl_teamid").setFocus();
                },
                function (error) {
                    console.log(error.message);
                }
            );

        }
        else {
            context.getAttribute("isescalated").setValue(true);
            context.data.save();
        }
    }
};
