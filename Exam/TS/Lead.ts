namespace Exam.TS {

    let Form: D365.Sdk.FormContext;
    export async function OnLoad(executionContext: D365.Sdk.ExecutionContext) {
        Form = executionContext.getFormContext() as D365.Sdk.FormContext;
        TeamOnChange();
    }

    export async function TeamOnChange() {
        let Team = Form.getControl("cr260_teamid").getAttribute().getValue();

        if (Team) {
            if (Team[0].entityType == "cr260_skill") {
                const confirmStrings = {
                    text: "You can not select a skill for this field. Please select a record from Teams Entity!" ,
                    title: `Warning!` ,
                };

                const confirmOptions = {
                    height: 200,
                    width: 400,
                };
                Helper.CRM.getXrm().Navigation.openAlertDialog(confirmStrings, confirmOptions);
            }    
        }
    }
}