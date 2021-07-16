export class Utils {
  /**
   * Called on form load to pass the form context to the monaco web resource
   * @param executionContext Execution Context
   */
  static async passContextToWebResource(executionContext: Xrm.Events.EventContext): Promise<void> {
    const formContext = executionContext.getFormContext();
    const monacoWebResource =
      formContext.getControl("WebResource_adx_source_ace") ||
      formContext.getControl("WebResource_adx_registerstartupscript_editor");

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const contentWindow = await (monacoWebResource as any).getContentWindow();
    const waitForWebResourceToRender = setInterval(function () {
      if (typeof contentWindow.setClientApiContext === "function") {
        contentWindow.setClientApiContext(Xrm, formContext);
        clearInterval(waitForWebResourceToRender);
      }
    }, 100);
  }

  /**
   * Executes logic that creates source history record for entity forms, web form steps and web templates.
   * @param formContext Form context
   * @param recordId Id of the record for which source history is being created
   * @param entityLogicalName Logical name of the entity
   */
  static async commitSource(formContext: Xrm.FormContext, recordId: string, entityLogicalName: string): Promise<void> {
    try {
      const recordName = formContext.getAttribute("adx_name").getValue();

      const metadataRequestPromise = await Xrm.Utility.getEntityMetadata(entityLogicalName);

      const entityDisplayName = metadataRequestPromise.DisplayName;

      const recordSourceAttribute =
        formContext.getAttribute("adx_source") || formContext.getAttribute("adx_registerstartupscript");

      let webFormStepLookup: string, entityFormLookup: string, webtemplateLookup: string;

      switch (entityLogicalName) {
        case "adx_entityform": {
          entityFormLookup = `/adx_entityforms(${recordId.replace("{", "").replace("}", "")})`;
          break;
        }
        case "adx_webformstep": {
          webFormStepLookup = `/adx_webformsteps(${recordId.replace("{", "").replace("}", "")})`;
          break;
        }
        case "adx_webtemplate": {
          webtemplateLookup = `/adx_webtemplates(${recordId.replace("{", "").replace("}", "")})`;
          break;
        }
      }

      const commitMessage = window.prompt("Enter a message");
      if (commitMessage) {
        const historyName = `${entityDisplayName}: ${recordName} - ${new Date().toLocaleString()}`;

        const data = {
          mal_name: historyName,
          mal_source: recordSourceAttribute.getValue(),
          mal_message: commitMessage,
          "mal_WebFormStep@odata.bind": webFormStepLookup,
          "mal_EntityForm@odata.bind": entityFormLookup,
          "mal_WebTemplate@odata.bind": webtemplateLookup,
        };

        Xrm.WebApi.createRecord("mal_sourcehistory", data).then(
          () => {
            Xrm.Navigation.openAlertDialog({ text: "Commit executed successfuly." }, { width: 330, height: 126 });
          },
          (error) => {
            Xrm.Navigation.openAlertDialog(
              { text: `Commit could not be created due to the following error: ${error}.` },
              { width: 330, height: 126 },
            );
          },
        );
      } else {
        Xrm.Navigation.openAlertDialog(
          { text: "You need to enter a message to be able to commit" },
          { width: 330, height: 126 },
        );
      }
    } catch (e) {
      Xrm.Navigation.openErrorDialog(e);
    }
  }
}
