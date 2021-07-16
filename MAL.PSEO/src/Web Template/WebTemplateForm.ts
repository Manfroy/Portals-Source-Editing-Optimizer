export class WebTemplateForm {
  /**
   * Automatically switches to the source tab if the general settings have already been set
   * @param executionContext Execution Context
   */
  static switchToSourceIfSettingsAreSet(executionContext: Xrm.Events.EventContext): void {
    const formContext = executionContext.getFormContext();
    const webTemplateName = formContext.getAttribute("adx_name").getValue();
    const webTemplateWebiste = formContext.getAttribute("adx_websiteid").getValue();

    if (webTemplateName && webTemplateWebiste) {
      formContext.ui.tabs.get("SourceCode").setFocus();
    }
  }

  /**
   * Checks if the Page Components Navigation solution is installed and hides the Enabled for PCN field if so
   * @param executionContext Execution Context
   */
  static async checkIfPCNInstalled(executionContext: Xrm.Events.EventContext): Promise<void> {
    const formContext = executionContext.getFormContext();
    try {
      await Xrm.WebApi.retrieveRecord("adx_sitesetting", "19f848a3-2b6b-442d-9204-ccd00c0ad172", "?$select=adx_name");
    } catch {
      formContext.getAttribute("mal_enabledforpcn").setRequiredLevel("none");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (<any>formContext.getControl("mal_enabledforpcn")).setVisible(false);
    }
  }
}
