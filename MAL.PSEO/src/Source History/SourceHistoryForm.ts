export class SourceHistoryForm {
  /**
   * Shows the relevant lookup on the history record form
   * @param executionContext Execution Context
   */
  static showRelevantSourceLookup(executionContext: Xrm.Events.EventContext): void {
    const formContext = executionContext.getFormContext();
    if (formContext.getAttribute("mal_webtemplate").getValue()) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (<any>formContext.getControl("mal_webtemplate")).setVisible(true);
    } else if (formContext.getAttribute("mal_entityform").getValue()) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (<any>formContext.getControl("mal_entityform")).setVisible(true);
    } else if (formContext.getAttribute("mal_webformstep").getValue()) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (<any>formContext.getControl("mal_webformstep")).setVisible(true);
    }
  }
}
