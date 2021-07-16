export class MonacoEditorOptimizer {
  static currentTheme = "light";
  static sourceTab: Xrm.Controls.Tab;
  static formContext: Xrm.FormContext;
  static monacoContainer: HTMLElement;
  static moncaoEditorParentDocument: Document;
  static monacoEditorIframe: HTMLIFrameElement;

  /**
   * Obtain context, source tab, parent document and invokes optimizing logic
   * @param executionContext Execution Context
   */
  static optimizeEditor(executionContext: Xrm.Events.EventContext): void {
    this.formContext = executionContext.getFormContext();
    this.sourceTab = this.formContext.ui.tabs.get("SourceCode");
    this.moncaoEditorParentDocument = parent.document;
    this.setUpElementsAndCommandsAfterSectionDivRenders(this.createButtons());
  }

  /**
   * Creates the Expand, Contract, Theme Switcher, Save and Commit buttons
   * @returns Array of buttons added to Editor
   */
  static createButtons(): HTMLSpanElement[] {
    return [
      this.setUpCommitButton(),
      this.setUpThemeSwitcherButton(),
      this.setUpContractButton(),
      this.setUpExpandButton(),
      this.setUpSaveButton(),
    ];
  }

  /**
   * Executes logic to add buttons to Editor and sets up keyboard commands
   * @param buttonsArray Buttons to be added to Editor
   */
  static setUpElementsAndCommandsAfterSectionDivRenders(buttonsArray: HTMLSpanElement[]): void {
    const waitForRecordSpecificElementToRender = setInterval(function () {
      const formDivElement: HTMLElement =
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector('div[id^="tab-section"]');
      if (formDivElement) {
        formDivElement.style.cssText = "margin-top: -125px; padding-top: 125px";
        formDivElement.append(...buttonsArray);

        // Remove Commit and Save buttons for diff checker
        if (MonacoEditorOptimizer.formContext.data.entity.getEntityName() === "mal_sourcehistory") {
          formDivElement.removeChild(buttonsArray[0]);
          formDivElement.removeChild(buttonsArray[4]);
        }
        MonacoEditorOptimizer.toggleIcons();
        MonacoEditorOptimizer.sourceTab.addTabStateChange(() => MonacoEditorOptimizer.toggleIcons());
        MonacoEditorOptimizer.setUpEditorCommands();
        clearInterval(waitForRecordSpecificElementToRender);
      }
    }, 100);
  }

  /**
   * Sets up Keyboard commands for Saving, expanding/contracting editor, Committing and switching theme
   */
  static setUpEditorCommands(): void {
    const waitForMonacoContainer = setInterval(function () {
      MonacoEditorOptimizer.monacoEditorIframe =
        (MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById(
          "WebResource_adx_source_ace",
        ) as HTMLIFrameElement) ||
        (MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById(
          "WebResource_adx_registerstartupscript_editor",
        ) as HTMLIFrameElement);

      MonacoEditorOptimizer.monacoContainer =
        MonacoEditorOptimizer.monacoEditorIframe &&
        MonacoEditorOptimizer.monacoEditorIframe.contentDocument.getElementById("container");

      if (MonacoEditorOptimizer.monacoContainer) {
        MonacoEditorOptimizer.AddCommands();
        clearInterval(waitForMonacoContainer);
      }
    }, 100);
  }

  /**
   * Hides/shows buttons when source tab is expanded/collapsed
   */
  static toggleIcons(): void {
    if (MonacoEditorOptimizer.sourceTab.getDisplayState() === "collapsed") {
      this.moncaoEditorParentDocument.getElementById("spanExpandEditor").style.display = "none";
      this.moncaoEditorParentDocument.getElementById("spanThemeSwitcher").style.display = "none";
      if (this.moncaoEditorParentDocument.getElementById("spanSaveEditor")) {
        this.moncaoEditorParentDocument.getElementById("spanSaveEditor").style.display = "none";
      }
      if (this.moncaoEditorParentDocument.getElementById("spanCommitButton")) {
        this.moncaoEditorParentDocument.getElementById("spanCommitButton").style.display = "none";
      }
    } else if (MonacoEditorOptimizer.sourceTab.getDisplayState() === "expanded") {
      this.moncaoEditorParentDocument.getElementById("spanExpandEditor").style.display = "inline";
      this.moncaoEditorParentDocument.getElementById("spanThemeSwitcher").style.display = "inline";
      if (this.moncaoEditorParentDocument.getElementById("spanSaveEditor")) {
        this.moncaoEditorParentDocument.getElementById("spanSaveEditor").style.display = "inline";
      }
      if (this.moncaoEditorParentDocument.getElementById("spanCommitButton")) {
        this.moncaoEditorParentDocument.getElementById("spanCommitButton").style.display = "inline";
      }
    }
  }

  /**
   * Sets up keyboard commands
   */
  static AddCommands(): void {
    if (MonacoEditorOptimizer.sourceTab.getDisplayState() === "expanded") {
      [this.monacoContainer, this.moncaoEditorParentDocument].forEach((e) =>
        e.addEventListener(
          "keydown",
          function (e: KeyboardEvent) {
            if (navigator.platform.match("Mac") ? e.metaKey : e.ctrlKey) {
              // Save Command
              if (e.code === "KeyS") {
                e.preventDefault();
                (
                  MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector(
                    '[data-id="edit-form-save-btn"]',
                  ) as HTMLElement
                ).click();
              } else if (e.shiftKey) {
                // Expand/Contract Command
                if (e.code === "KeyE") {
                  e.preventDefault();
                  const expandButton =
                    MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanExpandEditor");
                  const contractButton =
                    MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanContractEditor");
                  if (expandButton.style.display == "inline") {
                    expandButton.click();
                  } else if (contractButton.style.display == "inline") {
                    contractButton.click();
                  }
                }
                if (e.code === "KeyF") {
                  // Switch Theme Command
                  e.preventDefault();
                  MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanThemeSwitcher").click();
                }
                if (e.code === "KeyC") {
                  // Commit Command
                  e.preventDefault();
                  MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanCommitButton").click();
                }
              }
            }
          },
          false,
        ),
      );
    }
  }

  /**
   * Adds a custom header to the form for title and new buttons
   */
  static addCustomHeader(): void {
    const customHeader = document.createElement("div");
    customHeader.textContent = this.moncaoEditorParentDocument.title
      .replace(" - Power Apps", "")
      .replace(": Information:", ":")
      .replace(": PSEO Form:", ":");
    customHeader.id = "customHeader";
    customHeader.style.cssText = `
            height: 33px;
            background-color: #fff;
            position: absolute;
            top: 0px;
            left: 0px;
            width: 100%;
            padding: 5px 0 0 25px;
            font-size: 18px;
            color: #333;
        `;

    MonacoEditorOptimizer.monacoEditorIframe.parentElement.prepend(customHeader);
  }

  /**
   * Switches colors of buttons and header according the chosen theme
   */
  static setHeaderColorsWhenExpandedOnDarkTheme(): void {
    this.moncaoEditorParentDocument.getElementById("spanThemeSwitcher").style.color = "#ccc";
    this.moncaoEditorParentDocument.getElementById("spanContractEditor").style.color = "#ccc";
    if (this.moncaoEditorParentDocument.getElementById("spanSaveEditor")) {
      this.moncaoEditorParentDocument.getElementById("spanSaveEditor").style.color = "#ccc";
    }
    if (this.moncaoEditorParentDocument.getElementById("spanCommitButtonSvg")) {
      this.moncaoEditorParentDocument.getElementById("spanCommitButtonSvg").setAttribute("fill", "#ccc");
    }
    this.moncaoEditorParentDocument.getElementById("customHeader").style.backgroundColor = "#000";
    this.moncaoEditorParentDocument.getElementById("customHeader").style.color = "#aaa";
    MonacoEditorOptimizer.monacoContainer.style.cssText = "border: 0";
  }

  /**
   * Creates Commit Button
   * @returns Commit button
   */
  static setUpCommitButton(): HTMLSpanElement {
    const commitButton = document.createElement("span");
    commitButton.id = "spanCommitButton";
    commitButton.title = "Commit (Ctrl + Shift + C)";

    commitButton.style.cssText = `
        position: absolute; 
        right: 105px; 
        top: 7px; 
        font-size: 18px; 
        display: none; 
        z-index: 90000;
        cursor: pointer;
      `;

    const commitSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    commitSvg.id = "spanCommitButtonSvg";
    commitSvg.setAttribute("width", "22");
    commitSvg.setAttribute("height", "22");
    commitSvg.setAttribute("viewBox", "0 0 24 24");
    commitSvg.setAttribute("fill", "#666");
    commitSvg.setAttribute("style", "margin-top: -2px");

    const commitSvgPath1 = document.createElementNS("http://www.w3.org/2000/svg", "path");

    commitSvgPath1.setAttribute("fill-rule", "evenodd");
    commitSvgPath1.setAttribute(
      "d",
      "M17.5 11.75a.75.75 0 01.75-.75h5a.75.75 0 010 1.5h-5a.75.75 0 01-.75-.75zm-17.5 0A.75.75 0 01.75 11h5a.75.75 0 010 1.5h-5a.75.75 0 01-.75-.75z",
    );

    const commitSvgPath2 = document.createElementNS("http://www.w3.org/2000/svg", "path");

    commitSvgPath2.setAttribute("fill-rule", "evenodd");
    commitSvgPath2.setAttribute("d", "M12 16.25a4.5 4.5 0 100-9 4.5 4.5 0 000 9zm0 1.5a6 6 0 100-12 6 6 0 000 12z");

    commitSvg.appendChild(commitSvgPath1);
    commitSvg.appendChild(commitSvgPath2);
    commitButton.appendChild(commitSvg);

    commitButton.addEventListener("click", function () {
      (
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelectorAll(
          'button[aria-label="Commit"]',
        )[0] as HTMLElement
      ).click();
    });

    return commitButton;
  }

  /**
   * Creates Theme Switcher button
   * @returns Theme Switcher button
   */
  static setUpThemeSwitcherButton(): HTMLSpanElement {
    const themeSwitcher = document.createElement("span");
    themeSwitcher.id = "spanThemeSwitcher";
    themeSwitcher.title = "Switch Theme (Ctrl + Shift + F)";
    themeSwitcher.classList.add("pa-ac", "EditSeries-symbol", "symbolFont");
    themeSwitcher.style.cssText = `
        position: absolute; 
        right: 55px; 
        top: 7px; 
        font-size: 18px; 
        display: none; 
        z-index: 90000;
        cursor: pointer;
      `;

    themeSwitcher.addEventListener("click", function () {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const monacoEditor = (MonacoEditorOptimizer.monacoEditorIframe.contentWindow as any).monaco;

      if (MonacoEditorOptimizer.currentTheme === "dark") {
        if (
          MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanExpandEditor").style.display === "none"
        ) {
          MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanThemeSwitcher").style.color = "#666";
          if (MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanSaveEditor")) {
            MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanSaveEditor").style.color = "#666";
          }
          MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanContractEditor").style.color = "#666";
          if (MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanCommitButtonSvg")) {
            MonacoEditorOptimizer.moncaoEditorParentDocument
              .getElementById("spanCommitButtonSvg")
              .setAttribute("fill", "#666");
          }
          MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("customHeader").style.backgroundColor =
            "#fff";
          MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("customHeader").style.color = "#333";

          MonacoEditorOptimizer.monacoContainer.style.cssText = "border: 1px solid #ddd";
        }

        monacoEditor.editor.setTheme("vs-light");
        MonacoEditorOptimizer.currentTheme = "light";
      } else {
        if (
          MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanExpandEditor").style.display === "none"
        ) {
          MonacoEditorOptimizer.setHeaderColorsWhenExpandedOnDarkTheme();
        }
        monacoEditor.editor.setTheme("vs-dark");
        MonacoEditorOptimizer.currentTheme = "dark";
      }
    });
    return themeSwitcher;
  }

  /**
   * Creates Contract button
   * @returns Contract button
   */
  static setUpContractButton(): HTMLSpanElement {
    const contractBtn = document.createElement("span");
    contractBtn.id = "spanContractEditor";
    contractBtn.title = "Contract (Ctrl + Shift + E)";
    contractBtn.classList.add("pa-ac", "MiniContract-symbol", "symbolFont");
    contractBtn.style.cssText = `
        position: absolute; 
        right: 30px; 
        top: 7px; 
        display: none; 
        font-size: 18px; 
        z-index: 90000;
        cursor: pointer;
      `;

    contractBtn.addEventListener("click", function () {
      MonacoEditorOptimizer.moncaoEditorParentDocument
        .querySelectorAll(
          `div#mainContent,
            [id^="tab-section"] div,
            [data-id="form-footer"],
            iframe#WebResource_adx_source_ace, 
            iframe#WebResource_adx_registerstartupscript_editor`,
        )
        .forEach((e) => e.removeAttribute("style"));

      MonacoEditorOptimizer.monacoEditorIframe.contentDocument.getElementById("container").style.cssText =
        "border: 1px solid #ddd";

      MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("customHeader").remove();

      (
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector('div[id^="tab-section"]') as HTMLElement
      ).style.cssText = "margin-top: -125px; padding-top: 125px";

      this.style.display = "none";
      this.style.color = "#666";
      MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanExpandEditor").style.display = "inline";
      MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanThemeSwitcher").style.color = "#666";
      if (MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanSaveEditor")) {
        MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanSaveEditor").style.color = "#666";
      }
      if (MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanCommitButtonSvg")) {
        MonacoEditorOptimizer.moncaoEditorParentDocument
          .getElementById("spanCommitButtonSvg")
          .setAttribute("fill", "#666");
      }
    });
    return contractBtn;
  }

  /**
   * Creates Expand Button
   * @returns Expand Button
   */
  static setUpExpandButton(): HTMLSpanElement {
    const expandBtn = document.createElement("span");
    expandBtn.id = "spanExpandEditor";
    expandBtn.title = "Expand (Ctrl + Shift + E)";
    expandBtn.classList.add("pa-ac", "ExpandButton-symbol", "symbolFont");
    expandBtn.style.cssText = `
            position: absolute; 
            right: 30px; 
            top: 7px; 
            display: none; 
            font-size: 18px; 
            z-index: 90000;
            cursor: pointer;
        `;

    expandBtn.addEventListener("click", function () {
      (MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector("div#mainContent") as HTMLElement).style.cssText =
        "position: absolute; width: 100%; top: 0";

      (
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector('[id^="tab-section"]') as HTMLElement
      ).style.cssText = "position: absolute; z-index: 90000; height: 100%; overflow: hidden";

      (
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector('[id^="tab-section"] div') as HTMLElement
      ).style.overflowY = "hidden";

      (
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector('[data-id="form-footer"]') as HTMLElement
      ).style.display = "none";

      const customHeader = MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("customHeader");
      if (!customHeader) {
        MonacoEditorOptimizer.addCustomHeader();
      } else {
        customHeader.style.display = "inline";
      }

      MonacoEditorOptimizer.monacoEditorIframe.style.cssText = `
          position: absolute; 
          top: 33px; 
          left: 0px; 
          height: calc(100% - 33px);
        `;

      this.style.display = "none";
      MonacoEditorOptimizer.moncaoEditorParentDocument.getElementById("spanContractEditor").style.display = "inline";

      if (MonacoEditorOptimizer.currentTheme === "dark") {
        MonacoEditorOptimizer.setHeaderColorsWhenExpandedOnDarkTheme();
      }
    });
    return expandBtn;
  }

  /**
   * Creates Save Button
   * @returns Save Button
   */
  static setUpSaveButton(): HTMLSpanElement {
    const saveBtn = document.createElement("span");
    saveBtn.id = "spanSaveEditor";
    saveBtn.title = "Save (Ctrl + S)";
    saveBtn.classList.add("pa-ac", "Save-symbol", "symbolFont");
    saveBtn.style.cssText = `
            position: absolute; 
            right: 80px; 
            top: 7px; 
            display: none; 
            font-size: 18px; 
            z-index: 90000;
            cursor: pointer;
        `;

    saveBtn.addEventListener("click", function () {
      (
        MonacoEditorOptimizer.moncaoEditorParentDocument.querySelector('[data-id="edit-form-save-btn"]') as HTMLElement
      ).click();
    });

    return saveBtn;
  }
}
