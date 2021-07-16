/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
export class MonacoEditorCreator {
  static editorLoaded = false;

  /**
   * Obtains parameters passed through web resource and invokes logic to create editor
   * @param formContext Form Context object
   * @param Xrm Xrm object
   * @param monaco Monaco Object
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static async setUpEditor(formContext: Xrm.FormContext, Xrm: Xrm.XrmStatic, monaco: any): Promise<void> {
    if (this.editorLoaded) {
      return;
    }

    const urlQuery = new URL(window.document.location.href).searchParams;
    const dataQuery = new URLSearchParams(decodeURIComponent(urlQuery.get("data")));
    const sourceField = dataQuery.get("sourceField");
    const sourceTypeField = dataQuery.get("sourceTypeField");
    const sourceTypeAttribute: Xrm.Attributes.OptionSetAttribute = sourceTypeField
      ? formContext.getAttribute(sourceTypeField)
      : null;
    const sourceValue = formContext.getAttribute(sourceField).getValue();

    this.registerLiquidLanguage(monaco);
    this.registerAutocompleteProvider(monaco);

    if (formContext.data.entity.getEntityName() === "mal_sourcehistory") {
      const diffEditor = await this.createDiffEditor(formContext, Xrm, monaco, sourceValue);
      this.handleEditorResize(diffEditor);
    } else {
      const editor = this.createEditor(monaco, sourceValue, sourceTypeAttribute);
      this.handleLanguageSwitch(sourceTypeAttribute, monaco, editor);
      this.handleEditorResize(editor);
      this.handleSaving(formContext, sourceField, editor);
    }
    this.editorLoaded = true;
  }

  /**
   * Creates Diff Editor
   * @param formContext Form Context object
   * @param Xrm Xrm object
   * @param monaco Monaco object
   * @param sourceValue Source for editor
   * @returns Diff Editor created
   */
  static async createDiffEditor(
    formContext: Xrm.FormContext,
    Xrm: Xrm.XrmStatic,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    monaco: any,
    sourceValue: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ): Promise<any> {
    let targetSourceValue = "";
    let sourceType = "javascript";
    let selectQuery = "";
    let sourceField = "adx_source";

    const sourceRecord =
      formContext.getAttribute("mal_webtemplate").getValue() ||
      formContext.getAttribute("mal_entityform").getValue() ||
      formContext.getAttribute("mal_webformstep").getValue();

    if (sourceRecord[0].entityType != "adx_webtemplate") {
      sourceField = "adx_registerstartupscript";
      selectQuery = "?$select=" + sourceField;
    } else {
      selectQuery = "?$select=" + sourceField + ",mal_sourcetype";
    }

    const result = await Xrm.WebApi.retrieveRecord(sourceRecord[0].entityType, sourceRecord[0].id, selectQuery);
    targetSourceValue = result[sourceField];
    if (sourceRecord[0].entityType === "adx_webtemplate") {
      sourceType = result["mal_sourcetype@OData.Community.Display.V1.FormattedValue"] ?? "liquid";
    }

    const originalModel = monaco.editor.createModel(sourceValue, sourceType.toLocaleLowerCase());
    const modifiedModel = monaco.editor.createModel(targetSourceValue, sourceType.toLocaleLowerCase());

    const diffEditor = monaco.editor.createDiffEditor(document.getElementById("container"));
    diffEditor.setModel({
      original: originalModel,
      modified: modifiedModel,
    });

    diffEditor.getModifiedEditor().updateOptions({ readOnly: true });
    return diffEditor;
  }

  /**
   * Saves changes in editor code to source field
   * @param formContext Form Context object
   * @param sourceField Logical name of source field
   * @param editor Monaco editor
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static handleSaving(formContext: any, sourceField: string, editor: any): void {
    formContext.data.entity.addOnSave(() => formContext.getAttribute(sourceField).setValue(editor.getValue()));
  }

  /**
   * Resizes editor whenever its size changes to fit window
   * @param editor Monaco editor
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static handleEditorResize(editor: any): void {
    new ResizeObserver(() => editor.layout()).observe(document.getElementById("container"));
  }

  /**
   * Creates single editor
   * @param monaco Monaco object
   * @param sourceValue Value of source field
   * @param sourceTypeAttribute Source type for editor
   * @returns Editor created
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static createEditor(monaco: any, sourceValue: string, sourceTypeAttribute: Xrm.Attributes.OptionSetAttribute): any {
    return monaco.editor.create(document.getElementById("container"), {
      value: sourceValue,
      language: sourceTypeAttribute
        ? sourceTypeAttribute.getText()
          ? sourceTypeAttribute.getText().toLowerCase()
          : "liquid"
        : "javascript",
    });
  }

  /**
   * Sets the editor language based on selection made on language field for web templates
   * @param sourceTypeAttribute Source type for editor
   * @param monaco Monaco object
   * @param editor Monaco editor
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static handleLanguageSwitch(sourceTypeAttribute: Xrm.Attributes.OptionSetAttribute, monaco: any, editor: any): void {
    sourceTypeAttribute?.addOnChange(function () {
      const editorLanguage = sourceTypeAttribute.getText() ?? "liquid";
      monaco.editor.setModelLanguage(editor.getModel(), editorLanguage.toLowerCase());
    });
  }

  /**
   * Registers Liquid Language
   * @param monaco Monaco object
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static registerLiquidLanguage(monaco: any): void {
    // Register a new language
    monaco.languages.register({ id: "liquid" });

    monaco.languages.setLanguageConfiguration("liquid", {
      wordPattern: /(-?\d*\.\d\w*)|([^\`\~\!\@\$\^\&\*\(\)\=\+\[\{\]\}\\\|\;\:\'\"\,\.\<\>\/\s]+)/g,

      brackets: [
        ["{{-", "-}}"],
        ["{{", "}}"],
        ["{%-", "-%}"],
        ["{%", "%}"],
        ["{", "}"],
        ["[", "]"],
        ["(", ")"],
      ],
      autoClosingPairs: [
        { open: "{%-", close: "-%}" },
        { open: "{%", close: "%}" },
        { open: "{", close: "}" },
        { open: "[", close: "]" },
        { open: "(", close: ")" },
        { open: '"', close: '"' },
        { open: "'", close: "'" },
      ],

      surroundingPairs: [
        { open: "<", close: ">" },
        { open: '"', close: '"' },
        { open: "'", close: "'" },
      ],
    });

    // Register a tokens provider for the language
    monaco.languages.setMonarchTokensProvider("liquid", {
      defaultToken: "",
      tokenPostfix: "",
      // ignoreCase: true,
      keywords: [
        "assign",
        "capture",
        "endcapture",
        "increment",
        "decrement",
        "if",
        "else",
        "elsif",
        "endif",
        "for",
        "endfor",
        "break",
        "continue",
        "limit",
        "offset",
        "range",
        "reversed",
        "cols",
        "case",
        "endcase",
        "when",
        "block",
        "endblock",
        "true",
        "false",
        "in",
        "unless",
        "endunless",
        "cycle",
        "tablerow",
        "endtablerow",
        "contains",
        "startswith",
        "endswith",
        "comment",
        "endcomment",
        "raw",
        "endraw",
        "editable",
        "endentitylist",
        "endentityview",
        "endinclude",
        "endmarker",
        "entitylist",
        "entityview",
        "forloop",
        "image",
        "include",
        "marker",
        "outputcache",
        "plugin",
        "style",
        "text",
        "widget",
        "abs",
        "append",
        "at_least",
        "at_most",
        "capitalize",
        "ceil",
        "compact",
        "concat",
        "date",
        "default",
        "divided_by",
        "downcase",
        "escape",
        "escape_once",
        "first",
        "floor",
        "join",
        "last",
        "lstrip",
        "map",
        "minus",
        "modulo",
        "newline_to_br",
        "plus",
        "prepend",
        "remove",
        "remove_first",
        "replace",
        "replace_first",
        "reverse",
        "round",
        "rstrip",
        "size",
        "slice",
        "sort",
        "sort_natural",
        "split",
        "strip",
        "strip_html",
        "strip_newlines",
        "times",
        "truncate",
        "truncatewords",
        "uniq",
        "upcase",
        "url_decode",
        "url_encode",
      ],
      operators: ["==", ">", "<", "<=", ">=", "!=", "and", "or"],
      symbols: /[~!@#%\^&*-+=|\\:`<>.?\/]+/,
      exponent: /[eE][\-+]?[0-9]+/,
      // The main tokenizer for our languages
      tokenizer: {
        root: [
          // identifiers and keywords
          [/([a-zA-Z_\$][\w\$]*)(\s*)/, { cases: { "@default": "identifier" } }],
          [/<!--/, "comment.html", "@comment"],
          [/\{\%\s*comment\s*\%\}/, "comment.html", "@comment"],
          [/\{\{|\{\%/, { token: "@rematch", switchTo: "@liquidInSimpleState.root" }],
          [/(<)(\w+)(\/>)/, ["delimiter.html", "tag.html", "delimiter.html"]],
          [/(<)(script)/, ["delimiter.html", { token: "tag.html", next: "@script" }]],
          [/(<)(style)/, ["delimiter.html", { token: "tag.html", next: "@style" }]],
          [/(<)([:\w]+)/, ["delimiter.html", { token: "tag.html", next: "@liquidInHtmlAttribute" }]],
          [/(<\/)(\w+)/, ["delimiter.html", { token: "tag.html", next: "@liquidInHtmlAttribute" }]],
          [/</, "delimiter.html"],
          [/\{/, "delimiter.html"],
          { include: "numbers" },
          [/[ \t\r\n]+/], // whitespace
          [/[^<{]+/], // text
        ],
        // After <script
        script: [
          [/\{\{\-?|\{\%\-?/, { token: "@rematch", switchTo: "@liquidInSimpleState.script" }],
          [/type/, "attribute.name", "@scriptAfterType"],
          [/"([^"]*)"/, "attribute.value"],
          [/'([^']*)'/, "attribute.value"],
          [/[\w\-]+/, "attribute.name"],
          [/=/, "delimiter"],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@scriptEmbedded.text/javascript",
              nextEmbedded: "text/javascript",
            },
          ],
          [/[ \t\r\n]+/], // whitespace
          [/(<\/)(script\s*)(>)/, ["delimiter.html", "tag.html", { token: "delimiter.html", next: "@pop" }]],
        ],

        // After <script ... type
        scriptAfterType: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInSimpleState.scriptAfterType",
            },
          ],
          [/=/, "delimiter", "@scriptAfterTypeEquals"],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@scriptEmbedded.text/javascript",
              nextEmbedded: "text/javascript",
            },
          ], // cover invalid e.g. <script type>
          [/[ \t\r\n]+/], // whitespace
          [/<\/script\s*>/, { token: "@rematch", next: "@pop" }],
        ],

        // After <script ... type =
        scriptAfterTypeEquals: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInSimpleState.scriptAfterTypeEquals",
            },
          ],
          [
            /"([^"]*)"/,
            {
              token: "attribute.value",
              switchTo: "@scriptWithCustomType.$1",
            },
          ],
          [
            /'([^']*)'/,
            {
              token: "attribute.value",
              switchTo: "@scriptWithCustomType.$1",
            },
          ],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@scriptEmbedded.text/javascript",
              nextEmbedded: "text/javascript",
            },
          ], // cover invalid e.g. <script type=>
          [/[ \t\r\n]+/], // whitespace
          [/<\/script\s*>/, { token: "@rematch", next: "@pop" }],
        ],

        // After <script ... type = $S2
        scriptWithCustomType: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInSimpleState.scriptWithCustomType.$S2",
            },
          ],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@scriptEmbedded.$S2",
              nextEmbedded: "$S2",
            },
          ],
          [/"([^"]*)"/, "attribute.value"],
          [/'([^']*)'/, "attribute.value"],
          [/[\w\-]+/, "attribute.name"],
          [/=/, "delimiter"],
          [/[ \t\r\n]+/], // whitespace
          [/<\/script\s*>/, { token: "@rematch", next: "@pop" }],
        ],

        scriptEmbedded: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInEmbeddedState.scriptEmbedded.$S2",
              nextEmbedded: "@pop",
            },
          ],
          [/<\/script/, { token: "@rematch", next: "@pop", nextEmbedded: "@pop" }],
        ],

        // -- END <script> tags handling
        // -- BEGIN <style> tags handling

        // After <style
        style: [
          [/\{\{\-?|\{\%\-?/, { token: "@rematch", switchTo: "@liquidInSimpleState.style" }],
          [/type/, "attribute.name", "@styleAfterType"],
          [/"([^"]*)"/, "attribute.value"],
          [/'([^']*)'/, "attribute.value"],
          [/[\w\-]+/, "attribute.name"],
          [/=/, "delimiter"],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@styleEmbedded.text/css",
              nextEmbedded: "text/css",
            },
          ],
          [/[ \t\r\n]+/], // whitespace
          [/(<\/)(style\s*)(>)/, ["delimiter.html", "tag.html", { token: "delimiter.html", next: "@pop" }]],
        ],

        // After <style ... type
        styleAfterType: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInSimpleState.styleAfterType",
            },
          ],
          [/=/, "delimiter", "@styleAfterTypeEquals"],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@styleEmbedded.text/css",
              nextEmbedded: "text/css",
            },
          ], // cover invalid e.g. <style type>
          [/[ \t\r\n]+/], // whitespace
          [/<\/style\s*>/, { token: "@rematch", next: "@pop" }],
        ],

        // After <style ... type =
        styleAfterTypeEquals: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInSimpleState.styleAfterTypeEquals",
            },
          ],
          [
            /"([^"]*)"/,
            {
              token: "attribute.value",
              switchTo: "@styleWithCustomType.$1",
            },
          ],
          [
            /'([^']*)'/,
            {
              token: "attribute.value",
              switchTo: "@styleWithCustomType.$1",
            },
          ],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@styleEmbedded.text/css",
              nextEmbedded: "text/css",
            },
          ], // cover invalid e.g. <style type=>
          [/[ \t\r\n]+/], // whitespace
          [/<\/style\s*>/, { token: "@rematch", next: "@pop" }],
        ],

        // After <style ... type = $S2
        styleWithCustomType: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInSimpleState.styleWithCustomType.$S2",
            },
          ],
          [
            />/,
            {
              token: "delimiter.html",
              next: "@styleEmbedded.$S2",
              nextEmbedded: "$S2",
            },
          ],
          [/"([^"]*)"/, "attribute.value"],
          [/'([^']*)'/, "attribute.value"],
          [/[\w\-]+/, "attribute.name"],
          [/=/, "delimiter"],
          [/[ \t\r\n]+/], // whitespace
          [/<\/style\s*>/, { token: "@rematch", next: "@pop" }],
        ],

        styleEmbedded: [
          [
            /\{\{\-?|\{\%\-?/,
            {
              token: "@rematch",
              switchTo: "@liquidInEmbeddedState.styleEmbedded.$S2",
              nextEmbedded: "@pop",
            },
          ],
          [/<\/style/, { token: "@rematch", next: "@pop", nextEmbedded: "@pop" }],
        ],

        // -- END <style> tags handling

        comment: [
          [/\{\%\s*endcomment\s*\%\}/, "comment.hmtl", "@pop"],
          [/-->/, "comment.html", "@pop"],
          [/\b(?!-->|endcomment\s*\%\}\b)\w+/, "comment.content.html"],
          [/./, "comment.content.html"],
        ],
        // Support inside html tags
        liquidInHtmlAttribute: [
          [
            /"\{\{\-?|"\{\%\-?/,
            {
              token: "delimiter.liquid",
              next: "@liquidInHtmlAttributeEmbedded",
            },
          ],
          [/\/?>/, "delimiter.html", "@pop"],
          [/"([^"]*)"/, "attribute.value"],
          [/'([^']*)'/, "attribute.value"],
          [/[\w\-]+/, "attribute.name"],
          [/=/, "delimiter"],
          [/[ \t\r\n]+/], // whitespace
        ],
        liquidInHtmlAttributeEmbedded: [
          [/\'(.*)\'/],
          [/\-?\%\}"|\-?\}\}"/, "delimiter.liquid", "@pop"],
          [
            /([a-zA-Z_\$][\w\$]*)/,
            {
              cases: {
                "@keywords": "keyword",
                "@operators": "operator",
                "@default": "variable.parameter.liquid",
              },
            },
          ],
        ],
        liquidInSimpleState: [
          [/\{\{\-?|\{\%\-?/, "delimiter.liquid"],
          [/\-?\}\}|\-?\%\}/, { token: "delimiter.liquid", switchTo: "@$S2.$S3" }],
          { include: "liquidRoot" },
        ],
        liquidInEmbeddedState: [
          [/\{\{\-?|\{\%\-?/, "delimiter.liquid"],
          [
            /\-?\}\}|\-?\%\}/,
            {
              token: "delimiter.liquid",
              switchTo: "@$S2.$S3",
              nextEmbedded: "$S3",
            },
          ],
          { include: "liquidRoot" },
        ],
        liquidRoot: [
          [/\'(.*?)\'|\"(.*?)\"/],
          [
            /([a-zA-Z_\$][\w\$]*)/,
            {
              cases: {
                "@keywords": "keyword",
                "@operators": "operator",
                "@default": "variable.parameter.liquid",
              },
            },
          ],
          { include: "numbers" },
          [
            /@symbols/,
            {
              cases: {
                "@operators": "operator",
                "@default": "",
              },
            },
          ],
        ],
        numbers: [
          // numbers
          [/\d+\.\d*(@exponent)?/, "number.float"],
          [/\.\d+(@exponent)?/, "number.float"],
          [/\d+@exponent/, "number.float"],
          [/\d+/, "number"],
          [/[;,.]/, "delimiter"],
        ],
      },
    });
  }

  /**
   * Registers auto complete provider
   * @param monaco Monaco object
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static registerAutocompleteProvider(monaco: any): void {
    monaco.languages.registerCompletionItemProvider("liquid", {
      provideCompletionItems: function () {
        const autocompleteProviderItems = [];
        const keywords = [
          "assign",
          "capture",
          "endcapture",
          "increment",
          "decrement",
          "if",
          "else",
          "elsif",
          "endif",
          "for",
          "endfor",
          "break",
          "continue",
          "limit",
          "offset",
          "range",
          "reversed",
          "cols",
          "case",
          "endcase",
          "when",
          "block",
          "endblock",
          "true",
          "false",
          "in",
          "unless",
          "endunless",
          "cycle",
          "tablerow",
          "endtablerow",
          "contains",
          "startswith",
          "endswith",
          "comment",
          "endcomment",
          "raw",
          "endraw",
          "editable",
          "endentitylist",
          "endentityview",
          "endinclude",
          "endmarker",
          "entitylist",
          "entityview",
          "forloop",
          "image",
          "include",
          "marker",
          "outputcache",
          "plugin",
          "style",
          "text",
          "widget",
          "abs",
          "append",
          "at_least",
          "at_most",
          "capitalize",
          "ceil",
          "compact",
          "concat",
          "date",
          "default",
          "divided_by",
          "downcase",
          "escape",
          "escape_once",
          "first",
          "floor",
          "join",
          "last",
          "lstrip",
          "map",
          "minus",
          "modulo",
          "newline_to_br",
          "plus",
          "prepend",
          "remove",
          "remove_first",
          "replace",
          "replace_first",
          "reverse",
          "round",
          "rstrip",
          "size",
          "slice",
          "sort",
          "sort_natural",
          "split",
          "strip",
          "strip_html",
          "strip_newlines",
          "times",
          "truncate",
          "truncatewords",
          "uniq",
          "upcase",
          "url_decode",
          "url_encode",
        ];

        for (let i = 0; i < keywords.length; i++) {
          autocompleteProviderItems.push({
            label: keywords[i],
            kind: monaco.languages.CompletionItemKind.Keyword,
          });
        }

        return autocompleteProviderItems;
      },
    });
  }

  /**
   * Gets the user Language code for the editor
   * @returns User Language code
   */
  static getLanguageCode(): string {
    const localeId = parent.window.Xrm.Utility.getGlobalContext().userSettings.languageId;
    const defaultCode = "1033";
    const mapping = {
      1026: "bg",
      1027: "ca",
      1029: "cs",
      1030: "da",
      1031: "de",
      1032: "el",
      1033: "",
      3082: "es",
      1061: "et",
      1069: "eu",
      1035: "fi",
      1036: "fr",
      1110: "gl",
      1081: "hi",
      1050: "hr",
      1038: "hu",
      1057: "id",
      1040: "it",
      1041: "ja",
      1087: "kk",
      1042: "ko",
      1063: "lt",
      1062: "lv",
      1086: "ms",
      1044: "nb",
      1043: "nl",
      1045: "pl",
      2070: "pt",
      1046: "pt-br",
      1048: "ro",
      1049: "ru",
      1051: "sk",
      1060: "sl",
      3098: "sr",
      2074: "sr-latn",
      1053: "sv",
      1054: "th",
      1055: "tr",
      1058: "uk",
      1066: "vi",
      2052: "zh-cn",
      3076: "zh-hk",
      1028: "zh-tw",
    };
    return localeId in mapping ? mapping[localeId] : mapping[defaultCode];
  }
}
