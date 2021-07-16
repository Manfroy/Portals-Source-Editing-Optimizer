# Portals-Source-Editing-Optimizer

A tool for Power Apps Portals that aims to improve the coding experience for entity/basic forms, web/advanced form steps and web templates.

## Getting Started

### Setup

[Download](https://github.com/Manfroy/Portals-Source-Editing-Optimizer/releases/download/v1.0.0.0/PortalsSourceEditingOptimizer_1_0_0_0_managed.zip) the Portals Source Editing Optimizer solution and install it.

Once the solution is installed, make sure that when you open an entity/basic form, web/advanced form step or web template, you switch to the PSEO Form if it isn't already selected.

### Features

- **Source History**: Once a record/row is committed, a source history record/row that's accessible from the Source History tab of the PSEO form is created. Opening the source history record/row will show a side by side diff editor comparing the commited version of the code with the current version.

- **Command based saving**: It is possible to save changes to code with Ctrl + S.

- **Language Option for Web Templates**: It is possible to choose the language that should be applied to the web template editor which will enable intellisense support for the selected language. The options are: JavaScript, CSS, HTML and Liquid. OOTB, web templates only have intellisense support for the Liquid language. 

- **Theme Switching**: Ability to switch between default and dark theme.

- **Full Screen Editor**: Ability to expand the editor to fit the whole browser screen.

- **Tab for Source Only**: OOTB, for entity/basic forms and web/advanced form steps, you need to switch to a tab and then scroll down to get to the javascript section. You now have a source tab where only the editor shows. For web templates, the source tab is automatically shown on form load if the general settings have already been entered.


### Usage

Once you are on the Source tab of PSEO form for a record/row of any of the aforementioned entities/tables, four icons will show un the upper right corner of the form. They show in the following order from left to right:

- **Commit**: Commits the source code.

- **Save**: Saves changes to the record/row.

- **Switch Theme**: Switches the theme between default and dark themes.

- **Expand**: Expands the editor to fit the whole browser window. Only shows when editor has original size.

- **Contract**: Contracts the editor to original size. Only shows when editor is expanded.


### Shortcuts

- **Ctrl + S**: Save
- **Ctrl + Shift + C**: Commit
- **Ctrl + Shift + F**: Switch Theme
- **Ctrl + Shift + E**: Expand/Contract
