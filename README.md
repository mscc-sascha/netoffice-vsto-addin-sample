# netoffice-vsto-addin-sample
Sample to demonstrate an issue with VSTO Add-Ins and NetOffice Taskpanes

## Solution Structure

- **DocumentLevelAddIn**: A VSTO Add-In (Document Level) with a simple ribbon action.
- **NetOfficeSampleAddIn**: A TaskPane Add-In using the NetOffice library (abstract sample from [NetOfficeFw / Samples / 02 Ribbons And Panes](https://github.com/NetOfficeFw/Samples/tree/master/Word/02%20NetOffice%20Word%20COMAddin%20Sample/02%20Ribbons%20And%20Panes)).
- **WordStarter**: A Console Application that opens a Word document.

## Purpose

This sample demonstrates an issue with a VSTO Add-In in combination with a NetOffice TaskPane. When both are present and loaded simultaneously, Word becomes unresponsive, and loading the VSTO Add-In results in a timeout. This repository serves as a demonstration for a GitHub issue reported to the NetOffice project.

## Steps to Reproduce

1. Build the solution.
2. Register the TaskPane COM Add-In using the Developer Command Prompt:
    ```sh
    cd "C:\Path\To\Your\Project\NetOfficeSampleAddIn\bin\Debug\Debug"
    regasm /codebase Word02AddinCS4.dll
    ```
3. Add `C:\Path\To\Your\Project\` as a trusted location in Word (Word Options --> Trust Center, Trusted Locations --> Add new location).
4. Run the WordStarter console application.
    - Word opens and the DocumentLevelAddIn takes a very long time to load and blocks the Word document when opened!

The same behavior cannot be reproduced when using a Taskpane Add-In that was developed with a VSTO Taskpane Add-In.

[Screen Recording](./docs/Screenrecord2025-02-04%20103045.mp4)