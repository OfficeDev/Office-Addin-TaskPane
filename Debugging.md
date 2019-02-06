# Debugging

## Using Visual Studio Code

1. Open the folder in VS Code.
2. Run the `Watch` task using `Terminal`, `Run Task`.
3. Run the `Dev Server` task using `Terminal`, `Run Task`.

### Office Online (Edge - Windows 10)
1. Switch to the Debug view using `View`, `Debug` or press Ctrl+Shift+D.
2. Choose the `Office Online (Edge)` debug configuration.
3. Start debugging by pressing F5 or the green play icon.
4. When prompted, paste the share url for an Office document. You can obtain this by copying the link when sharing a document.
5. Use `Insert`, `Add-ins` to upload the manifest file. [More info](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing_)

### Office Online (Chrome - Mac)
1. Switch to the Debug view using `View`, `Debug` or press Ctrl+Shift+D.
2. Choose the `Office Online (Chrome)` debug configuration.
3. Start debugging by pressing F5 or the green play icon.
4. When prompted, paste the share url for an Office document. You can obtain this by copying the link when sharing a document.
5. Use `Insert`, `Add-ins` to upload the manifest file. [More info](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing_)


### Excel / PowerPoint / Word (Windows / Mac)
1. Switch to the Debug view.
2. Choose the desired debug configuration from the list: 
   * `Excel Desktop`
   * `PowerPoint Desktop`
   * `Word Desktop`
3. Choose `Start Without Debugging` from the `Debug` menu.
   NOTE: The integrated VSCode debugger cannot debug the Office Add-in running in the task pane. 
4. To debug, you need to use another debugger:
    * `Edge DevTools` (Windows 10)
    * `F12 DevTools` (Windows)
    * `Visual Studio` (Windows)
    * `Safari Inspector` (Mac)



