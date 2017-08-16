# Color me

This project is showing how you can use Office.js and React to build an Excel add-in.

## How to run

1. To run the add-in, you need side-load the add-in within the Excel application. The section below describes the way of side-loading of manifest file in different platforms.

    - On Windows, follow [this tutorial](https://dev.office.com/docs/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

    - On macOS, move the manifest file `office-add-in-react-manifest.xml` to the folder `/Users/{username}/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (if not exist, create one)

    - For Excel Online, use the upload my add-in button from the add-in command dialog to upload the manifest file. 

2. Run this in the terminal for a dev server.

    - Windows
    
        ```bash
        set HTTPS=true&&npm start
        ```
    
    -  macOS
    
        ```bash
        HTTPS=true npm start
        ```

3. Open Excel and click the Add-in to load.

<img width="1156" alt="screenshot" src="https://user-images.githubusercontent.com/3375461/27309424-87a21b06-5508-11e7-91c6-c9468efdbad0.png">

## How to create a new project by yourself

Follow the step by step tutorial [here](https://hongbo-miao.gitbooks.io/excel/content/quick-start/react.html).

## Learn more 

To learn more about JavaScript API for Office (Office.js), please check [here](https://dev.office.com/reference/add-ins/javascript-api-for-office).
