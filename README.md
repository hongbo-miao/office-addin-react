# Color me

This project is showing how you can use Office.js and React to build an Excel add-in.

## How to run

1. To run the add-in, you need side-load the add-in within the Excel application. The section below describes the way of side-loading of manifest file in different platforms.

    - On Windows, follow [this tutorial](https://dev.office.com/docs/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

    - On macOS, move the manifest file `office-add-in-react-manifest.xml` to the folder `/Users/{username}/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (if not exist, create one)

    - For Excel Online, use the upload my add-in button from the add-in command dialog to upload the manifest file. 

2. Run `yarn start` or `npm start` in the terminal for a dev server.

3. Open Excel and click the Add-in to load.

<img width="1156" alt="screenshot" src="https://user-images.githubusercontent.com/3375461/27309424-87a21b06-5508-11e7-91c6-c9468efdbad0.png">

## How to create a new project by yourself

1. Generate the React project using [Create React App](https://github.com/facebookincubator/create-react-app).

2. Generate the manifest file using [YO Office](https://github.com/OfficeDev/generator-office). When you generate, choose only generating the manifest file.
Then replace all the url in the generated manifest file from `https://localhost:3000` to `http://localhost:3000`. (Note the demo is to help you quick start, if you need use features like modal, you still need use `https`)

## Learn more 

To learn more about JavaScript API for Office (Office.js), please check [here](https://dev.office.com/reference/add-ins/javascript-api-for-office).
