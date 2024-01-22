/* global require, process, console */

const fs = require("fs");
const path = require("path");
const util = require("util");
const childProcess = require("child_process");

const host = process.argv[2];
const manifestType = process.argv[3];
const projectName = process.argv[4];
let appId = process.argv[5];
const hosts = ["excel", "onenote", "outlook", "powerpoint", "project", "word", "wxpo"];
const testPackages = [
  "@types/mocha",
  "@types/node",
  "mocha",
  "office-addin-mock",
  "office-addin-test-helpers",
  "office-addin-test-server",
  "ts-node",
];
const readFileAsync = util.promisify(fs.readFile);
const unlinkFileAsync = util.promisify(fs.unlink);
const writeFileAsync = util.promisify(fs.writeFile);

async function modifyProjectForSingleHost(host) {
  if (!host) {
    throw new Error("The host was not provided.");
  }
  if (!hosts.includes(host)) {
    throw new Error(`'${host}' is not a supported host.`);
  }
  if (host === "wxpo") {
    return;
  }
  await convertProjectToSingleHost(host);
  await updatePackageJsonForSingleHost(host);
  await updateLaunchJsonFile();
}

async function convertProjectToSingleHost(host) {
  // Copy host-specific manifest over manifest.xml
  const manifestContent = await readFileAsync(`./manifest.${host}.xml`, "utf8");
  await writeFileAsync(`./manifest.xml`, manifestContent);

  // Copy over host-specific taskpane code to taskpane.ts
  const srcContent = await readFileAsync(`./src/taskpane/${host}.ts`, "utf8");
  await writeFileAsync(`./src/taskpane/taskpane.ts`, srcContent);

  // // Special handling for json manifest
  // if (typeof manifestType !== "undefined" && manifestType == "json") {
  //   hosts = ["project", "onenote"];
  // }

  // Delete all host-specific files
  hosts.forEach(async function (host) {
    await unlinkFileAsync(`./manifest.${host}.xml`);
    await unlinkFileAsync(`./src/taskpane/${host}.ts`);
  });

  // Delete test folder
  deleteFolder(path.resolve(`./test`));

  // Delete the .github folder
  deleteFolder(path.resolve(`./.github`));

  // Delete CI/CD pipeline files
  deleteFolder(path.resolve(`./.azure-devops`));

  // Delete repo support files
  await deleteSupportFiles();
}

async function updatePackageJsonForSingleHost(host) {
  // Update package.json to reflect selected host
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Update 'config' section in package.json to use selected host
  content.config["app_to_debug"] = host;

  // Remove 'engines' section
  delete content.engines;

  // Remove scripts that are unrelated to the selected host
  Object.keys(content.scripts).forEach(function (key) {
    if (key === "convert-to-single-host" || key === "start:desktop:outlook") {
      delete content.scripts[key];
    }
  });

  // Remove test-related scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("test")) {
      delete content.scripts[key];
    }
  });

  // Remove test-related packages
  Object.keys(content.devDependencies).forEach(function (key) {
    if (testPackages.includes(key)) {
      delete content.devDependencies[key];
    }
  });

  // Write updated JSON to file
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

async function updateLaunchJsonFile() {
  // Remove 'Debug Tests' configuration from launch.json
  const launchJson = `.vscode/launch.json`;
  const launchJsonContent = await readFileAsync(launchJson, "utf8");
  const regex = /(.+{\r?\n.*"name": "Debug (?:UI|Unit) Tests",\r?\n(?:.*\r?\n)*?.*},.*\r?\n)/gm;
  const updatedContent = launchJsonContent.replace(regex, "");
  await writeFileAsync(launchJson, updatedContent);
}

function deleteFolder(folder) {
  try {
    if (fs.existsSync(folder)) {
      fs.readdirSync(folder).forEach(function (file) {
        const curPath = `${folder}/${file}`;

        if (fs.lstatSync(curPath).isDirectory()) {
          deleteFolder(curPath);
        } else {
          fs.unlinkSync(curPath);
        }
      });
      fs.rmdirSync(folder);
    }
  } catch (err) {
    throw new Error(`Unable to delete folder "${folder}".\n${err}`);
  }
}

async function deleteSupportFiles() {
  await unlinkFileAsync("CONTRIBUTING.md");
  await unlinkFileAsync("LICENSE");
  await unlinkFileAsync("README.md");
  await unlinkFileAsync("SECURITY.md");
  await unlinkFileAsync("./convertToSingleHost.js");
  await unlinkFileAsync(".npmrc");
  await unlinkFileAsync("package-lock.json");
  if (manifestType === "json") {
    deleteFolder(path.resolve(`./node_modules`));
  }
}

async function deleteJSONManifestRelatedFiles() {
  await unlinkFileAsync("manifest.json");
  await unlinkFileAsync("assets/color.png");
  await unlinkFileAsync("assets/outline.png");
}

async function deleteXMLManifestRelatedFiles() {
  await unlinkFileAsync("manifest.xml");
}

async function updatePackageJsonForXMLManifest() {
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Remove scripts that are only used with JSON manifest
  delete content.scripts["signin"];
  delete content.scripts["signout"];

  // Write updated JSON to file
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

async function updatePackageJsonForJSONManifest() {
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Remove special start scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("start:")) {
      delete content.scripts[key];
    }
  });

  // Change manifest file name extension
  content.scripts.start = "office-addin-debugging start manifest.json";
  content.scripts.stop = "office-addin-debugging stop manifest.json";
  content.scripts.validate = "office-addin-manifest validate manifest.json";

  // Write updated JSON to file
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

async function updateTasksJsonFileForJSONManifest() {
  const tasksJson = `.vscode/tasks.json`;
  const data = await readFileAsync(tasksJson, "utf8");
  let content = JSON.parse(data);

  content.tasks.forEach(function (task) {
    if (task.label.startsWith("Build")) {
      task.dependsOn = ["Install"];
    }
    if (task.label === "Debug: Outlook Desktop") {
      task.script = "start";
      task.dependsOn = ["Check OS", "Install"];
    }
  });

  const checkOSTask = {
    label: "Check OS",
    type: "shell",
    windows: {
      command: "echo 'Sideloading in Outlook on Windows is supported'",
    },
    linux: {
      command: "echo 'Sideloading on Linux is not supported' && exit 1",
    },
    osx: {
      command: "echo 'Sideloading in Outlook on Mac is not supported' && exit 1",
    },
    presentation: {
      clear: true,
      panel: "dedicated",
    },
  };

  content.tasks.push(checkOSTask);
  await writeFileAsync(tasksJson, JSON.stringify(content, null, 2));
}

async function updateWebpackConfigForJSONManifest() {
  const webPack = `webpack.config.js`;
  const webPackContent = await readFileAsync(webPack, "utf8");
  const updatedContent = webPackContent.replace(".xml", ".json");
  await writeFileAsync(webPack, updatedContent);
}

modifyProjectForMultiHostsWXPO(host).catch((err) => {
  console.error(`Error modify Project For Multi Hosts WXPO: ${err instanceof Error ? err.message : err}`);
  process.exitCode = 1;
});

async function modifyProjectForMultiHostsWXPO(host) {
  // await convertProjectToMultiHostsWXPO();
  // await updatePackageJsonForMultiHostsWXPO(host);
  await updatePackageJsonForJSONManifest();
  await updateLaunchJsonFile();
  // await deleteXMLManifestRelatedFiles();
}

async function modifyProjectForJSONManifest() {
  await updatePackageJsonForJSONManifest();
  await updateWebpackConfigForJSONManifest();
  await updateTasksJsonFileForJSONManifest();
  await deleteXMLManifestRelatedFiles();
}

async function convertProjectToMultiHostsWXPO() {
  let wxpoHosts = ["excel", "outlook", "powerpoint", "word"];
  wxpoHosts.forEach(async function (host) {
    if (!hosts.includes(host)) {
      await unlinkFileAsync(`./manifest.${host}.xml`);
      await unlinkFileAsync(`./src/taskpane/${host}.ts`);
    }
  });

  // // Copy host-specific manifest over manifest.xml
  // const manifestContent = await readFileAsync(`./manifest.${host}.xml`, "utf8");
  // await writeFileAsync(`./manifest.xml`, manifestContent);

  // // Copy over host-specific taskpane code to taskpane.ts
  // const srcContent = await readFileAsync(`./src/taskpane/${host}.ts`, "utf8");
  // await writeFileAsync(`./src/taskpane/taskpane.ts`, srcContent);

  // // Delete all host-specific files
  // wxpoHosts.forEach(async function (host) {
  //   await unlinkFileAsync(`./manifest.${host}.xml`);
  //   await unlinkFileAsync(`./src/taskpane/${host}.ts`);
  // });

  // Delete test folder
  deleteFolder(path.resolve(`./test`));

  // Delete the .github folder
  deleteFolder(path.resolve(`./.github`));

  // Delete CI/CD pipeline files
  deleteFolder(path.resolve(`./.azure-devops`));

  // Delete repo support files
  await deleteSupportFiles();
}

async function updatePackageJsonForMultiHostsWXPO(host) {
  // Update package.json to reflect selected host
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Update 'config' section in package.json to use selected host
  content.config["app_to_debug"] = host;

  // Remove 'engines' section
  delete content.engines;

  // Remove scripts that are unrelated to the selected host
  Object.keys(content.scripts).forEach(function (key) {
    if (key === "convert-to-single-host" || key === "start:desktop:outlook") {
      delete content.scripts[key];
    }
  });

  // Remove test-related scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("test")) {
      delete content.scripts[key];
    }
  });

  // Remove test-related packages
  Object.keys(content.devDependencies).forEach(function (key) {
    if (testPackages.includes(key)) {
      delete content.devDependencies[key];
    }
  });

  // Write updated JSON to file
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

/**
 * Modify the project so that it only supports a single host.
 * @param host The host to support.
 */
modifyProjectForSingleHost(host).catch((err) => {
  console.error(`Error modifying for single host: ${err instanceof Error ? err.message : err}`);
  process.exitCode = 1;
});

let manifestPath = "manifest.xml";

if (host !== "wxpo" || manifestType !== "json") {
  // Remove things that are only relevant to JSON manifest
  deleteJSONManifestRelatedFiles();
  updatePackageJsonForXMLManifest();
} else if (host === "wxpo" && manifestType === "json") {
  manifestPath = "manifest.json";
  modifyProjectForJSONManifestWXPO().catch((err) => {
    console.error(`Error modify Project For JSON Manifest WXPO: ${err instanceof Error ? err.message : err}`);
    process.exitCode = 1;
  });
} else {
  manifestPath = "manifest.json";
  modifyProjectForJSONManifest().catch((err) => {
    console.error(`Error modifying for JSON manifest: ${err instanceof Error ? err.message : err}`);
    process.exitCode = 1;
  });
}

if (projectName) {
  if (!appId) {
    appId = "random";
  }

  // Modify the manifest to include the name and id of the project
  const cmdLine = `npx office-addin-manifest modify ${manifestPath} -g ${appId} -d ${projectName}`;
  childProcess.exec(cmdLine, (error, stdout) => {
    if (error) {
      Promise.reject(stdout);
    } else {
      Promise.resolve();
    }
  });
}

async function modifyProjectForJSONManifestWXPO() {
  await updatePackageJsonForJSONManifestWXPO();
  await updateWebpackConfigForJSONManifest();
  await updateTasksJsonFileForJSONManifestWXPO();
  await updateSrcFolderForJSONManifestWXPO();
  await updateCommandsFileForJSONManifestWXPO();
  await deleteXMLManifestRelatedFilesWXPO();
}

async function updatePackageJsonForJSONManifestWXPO() {
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Remove 'app_to_debug' section
  delete content.config["app_to_debug"];

  // Remove special start scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("start:")) {
      delete content.scripts[key];
    }
  });

  // Remove special test scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("test:")) {
      delete content.scripts[key];
    }
  });

  // Change test script
  content.scripts.test = 'echo "No tests."';

  // Change manifest file name extension
  content.scripts.start = "office-addin-debugging start manifest.json";
  content.scripts.stop = "office-addin-debugging stop manifest.json";
  content.scripts.validate = "office-addin-manifest validate manifest.json";

  // Write updated JSON to file
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

async function updateTasksJsonFileForJSONManifestWXPO() {
  const tasksJson = `.vscode/tasks.json`;
  const data = await readFileAsync(tasksJson, "utf8");
  let content = JSON.parse(data);

  content.tasks.forEach(function (task) {
    const debugScripts = {
      "Debug: Excel Desktop": "start -- --app excel",
      "Debug: Outlook Desktop": "start -- --app outlook",
      "Debug: PowerPoint Desktop": "start -- --app powerpoint",
      "Debug: Word Desktop": "start -- --app word",
    };

    if (task.label.startsWith("Build")) {
      task.dependsOn = ["Install"];
    } else if (debugScripts[task.label]) {
      task.script = debugScripts[task.label];
      task.dependsOn = ["Check OS", "Install"];
    }
  });

  const checkOSTask = {
    label: "Check OS",
    type: "shell",
    windows: {
      command: "echo 'Sideloading on Windows is supported'",
    },
    linux: {
      command: "echo 'Sideloading on Linux is not supported' && exit 1",
    },
    osx: {
      command: "echo 'Sideloading on Mac is not supported' && exit 1",
    },
    presentation: {
      clear: true,
      panel: "dedicated",
    },
  };

  content.tasks.push(checkOSTask);
  await writeFileAsync(tasksJson, JSON.stringify(content, null, 2));
}

async function updateSrcFolderForJSONManifestWXPO() {
  const fs = require("fs");
  const path = require("path");

  const srcFolder = `./src/taskpane`;
  const files = ["excel.ts", "word.ts", "outlook.ts", "powerpoint.ts"];

  let content = `
  Office.onReady((info) => {
    const runFunctions = {
      [Office.HostType.Outlook]: runOutlook,
      [Office.HostType.Word]: runWord,
      [Office.HostType.Excel]: runExcel,
      [Office.HostType.PowerPoint]: runPowerPoint,
    };

    if (runFunctions[info.host]) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("run").onclick = runFunctions[info.host];
    }
  });
  `;

  const generateContent = (appName, runCode) => `
  export async function run${appName}() {
    try {
      await ${appName}.run(async (context) => {
        /**
         * Insert your ${appName} code here
         */
        ${runCode}
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
  `;

  const appCodes = {
    "outlook.ts":
      'const item = context.mailbox.item; item.body.set("Hello, world!", { coercionType: Office.CoercionType.Text }); item.subject.set("Hello, world!"); item.saveAsync();',
    "word.ts":
      'const range = context.document.getSelection(); range.insertText("Hello, world!", Word.InsertLocation.end); range.font.color = "red";',
    "excel.ts":
      'const range = context.workbook.getSelectedRange(); range.values = [["Hello, world!"]]; range.format.fill.color = "yellow";',
    "powerpoint.ts": 'const slide = context.presentation.slides.getFirst(); slide.insertText("Hello World!", "End");',
  };

  files.forEach((file) => {
    if (appCodes[file]) {
      let appName = file.split(".")[0].charAt(0).toUpperCase() + file.split(".")[0].slice(1);
      if (appName === "Powerpoint") {
        appName = "PowerPoint";
      }
      content += generateContent(appName, appCodes[file]);
    }
  });

  const taskpanePath = path.join(srcFolder, "taskpane.ts");
  fs.writeFileSync(taskpanePath, content);
}

async function deleteXMLManifestRelatedFilesWXPO() {
  await unlinkFileAsync("manifest.xml");
  for (const host of hosts) {
    await unlinkFileAsync(`manifest.${host}.xml`);
    await unlinkFileAsync(`./src/taskpane/${host}.ts`);
  }
}

async function updateCommandsFileForJSONManifestWXPO() {
  const fs = require("fs");
  const commandsFile = `./src/commands/commands.ts`;
  let content = await readFileAsync(commandsFile, "utf8");

  const officeApps = {
    Excel: `const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";`,
    Word: `const range = context.document.getSelection();
      range.font.color = "red";`,
    Outlook: `const item = context.mailbox.item;
      item.body.set("Hello, world!", { coercionType: Office.CoercionType.Text });
      item.subject.set("Hello, world!");
      item.saveAsync();`,
    PowerPoint: `const slide = context.presentation.slides.getFirst();
      slide.shapes.getFirst().text = "Hello, world!";`,
  };

  let codeToInsert = "";
  for (const [app, code] of Object.entries(officeApps)) {
    codeToInsert += `
    await ${app}.run(async (context) => {
      ${code}
      await context.sync();
    });
    `;
  }

  content = content.replace(
    "function action(event: Office.AddinCommands.Event) {",
    `function action(event: Office.AddinCommands.Event) {\n${codeToInsert}`
  );
  fs.writeFileSync(commandsFile, content);
}
