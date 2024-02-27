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
  if (
    (manifestType === "json" && (host === "onenote" || host === "project")) ||
    (manifestType === "xml" && host === "wxpo")
  ) {
    throw new Error(`'${host}' is not supported for ${manifestType} manifest.`);
  }

  await convertProjectToSingleHost(host, manifestType);

  await updatePackageJsonForSingleHost(host);
  await updateLaunchJsonFile();
}

async function convertProjectToSingleHost(host, manifestType) {
  // Copy host-specific manifest over manifest file
  const manifestPath = `./manifest.${host}.${manifestType}`;
  if (fs.existsSync(manifestPath)) {
    let manifestContent = await readFileAsync(manifestPath, "utf8");
    await writeFileAsync(`./manifest.${manifestType}`, manifestContent);
  }

  // Copy over host-specific taskpane code to taskpane.ts
  const taskpaneFilePath = "./src/taskpane/taskpane.ts";
  let taskpaneContent = await readFileAsync(taskpaneFilePath, "utf8");
  let targetHosts = host === "wxpo" ? ["word", "excel", "powerpoint", "outlook"] : [host];

  for (const host of hosts) {
    if (!targetHosts.includes(host)) {
      if (fs.existsSync(`./src/taskpane/${host}.ts`)) {
        await unlinkFileAsync(`./src/taskpane/${host}.ts`);
      }
      // remove unneeded imports
      taskpaneContent = taskpaneContent.replace(`import "./${host}";`, "").replace(/^\s*[\r\n]/gm, "");
    }
    // Remove unneeded manifest templates
    if (fs.existsSync(`./manifest.${host}.${manifestType}`)) {
      await unlinkFileAsync(`./manifest.${host}.${manifestType}`);
    }
  }
  await writeFileAsync(taskpaneFilePath, taskpaneContent);

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
  if (host === "wxpo") {
    // Specify a default debug host for json manifest
    content.config["app_to_debug"] = "excel";
  } else {
    content.config["app_to_debug"] = host;
  }

  // Remove 'engines' section
  delete content.engines;

  // Remove scripts that are unrelated to the selected host
  Object.keys(content.scripts).forEach(function (key) {
    if (key === "convert-to-single-host" || key === "start:desktop:outlook") {
      delete content.scripts[key];
    }
  });

  // Remove special start scripts
  Object.keys(content.scripts).forEach(function (key) {
    if (key.includes("start:")) {
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

  // Change manifest file name extension
  content.scripts.start = `office-addin-debugging start manifest.${manifestType}`;
  content.scripts.stop = `office-addin-debugging stop manifest.${manifestType}`;
  content.scripts.validate = `office-addin-manifest validate manifest.${manifestType}`;

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
}

async function deleteJSONManifestRelatedFiles() {
  await unlinkFileAsync("manifest.json");
  for (const host of hosts) {
    if (fs.existsSync(`./manifest.${host}.json`)) {
      await unlinkFileAsync(`manifest.${host}.json`);
    }
  }
  await unlinkFileAsync("assets/color.png");
  await unlinkFileAsync("assets/outline.png");
}

async function deleteXMLManifestRelatedFiles() {
  await unlinkFileAsync("manifest.xml");
  hosts.forEach(async function (host) {
    if (fs.existsSync(`./manifest.${host}.xml`)) {
      await unlinkFileAsync(`manifest.${host}.xml`);
    }
  });
}

async function updateWebpackConfigForJSONManifest() {
  const webPack = `webpack.config.js`;
  const webPackContent = await readFileAsync(webPack, "utf8");
  const updatedContent = webPackContent.replace(".xml", ".json");
  await writeFileAsync(webPack, updatedContent);
}

async function updateTasksJsonFileForJSONManifest() {
  const tasksJson = `.vscode/tasks.json`;
  const data = await readFileAsync(tasksJson, "utf8");
  let content = JSON.parse(data);

  content.tasks.forEach(function (task) {
    if (task.label.startsWith("Build") || task.label.startsWith("Debug:")) {
      task.dependsOn = ["Install"];
    }
  });

  await writeFileAsync(tasksJson, JSON.stringify(content, null, 2));
}

async function modifyProjectForJSONManifest() {
  await updateWebpackConfigForJSONManifest();
  await updateTasksJsonFileForJSONManifest();
  await deleteXMLManifestRelatedFiles();
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

if (manifestType !== "json") {
  // Remove things that are only relevant to JSON manifest
  deleteJSONManifestRelatedFiles();
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
