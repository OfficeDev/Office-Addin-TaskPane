/* global require process console */

const fs = require("fs");
const path = require("path");
const util = require("util");
const childProcess = require("child_process");

const readFileAsync = util.promisify(fs.readFile);
const unlinkFileAsync = util.promisify(fs.unlink);
const writeFileAsync = util.promisify(fs.writeFile);

const supportedHosts = ["excel", "outlook", "powerpoint", "word"];
const supporterHostsString = supportedHosts.join(", ");
const xmlHostTypes = { excel: "Workbook", outlook: "Mailbox", powerpoint: "Presentation", word: "Document" };
const supportedFeatures = ["commands", "functions", "events", "taskpane"];
const supportedFeatureString = supportedFeatures.join(", ");
const testPackages = [
  "@types/mocha",
  "@types/node",
  "mocha",
  "office-addin-mock",
  "office-addin-test-helpers",
  "office-addin-test-server",
  "ts-node",
];

// Help Text
if (process.argv[2] && process.argv[2].toLowerCase() === "help") {
  console.log("SYNTAX: convertForHosts.js <hosts> <features> <manifestType> <projectName> <appId>");
  console.log();
  console.log(`  hosts \tSpecifies which Office apps (comma seperated) will host the add-in: ${supporterHostsString}`);
  console.log(
    `  features \tSpecifies which features (comma seperated) to include in the add-in: ${supportedFeatureString}`
  );
  console.log(`  manifestType \tSpecify the type of manifest to use: 'xml' or 'json'.  Defaults to 'xml'`);
  console.log(
    `  projectName \tThe name of the project (use quotes when there are spaces in the name). Defaults to 'My Office Add-in'`
  );
  console.log(`  appId \tThe id of the project or 'random' to generate one.  Defaults to 'random'`);
  console.log();
  process.exit(1);
}

// Get arguments
let hosts = process.argv[2] || "excel,outlook,powerpoint,word";
let features = process.argv[3] || "commands,functions,events,taskpane";
let manifestType = process.argv[4] || "xml";
const projectName = process.argv[5] || "My Office Add-in";
let appId = process.argv[6] || "random";

// Define functions

async function convertProject() {
  // Validate arguments
  hosts = hosts.toLowerCase().split(",");
  if (!hosts.every((host) => supportedHosts.includes(host))) {
    throw new Error(`One or more specified hosts are not supported.  Supported hosts are ${supporterHostsString}`);
  }

  features = features.toLowerCase().split(",");
  if (!features.every((feature) => supportedFeatures.includes(feature))) {
    throw new Error(
      `One or more specified features are not supported.  Supported features are ${supportedFeatureString}`
    );
  }

  manifestType = manifestType.toLowerCase();
  if (manifestType !== "xml" && manifestType !== "json") {
    throw new Error(`Invalid manifest type.  Must be 'xml' or 'json'.`);
  }

  console.log(`Converting project for the following arguments:`);
  console.log(`  Hosts: ${hosts}`);
  console.log(`  Features: ${features}`);
  console.log(`  Manifest Type: ${manifestType}`);
  console.log(`  Project Name: ${projectName}`);
  console.log(`  App Id: ${appId}`);
  console.log();

  await updateSourceFiles();
  await updateFeatureFiles();
  await updateWebpackConfig();
  await updatePackageJson();
  await updateLaunchJsonFile();
  await deleteSupportFiles();

  // Make manifest type specific changes
  if (manifestType === "xml") {
    modifyProjectForXMLManifest();
  } else {
    modifyProjectForJsonManifest();
  }
}

async function updateSourceFiles() {
  // Delete unused source files
  const taskpaneFilePath = `./src/taskpane/taskpane.ts`;
  const commandsFilePath = `./src/commands/commands.ts`;
  let taskpaneContent = await readFileAsync(taskpaneFilePath, "utf8");
  let commandsContent = await readFileAsync(commandsFilePath, "utf8");

  supportedHosts.forEach(function (host) {
    if (!hosts.includes(host)) {
      deleteFileAsync(`./src/shared/${host}.ts`);
      deleteFileAsync(`./src/commands/${host}.ts`);
      deleteFileAsync(`./src/taskpane/${host}.ts`);
      taskpaneContent = taskpaneContent.replace(new RegExp(`import "\\./${host}";\r?\n?`, "gm"), "");
      commandsContent = commandsContent.replace(new RegExp(`import "\\./${host}";\r?\n?`, "gm"), "");
    }
  });

  await writeFileAsync(taskpaneFilePath, taskpaneContent);
  await writeFileAsync(commandsFilePath, commandsContent);
}

async function updateFeatureFiles() {
  // Delete unused features
  supportedFeatures.forEach(function (feature) {
    if (!features.includes(feature)) {
      deleteFolder(path.resolve(`./src/${feature}`));
    }
  });

  // Delete unused host specific features
  if (!hosts.includes("outlook")) {
    deleteFolder(path.resolve("./src/events"));
  }
  if (!hosts.includes("excel")) {
    deleteFolder(path.resolve("./src/functions"));
  }
}

async function updateWebpackConfig() {
  const webPack = `webpack.config.js`;
  const webPackContent = await readFileAsync(webPack, "utf8");
  const updatedContent = webPackContent.replace(".xml", `.${manifestType}`);
  await writeFileAsync(webPack, updatedContent);
}

async function updatePackageJson() {
  const packageJson = `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Update 'config' section in package.json to use first selected host
  content.config["app_to_debug"] = hosts[0].toLowerCase();

  // Remove 'engines' section
  delete content.engines;

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
  let content = JSON.parse(launchJsonContent);

  content.configurations = content.configurations.filter(function (config) {
    return hosts.some((host) => {
      return config.name.toLowerCase().startsWith(host);
    });
  });

  await writeFileAsync(launchJson, JSON.stringify(content, null, 2) + "\n");
}

async function deleteSupportFiles() {
  deleteFolder(path.resolve(`./test`));
  deleteFolder(path.resolve(`./.github`));
  deleteFolder(path.resolve(`./.azure-devops`));

  await deleteFileAsync("CONTRIBUTING.md");
  await deleteFileAsync("LICENSE");
  await deleteFileAsync("README.md");
  await deleteFileAsync("SECURITY.md");
  await deleteFileAsync("./convertForHosts.js");
  await deleteFileAsync(".npmrc");
  await deleteFileAsync("package-lock.json");
}

async function modifyProjectForXMLManifest() {
  // Remove JSON manifest related files
  await deleteFileAsync("manifest.json");
  supportedHosts.forEach(async function (host) {
    await deleteFileAsync(`manifest.${host}.json`);
  });
  await deleteFileAsync("assets/color.png");
  await deleteFileAsync("assets/outline.png");

  if (hosts.length === 1 && hosts[0] === "outlook") {
    // If only outlook is selected, use the outlook specific XML manifest
    const outlookManifestContent = await readFileAsync(`./manifest.outlook.xml`, "utf8");
    await writeFileAsync(`./manifest.xml`, outlookManifestContent);
    await deleteFileAsync(`manifest.outlook.xml`);
  } else {
    // Update based on selected hosts
    const manifestFilePath = `./manifest.xml`;
    let manifestContent = await readFileAsync(manifestFilePath, "utf8");
    supportedHosts.forEach(function (host) {
      if (host != "outlook" && !hosts.includes(host)) {
        let xmlHostType = xmlHostTypes[host];
        manifestContent = manifestContent
          .replace(new RegExp(`^\\s*<Host Name="${xmlHostType}"[^\\/]*\\/>\r?\n?`, "gm"), "")
          .replace(new RegExp(`^\\s*<Host xsi:type="${xmlHostType}"[^>]*>[\\s\\S]*?</Host>\r?\n?`, "gm"), "");
      }
    });
    await writeFileAsync(manifestFilePath, manifestContent);

    // Remove outlook specific XML manifest
    if (!hosts.includes("outlook")) {
      await deleteFileAsync(`manifest.outlook.xml`);
    }
  }
}

async function modifyProjectForJsonManifest() {
  // Remove XML manifest related files
  await deleteFileAsync("manifest.xml");
  supportedHosts.forEach(async function (host) {
    await deleteFileAsync(`manifest.${host}.xml`);
  });

  // Remove host specific JSON manifests
  supportedHosts.forEach(async function (host) {
    await deleteFileAsync(`manifest.${host}.json`);
  });
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

async function deleteFileAsync(file) {
  if (fs.existsSync(file)) {
    await unlinkFileAsync(file);
  }
}

/**
 * Modify the project so that it only supports indicated hosts.
 */
convertProject().catch((err) => {
  console.error(`Error modifying for hosts: ${err instanceof Error ? err.message : err}`);
  process.exitCode = 1;
});

// Modify the manifest to include the name and id of the project
const cmdLine = `npx office-addin-manifest modify manifest.${manifestType} -g ${appId} -d "${projectName}"`;
childProcess.exec(cmdLine, (error, stdout) => {
  if (error) {
    console.error(`Error updating the manifest: ${error}`);
    process.exitCode = 1;
  } else {
    console.log(stdout);
  }
});
