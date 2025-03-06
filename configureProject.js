/* global require process console */

const fs = require("fs");
const path = require("path");
const util = require("util");
const childProcess = require("child_process");

const readFileAsync = util.promisify(fs.readFile);
const unlinkFileAsync = util.promisify(fs.unlink);
const writeFileAsync = util.promisify(fs.writeFile);
const renameAsync = util.promisify(fs.rename);

const supportedHosts = ["excel", "outlook", "powerpoint", "word"];
const xmlHostTypes = { excel: "Workbook", outlook: "Mailbox", powerpoint: "Presentation", word: "Document" };
const supportedFeatures = ["commands", "functions", "events", "taskpane", "sharedRuntime"];
const supportedExtras = ["react", "auth"];

const typescriptDevDependencies = ["typescript", "ts-node"];
const reactDependencies = ["@fluentui/react-components", "@fluentui/react-icons", "react", "react-dom"];
const reactDevDependencies = ["@types/react", "@types/react-dom", "eslint-plugin-react"];
const customFunctionsDevDependencies = ["@types/custom-functions-runtime"];
const authDependencies = ["@azure/msal-browser"];
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
  console.log("SYNTAX: configureProject.js <hosts> <features> <codeLanguage> <manifestType> <extras> <projectName> <appId>");
  console.log();
  console.log(`  hosts \tSpecifies which Office apps (comma seperated) will host the add-in: ${supporterHosts.join(", ")}`);
  console.log(
    `  features \tSpecifies which features (comma seperated) to include in the add-in: ${supportedFeature.join(", ")}`
  );
  console.log(`  codeLanguage \tSpecifies the language to use for the project: 'ts' or 'js'.  Defaults to 'ts'`);
  console.log(`  manifestType \tSpecify the type of manifest to use: 'xml' or 'json'.  Defaults to 'xml'`);
  console.log(`  extras \tSpecify any additional features to include in the project: ${supportedExtras.join(", ")}`);
  console.log(
    `  projectName \tThe name of the project (use quotes when there are spaces in the name). Defaults to 'My Office Add-in'`
  );
  console.log(`  appId \tThe id of the project or 'random' to generate one.  Defaults to 'random'`);
  console.log();
  process.exit(1);
}

// Get arguments
const hosts = (process.argv[2]?.toLowerCase() || "excel,outlook,powerpoint,word").split(",");
const features = (process.argv[3]?.toLowerCase() || "commands,functions,events,taskpane").split(",");
const codeLanguage = process.argv[4]?.toLowerCase() || "typescript";
const manifestType = process.argv[5]?.toLowerCase() || "xml";
const extras = (process.argv[6]?.toLowerCase() || "").split(",");
const projectName = process.argv[7] || "My Office Add-in";
let appId = process.argv[8] || "random";

// Define functions

async function convertProject() {
  // Validate arguments
  if (!hosts.every((host) => supportedHosts.includes(host))) {
    throw new Error(`One or more specified hosts (${hosts}) are not supported.  Supported hosts are ${supporterHosts.join(", ")}`);
  }

  if (!features.every((feature) => supportedFeatures.includes(feature))) {
    throw new Error(
      `One or more specified features (${features}) are not supported.  Supported features are ${supportedFeature.join(", ")}`
    );
  }

  if (codeLanguage !== "ts" && codeLanguage !== "typescript" && codeLanguage !== "js" && codeLanguage !== "javascript") {
    throw new Error(`Invalid code language "${codeLanguage}".  Must be 'ts', 'typescript', 'js', or 'javascript'.`);
  }

  if (manifestType !== "xml" && manifestType !== "json") {
    throw new Error(`Invalid manifest type "${manifestType}".  Must be 'xml' or 'json'.`);
  }

  console.log(`Converting project for the following arguments:`);
  console.log(`  Hosts: ${hosts}`);
  console.log(`  Features: ${features}`);
  console.log(`  Code Language: ${codeLanguage}`);
  console.log(`  Manifest Type: ${manifestType}`);
  console.log(`  Extras: ${extras}`);
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
      if(feature === "taskpane" && extras.includes("react")) {
        deleteFolder(path.resolve(`./src/reactTaskpane`));
      }
    }
  });

  // Delete unused host specific features
  if (!hosts.includes("outlook")) {
    deleteFolder(path.resolve("./src/events"));
  }
  if (!hosts.includes("excel")) {
    deleteFolder(path.resolve("./src/functions"));
  }

  // Update React files
  if (features.includes("taskpane")) {
    if (extras.includes("react")) {
      deleteFolder(path.resolve("./src/taskpane"));
      await renameFolder(path.resolve("./src/reactTaskpane"), path.resolve("./src/taskpane"));
    } else {
      deleteFolder(path.resolve("./src/reactTaskpane"));
    }
  }

  // Update Auth files
  if (!extras.includes("auth")) {
    await deleteFileAsync(path.resolve("./src/shared/naa.ts"));
    await deleteFileAsync(path.resolve("./src/shared/graph.ts"));
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
  for (const key in content.scripts) {
    if (key.includes("test")) {
      delete content.scripts[key];
    }
  }

  // Remove scripts that are unrelated to the selected host
  Object.keys(content.scripts).forEach(function (key) {
    if (key === "configure-project") {
      delete content.scripts[key];
    }
  });

  // Remove test-related packages
  for (const key in content.devDependencies) {
    if (testPackages.includes(key)) {
      delete content.devDependencies[key];
    }
  }

  // Remove TypeScript related packages if converting to JavaScript
  if (codeLanguage === "js" || codeLanguage === "javascript") {
    typescriptDevDependencies.forEach(function (dep) {
      delete content.devDependencies[dep];
    });
  }

  // Remove custom functions related packages if not using custom functions
  if (!features.includes("functions")) {
    customFunctionsDevDependencies.forEach(function (dep) {
      delete content.devDependencies[dep];
    });
  }

  // Remove React related packages if not using React
  if (!extras.includes("react")) {
    reactDependencies.forEach(function (dep) {
      delete content.dependencies[dep];
    });
    reactDevDependencies.forEach(function (dep) {
      delete content.devDependencies[dep];
    });
  }

  // Remove Auth related packages if not using Auth
  if (!extras.includes("auth")) {
    authDependencies.forEach(function (dep) {
      delete content.dependencies[dep];
    });
  }

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
  await deleteFileAsync("./configureProject.js");
  await deleteFileAsync("tsconfig.convert.json");
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

  await modifyXmlManifest();
}

async function modifyXmlManifest() {
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

    // Remove unneeded shared runtime
    if (!features.includes("sharedRuntime")) {
      manifestContent = manifestContent.replace(/^\s*<Runtimes[\s\S]*?<\/Runtimes>\s*$/gm, "");
      manifestContent = manifestContent.replace(/^\s*<Requirements[\s\S]*?<\/Requirements>\s*$/gm, "");
    }

    // Update custom functions entries
    if (!features.includes("functions")) {
      manifestContent = manifestContent
        .replace(/^\s*<AllFormFactors[\s\S]*?<\/AllFormFactors>\s*$/gm, "")
        .replace(/^\s*<bt:Url\s+id=\"Functions\.[^.]+\.Url\".*\/>\s*$/gm, "")
        .replace(/^\s*<bt:String\s+id=\"Functions\.Namespace\".*\/>\s*$/gm, "");
    } else if (!features.includes("sharedRuntime")) {
      manifestContent = manifestContent
        .replace(
          /(<Page>\s*<SourceLocation\s+resid=\")Taskpane\.Url(\"\s*\/>\s*<\/Page>)/gm, 
          "$1Functions.Page.Url$2"
        );
    } else {
      manifestContent = manifestContent
        .replace(/^\s*<bt:Url\s+id=\"Functions\.Page\.Url\".*\/>\s*$/gm, "");
    }

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

  await modifyJsonManifest();
}

async function modifyJsonManifest() {
  const manifestFilePath = `./manifest.json`;
  let manifestContent = await readFileAsync(manifestFilePath, "utf8");
  const manifest = JSON.parse(manifestContent);

  // Remove unneeded authorizations

  // Remove unneeded requirements

  // Remove unneeded runtimes

  // Remove unneeded ribbons

  // Wriet the updated manifest back to the file
  await writeFileAsync(manifestFilePath, manifestContent);
}

// Helper functions
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

async function renameFolder(oldPath, newPath) {
  if (fs.existsSync(oldPath)) {
    await renameAsync(oldPath, newPath);
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
// const cmdLine = `npx office-addin-manifest modify manifest.${manifestType} -g ${appId} -d "${projectName}"`;
// childProcess.exec(cmdLine, (error, stdout) => {
//   if (error) {
//     console.error(`Error updating the manifest: ${error}`);
//     process.exitCode = 1;
//   } else {
//     console.log(stdout);
//   }
// });
