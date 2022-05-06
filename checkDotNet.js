const chalk = require("chalk");
const childProcess = require("child_process");

/* global console require */

try {
  // Find the .Net runtimes installed
  let result = childProcess.execSync("dotnet --list-runtimes");
  const pattern = /(?<=Microsoft.NETCore.App )[\d.]+/g;
  const matches = result.toString("utf-8").match(pattern);
  let foundDotNet = false;

  // Look for version 5 or greater
  matches?.forEach((match) => {
    const major = parseInt(match.split(".")[0]);
    foundDotNet = foundDotNet || major >= 5;
  });

  console.log("Correct version of .Net is installed");
} catch (err) {
  console.log(chalk.bold.red(".Net 5 or greater is required for json manifests."));
}
