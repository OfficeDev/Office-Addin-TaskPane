const fs = require('fs');
const path = require('path');
const projectFolder = process.env['projectFolder'];
const projectType = process.env['projectType'];
const host = process.env['host'];
const typescript = (process.env['typescript'] == 'true') ? true : false;

const hosts = [
    "excel",
    "onenote",
    "outlook",
    "powerpoint",
    "project",
    "word"
];

modifyProjectForSingleHost(projectFolder, projectType, host, typescript);

async function modifyProjectForSingleHost(projectFolder, projectType, host, typescript) {
    return new Promise(async (resolve, reject) => {
        try {
            await convertProjectToSingleHost(projectFolder, projectType, host, typescript);
            await updatePackageJsonForSingleHost(projectFolder, host);
            return resolve();
        } catch (err) {
            return reject(err);
        }
    });
}

async function convertProjectToSingleHost(projectFolder, projectType, host, typescript) {
    try {
        let extension = typescript ? "ts" : "js";

        // copy host-specific manifest over manifest.xml
        await fs.readFile(path.resolve(`${projectFolder}/manifest.${host}.xml`), 'utf8', async (err, data) => {
            if (err) throw err;
            await fs.writeFile(path.resolve(`${projectFolder}/manifest.xml`), data, (err) => {
                if (err) throw err;
            });
        });

        // copy host-specific taskpane.ts[js] over src/taskpane/taskpane.ts[js]
        await fs.readFile(path.resolve(`${projectFolder}/src/taskpane/${host}.${extension}`), 'utf8', async (err, data) => {
            if (err) throw err;
            await fs.writeFile(path.resolve(path.resolve(`${projectFolder}/src/taskpane/taskpane.${extension}`)), data, (err) => {
                if (err) throw err;
            });
        });

        // delete all host specific files
        hosts.forEach(async function (host) {
            await fs.unlink(path.resolve(`${projectFolder}/manifest.${host}.xml`), function (err) {
                if (err) throw err;
            });
            await fs.unlink(path.resolve(`${projectFolder}/src/taskpane/${host}.${extension}`), function (err) {
                if (err) throw err;
            });
        });
    } catch (err) {
        throw err;
    }
}

async function updatePackageJsonForSingleHost(projectFolder, host) {
    try {
        // update package.json to reflect selected host
        const packageJson = path.resolve(`${projectFolder}/package.json`);

        // copy host-specific manifest over manifest.xml
        let content;
        await fs.readFile(packageJson, 'utf8', async (err, data) => {
            if (err) throw err;
            content = JSON.parse(data);

            // update 'config' section in package.json to use selected host
            content.config["app-to-debug"] = host;

            // remove scripts from package.json that are unrelated to selected host,
            // and update sideload and unload scripts to use selected host.
            Object.keys(content.scripts).forEach(function (key) {
                if (key.includes("sideload:") || key.includes("unload:") || key.includes("convert-to-single-host")) {
                    delete content.scripts[key];
                }
                switch (key) {
                    case "sideload":
                    case "unload":
                        content.scripts[key] = content.scripts[`${key}:${host}`];
                        break;
                }
            });

            await fs.writeFile(packageJson, JSON.stringify(content, null, 4), (err) => {
                if (err) throw err;
            });
        });
    } catch (err) {
        throw new Error(err);
    }
}