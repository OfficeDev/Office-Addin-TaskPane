const fs = require('fs');
const host = process.argv[2];
const hosts = [
    "excel",
    "onenote",
    "outlook",
    "powerpoint",
    "project",
    "word"
];
const util = require('util');
const readFileAsync = util.promisify(fs.readFile);
const unlinkFileAsync = util.promisify(fs.unlink);
const writeFileAsync = util.promisify(fs.writeFile);

modifyProjectForSingleHost(host);

async function modifyProjectForSingleHost(host) {
    try {
        await convertProjectToSingleHost(host);
        await updatePackageJsonForSingleHost(host);
    } catch (err) {
        throw err;
    }
}

async function convertProjectToSingleHost(host) {
    try {
        // copy host-specific manifest over manifest.xml
        // copy host-specific manifest over manifest.xml
        const manifestContent = await readFileAsync(`./manifest.${host}.xml`, 'utf8');
        await writeFileAsync(`./manifest.xml`, manifestContent);

        // copy over host-specific taskpane code to taskpane.ts
        const srcContent = await readFileAsync(`./src/taskpane/${host}.ts`, 'utf8');
        await writeFileAsync(`./src/taskpane/taskpane.ts`, srcContent);

        // delete all host-specific files
        hosts.forEach(async function (host) {
            await unlinkFileAsync(`./manifest.${host}.xml`);
            await unlinkFileAsync(`./src/taskpane/${host}.ts`);
        });
    } catch (err) {
        throw err;
    }
}

async function updatePackageJsonForSingleHost(host) {
    try {
        // update package.json to reflect selected host
        const packageJson = `./package.json`;
        const data = await readFileAsync(packageJson, 'utf8');
        let content = JSON.parse(data);

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

        // write updated json to file
        await writeFileAsync(packageJson, JSON.stringify(content, null, 4));
    } catch (err) {
        throw new Error(err);
    }
}