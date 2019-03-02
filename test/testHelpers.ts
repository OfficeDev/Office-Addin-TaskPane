let port: number = 8080;
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

export async function pingTestServer(portNumber: number | undefined): Promise<Object> {
    return new Promise<Object>(async (resolve, reject) => {
        if (portNumber !== undefined) {
            port = portNumber;
        }

        const serverResponse: any = {};
        const serverStatus: string = "status";
        const platform: string = "platform";
        const xhr = new XMLHttpRequest();
        const pingUrl: string = `https://localhost:${port}/ping`;
        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200) {
                serverResponse[serverStatus] = xhr.status;
                serverResponse[platform] = xhr.responseText;
                resolve(serverResponse);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.statusText.message === "XHR error") {
                reject(serverResponse);
            }
        };
        xhr.open("GET", pingUrl, true);
        xhr.send();
    });
}

export async function sendTestResults(data: Object, portNumber: number | undefined): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        if (portNumber !== undefined) {
            port = portNumber;
        }

        const json = JSON.stringify(data);
        const xhr = new XMLHttpRequest();
        const url: string = `https://localhost:${port}/results/`;
        const dataUrl: string = url + "?data=" + encodeURIComponent(json);

        xhr.open("POST", dataUrl, true);
        xhr.send();
        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200 && xhr.responseText === "200") {
                resolve(true);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.statusTest === "XHR error") {
                reject(false);
            }
        };
    });
}