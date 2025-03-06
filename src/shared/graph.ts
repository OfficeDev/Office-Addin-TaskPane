/**
 * @file : graph.js
 * @author : Microsoft Corporation
 */

import { getAuthentication } from "../shared/naa";

/* global Office fetch*/

async function fetchGraphData(url: string, accessToken: string) {
  const result = await fetch(url, { headers: { Authorization: accessToken } });
  const response = await result.text();

  if (result.ok) {
    return JSON.parse(response);
  } else {
    throw response;
  }
}

//Graph call with Naa token
export async function getUserInfo() {
  try {
    const authentication = await getAuthentication(["User.Read", "openid", "profile"]);
    const graphResponse = await fetchGraphData("https://graph.microsoft.com/v1.0/me", authentication.accessToken);
    return graphResponse || "";
  } catch (error) {
    console.log(`Failed to get user info: ${error}`);
    return "";
  }
}
