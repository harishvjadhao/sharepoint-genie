import { IUser } from "../models/ISidebarAgentProperties";
import { acquireToken } from "../utils/acquireToken";

declare const APP_CONFIG: any;

export async function initSession(
  siteUrl: string,
  user: IUser
): Promise<{ sessionId: string }> {
  const token = await acquireToken(APP_CONFIG, user.loginName);

  const response = await fetch(`${APP_CONFIG.directConnectUrl}/session/init`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      siteUrl,
      userName: user.displayName?.split(" ")[0] || "",
    }),
  });

  if (!response.ok) {
    throw new Error("Failed to initialize session");
  }

  return await response.json();
}
