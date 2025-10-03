import { acquireToken } from "../utils/acquireToken";

declare const APP_CONFIG: any;

export async function initSession(
  siteUrl: string,
  user: string
): Promise<{ sessionId: string }> {
  const token = await acquireToken(APP_CONFIG, user);

  const response = await fetch(`${APP_CONFIG.baseUrl}/session/init`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ siteUrl }),
  });

  if (!response.ok) {
    throw new Error("Failed to initialize session");
  }

  return await response.json();
}
