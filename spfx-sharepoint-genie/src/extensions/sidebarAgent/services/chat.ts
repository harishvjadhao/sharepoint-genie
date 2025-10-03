import { acquireToken } from "../utils/acquireToken";

declare const APP_CONFIG: any;

export async function chat(
  sessionId: string,
  message: string,
  user: string
): Promise<any> {
  const token = await acquireToken(APP_CONFIG, user);

  const response = await fetch(`${APP_CONFIG.baseUrl}/chat`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      sessionId,
      message,
    }),
  });

  if (!response.ok) {
    throw new Error("Chat request failed");
  }

  return await response.json();
}
