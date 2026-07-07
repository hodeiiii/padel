// Server-side persistence for published tournaments.
//
// Talks to an Upstash Redis / Vercel KV REST endpoint when the environment is
// configured. When it is not (e.g. local dev without credentials), the public
// page transparently falls back to the browser's localStorage, so nothing
// breaks before the store is wired up on the host.

const REST_URL =
  process.env.KV_REST_API_URL ?? process.env.UPSTASH_REDIS_REST_URL ?? "";
const REST_TOKEN =
  process.env.KV_REST_API_TOKEN ?? process.env.UPSTASH_REDIS_REST_TOKEN ?? "";

export const storeConfigured = Boolean(REST_URL && REST_TOKEN);

async function runCommand(command: string[]): Promise<unknown> {
  const response = await fetch(REST_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${REST_TOKEN}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(command),
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(`Store request failed (${response.status})`);
  }

  const data = (await response.json()) as { result?: unknown; error?: string };

  if (data.error) throw new Error(data.error);

  return data.result ?? null;
}

export async function storeGet(key: string): Promise<string | null> {
  if (!storeConfigured) return null;

  const result = await runCommand(["GET", key]);

  return typeof result === "string" ? result : null;
}

export async function storeSet(key: string, value: string): Promise<void> {
  if (!storeConfigured) {
    throw new Error("Persistent store is not configured");
  }

  await runCommand(["SET", key, value]);
}
