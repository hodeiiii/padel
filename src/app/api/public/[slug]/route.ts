import type { NextRequest } from "next/server";

import { storeConfigured, storeGet, storeSet } from "@/app/lib/store";
import { publicStorageKey } from "@/app/lib/tournament";

// Always run at request time: reads/writes live data in the shared store.
export const dynamic = "force-dynamic";

// Publishing is gated by the same admin password used by the panel. Override it
// in the host with the ADMIN_PASSWORD environment variable.
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD ?? "landerlander";

type RouteParams = { params: Promise<{ slug: string }> };

export async function GET(_request: NextRequest, { params }: RouteParams) {
  const { slug } = await params;

  if (!storeConfigured) {
    return Response.json({ configured: false, tournament: null });
  }

  const raw = await storeGet(publicStorageKey(slug));

  return Response.json({
    configured: true,
    tournament: raw ? JSON.parse(raw) : null,
  });
}

export async function POST(request: NextRequest, { params }: RouteParams) {
  const { slug } = await params;

  if (request.headers.get("x-admin-password") !== ADMIN_PASSWORD) {
    return Response.json({ error: "unauthorized" }, { status: 401 });
  }

  if (!storeConfigured) {
    return Response.json({ error: "store-not-configured" }, { status: 503 });
  }

  const body = await request.text();

  try {
    JSON.parse(body);
  } catch {
    return Response.json({ error: "invalid-json" }, { status: 400 });
  }

  await storeSet(publicStorageKey(slug), body);

  return Response.json({ ok: true });
}
