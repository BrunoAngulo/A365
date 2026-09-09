import { NextResponse } from "next/server";

export async function POST(request: Request) {
  const body = await request.json().catch(() => null) as { password?: string } | null;
  if (body?.password !== "Palabra!Aleatoria2026#") {
    return NextResponse.json({ ok: false }, { status: 401 });
  }

  const response = NextResponse.json({ ok: true });
  response.cookies.set("a365_secure_access", "1", {
    httpOnly: true,
    maxAge: 60 * 60 * 24 * 30,
    path: "/",
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
  });
  return response;
}
