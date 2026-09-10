import { SignJWT, jwtVerify } from "jose";
import bcrypt from "bcryptjs";
import { cookies } from "next/headers";

const COOKIE = "famossul_session";
const secret = () => new TextEncoder().encode(process.env.AUTH_SECRET || "dev-only-change-me");

export async function validateCredentials(user: string, password: string) {
  if (!process.env.ADMIN_USER || !process.env.ADMIN_PASSWORD_HASH) return false;
  return user === process.env.ADMIN_USER && bcrypt.compare(password, process.env.ADMIN_PASSWORD_HASH);
}

export async function createSession() {
  const token = await new SignJWT({ role: "admin" })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime("12h")
    .sign(secret());
  const jar = await cookies();
  jar.set(COOKIE, token, { httpOnly: true, sameSite: "lax", secure: process.env.NODE_ENV === "production", path: "/", maxAge: 60 * 60 * 12 });
}

export async function clearSession() {
  const jar = await cookies();
  jar.set(COOKIE, "", { httpOnly: true, path: "/", maxAge: 0 });
}

export async function isAuthenticated() {
  try {
    const jar = await cookies();
    const token = jar.get(COOKIE)?.value;
    if (!token) return false;
    await jwtVerify(token, secret());
    return true;
  } catch { return false; }
}
