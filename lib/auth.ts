import bcrypt from "bcryptjs";
import { cookies } from "next/headers";
import { SignJWT, jwtVerify } from "jose";
import { prisma } from "@/lib/prisma";

const COOKIE_NAME = "famossul_session";

function getSecret() {
  const secret = process.env.AUTH_SECRET;

  if (!secret) {
    throw new Error("AUTH_SECRET não configurado.");
  }

  return new TextEncoder().encode(secret);
}

/**
 * Valida usuário e senha.
 * Se ainda não existir usuário no Neon, cria automaticamente
 * usando ADMIN_USER e ADMIN_INITIAL_PASSWORD.
 */
export async function validateCredentials(
  username: string,
  password: string
): Promise<boolean> {
  try {
    const adminUser = process.env.ADMIN_USER;
    const initialPassword = process.env.ADMIN_INITIAL_PASSWORD;

    if (!adminUser) {
      console.error("ADMIN_USER não configurado.");
      return false;
    }

    let user = await prisma.user.findUnique({
      where: {
        username,
      },
    });

    // Primeiro acesso: cria o administrador automaticamente
    if (!user) {
      if (
        username !== adminUser ||
        !initialPassword ||
        password !== initialPassword
      ) {
        return false;
      }

      const passwordHash = await bcrypt.hash(password, 12);

      user = await prisma.user.create({
        data: {
          username,
          passwordHash,
        },
      });
    }

    return await bcrypt.compare(password, user.passwordHash);
  } catch (error) {
    console.error("Erro ao validar credenciais:", error);
    return false;
  }
}

/**
 * Cria a sessão após login.
 * Mantido sem parâmetros para funcionar com a API atual do projeto.
 */
export async function createSession(): Promise<void> {
  const token = await new SignJWT({
    authenticated: true,
  })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime("30d")
    .sign(getSecret());

  const cookieStore = await cookies();

  cookieStore.set(COOKIE_NAME, token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    maxAge: 60 * 60 * 24 * 30,
  });
}

/**
 * Verifica se existe uma sessão válida.
 */
export async function isAuthenticated(): Promise<boolean> {
  try {
    const cookieStore = await cookies();
    const token = cookieStore.get(COOKIE_NAME)?.value;

    if (!token) {
      return false;
    }

    const { payload } = await jwtVerify(token, getSecret());

    return payload.authenticated === true;
  } catch {
    return false;
  }
}

/**
 * Remove a sessão.
 */
export async function clearSession(): Promise<void> {
  const cookieStore = await cookies();

  cookieStore.set(COOKIE_NAME, "", {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    expires: new Date(0),
  });
}

/**
 * Alias para compatibilidade, caso alguma parte
 * do projeto utilize destroySession.
 */
export async function destroySession(): Promise<void> {
  await clearSession();
}
