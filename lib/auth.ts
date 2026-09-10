import bcrypt from "bcryptjs";
import { cookies } from "next/headers";
import { prisma } from "@/lib/prisma";
import { SignJWT, jwtVerify } from "jose";

const secret = new TextEncoder().encode(
  process.env.AUTH_SECRET || "change-this-secret"
);

const COOKIE_NAME = "famossul_session";

export async function authenticate(username: string, password: string) {
  const adminUser = process.env.ADMIN_USER;
  const initialPassword = process.env.ADMIN_INITIAL_PASSWORD;

  if (!adminUser) {
    throw new Error("ADMIN_USER não configurado.");
  }

  // Procura usuário no Neon
  let user = await prisma.user.findUnique({
    where: {
      username,
    },
  });

  // Primeiro acesso:
  // cria automaticamente o administrador no Neon
  if (!user) {
    if (
      username !== adminUser ||
      !initialPassword ||
      password !== initialPassword
    ) {
      return null;
    }

    const passwordHash = await bcrypt.hash(password, 12);

    user = await prisma.user.create({
      data: {
        username,
        passwordHash,
      },
    });
  }

  // Usuário já existe: verifica senha criptografada
  const validPassword = await bcrypt.compare(
    password,
    user.passwordHash
  );

  if (!validPassword) {
    return null;
  }

  return {
    id: user.id,
    username: user.username,
  };
}

export async function createSession(user: {
  id: string;
  username: string;
}) {
  const token = await new SignJWT({
    userId: user.id,
    username: user.username,
  })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime("30d")
    .sign(secret);

  const cookieStore = await cookies();

  cookieStore.set(COOKIE_NAME, token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    maxAge: 60 * 60 * 24 * 30,
  });
}

export async function getSession() {
  try {
    const cookieStore = await cookies();
    const token = cookieStore.get(COOKIE_NAME)?.value;

    if (!token) {
      return null;
    }

    const { payload } = await jwtVerify(token, secret);

    return {
      id: payload.userId as string,
      username: payload.username as string,
    };
  } catch {
    return null;
  }
}

export async function destroySession() {
  const cookieStore = await cookies();

  cookieStore.delete(COOKIE_NAME);
}
