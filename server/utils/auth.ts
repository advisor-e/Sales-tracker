import { createHash, randomBytes, scryptSync, timingSafeEqual } from "crypto";
import { getCookie, setCookie, deleteCookie, createError, getRequestProtocol, type H3Event } from "h3";
import { prisma } from "./db";

const SESSION_COOKIE = "st_session";
const SESSION_DAYS = 30;

function hashToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}

export function hashPassword(password: string): string {
  const salt = randomBytes(16).toString("hex");
  const hash = scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

export function verifyPassword(password: string, storedHash: string): boolean {
  const [salt, expected] = String(storedHash || "").split(":");
  if (!salt || !expected) {
    return false;
  }
  const actual = scryptSync(password, salt, 64).toString("hex");
  const actualBuf = Buffer.from(actual, "hex");
  const expectedBuf = Buffer.from(expected, "hex");
  if (actualBuf.length !== expectedBuf.length) {
    return false;
  }
  return timingSafeEqual(actualBuf, expectedBuf);
}

export async function createSession(event: H3Event, userId: number): Promise<void> {
  const token = randomBytes(32).toString("hex");
  const tokenHash = hashToken(token);
  const expiresAt = new Date(Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000);
  const isHttps = getRequestProtocol(event, { xForwardedProto: true }) === "https";

  await prisma.session.create({
    data: { userId, tokenHash, expiresAt }
  });

  setCookie(event, SESSION_COOKIE, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: isHttps,
    expires: expiresAt,
    path: "/"
  });
}

export async function clearUserSession(event: H3Event): Promise<void> {
  const token = getCookie(event, SESSION_COOKIE);
  if (token) {
    await prisma.session.deleteMany({ where: { tokenHash: hashToken(token) } });
  }
  deleteCookie(event, SESSION_COOKIE, { path: "/" });
}

export async function requireUser(event: H3Event): Promise<{ id: number; email: string; displayName: string | null }> {
  const token = getCookie(event, SESSION_COOKIE);
  if (!token) {
    throw createError({ statusCode: 401, statusMessage: "Unauthorized" });
  }

  const session = await prisma.session.findFirst({
    where: {
      tokenHash: hashToken(token),
      expiresAt: { gt: new Date() }
    },
    include: { user: true }
  });

  if (!session?.user) {
    deleteCookie(event, SESSION_COOKIE, { path: "/" });
    throw createError({ statusCode: 401, statusMessage: "Unauthorized" });
  }

  return {
    id: session.user.id,
    email: session.user.email,
    displayName: session.user.displayName
  };
}

export async function getOptionalUser(event: H3Event): Promise<{ id: number; email: string; displayName: string | null } | null> {
  try {
    return await requireUser(event);
  } catch {
    return null;
  }
}
