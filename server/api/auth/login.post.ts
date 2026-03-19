import { z } from "zod";
import { defineEventHandler, readBody, createError } from "h3";
import { prisma } from "../../utils/db";
import { createSession, hashPassword, verifyPassword } from "../../utils/auth";

const schema = z.object({
  email: z.string().email().max(255),
  password: z.string().min(8).max(128)
});

export default defineEventHandler(async (event) => {
  const payload = schema.parse(await readBody(event));
  const email = payload.email.trim().toLowerCase();

  const adminEmail = String(process.env.ADMIN_EMAIL || "admin@example.com").trim().toLowerCase();
  const adminPassword = String(process.env.ADMIN_PASSWORD || "").trim();

  const userCount = await prisma.user.count();
  if (userCount === 0 && adminPassword && email === adminEmail && payload.password === adminPassword) {
    await prisma.user.create({
      data: {
        email,
        displayName: "Administrator",
        passwordHash: hashPassword(payload.password)
      }
    });
  }

  const user = await prisma.user.findUnique({ where: { email } });
  if (!user || !verifyPassword(payload.password, user.passwordHash)) {
    throw createError({ statusCode: 401, statusMessage: "Invalid credentials" });
  }

  await createSession(event, user.id);

  return {
    user: {
      id: user.id,
      email: user.email,
      displayName: user.displayName
    }
  };
});
