import prismaModule from "@prisma/client";
import { randomBytes, scryptSync } from "node:crypto";

const { PrismaClient } = prismaModule;

function hashPassword(password) {
  const salt = randomBytes(16).toString("hex");
  const hash = scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

const prisma = new PrismaClient();

async function main() {
  const email = String(process.env.ADMIN_EMAIL || "admin@example.com").trim().toLowerCase();
  const password = String(process.env.ADMIN_PASSWORD || "").trim();
  const displayName = String(process.env.ADMIN_DISPLAY_NAME || "Administrator").trim();

  if (!password) {
    throw new Error("ADMIN_PASSWORD is required");
  }

  const existing = await prisma.user.findUnique({ where: { email } });
  if (existing) {
    console.log(`Admin user already exists for ${email}`);
    return;
  }

  await prisma.user.create({
    data: {
      email,
      displayName,
      passwordHash: hashPassword(password)
    }
  });

  console.log(`Admin user created: ${email}`);
}

main()
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
