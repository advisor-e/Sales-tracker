import { z } from "zod";
import prismaModule from "@prisma/client";
import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

const { Prisma } = prismaModule;

const schema = z.object({
  signature: z.string().min(1).max(191),
  topic: z.string().min(1).max(255),
  audience: z.string().min(1).max(255),
  objective: z.string().min(1).max(255),
  tone: z.string().min(1).max(80),
  length: z.string().min(1).max(80),
  cta: z.string().min(1).max(255),
  principles: z.array(z.object({ title: z.string(), details: z.array(z.string()) })),
  selectedPerson: z.string().max(255).optional().nullable(),
  targetMode: z.string().max(120).optional().nullable(),
  styleStrength: z.string().max(40).optional().nullable(),
  styleTitles: z.array(z.string()).optional(),
  lengthRanges: z.unknown().optional(),
  styleThresholds: z.unknown().optional()
});

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const payload = schema.parse(await readBody(event));

  const existing = await prisma.blogInput.findFirst({
    where: { signature: payload.signature, userId: user.id },
    orderBy: { updatedAt: "desc" }
  });

  const data = {
    userId: user.id,
    signature: payload.signature,
    topic: payload.topic,
    audience: payload.audience,
    objective: payload.objective,
    tone: payload.tone,
    length: payload.length,
    cta: payload.cta,
    principlesJson: payload.principles,
    selectedPerson: payload.selectedPerson || null,
    targetMode: payload.targetMode || null,
    styleStrength: payload.styleStrength || null,
    styleTitlesJson: payload.styleTitles ?? Prisma.JsonNull,
    lengthRangesJson: (payload.lengthRanges as Prisma.InputJsonValue | undefined) ?? Prisma.JsonNull,
    styleThresholdsJson: (payload.styleThresholds as Prisma.InputJsonValue | undefined) ?? Prisma.JsonNull
  };

  const item = existing
    ? await prisma.blogInput.update({ where: { id: existing.id }, data })
    : await prisma.blogInput.create({ data });

  return { item };
});
