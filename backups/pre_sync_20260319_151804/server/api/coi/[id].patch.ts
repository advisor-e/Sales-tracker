import { z } from "zod";
import { prisma } from "~/server/utils/db";

const schema = z.object({
  coiName: z.string().min(1).max(255).optional(),
  email: z.string().max(255).optional().nullable(),
  cell: z.string().max(80).optional().nullable(),
  entity: z.string().max(255).optional().nullable(),
  position: z.string().max(255).optional().nullable(),
  industry: z.string().max(120).optional().nullable(),
  leadRelationshipPartner: z.string().max(255).optional().nullable(),
  relationshipSupport: z.string().max(255).optional().nullable(),
  couldWe: z.boolean().optional(),
  howWouldWe: z.boolean().optional(),
  willWe: z.boolean().optional(),
  testReview: z.boolean().optional(),
  totalReferrals: z.number().int().optional(),
  totalConverted: z.number().int().optional(),
  feeValue: z.number().optional()
});

export default defineEventHandler(async (event) => {
  const id = Number(getRouterParam(event, "id"));
  if (!Number.isFinite(id) || id <= 0) {
    throw createError({ statusCode: 400, statusMessage: "Invalid id" });
  }

  const payload = schema.parse(await readBody(event));
  const item = await prisma.coiEntry.update({ where: { id }, data: payload });
  return { item: { ...item, feeValue: Number(item.feeValue || 0) } };
});
