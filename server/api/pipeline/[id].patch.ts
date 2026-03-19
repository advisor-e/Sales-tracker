import { z } from "zod";
import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

const schema = z.object({
  prospectName: z.string().min(1).max(255).optional(),
  businessName: z.string().max(255).optional().nullable(),
  partner: z.string().max(255).optional().nullable(),
  leadStaff: z.string().max(255).optional().nullable(),
  prospectStatus: z.string().min(1).max(80).optional(),
  relationshipType: z.string().max(80).optional().nullable(),
  prospectSource: z.string().max(120).optional().nullable(),
  approachStyle: z.string().max(120).optional().nullable(),
  approachDate: z.string().optional().nullable(),
  secureMeeting: z.boolean().optional(),
  proposalSent: z.boolean().optional(),
  proposalValue: z.number().optional(),
  jobSecured: z.boolean().optional(),
  jobSecuredValue: z.number().optional(),
  comments: z.string().optional().nullable(),
  coiInvolved: z.string().max(255).optional().nullable()
});

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const id = Number(getRouterParam(event, "id"));
  if (!Number.isFinite(id) || id <= 0) {
    throw createError({ statusCode: 400, statusMessage: "Invalid id" });
  }

  const existing = await prisma.pipelineEntry.findFirst({ where: { id, userId: user.id } });
  if (!existing) {
    throw createError({ statusCode: 404, statusMessage: "Not found" });
  }

  const payload = schema.parse(await readBody(event));
  const item = await prisma.pipelineEntry.update({
    where: { id },
    data: {
      ...payload,
      approachDate: payload.approachDate === undefined ? undefined : payload.approachDate ? new Date(payload.approachDate) : null
    }
  });

  return {
    item: {
      ...item,
      proposalValue: Number(item.proposalValue || 0),
      jobSecuredValue: Number(item.jobSecuredValue || 0)
    }
  };
});
