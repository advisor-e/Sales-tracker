import { z } from "zod";
import { prisma } from "~/server/utils/db";

const schema = z.object({
  prospectName: z.string().min(1).max(255),
  businessName: z.string().max(255).optional().nullable(),
  partner: z.string().max(255).optional().nullable(),
  leadStaff: z.string().max(255).optional().nullable(),
  prospectStatus: z.string().min(1).max(80),
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
  const payload = schema.parse(await readBody(event));

  const item = await prisma.pipelineEntry.create({
    data: {
      prospectName: payload.prospectName,
      businessName: payload.businessName || null,
      partner: payload.partner || null,
      leadStaff: payload.leadStaff || null,
      prospectStatus: payload.prospectStatus,
      relationshipType: payload.relationshipType || null,
      prospectSource: payload.prospectSource || null,
      approachStyle: payload.approachStyle || null,
      approachDate: payload.approachDate ? new Date(payload.approachDate) : null,
      secureMeeting: payload.secureMeeting ?? false,
      proposalSent: payload.proposalSent ?? false,
      proposalValue: payload.proposalValue ?? 0,
      jobSecured: payload.jobSecured ?? false,
      jobSecuredValue: payload.jobSecuredValue ?? 0,
      comments: payload.comments || null,
      coiInvolved: payload.coiInvolved || null
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
