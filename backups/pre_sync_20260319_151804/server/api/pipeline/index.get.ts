import { prisma } from "~/server/utils/db";

export default defineEventHandler(async (event) => {
  const query = getQuery(event);
  const search = String(query.search || "").trim();
  const status = String(query.status || "").trim();
  const owner = String(query.owner || "").trim();

  const items = await prisma.pipelineEntry.findMany({
    where: {
      prospectStatus: status || undefined,
      leadStaff: owner || undefined,
      OR: search
        ? [
            { prospectName: { contains: search } },
            { businessName: { contains: search } },
            { partner: { contains: search } },
            { leadStaff: { contains: search } }
          ]
        : undefined
    },
    orderBy: [{ updatedAt: "desc" }],
    take: 500
  });

  return {
    items: items.map((item) => ({
      ...item,
      proposalValue: Number(item.proposalValue || 0),
      jobSecuredValue: Number(item.jobSecuredValue || 0)
    }))
  };
});
