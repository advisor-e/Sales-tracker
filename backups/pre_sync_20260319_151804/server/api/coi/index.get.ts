import { prisma } from "~/server/utils/db";

export default defineEventHandler(async (event) => {
  const query = getQuery(event);
  const search = String(query.search || "").trim();

  const items = await prisma.coiEntry.findMany({
    where: {
      OR: search
        ? [
            { coiName: { contains: search } },
            { entity: { contains: search } },
            { industry: { contains: search } },
            { leadRelationshipPartner: { contains: search } }
          ]
        : undefined
    },
    orderBy: [{ updatedAt: "desc" }],
    take: 500
  });

  return {
    items: items.map((item) => ({
      ...item,
      feeValue: Number(item.feeValue || 0)
    }))
  };
});
