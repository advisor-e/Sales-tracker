import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const query = getQuery(event);
  const search = String(query.search || "").trim();

  const items = await prisma.coiEntry.findMany({
    where: {
      userId: user.id,
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
