import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const items = await prisma.blogInput.findMany({
    where: { userId: user.id },
    orderBy: [{ updatedAt: "desc" }],
    take: 120
  });
  return { items };
});
