import { prisma } from "~/server/utils/db";

export default defineEventHandler(async () => {
  const items = await prisma.blogInput.findMany({
    orderBy: [{ updatedAt: "desc" }],
    take: 120
  });
  return { items };
});
