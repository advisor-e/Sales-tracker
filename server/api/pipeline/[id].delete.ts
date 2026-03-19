import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

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

  await prisma.pipelineEntry.delete({ where: { id } });
  return { ok: true };
});
