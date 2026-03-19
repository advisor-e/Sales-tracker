import { prisma } from "~/server/utils/db";

export default defineEventHandler(async (event) => {
  const id = Number(getRouterParam(event, "id"));
  if (!Number.isFinite(id) || id <= 0) {
    throw createError({ statusCode: 400, statusMessage: "Invalid id" });
  }

  await prisma.pipelineEntry.delete({ where: { id } });
  return { ok: true };
});
