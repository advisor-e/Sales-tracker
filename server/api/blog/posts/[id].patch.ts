import { z } from "zod";
import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

const schema = z.object({
  isPinned: z.boolean().optional(),
  outlineText: z.string().optional(),
  finalText: z.string().nullable().optional(),
  title: z.string().min(1).max(255).optional()
});

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const id = Number(getRouterParam(event, "id"));
  if (!Number.isFinite(id) || id <= 0) {
    throw createError({ statusCode: 400, statusMessage: "Invalid id" });
  }

  const body = schema.parse(await readBody(event));

  const existing = await prisma.blogPost.findFirst({ where: { id, userId: user.id } });
  if (!existing) {
    throw createError({ statusCode: 404, statusMessage: "Not found" });
  }

  const item = await prisma.blogPost.update({
    where: { id },
    data: {
      isPinned: body.isPinned,
      outlineText: body.outlineText,
      finalText: body.finalText,
      title: body.title
    }
  });

  return { item };
});
