import prismaModule from "@prisma/client";
import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

const { BlogPostKind } = prismaModule;

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const query = getQuery(event);
  const kindParam = String(query.kind || "draft").toLowerCase();
  const kind = kindParam === "final" ? BlogPostKind.final : BlogPostKind.draft;
  const search = String(query.search || "").trim();
  const pinnedOnly = String(query.pinnedOnly || "false").toLowerCase() === "true";

  const posts = await prisma.blogPost.findMany({
    where: {
      userId: user.id,
      kind,
      isPinned: pinnedOnly ? true : undefined,
      OR: search
        ? [
            { title: { contains: search } },
            { topic: { contains: search } },
            { selectedPerson: { contains: search } }
          ]
        : undefined
    },
    orderBy: [{ isPinned: "desc" }, { updatedAt: "desc" }]
  });

  return { items: posts };
});
