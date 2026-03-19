import { BlogPostKind } from "@prisma/client";
import { z } from "zod";
import { prisma } from "~/server/utils/db";

const schema = z.object({
  kind: z.enum(["draft", "final"]),
  title: z.string().min(1).max(255),
  topic: z.string().min(1).max(255),
  audience: z.string().min(1).max(255),
  objective: z.string().min(1).max(255),
  tone: z.string().min(1).max(80),
  length: z.string().min(1).max(80),
  cta: z.string().min(1).max(255),
  selectedPerson: z.string().max(255).optional().nullable(),
  targetMode: z.string().max(120).optional().nullable(),
  outlineText: z.string().min(1),
  finalText: z.string().optional().nullable(),
  isPinned: z.boolean().optional(),
  metadata: z.any().optional()
});

export default defineEventHandler(async (event) => {
  const body = await readBody(event);
  const payload = schema.parse(body);

  const item = await prisma.blogPost.create({
    data: {
      kind: payload.kind === "final" ? BlogPostKind.final : BlogPostKind.draft,
      title: payload.title,
      topic: payload.topic,
      audience: payload.audience,
      objective: payload.objective,
      tone: payload.tone,
      length: payload.length,
      cta: payload.cta,
      selectedPerson: payload.selectedPerson || null,
      targetMode: payload.targetMode || null,
      outlineText: payload.outlineText,
      finalText: payload.finalText || null,
      isPinned: payload.isPinned ?? false,
      metadataJson: payload.metadata ?? null
    }
  });

  return { item };
});
