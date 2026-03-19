import { z } from "zod";
import { generateDraft } from "~/server/utils/openai";

const schema = z.object({
  topic: z.string().min(1),
  audience: z.string().min(1),
  objective: z.string().min(1),
  tone: z.enum(["Professional", "Friendly", "Confident", "Educational"]),
  length: z.enum(["Short", "Medium", "Long"]),
  cta: z.string().min(1),
  principles: z.array(z.object({ title: z.string(), details: z.array(z.string()) })).min(1)
});

export default defineEventHandler(async (event) => {
  const payload = schema.parse(await readBody(event));
  return await generateDraft(payload);
});
