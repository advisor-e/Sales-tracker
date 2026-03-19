import { z } from "zod";
import { generateFinal } from "~/server/utils/openai";

const schema = z.object({
  outlineText: z.string().min(1),
  topic: z.string().min(1),
  audience: z.string().min(1),
  objective: z.string().min(1),
  tone: z.enum(["Professional", "Friendly", "Confident", "Educational"]),
  cta: z.string().min(1),
  polishLevel: z.enum(["Standard", "Strong", "Premium"])
});

export default defineEventHandler(async (event) => {
  const payload = schema.parse(await readBody(event));
  return await generateFinal(payload);
});
