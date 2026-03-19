import OpenAI from "openai";
import { dedupeParagraphs } from "./text";
import type { BlogDraftRequest, BlogFinalRequest } from "~/types/blog";

function getClient() {
  const config = useRuntimeConfig();
  const key = String(config.openaiApiKey || "").trim();
  if (!key) {
    return null;
  }
  return new OpenAI({ apiKey: key });
}

function buildDraftTemplate(payload: BlogDraftRequest): string {
  const principleBlock = payload.principles
    .filter((p) => p.title.trim() || p.details.some((d) => d.trim()))
    .slice(0, 3)
    .map((p, index) => {
      const detailLines = p.details
        .filter((d) => d.trim())
        .slice(0, 3)
        .map((d) => `- ${d}`)
        .join("\n");
      return `## Principle ${index + 1}: ${p.title || "Practical focus"}\n${detailLines}`;
    })
    .join("\n\n");

  return dedupeParagraphs(`# ${payload.topic}: Practical guide for ${payload.audience}\n\nThis outline is focused on ${payload.objective.toLowerCase()}.\n\n${principleBlock}\n\n## Final takeaway\n${payload.cta}.`);
}

function buildFinalTemplate(payload: BlogFinalRequest): string {
  return dedupeParagraphs(`# ${payload.topic}\n\n${payload.outlineText}\n\nIn summary, the key objective is ${payload.objective.toLowerCase()}. ${payload.cta}.`);
}

export async function generateDraft(payload: BlogDraftRequest): Promise<{ text: string; source: "ai" | "template"; error?: string }> {
  const client = getClient();
  if (!client) {
    return { text: buildDraftTemplate(payload), source: "template", error: "OpenAI key not configured" };
  }

  const principleLines = payload.principles
    .map((p, i) => {
      const details = p.details.filter(Boolean).map((d) => `  - ${d}`).join("\n");
      return `Section ${i + 1}: ${p.title}\n${details}`;
    })
    .join("\n");

  const system = "You are an expert financial content writer producing substantive markdown blog outlines.";
  const user = `Create a blog outline in markdown.\nTopic: ${payload.topic}\nAudience: ${payload.audience}\nObjective: ${payload.objective}\nTone: ${payload.tone}\nLength: ${payload.length}\nCTA: ${payload.cta}\n\nUse these principles:\n${principleLines}`;

  try {
    const response = await client.chat.completions.create({
      model: "gpt-4o",
      temperature: 0.7,
      max_tokens: 1800,
      messages: [
        { role: "system", content: system },
        { role: "user", content: user }
      ]
    });

    const text = dedupeParagraphs(response.choices[0]?.message?.content || "").trim();
    if (!text) {
      return { text: buildDraftTemplate(payload), source: "template", error: "Empty AI response" };
    }
    return { text, source: "ai" };
  } catch (error: any) {
    return { text: buildDraftTemplate(payload), source: "template", error: String(error?.message || error) };
  }
}

export async function generateFinal(payload: BlogFinalRequest): Promise<{ text: string; source: "ai" | "template"; error?: string }> {
  const client = getClient();
  if (!client) {
    return { text: buildFinalTemplate(payload), source: "template", error: "OpenAI key not configured" };
  }

  const system = "You are an elite financial content writer. Convert outlines into polished markdown articles without filler.";
  const user = `Turn this outline into a complete markdown article.\nTopic: ${payload.topic}\nAudience: ${payload.audience}\nObjective: ${payload.objective}\nTone: ${payload.tone}\nPolish level: ${payload.polishLevel}\nCTA: ${payload.cta}\n\nOutline:\n${payload.outlineText}`;

  try {
    const response = await client.chat.completions.create({
      model: "gpt-4o",
      temperature: 0.65,
      max_tokens: 2200,
      messages: [
        { role: "system", content: system },
        { role: "user", content: user }
      ]
    });

    const text = dedupeParagraphs(response.choices[0]?.message?.content || "").trim();
    if (!text) {
      return { text: buildFinalTemplate(payload), source: "template", error: "Empty AI response" };
    }
    return { text, source: "ai" };
  } catch (error: any) {
    return { text: buildFinalTemplate(payload), source: "template", error: String(error?.message || error) };
  }
}
