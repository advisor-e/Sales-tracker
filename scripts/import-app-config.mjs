import { readFile } from "node:fs/promises";
import prismaModule from "@prisma/client";

const { PrismaClient, BlogPostKind } = prismaModule;
const prisma = new PrismaClient();

function asString(value, fallback = "") {
  if (value === null || value === undefined) {
    return fallback;
  }
  return String(value).trim();
}

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function signatureForInput(item) {
  if (asString(item._signature)) {
    return asString(item._signature);
  }
  return JSON.stringify({
    topic: asString(item.topic),
    audience: asString(item.audience),
    objective: asString(item.objective),
    tone: asString(item.tone),
    length: asString(item.length),
    cta: asString(item.cta),
    principles: asArray(item.principles)
  });
}

async function main() {
  const configPath = process.argv[2] || "../sales-tracker-app/app_config.json";
  const userEmail = asString(process.env.IMPORT_USER_EMAIL || process.env.ADMIN_EMAIL || "admin@example.com").toLowerCase();

  const user = await prisma.user.findUnique({ where: { email: userEmail } });
  if (!user) {
    throw new Error(`User not found for email ${userEmail}. Create user first.`);
  }

  const raw = await readFile(configPath, "utf8");
  const parsed = JSON.parse(raw);

  const inputHistory = asArray(parsed.blog_input_history).filter((item) => item && typeof item === "object");
  const draftHistory = asArray(parsed.blog_draft_history).filter((item) => item && typeof item === "object");

  for (const item of inputHistory) {
    const signature = signatureForInput(item);
    const existing = await prisma.blogInput.findFirst({ where: { userId: user.id, signature } });

    const payload = {
      userId: user.id,
      signature,
      topic: asString(item.topic, "Untitled"),
      audience: asString(item.audience, "General audience"),
      objective: asString(item.objective, "Inform readers"),
      tone: asString(item.tone, "Professional"),
      length: asString(item.length, "Medium"),
      cta: asString(item.cta, "book a short strategy call"),
      principlesJson: asArray(item.principles),
      selectedPerson: asString(item.selected_person) || null,
      targetMode: asString(item.target_mode) || null,
      styleStrength: asString(item.style_strength) || null,
      styleTitlesJson: asArray(item.style_example_titles),
      lengthRangesJson: item.length_ranges || null,
      styleThresholdsJson: item.style_conf_thresholds || null
    };

    if (existing) {
      await prisma.blogInput.update({ where: { id: existing.id }, data: payload });
    } else {
      await prisma.blogInput.create({ data: payload });
    }
  }

  for (const item of draftHistory) {
    const createdAt = asString(item.created_at);
    const title = asString(item.title, asString(item.topic, "Imported post"));
    const outlineText = asString(item.outline_text || item.blog_text);
    const finalText = asString(item.final_post_text);

    if (outlineText) {
      await prisma.blogPost.create({
        data: {
          userId: user.id,
          kind: BlogPostKind.draft,
          title,
          topic: asString(item.topic, "Imported topic"),
          audience: asString(item.audience, "General audience"),
          objective: asString(item.objective, "Inform readers"),
          tone: asString(item.tone, "Professional"),
          length: asString(item.length, "Medium"),
          cta: asString(item.cta, "book a short strategy call"),
          selectedPerson: asString(item.selected_person) || null,
          targetMode: asString(item.target_mode) || null,
          outlineText,
          finalText: null,
          isPinned: Boolean(item.is_pinned),
          metadataJson: {
            importedFrom: "app_config.json",
            createdAtLegacy: createdAt || null
          }
        }
      });
    }

    if (finalText) {
      await prisma.blogPost.create({
        data: {
          userId: user.id,
          kind: BlogPostKind.final,
          title,
          topic: asString(item.topic, "Imported topic"),
          audience: asString(item.audience, "General audience"),
          objective: asString(item.objective, "Inform readers"),
          tone: asString(item.tone, "Professional"),
          length: asString(item.length, "Medium"),
          cta: asString(item.cta, "book a short strategy call"),
          selectedPerson: asString(item.selected_person) || null,
          targetMode: asString(item.target_mode) || null,
          outlineText: outlineText || finalText,
          finalText,
          isPinned: Boolean(item.is_pinned),
          metadataJson: {
            importedFrom: "app_config.json",
            createdAtLegacy: createdAt || null
          }
        }
      });
    }
  }

  console.log(`Imported ${inputHistory.length} blog inputs and ${draftHistory.length} draft-history entries for ${userEmail}`);
}

main()
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
