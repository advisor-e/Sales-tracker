import prismaModule from "@prisma/client";
import xlsx from "xlsx";

const { PrismaClient } = prismaModule;

const prisma = new PrismaClient();

const STATUS_OPTIONS = new Set(["Active", "Await Research", "Completed", "Dead", "On Hold"]);

function asString(value, fallback = "") {
  if (value === null || value === undefined) {
    return fallback;
  }
  return String(value).trim();
}

function toBool(value) {
  const text = asString(value).toLowerCase();
  return text === "yes" || text === "true" || text === "1";
}

function toNumber(value) {
  const text = asString(value).replace(/[$,\s]/g, "");
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function bestHeaderRow(rows, markers) {
  let bestIndex = 0;
  let bestScore = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i += 1) {
    const rowSet = new Set((rows[i] || []).map((value) => asString(value)).filter(Boolean));
    let score = 0;
    for (const marker of markers) {
      if (rowSet.has(marker)) {
        score += 1;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestIndex;
}

function sheetRows(workbook, name, markers) {
  const ws = workbook.Sheets[name];
  if (!ws) {
    return [];
  }
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) {
    return [];
  }
  const headerIndex = bestHeaderRow(rows, markers);
  const headers = rows[headerIndex].map((value) => asString(value));
  const out = [];

  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const source = rows[i] || [];
    const item = {};
    let nonEmpty = 0;
    for (let col = 0; col < headers.length; col += 1) {
      const key = headers[col];
      if (!key) {
        continue;
      }
      item[key] = source[col];
      if (asString(source[col])) {
        nonEmpty += 1;
      }
    }
    if (nonEmpty > 0) {
      out.push(item);
    }
  }

  return out;
}

async function main() {
  const workbookPath = process.argv[2] || "../Get.1a.Sales Tracker.polished.xlsx";
  const userEmail = asString(process.env.IMPORT_USER_EMAIL || process.env.ADMIN_EMAIL || "admin@example.com").toLowerCase();

  const user = await prisma.user.findUnique({ where: { email: userEmail } });
  if (!user) {
    throw new Error(`User not found for email ${userEmail}. Create user first.`);
  }

  const wb = xlsx.readFile(workbookPath);

  const salesRows = sheetRows(wb, "Sales Activity", [
    "Prospect Name",
    "Business Name",
    "Lead Staff (Client Manager)",
    "Prospect Status",
    "Approach Date"
  ]);

  const coiRows = sheetRows(wb, "COI Development", [
    "COI Name",
    "Email",
    "Entity",
    "Lead Relationship Partner",
    "Total Referrals"
  ]);

  await prisma.pipelineEntry.deleteMany({ where: { userId: user.id } });
  await prisma.coiEntry.deleteMany({ where: { userId: user.id } });

  for (const row of salesRows) {
    const prospectName = asString(row["Prospect Name"]);
    if (!prospectName) {
      continue;
    }

    const status = asString(row["Prospect Status"], "Active");
    await prisma.pipelineEntry.create({
      data: {
        userId: user.id,
        prospectName,
        businessName: asString(row["Business Name"]) || null,
        partner: asString(row["Partner"]) || null,
        leadStaff: asString(row["Lead Staff (Client Manager)"]) || null,
        prospectStatus: STATUS_OPTIONS.has(status) ? status : "Active",
        relationshipType: asString(row["Relationship Type"]) || null,
        prospectSource: asString(row["Prospect Source"]) || null,
        approachStyle: asString(row["Approach Style"]) || null,
        approachDate: parseDate(row["Approach Date"]),
        secureMeeting: toBool(row["Secure Meeting"]),
        proposalSent: toBool(row["Proposal Sent"]),
        proposalValue: toNumber(row["Proposal Value"]),
        jobSecured: toBool(row["Job Secured"]),
        jobSecuredValue: toNumber(row["Job Secured Value"]),
        comments: asString(row["Comments / Next Move"]) || null,
        coiInvolved: asString(row["COI Involved"]) || null
      }
    });
  }

  for (const row of coiRows) {
    const name = asString(row["COI Name"]);
    if (!name) {
      continue;
    }

    await prisma.coiEntry.create({
      data: {
        userId: user.id,
        coiName: name,
        email: asString(row["Email"]) || null,
        cell: asString(row["Cell #"]) || null,
        entity: asString(row["Entity"]) || null,
        position: asString(row["Position"]) || null,
        industry: asString(row["Industry"]) || null,
        leadRelationshipPartner: asString(row["Lead Relationship Partner"]) || null,
        relationshipSupport: asString(row["Relationship Support"]) || null,
        couldWe: toBool(row["Could We"]),
        howWouldWe: toBool(row["How Would We"]),
        willWe: toBool(row["Will We"]),
        testReview: toBool(row["Test/ Review"]),
        totalReferrals: Math.trunc(toNumber(row["Total Referrals"])),
        totalConverted: Math.trunc(toNumber(row["Total Converted"])),
        feeValue: toNumber(row["Fee Value"])
      }
    });
  }

  console.log(`Imported ${salesRows.length} pipeline rows and ${coiRows.length} COI rows for ${userEmail}`);
}

main()
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
