import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const [
    totalProspects,
    activeProspects,
    securedJobs,
    proposalAgg,
    securedAgg,
    totalCoi,
    coiAgg,
    statusRows,
    sourceRows,
    approachCount,
    meetingCount,
    proposalCount,
    staffRows,
    monthlyRows,
    coiIndustryRows
  ] = await Promise.all([
    prisma.pipelineEntry.count({ where: { userId: user.id } }),
    prisma.pipelineEntry.count({ where: { userId: user.id, prospectStatus: "Active" } }),
    prisma.pipelineEntry.count({ where: { userId: user.id, jobSecured: true } }),
    prisma.pipelineEntry.aggregate({ where: { userId: user.id }, _sum: { proposalValue: true } }),
    prisma.pipelineEntry.aggregate({ where: { userId: user.id }, _sum: { jobSecuredValue: true } }),
    prisma.coiEntry.count({ where: { userId: user.id } }),
    prisma.coiEntry.aggregate({ where: { userId: user.id }, _sum: { totalReferrals: true, totalConverted: true } }),
    prisma.pipelineEntry.groupBy({ by: ["prospectStatus"], where: { userId: user.id }, _count: { _all: true }, orderBy: { prospectStatus: "asc" } }),
    prisma.pipelineEntry.groupBy({ by: ["prospectSource"], where: { userId: user.id }, _count: { _all: true } }),
    prisma.pipelineEntry.count({ where: { userId: user.id, NOT: { approachStyle: null } } }),
    prisma.pipelineEntry.count({ where: { userId: user.id, secureMeeting: true } }),
    prisma.pipelineEntry.count({ where: { userId: user.id, proposalSent: true } }),
    prisma.pipelineEntry.findMany({ where: { userId: user.id }, select: { leadStaff: true, jobSecuredValue: true } }),
    prisma.pipelineEntry.findMany({ where: { userId: user.id, approachDate: { not: null } }, select: { approachDate: true, jobSecuredValue: true }, orderBy: { approachDate: "asc" } }),
    prisma.coiEntry.findMany({ where: { userId: user.id }, select: { industry: true } })
  ]);

  const sourceBreakdown = sourceRows
    .map((row) => {
      const raw = String(row.prospectSource || "").trim();
      return {
        source: raw || "Unknown",
        count: row._count._all
      };
    })
    .sort((a, b) => b.count - a.count);

  const staffTotals = new Map<string, number>();
  for (const row of staffRows) {
    const key = String(row.leadStaff || "").trim() || "Unassigned";
    staffTotals.set(key, (staffTotals.get(key) || 0) + Number(row.jobSecuredValue || 0));
  }
  const staffSecuredBreakdown = Array.from(staffTotals.entries())
    .map(([leadStaff, value]) => ({ leadStaff, value }))
    .sort((a, b) => b.value - a.value);

  const monthlyMap = new Map<string, number>();
  for (const row of monthlyRows) {
    if (!row.approachDate) {
      continue;
    }
    const d = new Date(row.approachDate);
    const month = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    monthlyMap.set(month, (monthlyMap.get(month) || 0) + Number(row.jobSecuredValue || 0));
  }
  const monthlySecuredTrend = Array.from(monthlyMap.entries())
    .map(([month, value]) => ({ month, value }))
    .sort((a, b) => a.month.localeCompare(b.month));

  const industryMap = new Map<string, number>();
  for (const row of coiIndustryRows) {
    const key = String(row.industry || "").trim();
    if (!key) {
      continue;
    }
    industryMap.set(key, (industryMap.get(key) || 0) + 1);
  }
  const coiIndustryBreakdown = Array.from(industryMap.entries())
    .map(([industry, relationships]) => ({ industry, relationships }))
    .sort((a, b) => b.relationships - a.relationships);

  return {
    approaches: approachCount,
    meetingsSecured: meetingCount,
    proposalsSent: proposalCount,
    workSecured: Number(securedAgg._sum.jobSecuredValue || 0),
    totalProspects,
    activeProspects,
    securedJobs,
    totalProposalValue: Number(proposalAgg._sum.proposalValue || 0),
    totalSecuredValue: Number(securedAgg._sum.jobSecuredValue || 0),
    totalCoi,
    totalReferrals: coiAgg._sum.totalReferrals || 0,
    totalConverted: coiAgg._sum.totalConverted || 0,
    statusBreakdown: statusRows.map((row) => ({ status: row.prospectStatus, count: row._count._all })),
    sourceBreakdown,
    staffSecuredBreakdown,
    monthlySecuredTrend,
    coiIndustryBreakdown
  };
});
