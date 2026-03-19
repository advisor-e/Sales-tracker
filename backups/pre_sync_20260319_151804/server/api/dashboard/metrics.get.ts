import { prisma } from "~/server/utils/db";

export default defineEventHandler(async () => {
  const [totalProspects, activeProspects, securedJobs, proposalAgg, securedAgg, totalCoi, coiAgg, statusRows] = await Promise.all([
    prisma.pipelineEntry.count(),
    prisma.pipelineEntry.count({ where: { prospectStatus: "Active" } }),
    prisma.pipelineEntry.count({ where: { jobSecured: true } }),
    prisma.pipelineEntry.aggregate({ _sum: { proposalValue: true } }),
    prisma.pipelineEntry.aggregate({ _sum: { jobSecuredValue: true } }),
    prisma.coiEntry.count(),
    prisma.coiEntry.aggregate({ _sum: { totalReferrals: true, totalConverted: true } }),
    prisma.pipelineEntry.groupBy({ by: ["prospectStatus"], _count: { _all: true }, orderBy: { prospectStatus: "asc" } })
  ]);

  return {
    totalProspects,
    activeProspects,
    securedJobs,
    totalProposalValue: Number(proposalAgg._sum.proposalValue || 0),
    totalSecuredValue: Number(securedAgg._sum.jobSecuredValue || 0),
    totalCoi,
    totalReferrals: coiAgg._sum.totalReferrals || 0,
    totalConverted: coiAgg._sum.totalConverted || 0,
    statusBreakdown: statusRows.map((row) => ({ status: row.prospectStatus, count: row._count._all }))
  };
});
