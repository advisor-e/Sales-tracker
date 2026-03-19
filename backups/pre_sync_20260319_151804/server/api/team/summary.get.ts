import { prisma } from "~/server/utils/db";

export default defineEventHandler(async () => {
  const rows = await prisma.pipelineEntry.findMany({
    select: {
      leadStaff: true,
      approachStyle: true,
      secureMeeting: true,
      proposalSent: true,
      proposalValue: true,
      jobSecured: true,
      jobSecuredValue: true
    }
  });

  const byTeam = new Map<string, {
    prospects: number;
    approachesMade: number;
    secureMeetings: number;
    proposalsSent: number;
    totalProposalValue: number;
    engagementsSecured: number;
    totalSecuredValue: number;
  }>();

  for (const row of rows) {
    const key = (row.leadStaff || "Unassigned").trim() || "Unassigned";
    if (!byTeam.has(key)) {
      byTeam.set(key, {
        prospects: 0,
        approachesMade: 0,
        secureMeetings: 0,
        proposalsSent: 0,
        totalProposalValue: 0,
        engagementsSecured: 0,
        totalSecuredValue: 0
      });
    }
    const item = byTeam.get(key)!;
    item.prospects += 1;
    item.approachesMade += row.approachStyle && row.approachStyle.trim() ? 1 : 0;
    item.secureMeetings += row.secureMeeting ? 1 : 0;
    item.proposalsSent += row.proposalSent ? 1 : 0;
    item.totalProposalValue += Number(row.proposalValue || 0);
    item.engagementsSecured += row.jobSecured ? 1 : 0;
    item.totalSecuredValue += Number(row.jobSecuredValue || 0);
  }

  const items = Array.from(byTeam.entries())
    .map(([leadStaff, item]) => {
      const avgApproachConversion = item.prospects > 0 ? item.approachesMade / item.prospects : 0;
      const avgSecuredConversion = item.proposalsSent > 0 ? item.engagementsSecured / item.proposalsSent : 0;
      return {
        leadStaff,
        ...item,
        avgApproachConversion,
        avgSecuredConversion
      };
    })
    .sort((a, b) => a.leadStaff.localeCompare(b.leadStaff));

  return { items };
});
