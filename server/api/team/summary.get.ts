import { prisma } from "~/server/utils/db";
import { requireUser } from "~/server/utils/auth";

export default defineEventHandler(async (event) => {
  const user = await requireUser(event);
  const rows = await prisma.pipelineEntry.findMany({
    where: { userId: user.id },
    select: {
      leadStaff: true,
      prospectStatus: true,
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
    active: number;
    awaitResearch: number;
    completed: number;
    dead: number;
    onHold: number;
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
        totalSecuredValue: 0,
        active: 0,
        awaitResearch: 0,
        completed: 0,
        dead: 0,
        onHold: 0
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

    const status = String(row.prospectStatus || "").trim();
    if (status === "Active") item.active += 1;
    else if (status === "Await Research") item.awaitResearch += 1;
    else if (status === "Completed") item.completed += 1;
    else if (status === "Dead") item.dead += 1;
    else if (status === "On Hold") item.onHold += 1;
  }

  const items = Array.from(byTeam.entries())
    .map(([leadStaff, item]) => {
      const avgApproachConversion = item.prospects > 0 ? item.approachesMade / item.prospects : 0;
      const avgProposalValue = item.proposalsSent > 0 ? item.totalProposalValue / item.proposalsSent : 0;
      const avgSecuredConversion = item.proposalsSent > 0 ? item.engagementsSecured / item.proposalsSent : 0;
      return {
        leadStaff,
        ...item,
        avgApproachConversion,
        avgProposalValue,
        avgSecuredConversion
      };
    })
    .sort((a, b) => a.leadStaff.localeCompare(b.leadStaff));

  return { items };
});
