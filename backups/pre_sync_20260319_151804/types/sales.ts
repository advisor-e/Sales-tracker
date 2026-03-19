export type PipelineEntry = {
  id: number;
  prospectName: string;
  businessName?: string | null;
  partner?: string | null;
  leadStaff?: string | null;
  prospectStatus: string;
  relationshipType?: string | null;
  prospectSource?: string | null;
  approachStyle?: string | null;
  approachDate?: string | null;
  secureMeeting: boolean;
  proposalSent: boolean;
  proposalValue: number;
  jobSecured: boolean;
  jobSecuredValue: number;
  comments?: string | null;
  coiInvolved?: string | null;
  createdAt: string;
  updatedAt: string;
};

export type CoiEntry = {
  id: number;
  coiName: string;
  email?: string | null;
  cell?: string | null;
  entity?: string | null;
  position?: string | null;
  industry?: string | null;
  leadRelationshipPartner?: string | null;
  relationshipSupport?: string | null;
  couldWe: boolean;
  howWouldWe: boolean;
  willWe: boolean;
  testReview: boolean;
  totalReferrals: number;
  totalConverted: number;
  feeValue: number;
  createdAt: string;
  updatedAt: string;
};

export type TeamSummaryRow = {
  leadStaff: string;
  prospects: number;
  approachesMade: number;
  secureMeetings: number;
  proposalsSent: number;
  totalProposalValue: number;
  engagementsSecured: number;
  totalSecuredValue: number;
  avgApproachConversion: number;
  avgSecuredConversion: number;
};

export type DashboardMetrics = {
  totalProspects: number;
  activeProspects: number;
  securedJobs: number;
  totalProposalValue: number;
  totalSecuredValue: number;
  totalCoi: number;
  totalReferrals: number;
  totalConverted: number;
  statusBreakdown: Array<{ status: string; count: number }>;
};
