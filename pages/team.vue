<script setup lang="ts">
import type { TeamSummaryRow } from "~/types/sales";

const rows = ref<TeamSummaryRow[]>([]);
const loading = ref(false);
const errorText = ref("");

const totalSecuredValue = computed(() => rows.value.reduce((sum, row) => sum + row.totalSecuredValue, 0));
const totalProspects = computed(() => rows.value.reduce((sum, row) => sum + row.prospects, 0));
const totalProposals = computed(() => rows.value.reduce((sum, row) => sum + row.proposalsSent, 0));

async function loadRows() {
  loading.value = true;
  errorText.value = "";
  try {
    const res = await $fetch<{ items: TeamSummaryRow[] }>("/api/team/summary");
    rows.value = res.items;
  } catch (error: unknown) {
    const e = error as { statusCode?: number; data?: { statusCode?: number; statusMessage?: string; message?: string }; message?: string };
    const status = Number(e?.statusCode || e?.data?.statusCode || 0);
    if (status === 401) {
      errorText.value = "Session expired. Redirecting to sign in...";
      await navigateTo("/login");
      return;
    }
    errorText.value = String(e?.data?.statusMessage || e?.data?.message || e?.message || "Failed to load team summary");
  } finally {
    loading.value = false;
  }
}

onMounted(loadRows);
</script>

<template>
  <section class="page-wrap">
    <header class="section-banner team-banner">
      <div>
        <p class="banner-kicker">Team</p>
        <h1>Team</h1>
        <p class="subhead">The original team report formulas are now surfaced as one dense management view instead of a plain export table.</p>
      </div>
      <button @click="loadRows">Refresh</button>
    </header>

    <section class="summary-strip">
      <article><span>Team Members</span><strong>{{ rows.length }}</strong></article>
      <article><span>Total Prospects</span><strong>{{ totalProspects }}</strong></article>
      <article><span>Proposals Sent</span><strong>{{ totalProposals }}</strong></article>
      <article><span>Secured Value</span><strong>${{ totalSecuredValue.toLocaleString() }}</strong></article>
    </section>

    <p v-if="errorText" class="error">{{ errorText }}</p>
    <p v-if="loading">Loading team summary...</p>

    <section class="card">
      <table>
        <thead>
          <tr>
            <th>Team Member</th>
            <th>Prospects</th>
            <th>Approaches</th>
            <th>Meetings</th>
            <th>Proposals</th>
            <th>Proposal Value</th>
            <th>Secured</th>
            <th>Secured Value</th>
            <th>Approach Conv</th>
            <th>Avg Proposal</th>
            <th>Secured Conv</th>
            <th>Active</th>
            <th>Await Research</th>
            <th>Completed</th>
            <th>Dead</th>
            <th>On Hold</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="row in rows" :key="row.leadStaff">
            <td>{{ row.leadStaff }}</td>
            <td>{{ row.prospects }}</td>
            <td>{{ row.approachesMade }}</td>
            <td>{{ row.secureMeetings }}</td>
            <td>{{ row.proposalsSent }}</td>
            <td>${{ row.totalProposalValue.toLocaleString() }}</td>
            <td>{{ row.engagementsSecured }}</td>
            <td>${{ row.totalSecuredValue.toLocaleString() }}</td>
            <td>{{ (row.avgApproachConversion * 100).toFixed(1) }}%</td>
            <td>${{ row.avgProposalValue.toLocaleString(undefined, { maximumFractionDigits: 0 }) }}</td>
            <td>{{ (row.avgSecuredConversion * 100).toFixed(1) }}%</td>
            <td>{{ row.active }}</td>
            <td>{{ row.awaitResearch }}</td>
            <td>{{ row.completed }}</td>
            <td>{{ row.dead }}</td>
            <td>{{ row.onHold }}</td>
          </tr>
        </tbody>
      </table>
    </section>
  </section>
</template>

<style scoped>
.page-wrap { display: grid; gap: 0.72rem; }
.section-banner {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 1rem;
  padding: 0.92rem 1rem;
  border-radius: 22px;
  color: #112c7a;
  box-shadow: 0 16px 38px rgba(67, 97, 238, 0.18);
}
.team-banner { background: linear-gradient(135deg, #4cc9f0, #4361ee); }
.banner-kicker { margin: 0 0 0.25rem; font-size: 0.8rem; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; }
.section-banner h1 { margin: 0; font-size: 1.65rem; }
.subhead { margin: 0.2rem 0 0; color: rgba(16, 35, 98, 0.84); font-size: 0.86rem; }
.card {
  border: 1px solid rgba(114, 135, 161, 0.25);
  border-radius: 14px;
  background: linear-gradient(175deg, #ffffff 0%, #f7fbff 100%);
  padding: 0.72rem;
  overflow: auto;
  box-shadow: 0 10px 24px rgba(17, 37, 63, 0.07);
}
.summary-strip { display: grid; gap: 0.58rem; grid-template-columns: repeat(4, minmax(0, 1fr)); }
.summary-strip article {
  border-radius: 18px;
  padding: 0.62rem 0.7rem;
  background: linear-gradient(180deg, #ebf5ff, #dbeafe);
  box-shadow: 0 12px 28px rgba(67, 97, 238, 0.12);
}
.summary-strip span { display: block; font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.08em; color: #1d4ed8; }
.summary-strip strong { display: block; margin-top: 0.2rem; font-size: 1.2rem; color: #1e3a8a; }
button {
  border: 1px solid rgba(15, 122, 138, 0.35);
  border-radius: 10px;
  padding: 0.34rem 0.48rem;
  background: rgba(255, 255, 255, 0.82);
  color: #1e3a8a;
  font-size: 0.8rem;
  font-weight: 700;
  line-height: 1.1;
  cursor: pointer;
}
table { width: 100%; border-collapse: collapse; min-width: 980px; }
th, td { border-bottom: 1px solid #e2e8f0; padding: 0.38rem 0.45rem; text-align: left; font-size: 0.82rem; line-height: 1.25; }
thead th {
  background: #f1f6fb;
  color: #1e3d60;
  font-size: 0.72rem;
  text-transform: uppercase;
  letter-spacing: 0.06em;
}
.error { color: #b91c1c; font-weight: 700; }
@media (max-width: 980px) { .summary-strip { grid-template-columns: 1fr; } }
</style>
