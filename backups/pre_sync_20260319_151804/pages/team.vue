<script setup lang="ts">
import type { TeamSummaryRow } from "~/types/sales";

const rows = ref<TeamSummaryRow[]>([]);
const loading = ref(false);
const errorText = ref("");

async function loadRows() {
  loading.value = true;
  errorText.value = "";
  try {
    const res = await $fetch<{ items: TeamSummaryRow[] }>("/api/team/summary");
    rows.value = res.items;
  } catch (error: any) {
    errorText.value = String(error?.data?.message || error?.message || "Failed to load team summary");
  } finally {
    loading.value = false;
  }
}

onMounted(loadRows);
</script>

<template>
  <section class="page-wrap">
    <header class="page-head">
      <h1>Team</h1>
      <button @click="loadRows">Refresh</button>
    </header>

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
            <th>Secured Conv</th>
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
            <td>{{ (row.avgSecuredConversion * 100).toFixed(1) }}%</td>
          </tr>
        </tbody>
      </table>
    </section>
  </section>
</template>

<style scoped>
.page-wrap { display: grid; gap: 1rem; }
.page-head { display: flex; justify-content: space-between; align-items: center; }
.card { border: 1px solid #dbeafe; border-radius: 10px; background: #fff; padding: 0.8rem; overflow: auto; }
table { width: 100%; border-collapse: collapse; min-width: 980px; }
th, td { border-bottom: 1px solid #e2e8f0; padding: 0.5rem; text-align: left; }
.error { color: #b91c1c; }
</style>
