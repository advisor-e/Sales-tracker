<script setup lang="ts">
import type { DashboardMetrics } from "~/types/sales";

const metrics = ref<DashboardMetrics | null>(null);
const loading = ref(false);
const errorText = ref("");

async function loadMetrics() {
  loading.value = true;
  errorText.value = "";
  try {
    metrics.value = await $fetch<DashboardMetrics>("/api/dashboard/metrics");
  } catch (error: any) {
    errorText.value = String(error?.data?.message || error?.message || "Failed to load dashboard metrics");
  } finally {
    loading.value = false;
  }
}

onMounted(loadMetrics);
</script>

<template>
  <section class="page-wrap">
    <header class="page-head">
      <h1>Dashboard</h1>
      <button @click="loadMetrics">Refresh</button>
    </header>

    <p v-if="errorText" class="error">{{ errorText }}</p>
    <p v-if="loading">Loading metrics...</p>

    <div v-if="metrics" class="kpi-grid">
      <article class="kpi"><h3>Total Prospects</h3><p>{{ metrics.totalProspects }}</p></article>
      <article class="kpi"><h3>Active Prospects</h3><p>{{ metrics.activeProspects }}</p></article>
      <article class="kpi"><h3>Secured Jobs</h3><p>{{ metrics.securedJobs }}</p></article>
      <article class="kpi"><h3>Total Proposal Value</h3><p>${{ metrics.totalProposalValue.toLocaleString() }}</p></article>
      <article class="kpi"><h3>Total Secured Value</h3><p>${{ metrics.totalSecuredValue.toLocaleString() }}</p></article>
      <article class="kpi"><h3>COI Relationships</h3><p>{{ metrics.totalCoi }}</p></article>
      <article class="kpi"><h3>Total Referrals</h3><p>{{ metrics.totalReferrals }}</p></article>
      <article class="kpi"><h3>Total Converted</h3><p>{{ metrics.totalConverted }}</p></article>
    </div>

    <section v-if="metrics" class="card">
      <h2>Pipeline Status Breakdown</h2>
      <table>
        <thead>
          <tr><th>Status</th><th>Count</th></tr>
        </thead>
        <tbody>
          <tr v-for="row in metrics.statusBreakdown" :key="row.status">
            <td>{{ row.status }}</td>
            <td>{{ row.count }}</td>
          </tr>
        </tbody>
      </table>
    </section>
  </section>
</template>

<style scoped>
.page-wrap { display: grid; gap: 1rem; }
.page-head { display: flex; justify-content: space-between; align-items: center; }
.kpi-grid { display: grid; gap: 0.8rem; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); }
.kpi { border: 1px solid #dbeafe; border-radius: 10px; padding: 0.8rem; background: #fff; }
.kpi h3 { margin: 0; font-size: 0.9rem; color: #334155; }
.kpi p { margin: 0.4rem 0 0; font-size: 1.25rem; font-weight: 700; }
.card { border: 1px solid #dbeafe; border-radius: 10px; background: #fff; padding: 0.8rem; }
.error { color: #b91c1c; }
</style>
