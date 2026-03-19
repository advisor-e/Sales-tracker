<script setup lang="ts">
import type { DashboardMetrics } from "~/types/sales";

const metrics = ref<DashboardMetrics | null>(null);
const loading = ref(false);
const errorText = ref("");

const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });

const statusChartRows = computed(() => {
  const rows = metrics.value?.statusBreakdown || [];
  const max = Math.max(...rows.map((row) => row.count), 1);
  return rows.map((row) => ({ ...row, width: Math.max(8, Math.round((row.count / max) * 100)) }));
});

const sourceChartRows = computed(() => {
  const rows = metrics.value?.sourceBreakdown || [];
  const max = Math.max(...rows.map((row) => row.count), 1);
  return rows.map((row) => ({ ...row, width: Math.max(8, Math.round((row.count / max) * 100)) }));
});

const staffChartRows = computed(() => {
  const rows = metrics.value?.staffSecuredBreakdown || [];
  const max = Math.max(...rows.map((row) => row.value), 1);
  return rows.map((row) => ({ ...row, width: Math.max(8, Math.round((row.value / max) * 100)) }));
});

const coiIndustryRows = computed(() => {
  const rows = metrics.value?.coiIndustryBreakdown || [];
  const max = Math.max(...rows.map((row) => row.relationships), 1);
  return rows.map((row) => ({ ...row, width: Math.max(8, Math.round((row.relationships / max) * 100)) }));
});

const securedRate = computed(() => {
  const total = metrics.value?.totalProspects || 0;
  if (!total) {
    return 0;
  }
  return Math.round(((metrics.value?.securedJobs || 0) / total) * 100);
});

const conversionRate = computed(() => {
  const total = metrics.value?.totalReferrals || 0;
  if (!total) {
    return 0;
  }
  return Math.round(((metrics.value?.totalConverted || 0) / total) * 100);
});

const securedOffset = computed(() => {
  const circumference = 2 * Math.PI * 50;
  return circumference - (Math.min(securedRate.value, 100) / 100) * circumference;
});

const conversionOffset = computed(() => {
  const circumference = 2 * Math.PI * 50;
  return circumference - (Math.min(conversionRate.value, 100) / 100) * circumference;
});

const monthlyLinePoints = computed(() => {
  const rows = metrics.value?.monthlySecuredTrend || [];
  if (!rows.length) {
    return "";
  }
  const max = Math.max(...rows.map((row) => row.value), 1);
  const step = rows.length > 1 ? 220 / (rows.length - 1) : 220;
  return rows
    .map((row, idx) => {
      const x = idx * step;
      const y = 110 - (row.value / max) * 100;
      return `${x},${y}`;
    })
    .join(" ");
});

async function loadMetrics() {
  loading.value = true;
  errorText.value = "";
  try {
    metrics.value = await $fetch<DashboardMetrics>("/api/dashboard/metrics");
  } catch (error: unknown) {
    const e = error as { statusCode?: number; data?: { statusCode?: number; statusMessage?: string; message?: string }; message?: string };
    const status = Number(e?.statusCode || e?.data?.statusCode || 0);
    if (status === 401) {
      errorText.value = "Session expired. Redirecting to sign in...";
      await navigateTo("/login");
      return;
    }
    errorText.value = String(e?.data?.statusMessage || e?.data?.message || e?.message || "Failed to load dashboard metrics");
  } finally {
    loading.value = false;
  }
}

onMounted(loadMetrics);
</script>

<template>
  <section class="page-wrap">
    <header class="section-banner dashboard-banner">
      <div>
        <p class="banner-kicker">Dashboard</p>
        <h1>Dashboard</h1>
        <p class="subhead">Headline numbers and charts now sit together so the whole picture reads in one pass.</p>
      </div>
      <button @click="loadMetrics">Refresh Data</button>
    </header>

    <p v-if="errorText" class="error">{{ errorText }}</p>
    <p v-if="loading" class="loading">Loading metrics...</p>

    <section v-if="metrics" class="kpi-main-grid">
      <article class="kpi tone-main"><h3>Approaches</h3><p>{{ metrics.approaches }}</p></article>
      <article class="kpi tone-main"><h3>Meetings Secured</h3><p>{{ metrics.meetingsSecured }}</p></article>
      <article class="kpi tone-main"><h3>Proposals Sent</h3><p>{{ metrics.proposalsSent }}</p></article>
      <article class="kpi tone-main"><h3>Work Secured</h3><p>{{ money.format(metrics.workSecured) }}</p></article>
    </section>

    <section v-if="metrics" class="kpi-grid">
      <article class="kpi tone-blue"><h3>Total Prospects</h3><p>{{ metrics.totalProspects }}</p><span>Active: {{ metrics.activeProspects }}</span></article>
      <article class="kpi tone-green"><h3>Secured Jobs</h3><p>{{ metrics.securedJobs }}</p><span>Win rate: {{ securedRate }}%</span></article>
      <article class="kpi tone-cyan"><h3>Total Proposal Value</h3><p>{{ money.format(metrics.totalProposalValue) }}</p><span>Pipeline value</span></article>
      <article class="kpi tone-indigo"><h3>Total Secured Value</h3><p>{{ money.format(metrics.totalSecuredValue) }}</p><span>Closed value</span></article>
      <article class="kpi tone-gold"><h3>COI Relationships</h3><p>{{ metrics.totalCoi }}</p><span>Referrals: {{ metrics.totalReferrals }}</span></article>
      <article class="kpi tone-rose"><h3>Total Converted</h3><p>{{ metrics.totalConverted }}</p><span>Conversion: {{ conversionRate }}%</span></article>
    </section>

    <section v-if="metrics" class="chart-grid">
      <article class="card chart-card">
        <h2>Work Secured by Staff</h2>
        <ul class="bar-list">
          <li v-for="row in staffChartRows" :key="row.leadStaff">
            <div class="bar-label"><strong>{{ row.leadStaff }}</strong><span>{{ money.format(row.value) }}</span></div>
            <div class="bar-track"><span class="bar-fill" :style="{ width: `${row.width}%` }" /></div>
          </li>
        </ul>
      </article>

      <article class="card chart-card">
        <h2>Monthly Secured Value Trend</h2>
        <div v-if="metrics.monthlySecuredTrend.length" class="trend-wrap">
          <svg viewBox="0 0 220 120" class="trend-svg" aria-label="Monthly secured value trend">
            <polyline :points="monthlyLinePoints" class="trend-line" />
          </svg>
          <div class="trend-labels">
            <span>{{ metrics.monthlySecuredTrend[0].month }}</span>
            <span>{{ metrics.monthlySecuredTrend[metrics.monthlySecuredTrend.length - 1].month }}</span>
          </div>
        </div>
        <p v-else class="empty">No dated pipeline records yet.</p>
      </article>

      <article class="card chart-card">
        <h2>Prospect Status Mix</h2>
        <ul class="bar-list">
          <li v-for="row in statusChartRows" :key="row.status">
            <div class="bar-label"><strong>{{ row.status }}</strong><span>{{ row.count }}</span></div>
            <div class="bar-track"><span class="bar-fill" :style="{ width: `${row.width}%` }" /></div>
          </li>
        </ul>
      </article>

      <article class="card chart-card">
        <h2>New Prospect Source</h2>
        <ul class="bar-list">
          <li v-for="row in sourceChartRows" :key="row.source">
            <div class="bar-label"><strong>{{ row.source }}</strong><span>{{ row.count }}</span></div>
            <div class="bar-track"><span class="bar-fill" :style="{ width: `${row.width}%` }" /></div>
          </li>
        </ul>
      </article>
    </section>

    <section v-if="metrics" class="chart-grid two-col">
      <article class="card chart-card donuts">
        <h2>Conversion Gauges</h2>
        <div class="donut-row">
          <div class="donut-wrap">
            <svg viewBox="0 0 120 120" class="donut-svg" aria-label="Secured jobs rate">
              <circle class="ring-bg" cx="60" cy="60" r="50" />
              <circle class="ring-value ring-secured" cx="60" cy="60" r="50" :stroke-dasharray="2 * Math.PI * 50" :stroke-dashoffset="securedOffset" />
            </svg>
            <div class="donut-text"><strong>{{ securedRate }}%</strong><span>Secured</span></div>
          </div>
          <div class="donut-wrap">
            <svg viewBox="0 0 120 120" class="donut-svg" aria-label="Referral conversion rate">
              <circle class="ring-bg" cx="60" cy="60" r="50" />
              <circle class="ring-value ring-conversion" cx="60" cy="60" r="50" :stroke-dasharray="2 * Math.PI * 50" :stroke-dashoffset="conversionOffset" />
            </svg>
            <div class="donut-text"><strong>{{ conversionRate }}%</strong><span>Referral conversion</span></div>
          </div>
        </div>
      </article>

      <article class="card chart-card">
        <h2>COI Relationships by Industry</h2>
        <ul class="bar-list">
          <li v-for="row in coiIndustryRows" :key="row.industry">
            <div class="bar-label"><strong>{{ row.industry }}</strong><span>{{ row.relationships }}</span></div>
            <div class="bar-track"><span class="bar-fill" :style="{ width: `${row.width}%` }" /></div>
          </li>
        </ul>
      </article>
    </section>
  </section>
</template>

<style scoped>
.page-wrap { display: grid; gap: 0.82rem; }
.section-banner {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 1rem;
  padding: 0.92rem 1rem;
  border-radius: 22px;
  color: #4a2500;
  box-shadow: 0 16px 38px rgba(255, 159, 28, 0.2);
}
.dashboard-banner { background: linear-gradient(135deg, #ffd166, #ff9f1c); }
.banner-kicker { margin: 0 0 0.25rem; font-size: 0.8rem; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; }
.section-banner h1 { margin: 0; font-size: 1.65rem; letter-spacing: 0.02em; }
.subhead { margin: 0.2rem 0 0; color: rgba(74, 37, 0, 0.82); font-size: 0.86rem; }

button {
  border: 1px solid rgba(15, 122, 138, 0.35);
  border-radius: 10px;
  padding: 0.55rem 0.85rem;
  background: linear-gradient(135deg, #e0f6fa, #d8f0f4);
  color: #0a4752;
  font-weight: 700;
  cursor: pointer;
}

.section-banner button {
  background: rgba(255, 255, 255, 0.74);
  border-color: rgba(124, 45, 18, 0.18);
  color: #7c2d12;
}

.loading { color: #4a617f; font-weight: 600; }

.kpi-main-grid { display: grid; gap: 0.9rem; grid-template-columns: repeat(4, minmax(160px, 1fr)); }
.kpi-grid { display: grid; gap: 0.9rem; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); }
.kpi {
  border-radius: 14px;
  padding: 1rem;
  color: #0f2340;
  box-shadow: 0 10px 26px rgba(17, 37, 63, 0.08);
}
.kpi h3 { margin: 0; font-size: 0.76rem; text-transform: uppercase; letter-spacing: 0.06em; opacity: 0.75; }
.kpi p { margin: 0.3rem 0 0.2rem; font-size: 1.45rem; font-weight: 800; }
.kpi span { font-size: 0.78rem; opacity: 0.9; }

.tone-main { background: linear-gradient(155deg, #d8f1f4, #c8ebf0); }
.tone-blue { background: linear-gradient(155deg, #dbf0ff, #cbe9ff); }
.tone-green { background: linear-gradient(155deg, #d8f5ea, #c8eedf); }
.tone-cyan { background: linear-gradient(155deg, #d7f6f8, #c5eef2); }
.tone-indigo { background: linear-gradient(155deg, #e2e8ff, #d5dcff); }
.tone-gold { background: linear-gradient(155deg, #fef1ce, #fde9bc); }
.tone-rose { background: linear-gradient(155deg, #ffe2e2, #ffd3d3); }

.chart-grid { display: grid; gap: 0.9rem; grid-template-columns: 1fr 1fr; }
.chart-grid.two-col { grid-template-columns: 1fr 1fr; }
.card {
  border: 1px solid rgba(114, 135, 161, 0.25);
  border-radius: 14px;
  background: linear-gradient(175deg, #ffffff 0%, #f7fbff 100%);
  box-shadow: 0 10px 30px rgba(17, 37, 63, 0.08);
  padding: 1rem;
}
.chart-card h2 { margin: 0 0 0.8rem; font-size: 1rem; color: #173155; }

.bar-list { list-style: none; padding: 0; margin: 0; display: grid; gap: 0.45rem; }
.bar-label { display: flex; justify-content: space-between; font-size: 0.8rem; color: #1f3c60; margin-bottom: 0.18rem; gap: 0.6rem; }
.bar-track {
  height: 10px;
  border-radius: 999px;
  background: #e8edf3;
  overflow: hidden;
}
.bar-fill {
  display: block;
  height: 100%;
  border-radius: inherit;
  background: linear-gradient(90deg, #0f7a8a 0%, #1ca4b8 100%);
}

.trend-wrap { display: grid; gap: 0.4rem; }
.trend-svg {
  width: 100%;
  max-height: 160px;
  border-radius: 10px;
  background: #f2f7fb;
  border: 1px solid #d9e5f0;
  padding: 0.3rem;
}
.trend-line {
  fill: none;
  stroke: #0f7a8a;
  stroke-width: 3;
  stroke-linecap: round;
  stroke-linejoin: round;
}
.trend-labels { display: flex; justify-content: space-between; font-size: 0.72rem; color: #607692; }
.empty { color: #607692; margin: 0.25rem 0 0; }

.donuts { display: grid; align-content: start; }
.donut-row { display: grid; grid-template-columns: 1fr 1fr; gap: 0.7rem; }
.donut-wrap {
  border-radius: 12px;
  background: #f9fbfd;
  border: 1px solid #e5edf5;
  padding: 0.6rem;
  display: grid;
  place-items: center;
  gap: 0.35rem;
}
.donut-svg { width: 118px; height: 118px; transform: rotate(-90deg); }
.ring-bg { fill: none; stroke: #dbe5ef; stroke-width: 10; }
.ring-value {
  fill: none;
  stroke-width: 10;
  stroke-linecap: round;
  transition: stroke-dashoffset 0.35s ease;
}
.ring-secured { stroke: #10b981; }
.ring-conversion { stroke: #0f7a8a; }
.donut-text { display: grid; place-items: center; margin-top: -84px; pointer-events: none; }
.donut-text strong { font-size: 1.3rem; color: #0f2340; }
.donut-text span { font-size: 0.78rem; color: #4a617f; text-align: center; }

.error { color: #b91c1c; font-weight: 700; }

@media (max-width: 1080px) {
  .kpi-main-grid { grid-template-columns: repeat(2, minmax(140px, 1fr)); }
}

@media (max-width: 960px) {
  .chart-grid,
  .chart-grid.two-col { grid-template-columns: 1fr; }
  .section-banner { flex-direction: column; align-items: flex-start; }
}
</style>
