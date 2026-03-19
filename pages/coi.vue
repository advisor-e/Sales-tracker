<script setup lang="ts">
import type { CoiEntry } from "~/types/sales";

const search = ref("");
const items = ref<CoiEntry[]>([]);
const loading = ref(false);
const errorText = ref("");

const draft = reactive({
  coiName: "",
  email: "",
  entity: "",
  industry: "",
  leadRelationshipPartner: "",
  totalReferrals: 0,
  totalConverted: 0,
  feeValue: 0
});

const totalFees = computed(() => items.value.reduce((sum, item) => sum + Number(item.feeValue || 0), 0));
const totalReferrals = computed(() => items.value.reduce((sum, item) => sum + Number(item.totalReferrals || 0), 0));
const totalConverted = computed(() => items.value.reduce((sum, item) => sum + Number(item.totalConverted || 0), 0));

async function loadItems() {
  loading.value = true;
  errorText.value = "";
  try {
    const res = await $fetch<{ items: CoiEntry[] }>("/api/coi", { query: { search: search.value || undefined } });
    items.value = res.items;
  } catch (error: unknown) {
    const e = error as { statusCode?: number; data?: { statusCode?: number; statusMessage?: string; message?: string }; message?: string };
    const status = Number(e?.statusCode || e?.data?.statusCode || 0);
    if (status === 401) {
      errorText.value = "Session expired. Redirecting to sign in...";
      await navigateTo("/login");
      return;
    }
    errorText.value = String(e?.data?.statusMessage || e?.data?.message || e?.message || "Failed to load COI entries");
  } finally {
    loading.value = false;
  }
}

async function createItem() {
  await $fetch("/api/coi", { method: "POST", body: draft });
  draft.coiName = "";
  draft.email = "";
  await loadItems();
}

async function removeItem(item: CoiEntry) {
  await $fetch(`/api/coi/${item.id}`, { method: "DELETE" });
  await loadItems();
}

onMounted(loadItems);
</script>

<template>
  <section class="page-wrap">
    <header class="section-banner coi-banner">
      <div>
        <p class="banner-kicker">COI</p>
        <h1>COI</h1>
        <p class="subhead">Referral development sits in its own screen again with clear commercial totals, not just a plain form and grid.</p>
      </div>
      <button @click="loadItems">Refresh</button>
    </header>

    <section class="summary-strip">
      <article><span>Relationships</span><strong>{{ items.length }}</strong></article>
      <article><span>Total Referrals</span><strong>{{ totalReferrals }}</strong></article>
      <article><span>Total Converted</span><strong>{{ totalConverted }}</strong></article>
      <article><span>Fee Value</span><strong>${{ totalFees.toLocaleString() }}</strong></article>
    </section>

    <section class="card">
      <h2>Add COI</h2>
      <div class="grid">
        <label>COI Name<input v-model="draft.coiName" /></label>
        <label>Email<input v-model="draft.email" /></label>
        <label>Entity<input v-model="draft.entity" /></label>
        <label>Industry<input v-model="draft.industry" /></label>
        <label>Lead Partner<input v-model="draft.leadRelationshipPartner" /></label>
        <label>Total Referrals<input v-model.number="draft.totalReferrals" type="number" min="0" /></label>
        <label>Total Converted<input v-model.number="draft.totalConverted" type="number" min="0" /></label>
        <label>Fee Value<input v-model.number="draft.feeValue" type="number" min="0" step="100" /></label>
      </div>
      <button :disabled="!draft.coiName.trim()" @click="createItem">Save COI</button>
    </section>

    <section class="card">
      <h2>COI List</h2>
      <div class="filters">
        <input v-model="search" placeholder="Search COI" @keyup.enter="loadItems" />
        <button @click="loadItems">Apply</button>
      </div>
      <p v-if="errorText" class="error">{{ errorText }}</p>
      <p v-if="loading">Loading COI entries...</p>
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Email</th>
            <th>Entity</th>
            <th>Industry</th>
            <th>Lead Partner</th>
            <th>Referrals</th>
            <th>Converted</th>
            <th>Fee Value</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="item in items" :key="item.id">
            <td>{{ item.coiName }}</td>
            <td>{{ item.email }}</td>
            <td>{{ item.entity }}</td>
            <td>{{ item.industry }}</td>
            <td>{{ item.leadRelationshipPartner || 'Unassigned' }}</td>
            <td>{{ item.totalReferrals }}</td>
            <td>{{ item.totalConverted }}</td>
            <td>${{ Number(item.feeValue || 0).toLocaleString() }}</td>
            <td><button @click="removeItem(item)">Delete</button></td>
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
  color: #4b136f;
  box-shadow: 0 16px 38px rgba(147, 51, 234, 0.18);
}
.coi-banner { background: linear-gradient(135deg, #9b5de5, #6a4c93); }
.banner-kicker { margin: 0 0 0.25rem; font-size: 0.8rem; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; }
.section-banner h1 { margin: 0; font-size: 1.65rem; }
.subhead { margin: 0.2rem 0 0; color: rgba(55, 13, 87, 0.86); font-size: 0.86rem; }
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
  background: linear-gradient(180deg, #f3e8ff, #e9d5ff);
  box-shadow: 0 12px 28px rgba(147, 51, 234, 0.12);
}
.summary-strip span { display: block; font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.08em; color: #7e22ce; }
.summary-strip strong { display: block; margin-top: 0.2rem; font-size: 1.2rem; color: #581c87; }
.grid { display: grid; gap: 0.48rem; grid-template-columns: repeat(4, minmax(0, 1fr)); margin-bottom: 0.58rem; }
label { display: flex; flex-direction: column; gap: 0.2rem; font-size: 0.75rem; font-weight: 700; color: #274568; }
input, button {
  border: 1px solid #c8d6e5;
  border-radius: 8px;
  padding: 0.36rem 0.48rem;
  font: inherit;
  font-size: 0.84rem;
  line-height: 1.2;
}
button {
  background: rgba(255, 255, 255, 0.84);
  color: #581c87;
  font-size: 0.8rem;
  font-weight: 700;
  line-height: 1.1;
  padding: 0.34rem 0.48rem;
  cursor: pointer;
}
.filters { display: flex; gap: 0.42rem; margin-bottom: 0.58rem; }
table { width: 100%; border-collapse: collapse; min-width: 900px; }
th, td { border-bottom: 1px solid #e2e8f0; padding: 0.38rem 0.45rem; text-align: left; font-size: 0.82rem; line-height: 1.25; }
thead th {
  background: #f1f6fb;
  color: #1e3d60;
  font-size: 0.72rem;
  text-transform: uppercase;
  letter-spacing: 0.06em;
}
.error { color: #b91c1c; font-weight: 700; }
@media (max-width: 980px) { .grid, .summary-strip { grid-template-columns: 1fr; } }
</style>
