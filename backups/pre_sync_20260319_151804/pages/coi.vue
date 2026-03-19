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

async function loadItems() {
  loading.value = true;
  errorText.value = "";
  try {
    const res = await $fetch<{ items: CoiEntry[] }>("/api/coi", { query: { search: search.value || undefined } });
    items.value = res.items;
  } catch (error: any) {
    errorText.value = String(error?.data?.message || error?.message || "Failed to load COI entries");
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
    <header class="page-head">
      <h1>COI</h1>
      <button @click="loadItems">Refresh</button>
    </header>

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
.page-wrap { display: grid; gap: 1rem; }
.page-head { display: flex; justify-content: space-between; align-items: center; }
.card { border: 1px solid #dbeafe; border-radius: 10px; background: #fff; padding: 0.8rem; overflow: auto; }
.grid { display: grid; gap: 0.6rem; grid-template-columns: repeat(4, minmax(0, 1fr)); margin-bottom: 0.7rem; }
label { display: flex; flex-direction: column; gap: 0.25rem; font-size: 0.88rem; font-weight: 600; }
input, button { border: 1px solid #cbd5e1; border-radius: 8px; padding: 0.45rem 0.55rem; }
.filters { display: flex; gap: 0.6rem; margin-bottom: 0.7rem; }
table { width: 100%; border-collapse: collapse; min-width: 900px; }
th, td { border-bottom: 1px solid #e2e8f0; padding: 0.45rem; text-align: left; }
.error { color: #b91c1c; }
@media (max-width: 980px) { .grid { grid-template-columns: 1fr; } }
</style>
