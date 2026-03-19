<script setup lang="ts">
import type { PipelineEntry } from "~/types/sales";

const search = ref("");
const status = ref("");
const items = ref<PipelineEntry[]>([]);
const errorText = ref("");
const loading = ref(false);

const draft = reactive({
  prospectName: "",
  businessName: "",
  partner: "",
  leadStaff: "",
  prospectStatus: "Active",
  approachStyle: "",
  secureMeeting: false,
  proposalSent: false,
  proposalValue: 0,
  jobSecured: false,
  jobSecuredValue: 0,
  comments: ""
});

async function loadItems() {
  loading.value = true;
  errorText.value = "";
  try {
    const res = await $fetch<{ items: PipelineEntry[] }>("/api/pipeline", {
      query: {
        search: search.value || undefined,
        status: status.value || undefined
      }
    });
    items.value = res.items;
  } catch (error: any) {
    errorText.value = String(error?.data?.message || error?.message || "Failed to load pipeline");
  } finally {
    loading.value = false;
  }
}

async function createItem() {
  await $fetch("/api/pipeline", {
    method: "POST",
    body: draft
  });
  draft.prospectName = "";
  draft.businessName = "";
  draft.comments = "";
  await loadItems();
}

async function toggleSecure(item: PipelineEntry) {
  await $fetch(`/api/pipeline/${item.id}`, {
    method: "PATCH",
    body: { secureMeeting: !item.secureMeeting }
  });
  await loadItems();
}

async function toggleSecured(item: PipelineEntry) {
  await $fetch(`/api/pipeline/${item.id}`, {
    method: "PATCH",
    body: { jobSecured: !item.jobSecured }
  });
  await loadItems();
}

async function removeItem(item: PipelineEntry) {
  await $fetch(`/api/pipeline/${item.id}`, { method: "DELETE" });
  await loadItems();
}

onMounted(loadItems);
</script>

<template>
  <section class="page-wrap">
    <header class="page-head">
      <h1>Pipeline</h1>
      <button @click="loadItems">Refresh</button>
    </header>

    <section class="card">
      <h2>Add Prospect</h2>
      <div class="grid">
        <label>Prospect Name<input v-model="draft.prospectName" /></label>
        <label>Business Name<input v-model="draft.businessName" /></label>
        <label>Partner<input v-model="draft.partner" /></label>
        <label>Team Member<input v-model="draft.leadStaff" /></label>
        <label>Status<input v-model="draft.prospectStatus" /></label>
        <label>Approach Style<input v-model="draft.approachStyle" /></label>
        <label>Proposal Value<input v-model.number="draft.proposalValue" type="number" min="0" step="100" /></label>
        <label>Secured Value<input v-model.number="draft.jobSecuredValue" type="number" min="0" step="100" /></label>
      </div>
      <label>Comments<textarea v-model="draft.comments" rows="3" /></label>
      <div class="actions">
        <label><input v-model="draft.secureMeeting" type="checkbox" /> Secure Meeting</label>
        <label><input v-model="draft.proposalSent" type="checkbox" /> Proposal Sent</label>
        <label><input v-model="draft.jobSecured" type="checkbox" /> Job Secured</label>
        <button :disabled="!draft.prospectName.trim()" @click="createItem">Save Prospect</button>
      </div>
    </section>

    <section class="card">
      <h2>Pipeline List</h2>
      <div class="filters">
        <input v-model="search" placeholder="Search by name, business, owner" @keyup.enter="loadItems" />
        <input v-model="status" placeholder="Filter status" @keyup.enter="loadItems" />
        <button @click="loadItems">Apply</button>
      </div>
      <p v-if="errorText" class="error">{{ errorText }}</p>
      <p v-if="loading">Loading pipeline...</p>
      <table>
        <thead>
          <tr>
            <th>Prospect</th>
            <th>Business</th>
            <th>Owner</th>
            <th>Status</th>
            <th>Meeting</th>
            <th>Secured</th>
            <th>Proposal Value</th>
            <th>Secured Value</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="item in items" :key="item.id">
            <td>{{ item.prospectName }}</td>
            <td>{{ item.businessName }}</td>
            <td>{{ item.leadStaff || item.partner || 'Unassigned' }}</td>
            <td>{{ item.prospectStatus }}</td>
            <td><button @click="toggleSecure(item)">{{ item.secureMeeting ? 'Yes' : 'No' }}</button></td>
            <td><button @click="toggleSecured(item)">{{ item.jobSecured ? 'Yes' : 'No' }}</button></td>
            <td>${{ Number(item.proposalValue || 0).toLocaleString() }}</td>
            <td>${{ Number(item.jobSecuredValue || 0).toLocaleString() }}</td>
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
.grid { display: grid; gap: 0.6rem; grid-template-columns: repeat(4, minmax(0, 1fr)); }
label { display: flex; flex-direction: column; gap: 0.25rem; font-size: 0.88rem; font-weight: 600; }
input, textarea, button { border: 1px solid #cbd5e1; border-radius: 8px; padding: 0.45rem 0.55rem; }
.actions { display: flex; gap: 0.6rem; align-items: center; flex-wrap: wrap; margin-top: 0.7rem; }
.filters { display: flex; gap: 0.6rem; margin-bottom: 0.7rem; }
table { width: 100%; border-collapse: collapse; min-width: 900px; }
th, td { border-bottom: 1px solid #e2e8f0; padding: 0.45rem; text-align: left; }
.error { color: #b91c1c; }
@media (max-width: 980px) { .grid { grid-template-columns: 1fr; } }
</style>
