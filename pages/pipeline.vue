<script setup lang="ts">
import type { PipelineEntry } from "~/types/sales";

const search = ref("");
const status = ref("");
const items = ref<PipelineEntry[]>([]);
const errorText = ref("");
const loading = ref(false);

const statusOptions = ["Active", "Await Research", "Completed", "Dead", "On Hold"];
const relationshipOptions = ["Existing Client", "New Prospect"];
const sourceOptions = ["Social Media", "Web Enquiry", "Walk-In", "Phone-In", "Referral", "Cold Target", "Networking", "Pers' Relations"];
const approachOptions = ["Direct Contact", "Pre Approach - Single", "Pre Approach - Sequence", "Pre Approach Gift", "Group Positioning", "Quiz Link Sent"];

const draft = reactive({
  prospectName: "",
  businessName: "",
  partner: "",
  leadStaff: "",
  prospectStatus: "Active",
  relationshipType: "New Prospect",
  prospectSource: "",
  approachStyle: "",
  approachDate: "",
  secureMeeting: false,
  proposalSent: false,
  proposalValue: 0,
  jobSecured: false,
  jobSecuredValue: 0,
  comments: "",
  coiInvolved: "N/A"
});

const totalSecuredValue = computed(() => items.value.reduce((sum, item) => sum + Number(item.jobSecuredValue || 0), 0));
const meetingsCount = computed(() => items.value.filter((item) => item.secureMeeting).length);
const securedCount = computed(() => items.value.filter((item) => item.jobSecured).length);

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
  } catch (error: unknown) {
    const e = error as { statusCode?: number; data?: { statusCode?: number; statusMessage?: string; message?: string }; message?: string };
    const status = Number(e?.statusCode || e?.data?.statusCode || 0);
    if (status === 401) {
      errorText.value = "Session expired. Redirecting to sign in...";
      await navigateTo("/login");
      return;
    }
    errorText.value = String(e?.data?.statusMessage || e?.data?.message || e?.message || "Failed to load pipeline");
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
  draft.approachDate = "";
  draft.prospectSource = "";
  draft.coiInvolved = "N/A";
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
    <header class="section-banner pipeline-banner">
      <div>
        <p class="banner-kicker">Pipeline</p>
        <h1>Pipeline</h1>
        <p class="subhead">Review active prospects and keep the working pipeline visible while you update new business activity.</p>
      </div>
      <button @click="loadItems">Refresh</button>
    </header>

    <section class="summary-strip">
      <article><span>Visible Prospects</span><strong>{{ items.length }}</strong></article>
      <article><span>Meetings Secured</span><strong>{{ meetingsCount }}</strong></article>
      <article><span>Jobs Secured</span><strong>{{ securedCount }}</strong></article>
      <article><span>Secured Value</span><strong>${{ totalSecuredValue.toLocaleString() }}</strong></article>
    </section>

    <section class="card">
      <h2>Add Prospect</h2>
      <div class="grid">
        <label>Prospect Name<input v-model="draft.prospectName" /></label>
        <label>Business Name<input v-model="draft.businessName" /></label>
        <label>Partner<input v-model="draft.partner" /></label>
        <label>Team Member<input v-model="draft.leadStaff" /></label>
        <label>Status
          <select v-model="draft.prospectStatus">
            <option v-for="opt in statusOptions" :key="opt" :value="opt">{{ opt }}</option>
          </select>
        </label>
        <label>Relationship Type
          <select v-model="draft.relationshipType">
            <option v-for="opt in relationshipOptions" :key="opt" :value="opt">{{ opt }}</option>
          </select>
        </label>
        <label>Prospect Source
          <select v-model="draft.prospectSource">
            <option value="">Select source</option>
            <option v-for="opt in sourceOptions" :key="opt" :value="opt">{{ opt }}</option>
          </select>
        </label>
        <label>Approach Style
          <select v-model="draft.approachStyle">
            <option value="">Select style</option>
            <option v-for="opt in approachOptions" :key="opt" :value="opt">{{ opt }}</option>
          </select>
        </label>
        <label>Approach Date<input v-model="draft.approachDate" type="date" /></label>
        <label>COI Involved<input v-model="draft.coiInvolved" placeholder="N/A" /></label>
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
.page-wrap { display: grid; gap: 0.72rem; }
.section-banner {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 1rem;
  padding: 0.92rem 1rem;
  border-radius: 22px;
  color: #5f1130;
  box-shadow: 0 16px 38px rgba(244, 63, 94, 0.18);
}
.pipeline-banner { background: linear-gradient(135deg, #ff6b6b, #f72585); }
.banner-kicker { margin: 0 0 0.25rem; font-size: 0.8rem; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; }
.section-banner h1 { margin: 0; font-size: 1.65rem; }
.subhead { margin: 0.2rem 0 0; color: rgba(76, 5, 25, 0.84); font-size: 0.86rem; }
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
  background: linear-gradient(180deg, #fff0f5, #ffe0ec);
  box-shadow: 0 12px 28px rgba(244, 63, 94, 0.12);
}
.summary-strip span { display: block; font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.08em; color: #9d174d; }
.summary-strip strong { display: block; margin-top: 0.2rem; font-size: 1.2rem; color: #831843; }
.grid { display: grid; gap: 0.48rem; grid-template-columns: repeat(4, minmax(0, 1fr)); }
label { display: flex; flex-direction: column; gap: 0.2rem; font-size: 0.75rem; font-weight: 700; color: #274568; }
input, textarea, select, button {
  border: 1px solid #c8d6e5;
  border-radius: 8px;
  padding: 0.36rem 0.48rem;
  font: inherit;
  font-size: 0.84rem;
  line-height: 1.2;
}
button {
  background: linear-gradient(135deg, #e0f6fa, #d8f0f4);
  color: #831843;
  font-size: 0.8rem;
  font-weight: 700;
  line-height: 1.1;
  padding: 0.34rem 0.48rem;
  cursor: pointer;
}
.section-banner button { background: rgba(255, 255, 255, 0.82); border-color: rgba(131, 24, 67, 0.16); }
.actions { display: flex; gap: 0.42rem; align-items: center; flex-wrap: wrap; margin-top: 0.58rem; }
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
