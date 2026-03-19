<script setup lang="ts">
import type { Principle } from "~/types/blog";

type BlogPost = {
  id: number;
  kind: "draft" | "final";
  title: string;
  topic: string;
  audience: string;
  objective: string;
  tone: string;
  length: string;
  cta: string;
  selectedPerson?: string | null;
  targetMode?: string | null;
  outlineText: string;
  finalText?: string | null;
  isPinned: boolean;
  createdAt: string;
  updatedAt: string;
};

type BlogInput = {
  id: number;
  signature: string;
  topic: string;
  audience: string;
  objective: string;
  tone: string;
  length: string;
  cta: string;
  principlesJson: Principle[];
  selectedPerson?: string | null;
  targetMode?: string | null;
  styleStrength?: string | null;
  updatedAt: string;
};

const tones = ["Professional", "Friendly", "Confident", "Educational"];
const lengths = ["Short", "Medium", "Long"];
const polishLevels = ["Standard", "Strong", "Premium"];

const form = reactive({
  topic: "",
  audience: "Business owners",
  objective: "Help readers make practical financial decisions",
  tone: "Professional",
  length: "Medium",
  cta: "book a short strategy call",
  selectedPerson: "",
  targetMode: "Lead Staff (Client Manager)",
  polishLevel: "Strong"
});

const principles = ref<Principle[]>([
  { title: "Market context", details: ["What changed and why it matters", "Which indicators to watch", "Where uncertainty is highest"] },
  { title: "Action plan", details: ["What to do this month", "How to sequence decisions", "How to avoid overreaction"] },
  { title: "Review cadence", details: ["What to review regularly", "When to adjust", "How to measure progress"] }
]);

const draftText = ref("");
const finalText = ref("");
const aiSource = ref<"ai" | "template" | "">("");
const aiError = ref("");
const busy = ref(false);
const startupError = ref("");

const inputs = ref<BlogInput[]>([]);
const draftPosts = ref<BlogPost[]>([]);
const finalPosts = ref<BlogPost[]>([]);

const draftSearch = ref("");
const finalSearch = ref("");
const draftPinnedOnly = ref(false);
const finalPinnedOnly = ref(false);

const filteredDrafts = computed(() => {
  return draftPosts.value.filter((item) => {
    if (draftPinnedOnly.value && !item.isPinned) {
      return false;
    }
    if (!draftSearch.value.trim()) {
      return true;
    }
    const q = draftSearch.value.toLowerCase();
    return `${item.title} ${item.topic} ${item.selectedPerson || ""}`.toLowerCase().includes(q);
  });
});

const filteredFinals = computed(() => {
  return finalPosts.value.filter((item) => {
    if (finalPinnedOnly.value && !item.isPinned) {
      return false;
    }
    if (!finalSearch.value.trim()) {
      return true;
    }
    const q = finalSearch.value.toLowerCase();
    return `${item.title} ${item.topic} ${item.selectedPerson || ""}`.toLowerCase().includes(q);
  });
});

function signatureForInputs() {
  return JSON.stringify(
    {
      topic: form.topic.trim(),
      audience: form.audience.trim(),
      objective: form.objective.trim(),
      tone: form.tone,
      length: form.length,
      cta: form.cta.trim(),
      principles: principles.value,
      selectedPerson: form.selectedPerson.trim(),
      targetMode: form.targetMode.trim(),
      styleStrength: "Balanced"
    },
    Object.keys(
      {
        topic: true,
        audience: true,
        objective: true,
        tone: true,
        length: true,
        cta: true,
        principles: true,
        selectedPerson: true,
        targetMode: true,
        styleStrength: true
      }
    )
  );
}

function buildPostTitle(kind: "draft" | "final") {
  const prefix = kind === "draft" ? "Draft" : "Final";
  return `${prefix}: ${form.topic || "Untitled"}`;
}

async function loadInputs() {
  const res = await $fetch<{ items: BlogInput[] }>("/api/blog/inputs");
  inputs.value = res.items;
}

async function loadPosts() {
  const drafts = await $fetch<{ items: BlogPost[] }>("/api/blog/posts", { query: { kind: "draft" } });
  const finals = await $fetch<{ items: BlogPost[] }>("/api/blog/posts", { query: { kind: "final" } });
  draftPosts.value = drafts.items;
  finalPosts.value = finals.items;
}

async function generateDraft() {
  busy.value = true;
  aiError.value = "";
  try {
    const res = await $fetch<{ text: string; source: "ai" | "template"; error?: string }>("/api/blog/generate/draft", {
      method: "POST",
      body: {
        topic: form.topic,
        audience: form.audience,
        objective: form.objective,
        tone: form.tone,
        length: form.length,
        cta: form.cta,
        principles: principles.value
      }
    });
    draftText.value = res.text;
    aiSource.value = res.source;
    aiError.value = res.error || "";
    await saveInputs();
  } finally {
    busy.value = false;
  }
}

async function generateFinal() {
  busy.value = true;
  aiError.value = "";
  try {
    const res = await $fetch<{ text: string; source: "ai" | "template"; error?: string }>("/api/blog/generate/final", {
      method: "POST",
      body: {
        outlineText: draftText.value,
        topic: form.topic,
        audience: form.audience,
        objective: form.objective,
        tone: form.tone,
        cta: form.cta,
        polishLevel: form.polishLevel
      }
    });
    finalText.value = res.text;
    aiSource.value = res.source;
    aiError.value = res.error || "";
  } finally {
    busy.value = false;
  }
}

async function saveInputs() {
  await $fetch("/api/blog/inputs", {
    method: "POST",
    body: {
      signature: signatureForInputs(),
      topic: form.topic,
      audience: form.audience,
      objective: form.objective,
      tone: form.tone,
      length: form.length,
      cta: form.cta,
      principles: principles.value,
      selectedPerson: form.selectedPerson,
      targetMode: form.targetMode,
      styleStrength: "Balanced",
      styleTitles: []
    }
  });
  await loadInputs();
}

async function savePost(kind: "draft" | "final") {
  await $fetch("/api/blog/posts", {
    method: "POST",
    body: {
      kind,
      title: buildPostTitle(kind),
      topic: form.topic,
      audience: form.audience,
      objective: form.objective,
      tone: form.tone,
      length: form.length,
      cta: form.cta,
      selectedPerson: form.selectedPerson,
      targetMode: form.targetMode,
      outlineText: draftText.value,
      finalText: kind === "final" ? finalText.value : null,
      metadata: {
        aiSource: aiSource.value,
        aiError: aiError.value
      }
    }
  });
  await loadPosts();
}

function restoreFromInput(item: BlogInput) {
  form.topic = item.topic;
  form.audience = item.audience;
  form.objective = item.objective;
  form.tone = item.tone;
  form.length = item.length;
  form.cta = item.cta;
  form.selectedPerson = item.selectedPerson || "";
  form.targetMode = item.targetMode || "Lead Staff (Client Manager)";
  principles.value = Array.isArray(item.principlesJson) ? item.principlesJson : principles.value;
}

function restoreDraft(item: BlogPost) {
  form.topic = item.topic;
  form.audience = item.audience;
  form.objective = item.objective;
  form.tone = item.tone;
  form.length = item.length;
  form.cta = item.cta;
  form.selectedPerson = item.selectedPerson || "";
  form.targetMode = item.targetMode || "Lead Staff (Client Manager)";
  draftText.value = item.outlineText || "";
}

function restoreFinal(item: BlogPost) {
  restoreDraft(item);
  finalText.value = item.finalText || "";
}

function duplicateFinalToDraft(item: BlogPost) {
  draftText.value = item.finalText || item.outlineText || "";
}

async function setPinned(item: BlogPost, value: boolean) {
  await $fetch(`/api/blog/posts/${item.id}`, {
    method: "PATCH",
    body: { isPinned: value }
  });
  await loadPosts();
}

async function deletePost(item: BlogPost) {
  await $fetch(`/api/blog/posts/${item.id}`, { method: "DELETE" });
  await loadPosts();
}

async function deleteInput(item: BlogInput) {
  await $fetch(`/api/blog/inputs/${item.id}`, { method: "DELETE" });
  await loadInputs();
}

onMounted(async () => {
  try {
    await Promise.all([loadInputs(), loadPosts()]);
  } catch (error: unknown) {
    const e = error as { data?: { message?: string }; message?: string };
    startupError.value = String(e?.data?.message || e?.message || "Failed to load initial data");
  }
});
</script>

<template>
  <main class="page">
    <section class="hero">
      <div>
        <p class="hero-kicker">Blog</p>
        <h1>Sales Tracker Blog Studio</h1>
        <p>Keep blog inputs, draft outlines, and polished final posts in one screen with a clearer editorial command layout.</p>
      </div>
      <div class="hero-stats">
        <article><span>Saved Inputs</span><strong>{{ inputs.length }}</strong></article>
        <article><span>Drafts</span><strong>{{ draftPosts.length }}</strong></article>
        <article><span>Finals</span><strong>{{ finalPosts.length }}</strong></article>
      </div>
      <p v-if="startupError" class="error">Startup warning: {{ startupError }}</p>
    </section>

    <section class="card form-card">
      <h2>Blog Inputs</h2>
      <div class="grid">
        <label>
          Topic
          <input v-model="form.topic" placeholder="e.g., Interest rates and small business cashflow" />
        </label>
        <label>
          Audience
          <input v-model="form.audience" />
        </label>
        <label>
          Objective
          <input v-model="form.objective" />
        </label>
        <label>
          Tone
          <select v-model="form.tone">
            <option v-for="tone in tones" :key="tone" :value="tone">{{ tone }}</option>
          </select>
        </label>
        <label>
          Length
          <select v-model="form.length">
            <option v-for="length in lengths" :key="length" :value="length">{{ length }}</option>
          </select>
        </label>
        <label>
          CTA
          <input v-model="form.cta" />
        </label>
        <label>
          Target Mode
          <input v-model="form.targetMode" />
        </label>
        <label>
          Selected Person
          <input v-model="form.selectedPerson" />
        </label>
      </div>

      <h3>Principles</h3>
      <div class="principles">
        <article v-for="(p, index) in principles" :key="index" class="principle">
          <label>
            Title
            <input v-model="p.title" />
          </label>
          <label v-for="(_, detailIndex) in p.details" :key="detailIndex">
            Detail {{ detailIndex + 1 }}
            <input v-model="p.details[detailIndex]" />
          </label>
        </article>
      </div>

      <div class="actions">
        <button :disabled="busy" @click="saveInputs">Save Inputs</button>
        <button :disabled="busy || !form.topic.trim()" @click="generateDraft">Generate Draft</button>
        <button :disabled="busy || !draftText.trim()" @click="generateFinal">Generate Final</button>
        <button :disabled="busy || !draftText.trim()" @click="savePost('draft')">Save Draft</button>
        <button :disabled="busy || !finalText.trim()" @click="savePost('final')">Save Final</button>
      </div>

      <p v-if="aiSource" class="status">Last generation source: <strong>{{ aiSource }}</strong></p>
      <p v-if="aiError" class="error">{{ aiError }}</p>
    </section>

    <section class="columns">
      <article class="card">
        <h2>Saved Inputs</h2>
        <div class="list">
          <div v-for="item in inputs" :key="item.id" class="list-item">
            <div>
              <strong>{{ item.topic }}</strong>
              <p>{{ new Date(item.updatedAt).toLocaleString() }}</p>
            </div>
            <div class="row-actions">
              <button @click="restoreFromInput(item)">Restore</button>
              <button @click="deleteInput(item)">Delete</button>
            </div>
          </div>
        </div>
      </article>

      <article class="card">
        <h2>Draft Outlines</h2>
        <div class="filters">
          <input v-model="draftSearch" placeholder="Search drafts" />
          <label><input v-model="draftPinnedOnly" type="checkbox" /> Pinned only</label>
        </div>
        <div class="list">
          <div v-for="item in filteredDrafts" :key="item.id" class="list-item">
            <div>
              <strong>{{ item.title }}</strong>
              <p>{{ new Date(item.updatedAt).toLocaleString() }} • {{ item.isPinned ? 'Pinned' : 'Unpinned' }}</p>
            </div>
            <div class="row-actions wrap">
              <button @click="restoreDraft(item)">Restore</button>
              <button @click="setPinned(item, !item.isPinned)">{{ item.isPinned ? 'Unpin' : 'Pin' }}</button>
              <button @click="deletePost(item)">Delete</button>
            </div>
          </div>
        </div>
      </article>

      <article class="card">
        <h2>Final Posts</h2>
        <div class="filters">
          <input v-model="finalSearch" placeholder="Search finals" />
          <label><input v-model="finalPinnedOnly" type="checkbox" /> Pinned only</label>
        </div>
        <div class="list">
          <div v-for="item in filteredFinals" :key="item.id" class="list-item">
            <div>
              <strong>{{ item.title }}</strong>
              <p>{{ new Date(item.updatedAt).toLocaleString() }} • {{ item.isPinned ? 'Pinned' : 'Unpinned' }}</p>
            </div>
            <div class="row-actions wrap">
              <button @click="restoreFinal(item)">Restore</button>
              <button @click="duplicateFinalToDraft(item)">Duplicate to Draft</button>
              <button @click="setPinned(item, !item.isPinned)">{{ item.isPinned ? 'Unpin' : 'Pin' }}</button>
              <button @click="deletePost(item)">Delete</button>
            </div>
          </div>
        </div>
      </article>
    </section>

    <section class="card editor-card">
      <h2>Draft Outline</h2>
      <textarea v-model="draftText" rows="14" />
      <h2>Final Post</h2>
      <label>
        Polish Level
        <select v-model="form.polishLevel">
          <option v-for="level in polishLevels" :key="level" :value="level">{{ level }}</option>
        </select>
      </label>
      <textarea v-model="finalText" rows="16" />
    </section>
  </main>
</template>

<style scoped>
.page {
  display: grid;
  gap: 0.72rem;
  color: #0f172a;
  font-family: "Segoe UI", "Arial", sans-serif;
}

.hero {
  display: grid;
  gap: 1rem;
  grid-template-columns: 1.4fr 1fr;
  align-items: end;
  padding: 0.92rem 1rem;
  border-radius: 22px;
  background: linear-gradient(135deg, #86efac 0%, #34d399 100%);
  border: 1px solid rgba(16, 185, 129, 0.18);
  box-shadow: 0 16px 38px rgba(16, 185, 129, 0.16);
}

.hero-kicker {
  margin: 0 0 0.25rem;
  font-size: 0.8rem;
  font-weight: 800;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}

.hero h1 {
  margin: 0;
  font-size: 1.7rem;
}

.hero p {
  margin: 0.4rem 0 0;
}

.hero-stats {
  display: grid;
  gap: 0.52rem;
  grid-template-columns: repeat(3, minmax(0, 1fr));
}

.hero-stats article {
  padding: 0.6rem 0.68rem;
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.78);
}

.hero-stats span {
  display: block;
  font-size: 0.8rem;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: #166534;
}

.hero-stats strong {
  display: block;
  margin-top: 0.2rem;
  font-size: 1.2rem;
}

.card {
  background: #ffffff;
  border: 1px solid rgba(16, 185, 129, 0.14);
  border-radius: 18px;
  padding: 0.72rem;
  box-shadow: 0 12px 28px rgba(16, 50, 74, 0.08);
}

.grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 0.45rem;
}

label {
  display: flex;
  flex-direction: column;
  gap: 0.2rem;
  font-size: 0.78rem;
  font-weight: 600;
}

input,
select,
textarea,
button {
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  padding: 0.36rem 0.48rem;
  font-size: 0.84rem;
  line-height: 1.2;
}

textarea {
  width: 100%;
  margin: 0.25rem 0 0.55rem;
  font-family: "Consolas", "Courier New", monospace;
}

.actions {
  display: flex;
  gap: 0.38rem;
  flex-wrap: wrap;
  margin-top: 0.55rem;
}

button {
  cursor: pointer;
  background: #f8fafc;
  font-size: 0.8rem;
  line-height: 1.1;
  padding: 0.34rem 0.48rem;
}

button:hover:enabled {
  background: #f1f5f9;
}

button:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.principles {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 0.46rem;
  margin-top: 0.45rem;
}

.principle {
  border: 1px solid #e2e8f0;
  border-radius: 10px;
  padding: 0.48rem;
}

.columns {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 0.6rem;
}

.list {
  display: flex;
  flex-direction: column;
  gap: 0.32rem;
  max-height: 360px;
  overflow: auto;
}

.list-item {
  border: 1px solid #e2e8f0;
  border-radius: 10px;
  padding: 0.4rem 0.48rem;
}

.list-item p {
  margin: 0.16rem 0 0;
  color: #475569;
  font-size: 0.75rem;
}

.row-actions {
  display: flex;
  gap: 0.3rem;
  margin-top: 0.3rem;
}

.row-actions.wrap {
  flex-wrap: wrap;
}

.filters {
  display: flex;
  gap: 0.4rem;
  margin-bottom: 0.42rem;
  align-items: center;
}

.status {
  color: #166534;
}

.error {
  color: #b91c1c;
}

@media (max-width: 980px) {
  .hero,
  .hero-stats,
  .grid,
  .principles,
  .columns {
    grid-template-columns: 1fr;
  }
}
</style>
