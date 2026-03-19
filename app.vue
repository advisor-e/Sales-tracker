<script setup lang="ts">
const route = useRoute();
const navItems = [
  { to: "/home", label: "Home", className: "nav-home" },
  { to: "/dashboard", label: "Dashboard", className: "nav-dashboard" },
  { to: "/pipeline", label: "Pipeline", className: "nav-pipeline" },
  { to: "/team", label: "Team", className: "nav-team" },
  { to: "/coi", label: "COI", className: "nav-coi" },
  { to: "/", label: "Blog", className: "nav-blog" }
] as const;
const auth = ref<{ authenticated: boolean; user: { email: string; displayName: string | null } | null }>({
  authenticated: false,
  user: null
});

async function refreshAuth() {
  try {
    auth.value = await $fetch("/api/auth/me");
  } catch {
    auth.value = { authenticated: false, user: null };
  }
}

async function logout() {
  await $fetch("/api/auth/logout", { method: "POST" });
  await refreshAuth();
  if (route.path !== "/login") {
    await navigateTo("/login");
  }
}

onMounted(refreshAuth);
watch(
  () => route.path,
  () => {
    refreshAuth();
  }
);
</script>

<template>
  <div class="app-shell">
    <header class="top-nav">
      <div class="brand-block">
        <strong>Sales Command Center</strong>
        <span v-if="auth.authenticated">{{ auth.user?.displayName || auth.user?.email }}</span>
      </div>
      <template v-if="auth.authenticated">
        <nav class="nav-links">
          <NuxtLink v-for="item in navItems" :key="item.to" :to="item.to" :class="item.className">{{ item.label }}</NuxtLink>
        </nav>
        <button class="auth-btn" @click="logout">Sign Out</button>
      </template>
      <template v-else>
        <NuxtLink to="/login">Sign In</NuxtLink>
      </template>
    </header>
    <main class="page-content">
      <NuxtPage />
    </main>
  </div>
</template>

<style>
:root {
  --bg: #f2f5f8;
  --surface: #ffffff;
  --surface-2: #f8fafc;
  --ink: #11253f;
  --ink-muted: #4a617f;
  --line: #d4dde7;
  --brand: #0f7a8a;
  --brand-2: #1ca4b8;
  --accent: #f59e0b;
  --good: #10b981;
  --bad: #dc2626;
  --shadow: 0 12px 35px rgba(17, 37, 63, 0.08);
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  font-family: "Segoe UI", "Avenir Next", "Trebuchet MS", sans-serif;
  color: var(--ink);
  background:
    radial-gradient(circle at top right, rgba(255, 209, 102, 0.35) 0%, rgba(242, 245, 248, 0) 24%),
    radial-gradient(circle at left top, rgba(76, 201, 240, 0.18) 0%, rgba(242, 245, 248, 0) 30%),
    radial-gradient(circle at bottom right, rgba(255, 107, 107, 0.14) 0%, rgba(242, 245, 248, 0) 26%),
    linear-gradient(180deg, #f5f9fc 0%, #eef4f6 100%);
}

.app-shell {
  min-height: 100vh;
  background: transparent;
}

.top-nav {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 0.6rem;
  flex-wrap: wrap;
  padding: 0.9rem 1rem;
  border-bottom: 1px solid var(--line);
  background: linear-gradient(120deg, rgba(255, 255, 255, 0.96) 0%, rgba(244, 251, 252, 0.92) 100%);
  backdrop-filter: blur(8px);
  position: sticky;
  top: 0;
  z-index: 20;
}

.nav-links {
  display: flex;
  gap: 0.55rem;
  flex-wrap: wrap;
}

.brand-block {
  margin-right: auto;
  display: grid;
  gap: 0.1rem;
}

.brand-block strong {
  font-size: 0.98rem;
  letter-spacing: 0.02em;
}

.brand-block span {
  font-size: 0.78rem;
  color: var(--ink-muted);
}

.top-nav a {
  padding: 0.44rem 0.82rem;
  border: 1px solid rgba(148, 163, 184, 0.25);
  border-radius: 999px;
  text-decoration: none;
  color: var(--ink);
  font-weight: 600;
  background: rgba(255, 255, 255, 0.92);
  transition: all 0.2s ease;
}

.top-nav a.router-link-active {
  transform: translateY(-1px);
  box-shadow: 0 10px 22px rgba(17, 37, 63, 0.12);
}

.top-nav a:hover {
  transform: translateY(-1px);
  border-color: rgba(15, 122, 138, 0.35);
}

.nav-home.router-link-active { background: linear-gradient(120deg, #ffd6d0, #ffe7b0); border-color: rgba(255, 159, 28, 0.45); color: #7c2d12; }
.nav-dashboard.router-link-active { background: linear-gradient(120deg, #ffe3a4, #ffcc7b); border-color: rgba(255, 159, 28, 0.45); color: #7c2d12; }
.nav-pipeline.router-link-active { background: linear-gradient(120deg, #ffc0cb, #ff9ab7); border-color: rgba(244, 63, 94, 0.35); color: #831843; }
.nav-team.router-link-active { background: linear-gradient(120deg, #d8ecff, #bfdbfe); border-color: rgba(59, 130, 246, 0.35); color: #1d4ed8; }
.nav-coi.router-link-active { background: linear-gradient(120deg, #e9d5ff, #d8b4fe); border-color: rgba(147, 51, 234, 0.35); color: #6b21a8; }
.nav-blog.router-link-active { background: linear-gradient(120deg, #d1fae5, #a7f3d0); border-color: rgba(16, 185, 129, 0.35); color: #166534; }

.nav-home:hover { border-color: rgba(255, 159, 28, 0.35); }
.nav-dashboard:hover { border-color: rgba(255, 159, 28, 0.35); }
.nav-pipeline:hover { border-color: rgba(244, 63, 94, 0.28); }
.nav-team:hover { border-color: rgba(59, 130, 246, 0.28); }
.nav-coi:hover { border-color: rgba(147, 51, 234, 0.28); }
.nav-blog:hover { border-color: rgba(16, 185, 129, 0.28); }

.auth-btn {
  padding: 0.44rem 0.82rem;
  border: 1px solid rgba(220, 38, 38, 0.28);
  border-radius: 999px;
  background: #fff5f5;
  color: #9f1239;
  font-weight: 600;
  cursor: pointer;
}

.page-content {
  max-width: 1320px;
  margin: 0 auto;
  padding: 0.95rem 1rem;
}

.dense-table {
  width: 100%;
  border-collapse: collapse;
}

.dense-table th,
.dense-table td {
  border-bottom: 1px solid #dbe5ef;
  font-size: 0.82rem;
  line-height: 1.25;
  padding: 0.38rem 0.45rem;
  text-align: left;
  vertical-align: middle;
}

.dense-table thead th {
  background: #f1f6fb;
  color: #1e3d60;
  font-size: 0.72rem;
  font-weight: 800;
  letter-spacing: 0.06em;
  position: sticky;
  text-transform: uppercase;
  top: 0;
  z-index: 2;
}

.dense-control {
  border: 1px solid #c8d6e5;
  border-radius: 8px;
  font: inherit;
  font-size: 0.84rem;
  line-height: 1.2;
  padding: 0.36rem 0.48rem;
}

.dense-button {
  border-radius: 8px;
  font-size: 0.8rem;
  font-weight: 700;
  line-height: 1.1;
  padding: 0.34rem 0.48rem;
}

@media (max-width: 780px) {
  .top-nav {
    justify-content: stretch;
  }

  .nav-links {
    width: 100%;
  }
}
</style>
