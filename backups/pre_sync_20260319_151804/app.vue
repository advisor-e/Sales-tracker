<template>
  <div class="app-shell">
    <header class="top-nav">
      <NuxtLink to="/dashboard">Dashboard</NuxtLink>
      <NuxtLink to="/pipeline">Pipeline</NuxtLink>
      <NuxtLink to="/team">Team</NuxtLink>
      <NuxtLink to="/coi">COI</NuxtLink>
      <NuxtLink to="/">Blog</NuxtLink>
    </header>
    <main class="page-content">
      <NuxtPage />
    </main>
  </div>
</template>

<style scoped>
.app-shell {
  min-height: 100vh;
  background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
}

.top-nav {
  display: flex;
  gap: 0.6rem;
  flex-wrap: wrap;
  padding: 0.85rem 1rem;
  border-bottom: 1px solid #cbd5e1;
  background: #ffffff;
  position: sticky;
  top: 0;
  z-index: 10;
}

.top-nav a {
  padding: 0.38rem 0.72rem;
  border: 1px solid #cbd5e1;
  border-radius: 999px;
  text-decoration: none;
  color: #0f172a;
  font-weight: 600;
}

.top-nav a.router-link-active {
  background: #dbeafe;
  border-color: #93c5fd;
}

.page-content {
  max-width: 1320px;
  margin: 0 auto;
  padding: 1rem;
}
</style>
