<script setup lang="ts">
const email = ref("");
const password = ref("");
const errorText = ref("");
const busy = ref(false);

async function useMikeLogin() {
  email.value = "mike@advisor-e.com";
  password.value = "TempPass!234";
  await submit();
}

async function submit() {
  busy.value = true;
  errorText.value = "";
  try {
    await $fetch("/api/auth/login", {
      method: "POST",
      body: {
        email: email.value,
        password: password.value
      }
    });
    await navigateTo("/dashboard");
  } catch (error: unknown) {
    const e = error as { data?: { statusMessage?: string }; message?: string };
    errorText.value = String(e?.data?.statusMessage || e?.message || "Login failed");
  } finally {
    busy.value = false;
  }
}
</script>

<template>
  <section class="login-wrap">
    <article class="login-card">
      <h1>Sign In</h1>
      <p>Use your account to access your tenant data.</p>
      <form @submit.prevent="submit" class="login-form">
        <label>
          Email
          <input id="email" name="email" v-model="email" type="email" placeholder="you@example.com" autocomplete="username" />
        </label>
        <label>
          Password
          <input id="password" name="password" v-model="password" type="password" placeholder="********" autocomplete="current-password" />
        </label>
        <button type="submit" :disabled="busy || !email.trim() || !password.trim()">{{ busy ? "Signing in..." : "Sign In" }}</button>
      </form>
      <button type="button" class="quick-btn" :disabled="busy" @click="useMikeLogin">Use Mike Login</button>
      <p v-if="errorText" class="error">{{ errorText }}</p>
      <p class="hint">First-time bootstrap: set ADMIN_EMAIL and ADMIN_PASSWORD in .env and sign in with those values.</p>
    </article>
  </section>
</template>

<style scoped>
.login-wrap {
  min-height: calc(100vh - 120px);
  display: grid;
  place-items: center;
  padding: 1rem;
}

.login-card {
  width: min(420px, 96%);
  border: 1px solid rgba(114, 135, 161, 0.3);
  border-radius: 16px;
  background: linear-gradient(165deg, #ffffff 0%, #f7fbff 100%);
  box-shadow: 0 14px 38px rgba(17, 37, 63, 0.12);
  padding: 1.1rem;
  display: grid;
  gap: 0.75rem;
}

.login-card h1 {
  margin: 0;
  color: #123055;
  letter-spacing: 0.01em;
}

.login-card > p {
  margin: 0;
  color: #4a617f;
}

.login-form {
  display: grid;
  gap: 0.7rem;
}

label {
  display: flex;
  flex-direction: column;
  gap: 0.3rem;
  font-size: 0.86rem;
  font-weight: 700;
  color: #274568;
}

input,
button {
  border: 1px solid #c8d6e5;
  border-radius: 10px;
  padding: 0.55rem 0.65rem;
  font: inherit;
}

button {
  background: linear-gradient(135deg, #e0f6fa, #d8f0f4);
  color: #0a4752;
  font-weight: 700;
  cursor: pointer;
}

.quick-btn {
  background: linear-gradient(135deg, #ffeec8, #fde7bc);
  color: #7b4b07;
  border-color: rgba(190, 126, 15, 0.28);
}

.error {
  color: #b91c1c;
  font-weight: 700;
}

.hint {
  color: #506684;
  font-size: 0.82rem;
}
</style>
