export default defineNuxtRouteMiddleware(async (to) => {
  const authState = useState<boolean>("auth:authenticated", () => false);

  const checkAuth = async () => {
    try {
      const headers = process.server ? useRequestHeaders(["cookie"]) : undefined;
      const res = await $fetch<{ authenticated: boolean }>("/api/auth/me", { headers });
      authState.value = Boolean(res.authenticated);
    } catch {
      authState.value = false;
    }
  };

  await checkAuth();

  if (to.path === "/login") {
    if (authState.value) {
      return navigateTo("/dashboard");
    }
    return;
  }

  if (!authState.value) {
    return navigateTo("/login");
  }
});
