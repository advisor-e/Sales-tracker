export default defineNuxtConfig({
  compatibilityDate: "2026-03-20",
  devtools: { enabled: true },
  runtimeConfig: {
    openaiApiKey: process.env.OPENAI_API_KEY || "",
    public: {
      appName: "Sales Tracker (Nuxt)"
    }
  },
  nitro: {
    preset: "node-server"
  }
});
