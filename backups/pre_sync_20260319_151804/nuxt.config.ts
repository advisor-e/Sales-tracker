export default defineNuxtConfig({
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
