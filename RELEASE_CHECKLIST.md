# Sales Tracker Nuxt Release Checklist

## Status Snapshot

- Code sync from sales-tracker-nuxt-clean to sales-tracker-nuxt: completed
- Source backup before sync: completed
- DB-backed runtime validation: blocked (no MySQL available and no DATABASE_URL configured)

## 1. Environment Setup

1. Copy .env.example to .env
2. Set DATABASE_URL to a reachable MySQL instance
3. Set ADMIN_EMAIL and ADMIN_PASSWORD
4. Optionally set OPENAI_API_KEY

## 2. Dependency and Type Prep

1. npm install --ignore-scripts --no-fund
2. Set NUXT_TELEMETRY_DISABLED=1
3. npx nuxt prepare
4. npm run prisma:generate

## 3. Database Provisioning

1. npm run prisma:migrate -- --name init
2. npm run auth:create-admin
3. Optional imports:
   - npm run import:app-config -- ../sales-tracker-app/app_config.json
   - npm run import:workbook -- ..\Get.1a.Sales Tracker.polished.xlsx

## 4. Build and Smoke Tests

1. npm run build
2. npm run preview
3. Validate:
   - GET / returns 200
   - GET /api/auth/me returns JSON payload
   - Login with admin credentials succeeds
   - Pipeline and COI create/list/delete flows succeed
   - Dashboard and team summary load without errors

## 5. Production Start

1. Build once using npm run build
2. Start with node .output/server/index.mjs
3. Confirm process manager settings (restart, logs, env vars)

## 6. Pre-Release Gate

- No TypeScript/Pylance blocking diagnostics
- No runtime 500s in server logs during smoke tests
- Import scripts run idempotently on staging data
- Backup strategy for MySQL is documented
