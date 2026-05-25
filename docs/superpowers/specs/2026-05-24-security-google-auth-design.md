# Security Design — Google Account Whitelist

**Date:** 2026-05-24  
**Status:** Implemented

## Problem

The GAS webapp (`doGet()`) had no access control. Anyone with the URL could open the app.

## Solution

Server-side email whitelist in `doGet()` + deployment restricted to "Anyone with Google Account".

## What was changed

**`src/Code.gs`**
- Added `AUTHORIZED_EMAILS` constant with the client's email (`inelcoingeniero@gmail.com`)
- `doGet()` now calls `Session.getActiveUser().getEmail()` before rendering anything
- Non-matching emails get a plain "Acceso denegado" HTML page; the app never loads

**GAS Deployment settings (manual step)**
- "Who has access" must be set to "Anyone with Google Account" (not "Anyone")
- This forces Google to require login before `doGet()` runs — anonymous users blocked at the Google layer

## Adding more users (v2+)

Add emails to the `AUTHORIZED_EMAILS` array in `Code.gs`, then `clasp push` and redeploy.

## Trade-offs considered

| Option | Decision |
|--------|----------|
| Google account whitelist | Chosen — bulletproof, no passwords |
| PIN/password gate | Rejected — weaker, passwords can be shared |
| Both layers | Rejected — overkill for one user |
