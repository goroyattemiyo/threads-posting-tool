# AutoPostTool - Settings Screen Refactor Design

## Version
- Document Version: 1.0
- App Version: 2.0.0
- Date: 2026-02-27

---

## 1. Overview

### Current Issues
- Settings are scattered across multiple screens (welcome, setup, setup-auth, settings)
- "Back" button navigates to separate pages instead of staying in settings
- Meta Developer setup is on a dedicated page, not integrated into settings
- No manual/documentation link in settings
- No version info displayed

### Goal
Consolidate ALL setup and configuration into a single Settings screen with sections.
Flow: Splash -> Settings (first time) / Compose (returning user)

---

## 2. Screen Flow

### First-time User

Splash Screen | v Settings Screen (white setup guide at top) |-- Step 1: Spreadsheet Setup (create or connect) |-- Step 2: Meta Developer App (app_id, app_secret) |-- Step 3: Redirect URI (copy button) |-- Step 4: Authenticate (OAuth) | v (after successful auth) Compose Screen


### Returning User (authenticated)
Splash Screen | v Compose Screen (default) | +-- Nav: Compose | Schedule | History | Analytics | Settings


### Settings Screen (authenticated user)
Settings Screen |-- Setup Guide (collapsible, completed steps shown with checkmarks) |-- Account Management (switch, delete, import) |-- Connection Info (spreadsheet URL, app URL, redirect URI) |-- App Settings (app_id, app_secret - editable) |-- Authentication (re-auth button) |-- Manual & Help (Notion link, FAQ) |-- About (version, credits) |-- Actions (disconnect) |-- Navigation (quick links to other screens)


---

## 3. Screens to Remove

| Screen Name | Current Location | Action |
|---|---|---|
| `welcome` | app_ui.html: renderWelcomeScreen | Merge into Settings |
| `setup` | app_ui.html: renderSetupScreen | Merge into Settings |
| `setup-auth` | app_ui.html: renderSetupAuthScreen | Merge into Settings |
| `sheet-created` | app_ui.html: renderSheetCreatedScreen | Merge into Settings |
| `connect` | app_core.html: showScreen case | Alias to Settings |

After refactor, only these screens remain:
- `settings` (all setup + config)
- `compose` (post creation)
- `schedule` (scheduled posts)
- `history` (post history)
- `analytics` (insights)
- `loading` (transition state)

---

## 4. Settings Screen Sections (Detailed)

### 4.1 Setup Guide (Top - Collapsible)
- Shows for ALL users (collapsed by default for authenticated users)
- Steps with checkmark status:
  - [x] Step 1: Spreadsheet connected
  - [x] Step 2: App ID & Secret configured  
  - [x] Step 3: Redirect URI set
  - [x] Step 4: Authenticated
- "Start Posting" button appears when all steps complete

### 4.2 Manual & Help
- Link to Notion manual: https://sordid-echinacea-374.notion.site/Threads-2fab3545dbc680d697abd1c773d8aae6
- In-app FAQ (collapsible accordion):
  - "OAuth error 1349168" -> Check redirect URI trailing slash
  - "Token expired" -> Re-authenticate
  - "In-app browser warning" -> Open in external browser

### 4.3 Spreadsheet Connection
- Current spreadsheet URL (readonly + copy button)
- "Change Spreadsheet" button (shows input field)
- "Create New Spreadsheet" button (first-time only)

### 4.4 Meta Developer App Settings
- App ID input (editable, saved to spreadsheet)
- App Secret input (editable, saved to spreadsheet)
- Save button
- Link to Meta for Developers

### 4.5 Redirect URI
- Display current deployment URL
- Copy button
- Warning: "Trailing slash must match exactly"

### 4.6 Authentication
- Current status: Connected as @username / Not connected
- "Authenticate" button (opens OAuth in new window)
- "Re-authenticate" button (for token refresh)
- Token expiry warning (if applicable)

### 4.7 Account Management
- Account list (switch, delete)
- Import from another spreadsheet

### 4.8 App Info
- App Name: AutoPostTool
- Version: 2.0.0
- Icon + branding
- Credits / About link

### 4.9 Actions
- Disconnect button (with confirmation)
- Back to Top / Navigation buttons

---

## 5. Files to Modify

### app_core.html
- `App.init()`: Change initial screen logic
  - First time (no sheetId): -> `settings`
  - Has sheetId but no token: -> `settings`  
  - Has sheetId and token: -> `compose`
- `App.showScreen()`: Remove cases for `welcome`, `setup`, `setup-auth`, `sheet-created`, `connect`
  - All redirect to `settings`

### app_ui.html
- Remove: `renderWelcomeScreen`, `bindWelcomeEvents`
- Remove: `renderSetupScreen`, `bindSetupEvents`
- Remove: `renderSetupAuthScreen`, `bindSetupAuthEvents`
- Remove: `renderSheetCreatedScreen`, `bindSheetCreatedEvents`
- Keep: `renderMainLayout` (add settings to nav if not already)
- Keep: `showImportAccountModal`, `openEditModal`, etc.

### screen_settings.html
- Complete rewrite of `renderSettingsContent`:
  - Integrate spreadsheet creation/connection (from welcome)
  - Integrate Meta Developer setup (from setup)
  - Integrate OAuth authentication (from setup-auth)
  - Add manual links
  - Add version info
  - Add collapsible setup guide with step status
- Update `bindSettingsEvents` to handle all new functionality

### index.html
- No changes needed (splash -> App.init handles routing)

### screen_compose.html, screen_schedule.html, screen_history.html, screen_insights.html
- Remove any `showScreen('welcome')` or `showScreen('setup')` calls
- Replace with `showScreen('settings')`

---

## 6. Version Display

AutoPostTool v2.0.0 Build: 2026-02-27


Location: Bottom of Settings screen, styled as subtle footer text.

---

## 7. Manual Link

Primary: Notion page (external link, opens in new tab)
URL: https://sordid-echinacea-374.notion.site/Threads-2fab3545dbc680d697abd1c773d8aae6

Display as card with icon:
[Book Icon] Manual & Help View the full setup guide and FAQ on Notion ->


---

## 8. Migration Notes

### Backward Compatibility
- `showScreen('welcome')` calls throughout codebase -> redirect to `settings`
- `showScreen('setup')` -> redirect to `settings`
- `showScreen('setup-auth')` -> redirect to `settings` (scroll to auth section)
- `showScreen('connect')` -> redirect to `settings`
- localStorage keys remain unchanged

### Data Flow (unchanged)
Client (App.api) -> google.script.run.processApiRequest(params) -> Server (processApiRequest) -> SpreadsheetApp operations -> Return JSON response


---

## 9. Implementation Phases

### Phase A: Settings Screen Rewrite (Priority 1)
- Rewrite screen_settings.html with all sections
- Integrate spreadsheet creation + Meta Developer setup + OAuth

### Phase B: Remove Old Screens (Priority 2)  
- Remove render/bind functions from app_ui.html
- Update showScreen() routing in app_core.html

### Phase C: Navigation Cleanup (Priority 3)
- Update all showScreen() calls across all files
- Ensure no dead references to removed screens

### Phase D: Testing (Priority 4)
- Test first-time user flow
- Test returning user flow
- Test re-authentication flow
- Test account switching
- Test disconnect/reconnect

---

## 10. File Size Impact (Estimated)

| File | Before | After | Change |
|---|---|---|---|
| screen_settings.html | 16.8KB | ~35KB | +18KB (absorbs setup screens) |
| app_ui.html | 39.1KB | ~25KB | -14KB (removes setup screens) |
| app_core.html | 14.5KB | ~13KB | -1.5KB (simplified routing) |
| Total | 70.4KB | ~73KB | +2.6KB net |
