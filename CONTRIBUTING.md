# Contributing

Thank you for your interest in contributing to Fitness Tracker!

---

## Getting Started

The entire app lives in a single file: `fitness-tracker.html`. No build step is required to run it locally.

```bash
git clone https://github.com/hemo4321/fitness---app.git
cd fitness---app
open fitness-tracker.html   # or xdg-open on Linux
```

---

## Development Workflow

1. **Fork** the repository and create a branch from `main`:
   ```bash
   git checkout -b feat/your-feature-name
   ```

2. **Edit** `fitness-tracker.html` — React components are organized with `// ── SECTION NAME ──` dividers for navigation.

3. **Test** by refreshing the browser. Use browser DevTools (F12) to check the console for errors.

4. **AI features** — to test AI functionality locally, add your `sk-ant-...` key in the ⚙️ Settings tab.

5. **Submit a Pull Request** with a clear description of what changed and why.

---

## Code Style

- The file uses plain JSX-less React (`React.createElement`) — do not introduce a transpilation step or JSX.
- Inline styles only (no external CSS classes or stylesheets).
- No TypeScript — plain JavaScript only.
- Arabic UI strings belong alongside the component that uses them, not in a separate i18n file.
- Default to **no comments** — code should be self-explanatory through naming. Add a comment only when the *why* is non-obvious.

---

## Architecture Notes

| Concern | Approach |
|---------|---------|
| State management | Local component state (`useState`) — no Redux or context |
| Data persistence | `window.storage` wrapper around `localStorage` (see top of file) |
| API calls | Direct `fetch` to `api.anthropic.com` — key stored in `localStorage('fit_api_key')` |
| Routing | Single `tab` state string — no router library |
| Fonts | Google Fonts loaded at runtime inside the `<style>` tag in `FitnessApp` |

---

## Adding a New Tab

1. Create a new function component (e.g. `function MyTab() { ... }`) near the relevant section.
2. Add a new item to the `items` array in `BottomNav`.
3. Add `tab === 'mytab' && React.createElement(MyTab, ...)` in the `FitnessApp` return block.

---

## CI / APK Build

The build pipeline (`.github/workflows/build.yml`) requires an `EXPO_TOKEN` repository secret. It is not needed for browser-only development.

To trigger a build manually: GitHub → Actions → **Build APK** → Run workflow.

---

## Reporting Issues

Please open a GitHub Issue with:
- Browser and OS version
- Steps to reproduce
- Expected vs. actual behaviour
- Console errors (if any)
