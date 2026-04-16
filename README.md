# Fitness Tracker — تطبيق تتبع اللياقة البدنية

A mobile-first fitness tracking app built as a single HTML file. Tracks workouts, nutrition, weight progress, and muscle groups — with Arabic (RTL) UI and optional AI-powered features via the Claude API.

---

## Features

| Tab | Feature |
|-----|---------|
| ⚡ اليوم (Dashboard) | Daily calorie ring, net calories, weight logging, data export/import |
| 🥗 تغذية (Nutrition) | Manual meal entry, AI-powered calorie/protein estimation |
| 🏋️ تمارين (Workout) | Resistance training (60+ exercises) and cardio with calorie burn calculation |
| 💪 عضلات (Muscles) | Muscle group management, training frequency tracking, visual muscle map |
| 📈 تحليل (Progress) | Charts for calories, protein, and weight over 7/14/30/90 days |
| ⚙️ إعدادات (Settings) | Claude API key configuration for AI features |

---

## Quick Start

### Run in Browser

Open `fitness-tracker.html` directly in any modern browser. No build step required.

```bash
open fitness-tracker.html        # macOS
xdg-open fitness-tracker.html   # Linux
```

Data is persisted in `localStorage` — it stays between sessions automatically.

### AI Features (Optional)

The Nutrition tab (meal estimation) and Muscle Map tab (exercise classification) use the Claude API. To enable them:

1. Get an API key from [console.anthropic.com](https://console.anthropic.com) → API Keys → Create Key
2. Open the app → ⚙️ إعدادات tab
3. Paste your key (starts with `sk-ant-...`) and tap **حفظ المفتاح**

The key is stored locally in `localStorage` under `fit_api_key` and never sent anywhere except directly to `api.anthropic.com`.

---

## Build Android APK

The CI pipeline (`.github/workflows/build.yml`) builds an APK via Expo + EAS.

### Prerequisites

- GitHub repository with `EXPO_TOKEN` secret set
  - Get a token: [expo.dev](https://expo.dev) → Account → Access Tokens
  - Add as repo secret: Settings → Secrets → New secret → `EXPO_TOKEN`

### Trigger a Build

Push to `main` branch, or manually trigger via GitHub Actions → **Build APK** → Run workflow.

The build:
1. Creates an Expo project
2. Inlines React, ReactDOM, and Recharts libraries into the HTML (no CDN dependency)
3. Wraps the HTML in a React Native `WebView`
4. Builds an Android APK via EAS Build

Download the APK from the EAS dashboard or the Actions workflow artifact summary.

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| UI Framework | React 18.2.0 (UMD, no bundler) |
| Charts | Recharts 2.10.0 |
| Styling | CSS-in-JS (inline styles), RTL Arabic layout |
| Storage | `localStorage` API |
| AI | Anthropic Claude API (`claude-sonnet-4-5`) |
| Mobile Wrapper | Expo + React Native WebView |
| CI/CD | GitHub Actions + EAS Build |

---

## Project Structure

```
fitness---app/
├── fitness-tracker.html        # Entire app — single file
├── .github/
│   └── workflows/
│       └── build.yml           # Android APK build pipeline
├── README.md
├── CHANGELOG.md
└── CONTRIBUTING.md
```

---

## Data Model

All data is stored in `localStorage`:

| Key | Type | Description |
|-----|------|-------------|
| `fit_logs_v3` | `{ [date]: { meals: [], workouts: [] } }` | Daily logs |
| `fit_weight_v3` | `[{ date, weight }]` | Weight history |
| `fit_muscles_v1` | `[{ id, name, color, exercises }]` | Muscle group config |
| `fit_api_key` | `string` | Anthropic API key (local only) |

### Export / Import

On the Dashboard tab → export button generates a full JSON snapshot. Import accepts the same format to restore data.

---

## Configuration

User profile constants are defined at the top of `fitness-tracker.html`:

```js
const PROFILE = {
  weight: 100,        // kg
  height: 175,        // cm
  age: 30,
  gender: 'male',
  goalWeight: 80,     // kg
  startWeight: 100,   // kg
  tdee: 3021,         // Total Daily Energy Expenditure
  dailyCalories: 2100,
  dailyProtein: 160   // grams
};
```

Edit these values directly in the file to match your own profile.

---

## Browser Support

Any modern browser with `localStorage` and ES2020+ support:
- Chrome 88+ / Edge 88+
- Firefox 85+
- Safari 14+
- Android WebView (via the APK build)

---

## License

MIT
