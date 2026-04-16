# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Fixed
- Added missing `x-api-key`, `anthropic-version: 2023-06-01`, and `anthropic-dangerous-direct-browser-access: true` headers to all three Anthropic API fetch calls — AI features were silently failing with 401 errors
- Updated model identifier from `claude-sonnet-4-20250514` to `claude-sonnet-4-5` (correct stable ID)
- Improved AI error messages: now shows a clear prompt to configure the API key when none is set

### Added
- **Settings tab (⚙️ إعدادات)** — secure local storage of the Anthropic API key via the new `fit_api_key` localStorage key
- `README.md` — full project documentation including quick start, build instructions, data model, and configuration
- `CHANGELOG.md` — this file
- `CONTRIBUTING.md` — contribution guidelines
- `.gitignore` — excludes build artifacts, `node_modules`, `.env` files, and OS metadata

---

## [1.0.0] — Initial Release

### Added
- Single-file React 18 app (`fitness-tracker.html`) with full Arabic (RTL) UI
- Dashboard tab: daily calorie ring, calories burned, net calories, weight logging, export/import
- Nutrition tab: manual meal entry, AI-powered meal estimation (calories + protein)
- Workout tab: resistance training (60+ exercises in Arabic) and 7 cardio types with calorie burn formulas
- Muscle Map tab: customizable muscle groups, training frequency tracking, visual SVG muscle map, AI exercise classification
- Progress tab: area/bar/line charts for calories, protein, and weight over selectable date ranges
- `localStorage` persistence for all data with JSON export/import
- GitHub Actions CI/CD pipeline building an Android APK via Expo + EAS Build
