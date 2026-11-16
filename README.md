# Green PowerPoint 📊

An open-source, mobile & desktop-friendly **Microsoft Office 365 Presentation Editor** built for Android & Web.

Developed & Maintained by **Sachinthaka Sandaruwan** ([Green Soft Technologies](https://github.com/GreenSoftSachi)).

Part of the **Green Office Suite**:
- 📝 [Green Word](https://github.com/GreenSoftSachi/green-word)
- 📈 [Green Excel](https://github.com/GreenSoftSachi/green-excel)
- 📊 **Green PowerPoint**

---

## ✨ Features

- **Office 365 Ribbon Interface**: Seamless tabbed ribbon toolbar with File (Backstage), Home, Insert, Draw, Design, Transitions, Animation, and View tabs.
- **PowerPoint PPTX Import & Export**: Full compatibility with `.pptx` presentations using native parsing and generation.
- **Rich Elements**: Text boxes, shapes pool, lines, charts (eCharts), tables, math equations (LaTeX), and multimedia (audio/video).
- **Slide Show & Presenter Mode**: Fullscreen slide show presentation mode with keyboard / touch slide navigation.
- **Animations & Transitions**: Configurable slide transitions (Fade, Morph, Slide, Wipe) and element-level entrance/emphasis animations.
- **Offline First**: All presentations are saved locally via browser IndexedDB storage with zero server dependency.
- **Android APK & AAB**: Packaged into a fast, native Android WebView application.

---

## 🛠️ Architecture

- **Frontend Core**: Vue 3 + Vite + TypeScript (PPTist engine customized with Office 365 Ribbon UI)
- **Mobile Container**: Native Android WebView with AndroidX WebKit AssetLoader
- **Continuous Integration**: Automated GitHub Actions pipeline producing signed APK and Google Play Store App Bundle (AAB).

---

## 📱 Releases

Download the latest release APK directly from [GitHub Releases](https://github.com/GreenSoftSachi/green-powerpoint/releases).
