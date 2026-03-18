# Classroom Slide Builder — Install Guide

## Quick Install (Mac)

Open Terminal and paste this one command:

```
curl -sL https://raw.githubusercontent.com/BRANDONIBEN/classroom-slide-builder/main/install.sh | bash
```

This clones the plugin to `~/Documents/classroom-slide-builder/`.

## Load in Figma

1. Open **Figma** (desktop app)
2. Go to **Plugins** > **Development** > **Import plugin from manifest...**
3. Navigate to `Documents/classroom-slide-builder/` and select **manifest.json**
4. The plugin appears under **Plugins > Development > Classroom Slide Builder**

## Run the Plugin

In any Figma file: **Plugins** > **Development** > **Classroom Slide Builder**

Enter the team password when prompted. Ask Brandon if you don't have it.

## Updating

Double-click `update.command` in the plugin folder, or run:

```
cd ~/Documents/classroom-slide-builder && git pull
```

The plugin will show an "Update available" banner when a new version is released.

## Course Builder Web App

Author and manage slide content at: **classroom.brandoniben.com**

Sign in with your @passioncitychurch.com or @268generation.com Google account.

## How It Works

1. **Author slides** in the Course Builder web app
2. **Publish** — pushes a PDF to the Dropbox Watch Folder
3. **Auto-sync** — GitHub Actions converts the PDF to JSON every 15 minutes
4. **Build in Figma** — the plugin pulls the latest data automatically
