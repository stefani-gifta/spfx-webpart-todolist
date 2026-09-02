# Banner web part — setup guide

## 1. Delete the To Do List web part

In your SPFx solution, remove:

```
src/webparts/toDoList/ToDoListWebPart.ts
src/webparts/toDoList/ToDoListWebPart.module.scss
src/webparts/toDoList/ToDoListWebPart.manifest.json
src/webparts/toDoList/loc/en-us.js
src/webparts/toDoList/loc/mystrings.d.ts
```

(adjust the folder name if yours differs — it's whatever folder holds `ToDoListWebPart.ts`).

Then remove its entry from `config/config.json` under `"bundles"` (the key will be
something like `"to-do-list-web-part"`), and delete any leftover reference to it in
`config/serve.json` under `"initialPage"` if it points at that web part's preview.

## 2. Scaffold the new web part

From your solution root, run the Yeoman generator to create a properly wired-up
web part shell (this handles the manifest GUID, config.json bundle entry, and loc
files for you):

```
yo @microsoft/sharepoint
```

- **What to add**: WebPart
- **Web part name**: `Banner`
- **Framework**: `No JavaScript framework` (matches the style of the To Do List web part)

This creates `src/webparts/banner/...`. Now **replace** the generated files with the
ones in this package:

| Generated file | Replace with |
|---|---|
| `BannerWebPart.ts` | `BannerWebPart.ts` (this package) |
| `BannerWebPart.module.scss` | `BannerWebPart.module.scss` (this package) |
| `loc/en-us.js` | `loc/en-us.js` (this package) |
| `loc/mystrings.d.ts` | `loc/mystrings.d.ts` (this package) |

Add these new files into the same `src/webparts/banner/` folder:

| File | Purpose |
|---|---|
| `assets/company-logo.png` | Placeholder logo — **replace with your real company PNG**, same filename, or update the `import` path in `BannerWebPart.ts` |
| `png.d.ts` | Lets TypeScript import `.png` files. Skip this if your project already has a `declare module '*.png'` somewhere. |

Keep the manifest.json the generator created (it already has a valid GUID). The
`BannerWebPart.manifest.json` included here is just a reference/backup copy — only
use it if you hand-wire the web part instead of using the generator.

### If you'd rather wire it up by hand instead of using the generator

Copy the whole `banner` folder in this package to `src/webparts/banner/`, then add
this to `config/config.json` under `"bundles"`:

```json
"banner": {
  "components": [
    {
      "entrypoint": "./lib/webparts/banner/BannerWebPart.js",
      "manifest": "./src/webparts/banner/BannerWebPart.manifest.json"
    }
  ]
}
```

## 3. What the web part does

- **Greeting** — "Good morning" / "Good afternoon" / "Good evening" based on the
  visitor's local system clock.
- **Date** — formatted like `Wednesday, 2 September 2026`, using the site's UI
  culture.
- **Name** — first name shown large on the left; full display name shown under
  the avatar on the right. Both come from `context.pageContext.user`, so no API
  calls or permissions are needed.
- **Avatar** — tries the user's real SharePoint profile photo
  (`/_layouts/15/userphoto.aspx`); if that fails to load, it falls back to
  initials (e.g. "SG") automatically. The green dot is decorative, not live
  presence — wiring real presence needs Microsoft Graph permissions, which is a
  reasonable next step if you want it later.
- **Links** — Email and Calendar open Outlook on the web
  (`outlook.office.com`); OneDrive uses `/_layouts/15/onedrive.aspx` on the
  current site, which redirects to the visitor's own OneDrive. All open in a
  new tab.
- **Theme colors** — `onThemeChanged` reads the site's applied theme palette
  (`themeDarker` / `themeDark` / `themePrimary` etc.) and pushes them into CSS
  variables, so the left panel's gradient automatically matches whatever theme
  is applied to the SharePoint site — not a hardcoded red. The screenshot's red
  is simply what renders when a red-toned theme is active.
- **Company logo** — top-left of the banner, from `assets/company-logo.png`.
  Swap that file for your real logo (transparent-background PNG works best
  against the colored gradient). If you'd rather make it configurable per page
  instance instead of bundling a fixed file, add a `PropertyPaneTextField` for
  a logo URL and use that instead of the import — happy to add that if you
  want it.

No "here's what's happening..." subtext is included, per your request.
