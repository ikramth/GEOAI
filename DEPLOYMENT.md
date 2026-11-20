# Redeployment Steps (GeoAI web presentation)

## 1. Prepare the cleaned project
- Keep these runtime files:
  - `index.html`
  - `public/` (world-map.jpg, cartographie_drone.jpg, apple-layers.jpg, geoai-bands.jpg)
  - `generate_pptx.py` (optional, for USB .pptx regeneration)
  - `GeoAI_Presentation.pptx` (generated output)
- Removed legacy files: LICENSE, README.md, Recherche_Fondamentale_GeoAI_et_SIG.docx, mapinfo1.JPG, root-level duplicate images.

## 2. Verify locally
1. Open `index.html` directly in a modern browser (Chrome/Edge/Firefox).
2. Check navigation, language switcher (FR/EN/AR + RTL), Apple and Layers slides, author order.
3. Regenerate the PowerPoint if needed:
   ```
   pip install python-pptx
   python generate_pptx.py
   ```

## 3. Push changes (fresh, no stale artifacts)
```
git init
git add .
git commit -m "Clean GeoAI web build: i18n, slide order, cleanup"
git branch -M main
git remote add origin <YOUR_REPO_URL>
git push -u origin main
```

## 4. Trigger a fresh redeployment
- GitHub Pages: Settings → Pages → Source = `main` / root. Save, then re-run the Pages workflow or push an empty commit to force rebuild.
- Netlify: drag-and-drop the project folder or connect the repo; trigger "Clear cache and deploy site".
- Vercel: import the repo; set Framework Preset to "Other", Output Directory to `.`; redeploy with "Redeploy" (clear build cache).

## 5. Clear stale artifacts
- Bump asset filenames or add cache-busting query strings if images change.
- Hard-refresh the browser (Ctrl+F5) and test in an incognito window.
- Confirm `public/` assets resolve at their deployed URLs.
