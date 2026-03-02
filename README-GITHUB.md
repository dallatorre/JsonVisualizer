# Pubblicare l'add-in su GitHub + GitHub Pages

Questa guida pubblica il tuo add-in su GitHub Pages (HTTPS) senza IIS.

## 1) Crea repository GitHub e push

Dalla root del progetto:

```powershell
git init
git add .
git commit -m "Initial JsonVisualizer add-in"
git branch -M main
git remote add origin https://github.com/<TUO-UTENTE>/<TUO-REPO>.git
git push -u origin main
```

## 2) Abilita GitHub Pages via Actions

- Vai su `Settings > Pages`
- In `Build and deployment`, scegli `GitHub Actions`
- Il workflow `.github/workflows/deploy-pages.yml` farà build e deploy automatico ad ogni push su `main`

## 3) URL finale add-in

Il workflow imposta automaticamente `ADDIN_BASE_URL`:

- Repository standard: `https://<utente>.github.io/<repo>/`
- Repository user site (`<utente>.github.io`): `https://<utente>.github.io/`

Genera in `dist/manifest.json` gli URL corretti per `taskpane.html`, `commands.html`, `commands.js`.

## 4) Installa l'add-in in Outlook

Dopo il deploy, scarica `manifest.json` pubblicato (o usa quello di `dist` generato localmente con lo stesso URL) e installalo:

- New Outlook / Outlook Web: `Get Add-ins` -> `My add-ins` -> `Add a custom add-in` -> `Add from file`

## 5) Build locale equivalente (facoltativa)

Se vuoi testare localmente gli stessi URL GitHub Pages:

```powershell
$env:ADDIN_BASE_URL = "https://<utente>.github.io/<repo>/"
npm run build:ghpages
```

## Nota importante

Con New Outlook il Web Add-in richiede comunque connettività alla piattaforma Office/Outlook: GitHub Pages risolve l'hosting pubblico HTTPS ma non rende il client offline al 100%.
