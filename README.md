# Gestion des gains — installation

## Fichiers
- `index.html` : interface responsive avec van animé en CSS, coffre et porte conducteur interactifs.
- `Code_API.gs` : API Google Apps Script sécurisée.
- `manifest.webmanifest`, `sw.js` et trois icônes : PWA installable.

## Avant publication
1. Dans `Code_API.gs`, vérifiez `SPREADSHEET_ID` et lancez `setup()` depuis Apps Script.
2. Dans Apps Script, exécutez `setAdminPasswordFromPrompt()` et choisissez un mot de passe long et unique. Ne publiez jamais ce mot de passe dans GitHub.
3. Déployez le projet comme Application Web : exécution en tant que vous, accès « Tout le monde ». Copiez l'URL `/exec`.
4. Dans `index.html`, remplacez `REMPLACER_PAR_URL_APPS_SCRIPT_EXEC` dans `API_URL`.
5. Envoyez à la racine du dépôt GitHub les six fichiers du site : `index.html`, `manifest.webmanifest`, `sw.js`, `icon-192.png`, `icon-512.png`, `apple-touch-icon.png`. `Code_API.gs` doit rester dans Apps Script, pas dans GitHub Pages.
6. Activez GitHub Pages sur `main` / root.

## Sécurité
- Le mot de passe admin est stocké dans les Script Properties d'Apps Script, pas dans le site public.
- Le mot de passe employé est stocké sous forme de SHA-256 avec sel propre au compte.
- Cette application est conçue pour un usage personnel/petite équipe. Avant un usage professionnel sensible, faites auditer les accès et la configuration du déploiement.
- Gardez une copie régulière du Google Sheet.
