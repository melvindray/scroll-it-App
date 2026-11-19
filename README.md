# Scroll-it

Feed vidéo façon TikTok pour les cours de médecine (Reichman). Application web installable (PWA) servie par GitHub Pages.

- **Ce dépôt** : l'application (`index.html`, `sw.js`, `config.json`).
- **Un dépôt par système** : `scroll-it_Respiratory`, `scroll-it_cardio`, … Chacun contient `manifest.json` et `clips/<film>/cNN.mp4` (540p, sans son) + `cNN_he.m4a` / `cNN_en.m4a`.
- Ajouter un système : créer le dépôt `scroll-it_<Système>` (public, Pages activé), puis l'ajouter à `config.json`.
