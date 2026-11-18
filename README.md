# Agenda ICC — synchronisation automatique

Chaque jour, GitHub télécharge le fichier Excel des événements ICC
(lien SharePoint public) et publie un flux d'agenda :

    https://geswende.github.io/icc-agenda/icc-mtl.ics  (Montréal)
    https://geswende.github.io/icc-agenda/icc-lgl.ics  (Longueuil)
    https://geswende.github.io/icc-agenda/icc.ics      (les deux)

## Mise en place (tout dans le navigateur, aucune commande)

1. **Créer le dépôt** : github.com → **New repository**
   - Nom : `icc-agenda` — **Public** (requis pour GitHub Pages gratuit)
2. **Téléverser les fichiers** : *Add file → Upload files*, glisser tout le
   contenu du dossier (dont `.github`, `scripts`, `docs`) → *Commit*.
   - Si le dossier `.github` ne monte pas : *Add file → Create new file*,
     nom `.github/workflows/sync.yml`, coller le contenu.
3. **Ajouter le lien Excel en secret** : *Settings → Secrets and variables
   → Actions → New repository secret*
   - Names : `SOURCE_URL_MTL` (Montréal) et `SOURCE_URL_LGL` (Longueuil)
   - Secret : le lien SharePoint de chaque fichier
4. **Activer Pages** : *Settings → Pages → Source : Deploy from a branch*
   → Branch `main`, dossier `/docs` → *Save*
5. **Lancer une 1re fois** : onglet *Actions* → « Synchroniser agenda ICC »
   → *Run workflow*. Ensuite c'est automatique chaque nuit.
6. **S'abonner dans Google Agenda** (ordinateur) : *Autres agendas* →
   **+** → *À partir de l'URL* → coller l'URL `.ics` ci-dessus.
   - iPhone : *Réglages → Calendrier → Comptes → Ajouter un compte →
     Autre → Ajouter un calendrier avec abonnement*.

## Si les colonnes ne sont pas reconnues

L'onglet *Actions* affiche « colonnes introuvables ». Copier
`scripts/config.example.json` en `scripts/config.json` et y mettre les
titres exacts des colonnes du fichier Excel.

## Bon à savoir
- Google rafraîchit un abonnement toutes les 8 à 24 h (on ne peut pas forcer).
- Un événement sans heure devient « journée entière » ; sans heure de fin,
  il dure 2 h.
- Le lien Excel reste caché (secret), mais `icc.ics` est public.
- `docs/icc.json` contient les mêmes événements pour un tableau de bord.
