# Agenda ICC — synchronisation automatique

Chaque nuit, GitHub télécharge les fichiers Excel des calendriers ICC
Montréal et ICC Longueuil (liens SharePoint publics) et publie des agendas :

| Agenda | Lien d'abonnement |
|---|---|
| ICC Montréal — grands événements | https://geswende.github.io/icc-agenda/icc-mtl.ics |
| ICC Longueuil — grands événements | https://geswende.github.io/icc-agenda/icc-lgl.ics |
| Les deux ensemble | https://geswende.github.io/icc-agenda/icc.ics |
| ICC Montréal — détaillé | https://geswende.github.io/icc-agenda/icc-mtl-detail.ics |
| ICC Longueuil — détaillé | https://geswende.github.io/icc-agenda/icc-lgl-detail.ics |

S'abonner (Google Agenda sur ordinateur) : *Autres agendas* → **+** →
*À partir de l'URL*. iPhone : *Réglages → Calendrier → Comptes → Ajouter un
compte → Autre → Calendrier avec abonnement*.

## Fonctionnement
- Secrets : `SOURCE_URL_MTL`, `SOURCE_URL_LGL` (liens SharePoint, jamais affichés).
- « Grands événements » = onglet *Grands évènements* de chaque fichier.
- « Détaillé » = onglets mensuels (grilles), avec heure et salle quand elles
  sont écrites dans la case (ex. « Salle Zoé (21h à 3h) »).
- Lancer à la main : onglet *Actions* → « Synchroniser agenda ICC » → *Run workflow*.
- Google rafraîchit un abonnement toutes les 8 à 24 h.
