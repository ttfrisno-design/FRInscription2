# Site de test (préproduction)

Le site de test permet de faire des modifications en profondeur **sans toucher au
site en production** ni à ses données (feuille Google, Drive, emails, paiements).

Les fichiers `index.html` et `backend/Code.gs` sont **les mêmes** pour les deux sites :
ils détectent tout seuls s'ils tournent en test ou en production. Une fois une
modification validée en test, il suffit de copier le fichier côté production.

```
             SITE DE TEST                                  PRODUCTION
  github.io/FRInscription-test/                github.io/FRInscription/
            │  (bandeau orange)                            │
            ▼                                              ▼
  Apps Script « FRI Inscriptions TEST »       Apps Script actuel
   FRI_MODE=test                               (aucune propriété FRI_MODE)
            │                                              │
            ▼                                              ▼
  Copie de la feuille Google                   Feuille de production
  Dossiers Drive « TEST - … »                  Dossiers Drive réels
  Emails → FRI_EMAIL_TEST uniquement           Emails aux adhérents
  HelloAsso bac à sable                        HelloAsso réel
```

## Mise en place (une seule fois)

### 1. Copie de la feuille Google
1. Ouvrir la feuille de production › **Fichier › Créer une copie**, la nommer
   « FRI Inscriptions — TEST ».
2. (Conseillé) Vider les onglets de données de la copie pour ne pas manipuler de
   vraies données personnelles en test, en gardant les lignes d'en-tête.
3. Noter l'identifiant de la copie (la partie entre `/d/` et `/edit` dans l'adresse).

### 2. Projet Apps Script de test
1. Sur https://script.google.com, ouvrir le projet actuel › **Vue d'ensemble ›
   Créer une copie** ; renommer la copie « FRI Inscriptions TEST ».
2. Dans la copie : **Paramètres du projet › Propriétés du script** :

   | Propriété | Valeur |
   |---|---|
   | `FRI_MODE` | `test` |
   | `FRI_SHEET_ID` | identifiant de la copie de la feuille (étape 1) |
   | `FRI_EMAIL_TEST` | votre adresse : **tous** les emails du test y partent |
   | `FRI_SECRET_TOKEN` | une valeur **différente** de la production |
   | `HA_CLIENT_ID`, `HA_CLIENT_SECRET`, `HA_ORG_SLUG` | identifiants du compte HelloAsso **bac à sable** (facultatif) |

   Sans `FRI_SHEET_ID`, le script de test refuse de démarrer (garde-fou).
3. Exécuter `initAdminCredentials` dans la copie (avec des mots de passe de test),
   puis effacer les mots de passe du code.
4. **Déployer › Nouveau déploiement › Application Web** (Exécuter en tant que : Moi,
   Accès : Tout le monde). Copier l'URL `/exec`.
5. Ne **pas** installer les déclencheurs (sauvegarde, fin de saison, audit) dans la
   copie, sauf pour les tester.

### 3. Site de test sur GitHub Pages
1. Créer un dépôt GitHub **`FRInscription-test`** (le mot « test » dans le nom
   active le mode test).
2. Y déposer `index.html`, après avoir collé l'URL `/exec` de l'étape 2 dans
   `CONFIG_ENV.test.appsScriptUrl` (en haut du script) et, si besoin, l'ID de la
   copie de la feuille dans `CONFIG_ENV.test.sheetIdPublic`.
3. **Settings › Pages** : publier la branche `main`. Le site est disponible sur
   `https://ttfrisno-design.github.io/FRInscription-test/`.

## Vérifier que le test est bien isolé
- Le site affiche un bandeau orange « SITE DE TEST ».
- Dans l'admin, le test de connexion répond « PONG … (TEST) ».
- Une inscription de test apparaît dans la **copie** de la feuille, pas en production.
- Les emails arrivent sur `FRI_EMAIL_TEST`, avec un sujet `[TEST → adresse prévue]`.

## Passer une modification en production
1. Tester sur le site de test.
2. Copier le même `index.html` dans le dépôt `FRInscription` (sans toucher à
   `CONFIG_ENV`, qui contient les deux configurations).
3. Copier le même `Code.gs` dans le projet Apps Script de production, puis
   **Gérer les déploiements › Modifier › Nouvelle version** (même URL).

## Ce que le mode test change
| | Production | Test |
|---|---|---|
| Feuille Google | feuille réelle | `FRI_SHEET_ID` (refuse de démarrer si absente) |
| Dossiers Drive | noms réels | préfixés « TEST - » |
| Emails | adhérents, trésorier, admin | tous détournés vers `FRI_EMAIL_TEST` |
| HelloAsso | api.helloasso.com | api.helloasso-sandbox.com |
| Lien de paiement direct | réel | désactivé |
| Données du navigateur | clés `fri_…` | clés `TEST__fri_…` (séparées) |
| Serveur appelé par le site | URL de production | URL de test uniquement (jamais celle de production) |

Le fichier des adhérents 2025-26 et celui des licences FFTT restent lus (lecture
seule) pour la recherche d'adhérent : le test n'y écrit jamais.
