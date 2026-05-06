# Remove-SPOVersionHistory Nettoyage de l'historique de versions SharePoint Online

> 🇬🇧 [English version available here](README.md)

Script PowerShell pour nettoyer automatiquement l'historique de versions de fichiers sur tous les sites SharePoint Online de votre tenant.

>Ce script est un **outil d'optimisation et de gouvernance du stockage**.
> Il agit sur l'historique de versions des fichiers dans SharePoint Online
> et doit être utilisé en complément des politiques de rétention et de
> conformité Microsoft (SharePoint versioning, Microsoft Purview),
> sans les remplacer

---

## À qui s'adresse ce script

Ce script s'adresse aux administrateurs SharePoint Online, aux administrateurs Microsoft 365 et Entra ID, ainsi qu'aux équipes IT Gouvernance et Conformité.

Il n'est pas destiné aux utilisateurs finaux.

---

## Quand NE PAS utiliser ce script

N'utilisez pas ce script si vous n'avez pas validé la gouvernance avec les équipes métier, si vous n'avez pas analysé le rapport de simulation au préalable, ou si vous devez respecter une rétention légale stricte sans validation préalable de votre équipe conformité.

---

## Le problème

SharePoint Online conserve jusqu'à **500 versions par fichier** par défaut sur les bibliothèques existantes.
Le versioning intelligent a été introduit en 2024, mais il ne s'applique qu'aux nouvelles bibliothèques créées après son activation.

> **Important** : Activer le mode automatique ne nettoie PAS les versions existantes.
> Il s'applique uniquement aux nouvelles versions créées après le changement.
> Les anciennes bibliothèques continuent d'accumuler des versions silencieusement.

Un seul fichier PowerPoint modifié régulièrement peut consommer **plusieurs centaines de Mo** en versions inutilisées.
Multipliez ça par des milliers de fichiers et des dizaines de sites, vous avez un problème de stockage.

> **Avant d'acheter du stockage supplémentaire, auditez votre historique de versions.**

---

## Ce que ce script ne fait PAS

Ce script ne supprime pas de fichiers, ne contourne pas les étiquettes de rétention, ne remplace pas Microsoft Purview, ne modifie pas les paramètres de versioning SharePoint et ne remplace pas une politique de conformité ou de gestion des documents.

---

## Pourquoi pas New-SPOSiteFileVersionBatchDeleteJob ?

Microsoft propose une commande native pour nettoyer les versions existantes :

```powershell
New-SPOSiteFileVersionBatchDeleteJob -Identity $siteUrl -Automatic
```

> ⚠️ **Problème majeur** : Les versions supprimées par cette commande ne vont **PAS** dans la corbeille et ne sont **pas récupérables**.

Ce script offre un mode simulation avant toute suppression, un mode corbeille récupérable (~93 jours), un rapport CSV exportable dans Power BI, un résumé JSON à chaque run, une politique différenciée Normal vs Critique et des logs complets pour audit.

---

## Fonctionnalités

- Analyse **tout le tenant automatiquement** en une seule exécution (testé sur 200+ sites)
- **Deux modes d'authentification** : Interactif (laptop) ou Certificat (serveur)
- **Mode test sur 1 ou plusieurs sites** avant de lancer sur tout le tenant
- **Politique de rétention différenciée** : Sites normaux vs sites critiques
- **Mode simulation** avant toute suppression commencez toujours par là
- **Mode corbeille** pour le premier run en production (récupérable ~93 jours)
- **Retry automatique** sur timeout HTTP avec backoff progressif (10s, 20s, 30s)
- **Skip automatique des gros fichiers** en mode interactif (configurable via MaxFileSizeMB)
- **Rapport CSV** exportable directement dans Power BI
- **Résumé JSON** généré à chaque run
- **RunId / RunMode / RunDate** pour comparer les runs entre eux dans Power BI
- **SiteType** (Normal/Critique) dans chaque ligne du rapport
- **Ignore les bibliothèques système** automatiquement (SharePoint EN + FR)
- **Ignore les sites système** (OneDrive, portals, search)
- Sites avec accès refusé listés séparément dans le rapport final
- Confirmation humaine obligatoire avant tout run en production
- Historique complet dans un fichier log horodaté
- **Aucun secret dans le script** utilise des variables d'environnement

---

## Rapport Power BI inclus

Un fichier `.pbit` (template) est disponible dans le repository pour visualiser les résultats.
Il ne contient aucune donnée réelle il suffit de le connecter à votre dossier CSV.

```
Pages disponibles :
Executive Summary    : KPI globaux, top sites, répartition par extension
Analysis by Site     : Scatter plot GB vs Versions, table détaillée
Analysis by File     : Top 20 fichiers, table avec drill-down
Timeline             : Espace récupérable et versions par mois
Run Comparison       : Comparaison SIMULATION vs PRODUCTION-RECYCLE
```

Pour l'utiliser :
```
1. Copier vos fichiers CSV dans C:\Temp\SPOVersionCleanup\
2. Ouvrir le fichier .pbit dans Power BI Desktop
3. Mettre à jour le paramètre pReportFolder
4. Refresh
```
---
## Power BI Report Preview
![Executive Summary](images/executive-summary.png)

![Analysis by Site](images/analysis-by-site.png)
---

## Politique de rétention

| Type de site | Politique | Versions conservées |
|---|---|---|
| Normal | Option A : Total | 10 versions au total (9 historique + courante) |
| Critique | Option B : Historique | 50 versions historique + courante |

> **Note technique** : `Get-PnPFileVersion` ne retourne pas la version courante.
> Option A "10 au total" = 9 historique + courante = 10 versions conservées au total.

---

## Relation avec les paramètres natifs SharePoint

SharePoint Online dispose de contrôles natifs de versioning :
- **Basé sur l'âge** : versions plus vieilles que X jours supprimées
- **Basé sur le nombre** : maximum X versions majeures conservées

Ce script ajoute une **troisième couche** :
- **Basé sur le seuil métier** : politique Normal/Critique définie par votre organisation

Ces règles sont **complémentaires**, pas contradictoires.
SharePoint protège l'historique dans le temps.
Ce script protège le stockage avec des limites contrôlées et auditables.

---

## OneDrive

Les sites OneDrive personnels sont **ignorés par défaut**.

Raisons :
- Données personnelles : considérations de confidentialité
- Permissions différentes requises
- Recommandé de traiter séparément avec approbation RH/direction

Pour inclure OneDrive, commentez ou supprimez la condition de skip dans le script.

---

## Prérequis

- PowerShell 7.4+
- PnP.PowerShell 3.x+
- App Registration Microsoft Entra ID (voir section Authentification)
- Rôle **SharePoint Administrator** ou **Global Administrator**

---

## Authentification

Ce script supporte **deux modes d'authentification**.
Aucun secret n'est stocké dans le script toutes les valeurs sont passées via des variables d'environnement.

---

### Mode 1 : Interactif (laptop / test local)

Idéal pour : premier test, validation de la connexion, test sur 2-3 sites maximum.

> ⚠️ **Limitation importante** : Le mode interactif est conçu pour la **validation et les tests uniquement**.
> Pour les gros fichiers (.pbix, .pptx) ou les tenants avec 50+ sites,
> utilisez le mode certificat même sur un laptop pour éviter les timeouts.

#### Créer l'App Registration Entra ID

```
portal.azure.com
→ Microsoft Entra ID
→ App registrations
→ New registration
   Name         : SP-Version-Cleanup
   Redirect URI : https://login.microsoftonline.com/common/oauth2/nativeclient
→ Register
→ Copier l'Application (client) ID
```

#### Ajouter les permissions

```
→ API permissions → Add a permission

Permission 1 : obligatoire pour le mode interactif
   SharePoint → Delegated → AllSites.FullControl

Permission 2 : pour le mode certificat
   SharePoint → Application → Sites.FullControl.All

→ Grant admin consent for [votre organisation]
→ Vérifier que les deux permissions affichent : Granted ✅
```

> ⚠️ **Important** : La permission déléguée `AllSites.FullControl` est **obligatoire** pour le mode interactif.
> Sans elle, vous obtiendrez une erreur 403 sur `Get-PnPTenantSite`,
> même si vous avez le rôle SharePoint Administrator ou Global Administrator.

#### Configuration et lancement

```powershell
$env:SP_TENANT    = "votre-tenant.onmicrosoft.com"
$env:SP_CLIENT_ID = "votre-client-id-entra"

.\Remove-SPOVersionHistory.ps1 -UseInteractive
```

---

### Mode 2 : Certificat (serveur / automatisé / production)

Idéal pour : serveurs, tâches planifiées, scan complet du tenant. Tourne sans surveillance aucun popup.

#### Étape 1 : Créer le certificat et l'App Registration

```powershell
# Sur votre serveur PowerShell 7 en administrateur
Register-PnPEntraIDApp `
    -ApplicationName "SP-Version-Cleanup" `
    -Tenant "votre-tenant.onmicrosoft.com" `
    -OutPath "C:\Certs" `
    -CertificatePassword (ConvertTo-SecureString "VotreMotDePasse" -Force -AsPlainText) `
    -SharePointApplicationPermissions "Sites.FullControl.All" `
    -DeviceLogin
```

Cette commande génère :
```
C:\Certs\SP-Version-Cleanup.pfx  (clé privée garder secret)
C:\Certs\SP-Version-Cleanup.cer  (clé publique pour Entra ID)
App Registration créée dans Entra ID
Permission Sites.FullControl.All assignée
ClientId affiché dans la console le noter
```

#### Étape 2 : Uploader le certificat dans Entra ID

```
portal.azure.com
→ Microsoft Entra ID → App registrations
→ SP-Version-Cleanup
→ Certificates & secrets → Certificates
→ Upload certificate → Choisir SP-Version-Cleanup.cer
→ Add
```

#### Étape 3 : Donner le consentement admin

```
→ API permissions
→ Grant admin consent for [votre organisation] → Yes
→ Vérifier : Sites.FullControl.All → Granted ✅
```

#### Étape 4 : Configuration et lancement

```powershell
$env:SP_TENANT        = "votre-tenant.onmicrosoft.com"
$env:SP_CLIENT_ID     = "votre-client-id-entra"
$env:SP_CERT_PATH     = "C:\Certs\SP-Version-Cleanup.pfx"
$env:SP_CERT_PASSWORD = "VotreMotDePasse"

.\Remove-SPOVersionHistory.ps1
```

---

### Optionnel : Surcharger les URLs

Par défaut, les URLs sont dérivées automatiquement depuis SP_TENANT.
Si votre URL SharePoint diffère de votre nom de tenant :

```powershell
$env:SP_TENANT_URL = "https://votre-tenant.sharepoint.com"
$env:SP_ADMIN_URL  = "https://votre-tenant-admin.sharepoint.com"
```

---

## Démarrage rapide

### Étape 1 : Définir vos sites critiques

Ouvrez le script et personnalisez `$CriticalSites` :

```powershell
$CriticalSites = @(
    "Accounting",     # /sites/Accounting
    "LegalAffairs",   # /sites/LegalAffairs
    "PeopleOps",      # /sites/PeopleOps
    "RegulatoryDocs", # /sites/RegulatoryDocs
    "AuditReports"    # /sites/AuditReports
)
```

> **Important** : `$CriticalSites` utilise une correspondance simple par mots-clés.
> Validez la liste avec votre équipe conformité avant la production.

#### Comment identifier vos mots-clés de sites critiques

```
URL du site :
https://tenant.sharepoint.com/sites/AuditReports2024
                                      ↑
                          Utilisez cette partie : "AuditReports"
```

**3 façons de trouver vos noms de sites :**

```
Méthode 1 : URL du navigateur
   Ouvrez le site → copiez la partie après /sites/

Méthode 2 : SharePoint Admin Center
   Active sites → colonne URL → partie après /sites/

Méthode 3 : PowerShell
   Get-PnPTenantSite | Select-Object Url
```

**Bonnes pratiques :**
- Utilisez des mots-clés stables et descriptifs
- Évitez les valeurs trop courtes risque de faux positifs
- Validez avec votre équipe conformité

---

### Étape 2 : Tester sur 1 ou 2 sites avant tout (fortement recommandé)

```powershell
$env:SP_TENANT    = "votre-tenant.onmicrosoft.com"
$env:SP_CLIENT_ID = "votre-client-id"

# Tester sur un seul site
.\Remove-SPOVersionHistory.ps1 -UseInteractive -TestSite "AuditReports"

# Tester sur plusieurs sites
.\Remove-SPOVersionHistory.ps1 -UseInteractive -TestSites @("Accounting","LegalAffairs")
```

> **Note** : En mode interactif avec `-TestSite`, le script se connecte **directement** au site
> sans appeler `Get-PnPTenantSite`. Cela évite les erreurs 403 liées aux droits admin center.

---

### Étape 3 : Simulation sur tout le tenant

```powershell
# ModeTest = $true par défaut aucune suppression effectuée
.\Remove-SPOVersionHistory.ps1
```

Analysez le rapport CSV dans `C:\Temp\SPOVersionCleanup\`

---

### Étape 4 : Run en production mode corbeille

Modifiez le script :
```powershell
$ModeTest    = $false
$ModeRecycle = $true
```

Lancez :
```powershell
.\Remove-SPOVersionHistory.ps1
```

Attendez **2-3 semaines**. Vérifiez qu'aucun incident n'est signalé.

---

### Étape 5 : Suppression définitive (optionnel)

```powershell
$ModeTest    = $false
$ModeRecycle = $false
```

---

## Ordre d'exécution recommandé

```
Étape 1 : ModeTest=$true  + ModeRecycle=$true  : Simulation   (analyser le CSV)
Étape 2 : ModeTest=$false + ModeRecycle=$true  : Production   (corbeille ~93 jours)
Étape 3 : ModeTest=$false + ModeRecycle=$false : Définitif    (optionnel)
```

> `$ModeTest` et `$ModeRecycle` sont intentionnellement codés en dur comme mécanisme de sécurité.
> Vous devez les modifier explicitement avant de lancer en production.

---

## Gros fichiers et timeouts

Le paramètre `MaxFileSizeMB` est automatiquement configuré selon le mode d'authentification :

```
Mode interactif  : MaxFileSizeMB = 800  (évite les timeouts PnP 100s)
Mode certificat  : MaxFileSizeMB = 0    (pas de limite, connexion stable)
```

Les fichiers skippés apparaissent dans les logs :
```
[WARN] SKIP_LARGE : 1208 MB > limit 800 MB
[WARN] TIP : Set MaxFileSizeMB=0 with certificate mode to process this file
```

---

## Vérifier l'espace libéré après le recycle bin

> ⚠️ **Important** : Le mode corbeille ne libère PAS l'espace immédiatement.
> Les versions en corbeille comptent encore dans le quota SharePoint.

L'espace est réellement libéré quand la corbeille est vidée manuellement, après ~93 jours automatiquement ou en lançant `ModeRecycle=$false`.

> Après avoir vidé la corbeille, l'espace peut prendre **24-48 heures**
> pour se mettre à jour dans le SharePoint Admin Center.

**3 façons de vérifier l'espace libéré :**

```powershell
# 1. Via PowerShell (précis et immédiat)
Get-PnPTenantSite `
    -Url "https://votre-tenant.sharepoint.com/sites/VotreSite" |
    Select-Object Url, StorageUsageCurrent

# 2. SharePoint Admin Center
#    Active sites → colonne Storage used

# 3. Corbeille du site
#    /sites/VotreSite → Site contents → Recycle bin
```

---

## Comparer les runs dans Power BI

Chaque run génère un rapport CSV avec les colonnes :

| Colonne | Description |
|---|---|
| RunId | Identifiant unique (ex: 20260427_121235) |
| RunMode | SIMULATION / PRODUCTION-RECYCLE / PRODUCTION-DELETE |
| RunDate | Date du run |
| SiteType | Normal / Critical |
| Action | SIMULATION / RECYCLED / DELETED |

Pour comparer simulation vs production :
```
1. Copier tous les CSV dans C:\Temp\SPOVersionCleanup\
2. Power BI lit automatiquement tout le dossier
3. Filtrer par RunMode pour comparer
4. La colonne File est la clé de jointure entre les runs
```

---

## Colonnes du rapport CSV

| Colonne | Description |
|---|---|
| Site | URL du site SharePoint |
| Library | Nom de la bibliothèque |
| File | Chemin complet du fichier |
| SiteType | Normal / Critical |
| HistoryVersions | Versions historiques trouvées |
| ThresholdKept | Versions historiques conservées |
| Removed | Versions supprimées |
| SpaceMB | Espace récupéré en MB |
| LastModified | Date de dernière modification |
| Action | SIMULATION / RECYCLED / DELETED |
| RunId | Identifiant du run |
| RunMode | Mode du run |
| RunDate | Date du run |

---

## Note importante

Ce script réduit la consommation de stockage en supprimant les anciennes versions de fichiers.
Il ne **remplace pas** les étiquettes de rétention Microsoft Purview ni la gestion des documents pour la conformité réglementaire.

---

## Changelog

### v1.0.3
- NEW : Validation mutuelle `-TestSite` vs `-TestSites` (pas les deux simultanément)
- NEW : `MaxFileSizeMB` automatique selon le mode (800 interactif / 0 certificat)
- NEW : `-TestSite` et `-TestSites` se connectent directement au site sans `Get-PnPTenantSite`
- NEW : Log "PnP 100s timeout" plus explicite
- FIX : Version interne retirée du header

### v1.0.2
- NEW : RunId, RunMode, RunDate dans chaque ligne CSV
- NEW : SiteType (Normal/Critical) dans chaque ligne CSV
- NEW : Résumé JSON généré à chaque run
- NEW : Retry progressif sur timeout (10s, 20s, 30s)
- NEW : Skip automatique des gros fichiers en mode interactif
- FIX : Continue après timeout au lieu de bloquer indéfiniment

### v1.0.1
- NEW : `-TestSite` pour tester sur un site spécifique
- NEW : `-TestSites` pour tester sur plusieurs sites

### v1.0.0
- NEW : `-UseInteractive` pour tester sur laptop sans certificat
- Un seul script pour laptop (interactif) et serveur (certificat)

---

## Licence

Licence MIT libre d'utilisation, de modification et de distribution.

## Auteur

Abdoulaye Ndao
[LinkedIn](https://www.linkedin.com/in/abdoulaye-ndao-m-sc-055241126/) | [GitHub](https://github.com/abdoulayendao007/Remove-SPOVersionHistory)
