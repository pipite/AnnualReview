# Vue d'ensemble

**AnnualReview.ps1** est un script PowerShell de traitement des données RH pour la gestion des évaluations annuelles et des données employés.

Le script lit les données de diverses sources, y applique un traitement, puis met à jour les tables d'une base de données.

* Source de données provenant de plusieurs sources au format **CSV, XLS et XLSX.**
* Synchronise des tables d'une base de donnée **PostgreSQL**

## Source des données  

|                   Paramètres .ini                   |                    Fichiers                     |
| --------------------------------------------------- | ----------------------------------------------- |
| [**CSV_Salaries**][fichierCSV]                      | req22_Revue_annuelle_Admin-xxx.csv              |
| [**XLS_Salaries-complement**][fichierXLS]           | Progessi - Liste BM_xxx.xls                     |
| [**CSV_Salaires-primes**][fichierCSV_Histo]         | 1 - Revue annuelle - Primes et charges xxx.csv  |
| [**CSV_Salaires-primes**][fichierCSV_Current]       | req23_Revue_annuelle_primes_charges-xxx.csv     |
| [**XLS_CSV_Frais**][fichierXLS_Fusion_histo]        | DRH - Frais Policy_Fusion_Group_YTD_2024xxx     |
| [**XLS_CSV_Frais**][fichierXLS_Fusion_current]      | DRH - Frais Policy_Fusion_Group_YTD_2025xxx.xls |
| [**XLS_CSV_Frais**][fichierXLS_Fortil_histo]        | DRH - Frais Policy_FORTIL GROUP_YTD_2024xxx.xls |
| [**XLS_CSV_Frais**][fichierXLS_Fortil_current]      | DRH - Frais Policy_FORTIL GROUP_YTD_2025xxx.xls |
| [**XLS_CSV_Frais**][fichierXLS_Reservation-berceau] | Politique famille - réservation berceau.xlsx    |
| [**XLS_CSV_Frais**][fichierCSV_famille]             | req21_Revue_annuelle__Absences-xxx.csv          |
| [**CSV_entrepreneuriat**][fichierCSV]               | import_valorisation_entrepreneuriat.csv         |
| [**XLS_Projets-liste**][fichierXLS]                 | OLAP-Projets_liste_xxx.xls                      |
| [**XLS_Projets-affectation**][fichierXLS]           | OLAP-affectations_.xlsx                         |
| [**CSV_historiques_salaires**][fichierCSV]          | 1 - Revue annuelle - Salaires contractuels.csv  |

## Base de données PostgreSQL

|                   Paramètres .ini                   |            Tables             |
| --------------------------------------------------- | ----------------------------- |
| [SQL_Postgre_Review][utilisateurs]                  | utilisateurs                  |
| [SQL_Postgre_Review][remuneration]                  | remuneration                  |
| [SQL_Postgre_Review][performance_co]                | performance_co                |
| [SQL_Postgre_Review][protection sociale]            | protection sociale            |
| [SQL_Postgre_Review][conges_utilisateur]            | conges_utilisateur            |
| [SQL_Postgre_Review][Utilisateurs_frais]            | Utilisateurs_frais            |
| [SQL_Postgre_Review][vallorisation_entrepreneuriat] | vallorisation_entrepreneuriat |
| [SQL_Postgre_Review][projets]                       | projets                       |
| [SQL_Postgre_Review][affectations]                  | affectations                  |
| [SQL_Postgre_Review][historiques_salaires]          | historiques_salaires          |

## Principe de mise à jour

A partir des fichiers sources, traiter les datas et **générer des hashtables (src) ayant la même structure que les tables**.

* Charge les données depuis les fichiers sources **CSV, XLS** et **XLSX.**
* Applique un traitement (calcul, exceptions, règles), puis génère le résultat du traitement dans des hashtables (src)
* Charge des hashtables (dst) depuis les tables de la base de données.
* Compare les hashtables src/dst
* Synchronise les tables de la base de données si nécessaire. Seuls les enregistrements en écart sont modifiés.
    * Si nouvel enregistrement, envoi une commande SQL **INSERT INTO**
    * Si un enregistrement existant est modifié, envoi une commande SQL **UPDATE TABLE**
    * Ne fait rien si pas d'écart de données.
    
**Nota important** : *Ce script ne gère pas les éventuelles suppressions de données d'un fichier source. Une fois qu'un enregistrement a été créé en base de données, si on supprime ensuite cet enregistrement des fichiers sources, cet enregistrement restera en base de données.*

**Si vous modifiez les conditions d'exceptions d'inclusion/exclusion du fichier .ini, il peut être nécessaire selon les cas faire une trucate total de la base de données et relancer le script.**

|             Table             | Hashtable (src) |  Hashtable (dst)  |
| ----------------------------- | --------------- | ----------------- |
| utilisateurs                  | $script:USER    | $script:BDDuser   |
| remuneration                  | $script:REMUN   | $script:BDDRemun  |
| performance_co                | $script:PERF    | $script:BDDPerf   |
| protection sociale            | $script:PROT    | $script:BDDProt   |
| conges_utilisateur            | $script:CONGES  | $script:BDDConges |
| Utilisateurs_frais            | $script:FRAIS   | $script:BDDFrais  |
| vallorisation_entrepreneuriat | $script:VALOR   | $script:BDDValor  |
| projets                       | $script:PROJ    | $script:BDDProj   |
| affectations                  | $script:AFFECT  | $script:BDDAffect |
| historiques_salaires          | $script:HISTO   | $script:BDDHisto  |

# Traitements

## Conventions

Les fonctions **Query_XLS_xxx** chargent en mémoire les fichiers sources **XLS ou XLSX**

Les fonctions **Query_CSV_xxx** chargent en mémoire les fichiers sources **CSV**

Les fonctions **Compute_xxx** effectuent les traitements et verifications d'exclusions

Les fonctions **Query_BDD_xxx** chargent une table en memoire

Les fonctions **Update_BDD_xxx** effectuent la synchronisation d'une table

## Initialisation de la hashtable $script:USER

Cette hashtable contient information du fichier source pour chaque [**Matricule**]

La clé d'unicité est "**Matricule**" (format "**00000999**")

### Sources

|        Paramètres .ini         |              Fichiers              |
| ------------------------------ | ---------------------------------- |
| [**CSV_Salaries**][fichierCSV] | req22_Revue_annuelle_Admin-xxx.csv |

### Code dans le script

Les fonctions du script concernant l'Initialisation de la hashtable $script:**USER** sont celle ci: 
```
Query_CSV_Salaries
	Compute_USER
```

### Fonction Query_CSV_Salaries


La fonction Query_CSV_Salaries charge les users dans $script:**USER** et $script:**USERAll**

### Fonction Compute_USER

La fonction **Compute_USER** excluera un utilisateur de $script:**USER** si :

* [**InfosG Observations**] = **PROFIL SECONDAIRE**
* [**Date d'entrée**] > **Today**
* [**Date de sortie dans la société**] < **Today**
* [**Date d'entrée**] ou [**Date de sortie dans la société**] au mauvais format de date.

Chaque utilisateur exclu sera 

* supprimé de la hashtable $script:**USER**
* ajouté à la hashtable $script:**EXCLUS**

## Initialisation de la hashtable $script:ListeBM

Cette hashtable contient les informations du fichier source pour chaque [**Matricule**]

La clé d'unicité est "**Matricule**" (format "**999**")

### Sources

|              Paramètres .ini              |          Fichiers           |
| ----------------------------------------- | --------------------------- |
| [**XLS_Salaries-complement**][fichierXLS] | Progessi - Liste BM_xxx.xls |

Liste des utilisateurs non exclus de **$script:USER**

### Code dans le script

Les fonctions du script concernant l'Initialisation de la hashtable $script:**ListeBM** sont celle ci: 
```
Query_XLS_Salaries-complement
	Compute_ListeBM
```
### Fonction Query_XLS_Salaries-complement

Charge dans $script:**table_XLS_Salaries_complement** tous les enregistrements du fichier excel.

### Fonction Compute_ListeBM

Parcour tous les enregistrements de $script:**table_XLS_Salaries_complement** et les ajoute à $script:**ListeBM**
 
**Sauf si**

* [**Matricule**] est vide, égal à 0, ou n'est pas un nombre
* [**Matricule**] existe déjà dans $script:**ListeBM** (doublon)
* [**Type de contrat**] = Stagiaire,Apprendistato,... (**[Exclude][contrat]** du .ini)
* [**Dénomination sociale**] = FORTIL BELGIUM,...  (**[Exclude][denomination_sociale]** du .ini)
* [**Classe horaire de présence**] = TEMPS PARTIEL 1H,... (**[Exclude][Classe horaire de présence]** du .ini)
* [**UPN O365_act**] est vide

**Et seulement si**

* [**Matricule**] existe dans **$script:USER**
* [**statusId**] = PARTY_ENABLED,... ( **[Include Only][statusId]** du .ini)
* [**Interne/externe**] = Profil interne,... ( **[Include Only][Interne/externe]** du .ini)

La fonction **Compute_ListeBM** ajoute à la liste des exclus $script:**EXCLUS** tous les matricules présent dans $script:**USER**, sans correspondance dans les non exclus de $script:**ListeBM**

## Initialisation de la hashtable $script:Review

Cette hashtable contient les montants cumulés pour chaque [**Matricule**]  [**année**] [**CatRub**]  

Les clés d'unicité sont [**Matricule**] format "**00000999**", [**année**] et [**CatRub**]

### Sources

|                Paramètres .ini                |                    Fichiers                    |
| --------------------------------------------- | ---------------------------------------------- |
| [**CSV_Salaires-primes**][fichierCSV_Histo]   | 1 - Revue annuelle - Primes et charges xxx.csv |
| [**CSV_Salaires-primes**][fichierCSV_Current] | req23_Revue_annuelle_primes_charges-xxx.csv    |

liste des utilisateurs non exclus de $script:**USER**

### Code dans le script

```
Query_CSV_Salaires-primes
```
### Fonction Query_CSV_Salaires-primes

Charge et cumule les deux fichiers sources (historique et année courrante) dans $script:**Review**

Cumule les montants pour chaque [**Matricule**]  [**année**] [**CatRub**]

Les lignes qui ont [**CatRub**] vide ne sont pas comptabilisées.

## Hashtable auxiliaire $script:EXCLUS

Cette hashtable contient les matricules exclus du traitement avec la raison de leur exclusion.

La clé d'unicité est "**Matricule**" (format "**999**" sans zéros initiaux)

### Fonction Add-Exclusion

La fonction **Add-Exclusion** est appelée par diverses fonctions de traitement pour enregistrer les matricules exclus.

**Paramètres :**
* `$matricule` : Le matricule à exclure
* `$source` : La source de l'exclusion (nom de la fonction ou du fichier)
* `$raison` : La raison de l'exclusion
* `-USER` : Switch optionnel. Si présent, ajoute le matricule à `$script:EXCLUS`

**Comportement :**
* Enregistre l'exclusion dans le fichier de log INA (inactifs)
* Si le switch `-USER` est présent et que le matricule n'existe pas déjà dans `$script:EXCLUS`, l'ajoute avec sa raison
* Gère les cas particuliers : matricules vides, matricules non numériques, etc.

**Utilisation :**

La fonction est appelée dans plusieurs contextes :
* **Compute_USER** : Exclusion des profils secondaires, dates d'entrée/sortie invalides
* **Compute_ListeBM** : Matricules présents dans USER mais absents de ListeBM
* **Query_CSV_Salaires-primes** : Matricules invalides ou absents de USER
* **Autres fonctions de traitement** : Validation des données sources

**Initialisation :**

`$script:EXCLUS` est initialisée comme hashtable vide dans la fonction **Compute_USER** (ligne 180 du script).

## Initialisation de la hashtable $script:PRIMAIRE

Cette hashtable contient les matricules qui ont un ou plusieurs profil secondaire.

Les clés d'unicité sont [**MatriculePrimaire**] [**MatriculeSecondaire**]

### Sources

$script:**USER**

$script:**ListeBM**

### Code dans le script

```
Compute_Profil_secondaire
```

### Fonction Compute_Profil_secondaire

Pour chaque **Matricule** dans $script:**USER**, si le champ **InfosG Observations** contient le mot **PRIMAIRE**, alors on extrait la racine de l'adresse email à partir de $script:**ListeBMAll**[**Matricule**]['UPN O365_act'].

Ensuite, on cherche dans tous les **autresmatricules** de $script:**ListeBMAll** ceux dont l'adresse email possède la même racine, et s'ils existent et sont actifs, on les ajoute dans $script:**PRIMAIRE**[**Matricule**][**autresmatricules**].

Nota: On doit utiliser $script:**ListeBMAll** qui contient tous les users, y compris les innactifs, puisque dans $script:**ListeBM**, les comptes secondaires ont été exclus puisque exclus de $script:**USER**.
    
## Synchronisation de la table [utilisateur]

Synchronisation de la table utilisateurs de la base PostgreSQL avec les données RH en gérant les hiérarchies et rôles.

### Sources

$script:**BDDuser**

$script:**USER**

$script:**ListeBM**

### Code dans le script

```
Query_BDD_Utilisateurs
	Compute_Managers
	Compute_Administrator
	Prepare_two_pass
Update_BDD_utilisateurs
```

### Descriptif de l'enchaînement de ces 5 fonctions

#### Query_BDD_utilisateurs

Charge les données utilisateurs existantes depuis la base de données PostgreSQL dans $script:**BDDuser**

#### Compute_Managers

Analyse les relations hiérarchiques entre utilisateurs

Attribue les rôles "collaborateur" ou "manager" selon les UPN O365 des managers

Établit les liens id_manager entre collaborateurs et leurs managers

#### Compute_Administrator

Traite les exceptions administratives définies dans le fichier .ini

Attribue le rôle "admin" aux utilisateurs spécifiés dans la section [users_USER]

Gère les exceptions de désactivation dans [users_desactive_exception] du .ini

#### Prepare_two_pass

Prépare les données en 2 passes pour la mise à jour de la base :

Pass 1 : tous les attributs utilisateur (nom, prénom, email, rôle, etc.) sauf manager_id

Pass 2 : uniquement les manager_id (pour éviter les conflits de clés étrangères)

*Nota : Ces deux passes sont nécessaire afin d'eviter d'affecter à un user, un idmanager qui n'aurait pas encore été créé, ce qui causerai un échec de contrainte référentielle dans la base de données.*

#### Update_BDD_utilisateurs

Exécute la mise à jour en base de données en 2 étapes :

* Passe 1 : met à jour tous les attributs utilisateur
* Passe 2 : met à jour les relations manager après que tous les utilisateurs existent

## Synchronisation de la table [remuneration]

Calculer et synchroniser en base les rémunérations annuelles consolidées de tous les utilisateurs, en gérant les profils multiples et en agrégeant les différentes composantes salariales.

### Sources

$script:**BDDRemun**

$script:**USER**

$script:**PRIMAIRE**

$script:**Review**

### Code dans le script

```
Query_BDD_remuneration
	Compute_Remuneration
	Compute_multi_remuneration
Update_BDD_remuneration
```

### Descriptif de l'enchaînement de ces 4 fonctions

#### Query_BDD_remuneration

Charge les données de rémunération existantes depuis la base PostgreSQL dans $script:**BDDRemun**

Utilise une clé composite (utilisateur_id, annee) pour identifier chaque enregistrement

#### Compute_Remuneration

Calcule les rémunérations annuelles pour chaque utilisateur à partir des données $script:**Review**

Agrège par catégories : **salaire de base, primes, heures supplémentaires, cotisations salariales/patronales**

Arrondit les montants à l'entier supérieur

Stocke le résultat dans $script:**REMUN** structuré par matricule/année

#### Compute_multi_remuneration

Traite les cas de profils multiples (utilisateurs ayant plusieurs profils)

Additionne au profil primaire les rémunérations des profils secondaires

Gère les cas où le profil primaire n'a pas de rémunération pour une année donnée

Consolide toutes les rémunérations dans le profil primaire d'un même utilisateur en incluant ceux de ses profils secondaires.

Cumule le résultat dans $script:**REMUN**

#### Update_BDD_remuneration

Met à jour la table rémunération en base PostgreSQL

Compare $script:**REMUN** (données calculées) avec $script:**BDDRemun** (données existantes)

Effectue les INSERT/UPDATE nécessaires via **Update_BDDTable**

## Synchronisation de la table [Performance Co]

Calculer et synchroniser en base les données d'épargne salariale (performance collaborative) consolidées de tous les utilisateurs, en gérant les profils multiples.

### Sources

$script:**BDDPerf**

$script:**USER**

$script:**PRIMAIRE**

$script:**Review**

### Code dans le script

```
Query_BDD_performance_co
	Compute_performance_co
	Compute_multi_performance_co
Update_BDD_performance_co
```

### Descriptif de l'enchaînement de ces 4 fonctions

#### Query_BDD_performance_co

Charge les données de performance collaborative existantes depuis la base PostgreSQL dans $script:**BDDPerf**

Utilise une clé composite (utilisateur_id, annee) pour identifier chaque enregistrement

#### Compute_performance_co

Calcule les performances collaboratives annuelles pour chaque utilisateur à partir des données $script:**Review**

Extrait spécifiquement la catégorie "EPARGNE SAL" (épargne salariale)

Arrondit les montants à l'entier supérieur

Stocke le résultat dans $script:**PERF** structuré par matricule/année

#### Compute_multi_performance_co

Traite les cas de profils multiples (utilisateurs ayant plusieurs contrats/profils)

Additionne au profil primaire l'épargne salariale des profils secondaires

Gère les cas où le profil primaire n'a pas de performance collaborative pour une année donnée

Consolide toutes les données d'épargne salariale dans le profil primaire d'un même utilisateur de ses différents profils

Cumule le résultat dans $script:**PERF**

#### Update_BDD_performance_co

Met à jour la table utilisateur_performance_co en base PostgreSQL

Compare $script:**PERF** (données calculées) avec $script:**BDDPerf** (données existantes)

Effectue les INSERT/UPDATE nécessaires via **Update_BDDTable**

## Synchronisation de la table [protection_sociale_utilisateur]

Calculer et synchroniser en base les données de protection sociale consolidées de tous les utilisateurs, en gérant les profils multiples.

### Sources

$script:**BDDProt**

$script:**USER**

$script:**PRIMAIRE**

$script:**Review**

### Code dans le script

```
Query_BDD_protect_sociale
	Compute_protect_sociale
	Compute_multi_protect_sociale
Update_BDD_protect_sociale
```

### Descriptif de l'enchaînement de ces 4 fonctions

#### Query_BDD_protect_sociale

Charge les données de protection sociale existantes depuis la base PostgreSQL dans $script:**BDDProt**

Utilise une clé composite (utilisateur_id, annee) pour identifier chaque enregistrement

#### Compute_protect_sociale

Calcule les données de protection sociale annuelles pour chaque utilisateur à partir des données $script:**Review**

Extrait 6 catégories spécifiques :

* PREVOYANCE SAL/PAT (prévoyance salariale/patronale)
* MUTUELLE SAL/PAT (mutuelle salariale/patronale)
* RETRAITE SAL/PAT (retraite salariale/patronale)

Arrondit les montants à l'entier supérieur.

Gère l'option **BypassProtSocialeCurrentYear** pour exclure l'année courante si configuré

Stocke le résultat dans $script:**PROT** structuré par matricule/année

#### Compute_multi_protect_sociale

Traite les cas de profils multiples (utilisateurs ayant plusieurs contrats/profils)

Additionne au profil primaire toutes les composantes de protection sociale des profils secondaires

Gère les cas où le profil primaire n'a pas de protection sociale pour une année donnée

Consolide les 6 composantes (prévoyance, mutuelle, retraite × salariale/patronale) dans le profil primaire d'un même utilisateur en incluant ses profils secondaires

#### Update_BDD_protect_sociale

Met à jour la table protection_sociale_utilisateur en base PostgreSQL

Compare $script:**PROT** (données calculées) avec $script:**BDDProt** (données existantes)

Effectue les INSERT/UPDATE nécessaires via **Update_BDDTable**

## Synchronisation de la table [conges_utilisateur]

Calculer et synchroniser en base les droits aux congés annuels de tous les utilisateurs selon différentes conventions (légale, Syntec, Fortil), avec leurs valorisations monétaires respectives, en tenant compte de l'ancienneté et de la présence effective dans l'année.

### Sources

$script:**BDDConges**

$script:**USER**

$script:**Review**

$script:**BDDRemun**

### Code dans le script

```
Query_BDD_conges_utilisateur
	Compute_conges_utilisateur
Update_BDD_conges_utilisateur
```

### Descriptif de l'enchaînement de ces 3 fonctions

#### Query_BDD_conges_utilisateur

Charge les données de congés existantes depuis la base PostgreSQL dans $script:**BDDConges**

Utilise une clé composite (utilisateur_id, annee) pour identifier chaque enregistrement

Cible la table conges_utilisateur

#### Compute_conges_utilisateur

##### Dates de reference

```
Dates de référence par année :
Date début année : 01/01/$annee (reference Fortil)
Date milieu année : 01/06/$annee (référence Syntec)

Pour la date de fin d'année: 2 cas
Année antérieure à l'année courante >> $date_fin_annee : 31/12/$annee
Année courante >>  $date_fin_annee : Today
```
##### Formules de calcul de $nb_jours_presence_dans_annee

```
$nb_jours_presence_dans_annee = :Max(0,Min(365, ($date_fin_annee - $date_entree).Days))
```
##### Formules de calcul de $jours_total

```
$jours_total = 25/365 * $nb_jours_presence_dans_annee
```  
 
##### Principe de calcul du nombre d'année pleine entre deux dates

```
Annee_pleine($base, $entree)
   $year = année_de($base) - année_de($entree)

   SI (mois_de($base) < mois_de($entree)) OU (mois_de($base) = mois_de($entree) ET jour_de($base) < jour_de($entree)) ALORS
       $year = $year - 1
   FIN SI

   résultat = Maximum(0, year)
```

##### Ancienneté

$anciennete_annee_middle = Annee_pleine $date_middle_annee $date_entree

$anciennete_annee_debut  = Annee_pleine $date_debut_annee $date_entree

##### Calcul jours Syntec

```
SI     ( $anciennete_annee_middle <  5 )   { $jours_syntec = 0 } # 0-4 ans
SINON SI ( $anciennete_annee_middle < 10 ) { $jours_syntec = 1 } # 5-9 ans  
SINON SI ( $anciennete_annee_middle < 15 ) { $jours_syntec = 2 } # 10-14 ans
SINON SI ( $anciennete_annee_middle < 20 ) { $jours_syntec = 3 } # 15-19 ans
SINON                                      { $jours_syntec = 4 } # ≥20 ans
```

##### Calcul jours Fortil

```
$jours_fortil = 0
SI pas d'exceptions "NUMTECH", "GO CONCEPT"... ALORS
    SI Années 2023-2024 :
        SI       ( $anciennete_annee_debut ≥ 4 ) { $jours_fortil = 5 } # ≥4 ans
        SINON SI ( $anciennete_annee_debut ≥ 2 ) { $jours_fortil = 4 } # ≥2 ans

    SI Années 2025+ :
        SI       ( $anciennete_annee_debut ≥ 4 ) { $jours_fortil = 8 } # ≥4 ans
        SINON SI ( $anciennete_annee_debut ≥ 2 ) { $jours_fortil = 6 } # ≥2 ans  
        SINON SI ( $anciennete_annee_debut ≥ 1 ) { $jours_fortil = 4 } # ≥1 an
```

##### Calcul monétisation
```
Calcul_Monetisation ($SB, $annee, $nbjours)
	$nbr_jours_ouvrables_moyen_par_mois = 21.667
	$currentyear = Get-Date.Year
	$currentmonth = Get-Date.Month

	SI ( $annee = $currentyear) 
		SI ( $currentmonth = 1 ) { Renvoi 0 }
		$val = (( $SB / ( $currentmonth -1 ) ) * 12) / 12 / $nbr_jours_ouvrables_moyen_par_mois * $nbjours
	SINON
		$val = $SB / 12 / $nbr_jours_ouvrables_moyen_par_mois * $nbjours
		
	Renvoi $val
}
```

#### Update_BDD_conges_utilisateur

Met à jour la table conges_utilisateur en base PostgreSQL

Compare $script:**CONGES** (données calculées) avec $script:**BDDConges** (données existantes)

Effectue les INSERT/UPDATE nécessaires via **Update_BDDTable**

## Synchronisation de la table [utilisateurs_frais]

Consolide tous les types de frais (repas, transport, formation, primes, berceau, comité d'entreprise) depuis diverses sources (Excel Fusion/Fortil, CSV famille, données Review) et les synchroniser en base de données.

### Sources

|                   Paramètres .ini                   |                    Fichiers                     |
| --------------------------------------------------- | ----------------------------------------------- |
| [**XLS_CSV_Frais**][fichierXLS_Fusion_histo]        | DRH - Frais Policy_Fusion_Group_YTD_2024xxx     |
| [**XLS_CSV_Frais**][fichierXLS_Fusion_current]      | DRH - Frais Policy_Fusion_Group_YTD_2025xxx.xls |
| [**XLS_CSV_Frais**][fichierXLS_Fortil_histo]        | DRH - Frais Policy_FORTIL GROUP_YTD_2024xxx.xls |
| [**XLS_CSV_Frais**][fichierXLS_Fortil_current]      | DRH - Frais Policy_FORTIL GROUP_YTD_2025xxx.xls |
| [**XLS_CSV_Frais**][fichierXLS_Reservation-berceau] | Politique famille - réservation berceau.xlsx    |
| [**XLS_CSV_Frais**][fichierCSV_famille]             | req21_Revue_annuelle__Absences-xxx.csv          |

$script:**USER**

$script:**ListeBM**

$script:**Review**

$script:**BDDFrais**

### Code dans le script

```
Query_XLS_fusion
Query_XLS_fortil
	Fill_Empty_Frais 
	if ( [Options][Compute_FLOTTE_AUTO] = yes ) {
		Compute_FLOTTE_AUTO
	}
Query_CSV_prime_naissance
	Compute_prime_naissance
Query_XLS_Reservation_berceau
	Compute_Reservation_berceau
Query_BDD_Utilisateurs_frais
Update_BDD_Utilisateurs_frais

```

### Descriptif de l'enchaînement de ces 10 fonctions

#### Query_XLS_fusion

Charge les données de frais Fusion depuis 2 fichiers Excel (historique + courant) dans $script:**FRAIS**. Traite les catégories repas/transport.

#### Query_XLS_fortil

Charge les données de frais Fortil depuis 2 fichiers Excel (historique + courant) et les cumule avec les données Fusion. 

#### Fill_Empty_Frais

Complète $script:**FRAIS** avec tous les utilisateurs de ListeBM n'ayant pas de frais depuis 2019.

Initialise à 0 (repas, transport, formation, prime_naissance) et calcule le comité d'entreprise (0 ou 30€).

#### Compute_FLOTTE_AUTO (Fonction conditionnelle)

**Cette fonction n'est exécutée que si l'option `[Options][Compute_FLOTTE_AUTO] = yes` dans le fichier .ini**

Soustrait les montants "FLOTTE AUTO" des données Review du champ transport dans $script:**FRAIS**. 

Corrige les frais de transport en déduisant les avantages flotte automobile pour éviter une double comptabilisation.

#### Query_CSV_prime_naissance

Charge le fichier CSV famille contenant les dates de naissance des enfants des employés dans $script:**CSVFam**.

#### Compute_prime_naissance

Calcule les primes de naissance basées sur le PMSS annuel (montant × pourcentage). Gère les naissances multiples dans la même année et met à jour $script:**FRAIS**.

#### Query_XLS_Reservation_berceau

Charge les données de réservation berceau depuis un fichier Excel : matricule, dates début/fin, prix mensuel dans $script:**XLSBerceau**.

#### Compute_Reservation_berceau

Calcule les frais de réservation berceau par année selon les périodes d'occupation. Nettoie les prix (supprime €, espaces) et cumule dans $script:**FRAIS**.

#### calcul_berceau

##### Définition des bornes temporelles

* Début d'année : 1er janvier de l'année concernée
* Fin d'année : 31 décembre de l'année concernée
* Cas particulier année courante : La fin est limitée au dernier jour du mois précédent (pas le mois en cours)

##### Calcul de la période effective

Début de période : Le plus tardif entre la date d'entrée et le début d'année

Fin de période : Le plus précoce entre la date de sortie et la fin d'année

Si la période calculée est invalide (début > fin), retourne 0

##### Comptage des mois complets La fonction parcourt mois par mois depuis le début de période :

Règle importante : Un mois n'est comptabilisé que s'il est entièrement couvert par la période

Si la période se termine avant la fin d'un mois, ce mois n'est pas comptabilisé

Chaque mois complet ajoute le prix mensuel au montant total

##### Résultat

Retourne le montant total (nombre de mois complets × prix mensuel)

#### Query_BDD_Utilisateurs_frais

Charge les données de frais existantes depuis la table PostgreSQL utilisateur_frais dans $script:**BDDFrais** avec clé composite (utilisateur_id, annee).

#### Update_BDD_Utilisateurs_frais

Synchronise $script:**FRAIS** (calculé) avec $script:**BDDFrais** (existant) en base PostgreSQL. Effectue les INSERT/UPDATE nécessaires.

#### comite_entreprise

30€ à partir de l'année 2024, seulement pour les entités non exclues.

fichier .ini

```
[Exclude]
comite_entreprise_30_euros = FORTIL SUD EST,FORTIL SUD OUEST,FORTIL NORD OUEST,FORTIL LYON,FORTIL PARIS,GO CONCEPT,ALBERT & CO
```
0€ sinon.

## Synchronisation de la table [entrepreneuriat]

Gérer la valorisation des participations entrepreneuriales des employés (dividendes)

**Décalage temporel**

Applique un décalage d'une année : les dividendes déclarés pour l'année N dans le CSV sont comptabilisés pour l'année N-1 dans la base. Cela correspond à une logique comptable où les dividendes de l'année N sont basés sur les résultats de l'année N-1.

### Sources

|            Paramètres .ini            |                  Fichiers                  |
| ------------------------------------- | ------------------------------------------ |
| [**CSV_entrepreneuriat**][fichierCSV] | import_valorisation_entrepreneuriatxxx.csv |

$script:**USER**

### Code dans le script

```
Query_CSV_val_entrepreneuriat
	Compute_val_entrepreneuriat
Query_BDD_val_entrepreneuriat
Update_BDD_val_entrepreneuriat
```

### Descriptif de l'enchaînement de ces 4 fonctions

#### Query_CSV_val_entrepreneuriat

Lit le fichier CSV configuré dans ["CSV_entrepreneuriat"]["fichierCSV"]

Utilise un en-tête fixe : "**utilisateur_id;annee;cashout_dividendes**"

Stocke les données dans $script:**CSVEntrep**

#### Compute_val_entrepreneuriat

Parcourt chaque ligne du CSV chargé ($script:**CSVEntrep**)

Validation utilisateur : Vérifie que le matricule existe dans $script:**BDDuser**

**Décalage temporel **: Applique les dividendes à l'année N-1 (si CSV contient 2024, les dividendes sont affectés à 2023)

Conversion sécurisée : nettoie et valide les montants

Arrondit à l'entier supérieur

Stocke dans $script:**VALOR **avec la structure : **[matricule][année-1]["cashout_dividendes"]**

Exclusions : Ajoute les matricules inexistants dans $script:**EXCLUS** avec motif

#### Query_BDD_val_entrepreneuriat

Interroge la table valorisation_entrepreneuriat configurée

Utilise une clé composite (utilisateur_id, annee) pour identifier les enregistrements

Stocke les données existantes dans $script:**BDDValor**

#### Update_BDD_val_entrepreneuriat

Compare $script:**VALOR** (nouvelles données) avec $script:**BDDValor** (données existantes)

Utilise la clé composite (**utilisateur_id, annee**) pour identifier les changements

Effectue les opérations INSERT / UPDATE pour les enregistrements modifiés

## Synchronisation de la table [Projets-liste]

Maintenir un référentiel unifié des projets (réels + fictifs) pour permettre l'affectation et le suivi des temps, incluant les périodes d'absence et d'intercontrat.

Permet de tracer les temps non-projet (**congés, absences, intercontrat, etc...**)

Projets réels : Issus du fichier Excel OLAP (IDs normaux)

Projets fictifs : Définis en configuration (IDs > 1000000)

Normalisation des statuts : Conversion des codes techniques (PRJ_ACTIVE/CLOSED) vers des libellés métier (En cours/Terminé)

### Sources

|           Paramètres .ini           |          Fichiers          |
| ----------------------------------- | -------------------------- |
| [**XLS_Projets-liste**][fichierCSV] | OLAP-Projets_liste_xxx.xls |


### Code dans le script
```
Query_XLS_Projets_liste
	Compute_Projets_liste
	Compute_Fake_Projets
Query_BDD_projets
Update_BDD_projets
```

### Descriptif de l'enchaînement de ces 5 fonctions

#### Query_XLS_Projets_liste

Lit le fichier Excel configuré dans ["XLS_Projets-liste"]["fichierXLS"]

Extrait 4 colonnes spécifiques : id (col 4), nom (col 5), client (col 12), statut (col 3)

Utilise une requête SQL sur la feuille Sheet0$, et stocke les données dans $script:**XLSProj**

#### Compute_Projets_liste

Parcourt chaque ligne du fichier Excel chargé

Validation ID : Vérifie que l'ID projet est numérique, sinon exclusion

Détection doublons : Génère une erreur si un ID projet existe déjà

Normalisation statuts :

"**PRJ_ACTIVE**" → "En cours"

"**PRJ_CLOSED**" → "Terminé"

Autres valeurs → "Unknown" + erreur

Structure finale : [**id, nom, client, statut, description**]

Initialise le champ description à vide

Stocke dans $script:**PROJ**

#### Compute_Fake_Projets

Lit la configuration [Projets_factice] du fichier INI

Ajoute des projets avec des IDs > 1000000 (ex: 1000001 = INTERCONTRAT)

Types de projets fictifs : *INTERCONTRAT, ABS_JOURFERIE, ABS_LOA, CM, CONGE_MALADIE_CAN, CP, CSS, CSS_ITALIA, exceptionnalleave, PERMESSI_ROL, etc.*

Tous les projets fictifs ont le statut "**En cours**"

#### Query_BDD_projets

Interroge la table projets configurée

Utilise l'ID comme clé unique

Stocke les données existantes dans $script:**BDDProj**


#### Update_BDD_projets

Compare $script:**PROJ** (projets) avec $script:**BDDProj** (existants)

Utilise l'ID comme clé unique pour identifier les changements

Effectue les INSERT / UPDATE pour les projets modifiés

## Synchronisation de la table [Projets_affectation]

Charge les données d'affectation des projets depuis un fichier Excel, et affecte ces projets aux users.

### Sources

|              Paramètres .ini               |              Fichiers               |
| ------------------------------------------ | ----------------------------------- |
| ["XLS_Projets-affectation"]["fichierXLS"]  | Fichiers\OLAP-affectations_xxx.xlsx |

$script:**ListeBM**

$script:**BDDProj**

### Code dans le script
```
Query_XLS_Projets_affectation
	Compute_Projets_affectation
Query_BDD_affectations
Update_BDD_affectations
```

### Descriptif de l'enchaînement de ces 4 fonctions

#### Query_XLS_Projets_affectation

Lit le fichier Excel configuré dans $script:cfg["XLS_Projets-affectation"]["fichierXLS"]

Extrait les colonnes **[Matricule], [Date de la FDT (Année)], [Tâche], [Quantité]** de la feuille Sheet0

Place les données dans la variable $script:**XLSAffect**

#### Compute_Projets_affectation

Parcourt chaque ligne de $script:**XLSAffect** pour créer une structure hiérarchique dans $script:**AFFECT**

Vérifie que le matricule est numérique

Contrôle que le matricule existe dans $script:**ListeBM**

Valide que le projet_id existe dans $script:**BDDProj**

Transformation : Remplace les projets factices par leurs vrais IDs via Replace_Fake_affectation

Structure créée : $script:**AFFECT**[utilisateur_id][projet_id][annee]** contenant :
utilisateur_id, projet_id, annee, jours_passes (formaté en décimal)

Gestion d'erreurs : Exclut les matricules/projets invalides avec des messages d'erreur détaillés

#### Query_BDD_affectations

Cette fonction charge les affectations existantes depuis la base de données PostgreSQL :

Table configurée dans $script:cfg["SQL_Postgre_Review"]["tableutilisateur_projet"]

Clés composites : Utilise une clé composite ("utilisateur_id","projet_id","annee")

Place les données dans $script:**BDDAffect**

#### Update_BDD_affectations

Compare $script:**AFFECT** (données Excel traitées) avec $script:**BDDAffect** (données BDD)

Effectue les INSERT / UPDATE pour les données modifiés

## Synchronisation de la table [historiques_salaires]

Historique des salaires avec une clé composite (utilisateur + date), permettant de suivre l'évolution des salaires dans le temps pour chaque employé.

### Sources

|              Paramètres .ini               |                    Fichiers                    |
| ------------------------------------------ | ---------------------------------------------- |
| ["CSV_historiques_salaires"]["fichierCSV"] | 1 - Revue annuelle - Salaires contractuels.csv |

$script:**BDDuser**

$script:**BDDHisto**

### Code dans le script
```
Query_CSV_historiques_salaires
	Compute_historiques_salaires
Query_BDD_historiques_salaires
Update_BDD_historiques_salaires
```

### Descriptif de l'enchaînement de ces 4 fonctions

#### Query_CSV_historiques_salaires

Lit le fichier CSV configuré dans $script:cfg["CSV_historiques_salaires"]["fichierCSV"]

Place les données dans $script:**CSVHisto** avec une clé "_index"

#### Compute_historiques_salaires

Parcourt chaque ligne de $script:**CSVHisto** pour créer une structure hiérarchique

Validation : Vérifie que le matricule existe dans $script:**BDDuser** (table des utilisateurs)

Structure créée : $script:**HISTO**[matricule][date] contenant :

* utilisateur_id : le matricule sans zéros de tête
* date : la date d'effet formatée
* montant : le salaire annuel de la colonne "SALAIRE ANNUEL"

Exclut les matricules non trouvés dans la table utilisateurs

#### Query_BDD_historiques_salaires

Cette fonction charge l'historique des salaires existant depuis la base de données PostgreSQL dans $script:cfg["SQL_Postgre_Review"]["tablehistoriques_salaires"]

Clés composites : Utilise une clé composite ("utilisateur_id","date")

Place les données dans $script:**BDDHisto**

#### Update_BDD_historiques_salaires

Compare $script:**HISTO** avec $script:**BDDHisto** (données BDD)

Effectue les INSERT / UPDATE pour les données modifiés

# Architecture

## Structure générale
```
AnnualReview.ps1
├── Configuration (LoadIni)
├── Modules externes
├── Traitement des données
├── Mise à jour base de données
└── Génération des rapports
```

## Dépendances

### Modules externes

Les modules sont chargés au démarrage du script dans l'ordre suivant :

1. **`Ini.ps1`** - Gestion des fichiers de configuration INI
2. **`Log.ps1`** - Système de logging (LOG, ERR, WRN, MOD, DLT, INA, DBG, QUIT)
3. **`Encode.ps1`** - Encodage des données et gestion UTF-8
4. **`Csv.ps1`** - Traitement des fichiers CSV
5. **`XLSX.ps1`** - Traitement des fichiers Excel (XLS/XLSX)
6. **`StrConvert.ps1`** - Conversion et manipulation de chaînes
7. **`SendEmail.ps1`** - Envoi d'emails (SMTP ou Microsoft Graph)

> **Note :** La documentation détaillée de chaque module sera fournie séparément.

### Module de base de données (chargement conditionnel)

Le module de gestion PostgreSQL est chargé selon le paramètre `[start][TransacSQL]` du fichier .ini :

- **`PostgreSQL - TransactionAllInOne.ps1`** - Si `TransacSQL = AllInOne` (toutes les transactions en une seule)
- **`PostgreSQL - TransactionOneByOne.ps1`** - Si `TransacSQL = OneByOne` (transactions individuelles)

### Assemblys .NET

- **`System.Web`** - Fonctionnalités web .NET
- **Drivers PostgreSQL** :
  - `Npgsql.dll` - Driver PostgreSQL pour .NET
  - `Microsoft.Extensions.Logging.Abstractions.dll` - Abstractions de logging Microsoft

## Flux d'exécution principal

Le script s'exécute selon le flux suivant :

### 1. Initialisation
```powershell
$script:cfgFile = "$PSScriptRoot\AnnualReview.ini"

# Chargement des modules de base
. "$PSScriptRoot\Modules\Ini.ps1"
. "$PSScriptRoot\Modules\Log.ps1"
. "$PSScriptRoot\Modules\Encode.ps1"
. "$PSScriptRoot\Modules\Csv.ps1"
. "$PSScriptRoot\Modules\XLSX.ps1"
. "$PSScriptRoot\Modules\StrConvert.ps1"
. "$PSScriptRoot\Modules\SendEmail.ps1"

LoadIni
SetConsoleToUFT8

# Chargement des assemblys .NET
Add-Type -AssemblyName System.Web
Add-Type -Path $script:cfg["SQL_Postgre_Review"]["microsoftExt"]
Add-Type -Path $script:cfg["SQL_Postgre_Review"]["pathdll"]

# Chargement du module PostgreSQL (conditionnel)
if ($script:cfg["start"]["TransacSQL"] -eq "AllInOne") {
    . "$PSScriptRoot\Modules\PostgreSQL - TransactionAllInOne.ps1"
} else {
    . "$PSScriptRoot\Modules\PostgreSQL - TransactionOneByOne.ps1"
}
```

### 2. Traitement des données sources
```
Query_CSV_Salaries
    Compute_USER

Query_XLS_Salaries-complement
    Compute_ListeBM

Query_CSV_Salaires-primes
    Compute_Profil_secondaire
```

### 3. Synchronisation des tables
```
# Utilisateurs
Query_BDD_Utilisateurs
    Compute_Managers
    Compute_Administrator
    Prepare_two_pass
Update_BDD_utilisateurs

# Rémunération
Query_BDD_remuneration
    Compute_Remuneration
    Compute_multi_remuneration
Update_BDD_remuneration

# Performance Co
Query_BDD_performance_co
    Compute_performance_co
    Compute_multi_performance_co
Update_BDD_performance_co

# Protection sociale
Query_BDD_protect_sociale
    Compute_protect_sociale
    Compute_multi_protect_sociale
Update_BDD_protect_sociale

# Congés
Query_BDD_conges_utilisateur
    Compute_conges_utilisateur
Update_BDD_conges_utilisateur

# Frais
Query_XLS_fusion
Query_XLS_fortil
    Fill_Empty_Frais
    if ([Options][Compute_FLOTTE_AUTO] = yes) {
        Compute_FLOTTE_AUTO
    }
Query_CSV_prime_naissance
    Compute_prime_naissance
Query_XLS_Reservation_berceau
    Compute_Reservation_berceau
Query_BDD_Utilisateurs_frais
Update_BDD_Utilisateurs_frais

# Entrepreneuriat
Query_CSV_val_entrepreneuriat
    Compute_val_entrepreneuriat
Query_BDD_val_entrepreneuriat
Update_BDD_val_entrepreneuriat

# Projets
Query_XLS_Projets_liste
    Compute_Projets_liste
    Compute_Fake_Projets
Query_BDD_projets
Update_BDD_projets

# Affectations
Query_XLS_Projets_affectation
    Compute_Projets_affectation
Query_BDD_affectations
Update_BDD_affectations

# Historiques salaires
Query_CSV_historiques_salaires
    Compute_historiques_salaires
Query_BDD_historiques_salaires
Update_BDD_historiques_salaires
```

### 4. Finalisation
```
Log_Deltas
QUIT "Main" "Fin du process"
```

## Fonction LoadIni

La fonction **LoadIni** est appelée au démarrage du script pour initialiser la configuration et les variables globales.

### Responsabilités

1. **Initialisation des variables de logging**
   - `$script:pathfilelog` - Chemin du fichier de log principal
   - `$script:pathfileerr` - Chemin du fichier d'erreurs
   - `$script:pathfileina` - Chemin du fichier des inactifs
   - `$script:pathfiledlt` - Chemin du fichier des deltas
   - `$script:pathfilemod` - Chemin du fichier des modifications

2. **Chargement du fichier .ini**
   - Vérifie l'existence du fichier `AnnualReview.ini`
   - Initialise la structure de base de `$script:cfg` avec les sections principales
   - Charge toutes les sections du fichier .ini via `Add-IniFiles` (y compris `Exclude`, `Include Only`, `Options`, `PMSS`, etc.)

3. **Résolution des chemins de fichiers**
   - Résout tous les chemins de fichiers sources (CSV, XLS, XLSX)
   - Résout les chemins des DLL PostgreSQL et Microsoft
   - Vérifie l'existence des fichiers obligatoires (paramètre `-Needed`)
   - Remplace `$rootpath$` par le chemin réel du script

4. **Gestion des fichiers de log**
   - Supprime les fichiers de log "One Shot" (log, ina, dlt) s'ils existent
   - Crée les nouveaux fichiers de log "One Shot"
   - Crée les fichiers de log cumulatifs (err, mod) s'ils n'existent pas

5. **Initialisation des variables de filtrage**
   - `$script:ValidEntities` - Liste des entités valides pour le comité d'entreprise (depuis `[Exclude][comite_entreprise_30_euros]`)
   - `$script:ValidFrais` - Liste des types de frais à inclure (depuis `[Include Only][Frais]`)
   - `$script:ValidTransport` - Liste des catégories de transport valides (depuis `[Include Only][Categorie_transport]`)

6. **Initialisation des variables de contrôle**
   - `$script:execok = $false` - Indicateur d'exécution réussie
   - `$script:start` - Chronomètre pour mesurer le temps d'exécution
   - `$script:MailErr = $false` - Indicateur d'erreur pour l'envoi d'email
   - `$script:WARNING = 0` - Compteur de warnings
   - `$script:ERREUR = 0` - Compteur d'erreurs
   - `$script:emailtxt` - Liste des messages pour l'email de rapport

### Code dans le script
```powershell
LoadIni
```

## Fonction Log_Deltas

La fonction **Log_Deltas** est appelée en fin d'exécution pour générer un rapport de correspondance entre les différentes hashtables principales.

### Responsabilités

Cette fonction effectue trois analyses de correspondance et enregistre les résultats dans le fichier de log DLT (deltas) :

1. **Correspondance USER >> ListeBM**
   - Parcourt tous les matricules de `$script:USER`
   - Vérifie leur présence dans `$script:ListeBM`
   - Pour les matricules absents, affiche la raison d'exclusion depuis `$script:EXCLUS` si disponible
   - Affiche le nombre de matricules communs et le nombre d'absents

2. **Correspondance ListeBM >> Review**
   - Parcourt tous les matricules de `$script:ListeBM`
   - Vérifie leur présence dans `$script:Review`
   - Pour les matricules absents, affiche :
     * La raison d'exclusion depuis `$script:EXCLUS` si disponible
     * Sinon, la date d'entrée et la société depuis `$script:USER` et `$script:BDDuser`
   - Affiche le nombre de matricules communs et le nombre d'absents

3. **Correspondance Review >> ListeBM**
   - Parcourt tous les matricules de `$script:Review`
   - Vérifie leur présence dans `$script:ListeBM`
   - Pour les matricules absents, affiche :
     * La raison d'exclusion depuis `$script:EXCLUS` si disponible
     * Sinon, la date d'entrée depuis `$script:USER`
   - Affiche le nombre de matricules communs et le nombre d'absents

### Utilité

Cette fonction permet de :
* Identifier les écarts entre les différentes sources de données
* Comprendre pourquoi certains matricules sont présents dans une source mais pas dans une autre
* Valider la cohérence des exclusions appliquées
* Faciliter le débogage et l'audit des données traitées

### Code dans le script
```powershell
Log_Deltas
```

# Configuration

## Fichier de configuration
Le script utilise un fichier `AnnualReview.ini` avec les sections suivantes :

```
# -----------------------------------------------------------------------------------------------------------------------------
#    AnnualReview.ini - Necessite Powershell 7 ou +
#      Ce script met à jour la base AnnualReview avec les données de :
#  		fichierCSV  Admin.csv
#  		fichierXLS  Progessi - Liste BM_*.xls
#  		fichierCSV  1 - Revue annuelle - Primes et charges 2018 -2024.csv
#  		fichierCSV  req23_Revue_annuelle_primes_charges-*.csv
#  		fichierXLS  DRH - Frais Policy_FORTIL GROUP_YTD_*.xls
#  		fichierXLS  DRH - Frais Policy_Fusion_Group_YTD_*.xls
#  		fichierCSV  req_21_Revue_annuelle__Absences-*.csv
#  		fichierCSV  import_valorisation_entrepreneuriat*.csv
#  		fichierXLS  OLAP-Projets_liste_*.xls
#  		fichierXLS  OLAP-affectations_*.xlsx
# -----------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------
#     Parametrage du comportement de l'interface AnnualReview.ps1
# -------------------------------------------------------------------

[start]
# Le parametre "ApplyUpdate" yes/no : permet de simuler sans modifier la base AnnualReview si ApplyUpdate = no
ApplyUpdate = yes

# Le parametre BypassProtSocialeCurrentYear permet de ne pas remplir la table protection_sociale_utilisateur pour l'annee courante
BypassProtSocialeCurrentYear = yes

# TransacSQL : OneByOne ou AllInOne
TransacSQL = AllInOne

# Le parametre "logtoscreen" contrôle l'affichage de toutes les infos de log/error/warning dans la console
logtoscreen = yes

# Le parametre "debug" contrôle l'affichage des infos de debug dans la console
debug       = no

# Le parametre "warntoerr" permet d'inclure ou pas les warnings dans le fichier SynchroAD.err
warntoerr   = yes

# Le parametre "Desactive_login" modifie les adresse emails xxx.xxx@xxx.xxx en xxx.xxx1@xxx.xxx, sauf ceux de la liste [users_desactive_exception]
Desactive_login = yes

# -------------------------------------------------------------------
#     Chemin des fichiers de LOGS
# -------------------------------------------------------------------
[intf]
name        = Synchronisation AnnualReview

# Chemin du fichier log : 
pathfilelog = $rootpath$\logs\AnnualReview_One_Shot.log

# Chemin du fichier log des exceptions inactivés
pathfileina = $rootpath$\logs\AnnualReview_One_Shot.ina

# Chemin du fichier log des deltas entre fichiers
pathfiledlt = $rootpath$\logs\AnnualReview_One_Shot.dlt

# Chemin du fichier logs d'erreur
pathfileerr = $rootpath$\logs\AnnualReview_Cumul.err

# Chemin du fichier des logs modifications
pathfilemod = $rootpath$\logs\AnnualReview_Cumul.mod

# -------------------------------------------------------------------
#     Parametrage Fichiers ADMIN, ListeBM, Review, Frais, Valorisation
# -------------------------------------------------------------------
[CSV_Salaries]
fichierCSV      = $rootpath$\Fichiers\req22_Revue_annuelle_Admin-*.csv
HEADERstartline = 4

[XLS_Salaries-complement]
fichierXLS      = $rootpath$\Fichiers\Progessi - Liste BM_*.xls

[CSV_Salaires-primes]
fichierCSV_Histo      = $rootpath$\Fichiers\1 - Revue annuelle - Primes et charges 2018 -2024.csv
DATAstartline_Histo   = 7
fichierCSV_Current    = $rootpath$\Fichiers\req23_Revue_annuelle_primes_charges-*.csv
DATAstartline_Current = 43

[XLS_CSV_Frais]
fichierXLS_Fortil_histo        = $rootpath$\Fichiers\DRH - Frais Policy_FORTIL GROUP_YTD_2024*.xls
fichierXLS_Fortil_current      = $rootpath$\Fichiers\DRH - Frais Policy_FORTIL GROUP_YTD_2025*.xls
fichierXLS_Fusion_histo        = $rootpath$\Fichiers\DRH - Frais Policy_Fusion_Group_YTD_2024*.xls
fichierXLS_Fusion_current      = $rootpath$\Fichiers\DRH - Frais Policy_Fusion_Group_YTD_2025*.xls
fichierXLS_Reservation-berceau = $rootpath$\Fichiers\Politique famille - réservation berceau.xlsx
fichierCSV_famille             = $rootpath$\Fichiers\req21_Revue_annuelle__Absences-*.csv
HEADERstartline_famille = 2

# PPMS Année, Pourcentage (pour frais, politique famille)
[PMSS]
#2026 = xxxx, 6
2025 = 3925, 6
2024 = 3864, 6
2023 = 3666, 6
2022 = 3428, 6
2021 = 3428, 6
2020 = 3428, 6
2019 = 3377, 6

[CSV_entrepreneuriat]
fichierCSV      = $rootpath$\Fichiers\import_valorisation_entrepreneuriat*.csv

[XLS_Projets-liste]
fichierXLS      = $rootpath$\Fichiers\OLAP-Projets_liste_*.xls

[XLS_Projets-affectation]
fichierXLS      = $rootpath$\Fichiers\OLAP-affectations_*.xlsx

[CSV_historiques_salaires]
fichierCSV      = $rootpath$\Fichiers\1 - Revue annuelle - Salaires contractuels.csv
HEADERstartline = 4

# -------------------------------------------------------------------
#     Parametrage des exceptions
# -------------------------------------------------------------------
[Exclude]
zero_conge_jours_fortil    = NUMTECH,GO CONCEPT,ALBERT & CO
comite_entreprise_30_euros = FORTIL SUD EST,FORTIL SUD OUEST,FORTIL NORD OUEST,FORTIL LYON,FORTIL PARIS,GO CONCEPT,ALBERT & CO
denomination_sociale       = FORTIL BELGIUM,ALLIANCE,FORTIL GROUP Support VIE
contrat                    = Stagiaire,Apprendistato
Classe horaire             = TEMPS PARTIEL 1H

[Include Only]
Interne/externe            = Profil interne
statusId                   = PARTY_ENABLED
Categorie_transport        = 36. Abonnement Transports en commun
Frais                      = Forfait repas midi,Forfait repas soir

[Options]
Compute_FLOTTE_AUTO        = no

[users_ADMIN]
# adm1, adm2, adm3, etc...
adm1 = sandrine.mattio@fortil.group
adm2 = laura.battesti@fortil.group

[users_desactive_exception]
# ex1, ex2, ex3, etc...
ex1 = sandrine.mattio@fortil.group
ex2 = laura.battesti@fortil.group
ex3 = pierre-hugo.agrain@fortil.group
ex4 = christophe.hamon@fortil.group

[Projets_factice]
1000001 = INTERCONTRAT
1000002 = ABS_JOURFERIE
1000003 = ABS_LOA
1000004 = CM
1000005 = CONGE_MALADIE_CAN
1000006 = CP
1000007 = CSS
1000008 = CSS_ITALIA
1000009 = exceptionnalleave
1000010 = PERMESSI_ROL
1000011 = RTT
1000012 = RTT_ITALIA


# -------------------------------------------------------------------
#     Parametrage du serveur Postgree pour AnnualReview
# -------------------------------------------------------------------

# Parametre de connection à la base AnnualReview
[SQL_Postgre_Review]                                                                       
# frmtdateIN  = Format de date reçu depuis la BDD (dépend de la config regionale/Date du serveur de script)
# frmtdateOUT = Utiliser le format ISO yyyy-MM-dd HH:mm:ss, standard valable quelquesoit le parametrage de la BDD
frmtdateIN   = dd/MM/yyyy HH:mm:ss
frmtdateOUT  = yyyy-MM-dd HH:mm:ss
pathdll      = $rootpath$\Extensions\Postgre\net6.0\Npgsql.dll
microsoftExt = $rootpath$\Extensions\Microsoft\logging.abstractions\netstandard2.0\Microsoft.Extensions.Logging.Abstractions.dll

# Serveur Postgres LOCAL FR
server       = 192.168.1.47
# Serveur Postgres LINUX US
#server       = 192.168.1.87
database     = Review
login        = xxxxx
password     = xxxxxxxxxxxxxx

tableUtilisateurs                   = utilisateurs
tableRemuneration                   = remuneration
tableutilisateur_performance_co     = utilisateur_performance_co
tableprotection_sociale_utilisateur = protection_sociale_utilisateur
tableconges_utilisateur             = conges_utilisateur
tableutilisateur_frais              = utilisateur_frais
tablevalorisation_entrepreneuriat   = valorisation_entrepreneuriat
tableprojets                        = projets
tableutilisateur_projet             = utilisateur_projet
tablehistoriques_salaires           = historiques_salaires

# -------------------------------------------------------------------
#     Parametrage des Emails
# -------------------------------------------------------------------

# Parametre pour l'envoi de mails (Protocoles possible : Microsoft.Graph ou SMTP)
# Le parametre "emailmode" permet de choisir le mode d'emission d'un mail (GRAPH ou SMTP)
# Envoi de mail si sendemail = "yes" / "no"
[email]
sendemail    = no
destinataire = xxxxx.xxxxx@xxxxx.xxx
Subject      = Synchro AD BERTIN
emailmode    = SMTP
UseSSL       = false

# Login pour SMTP
expediteur   = xxxxx.xxxxx@xxxxx.xxx
server       = smtp.gmail.com
port         = 
password     = 

```
# LOG

**Switches disponibles :**

- `-DBG` : Mode debug
- `-LOG` : Log standard
- `-ERR` : Erreur
- `-WRN` : Avertissement
- `-MOD` : Modification
- `-DLT` : Suppression
- `-INA` : Inactif
- `-EMAIL` : Inclure dans l'email
- `-CRLF` : Saut de ligne
- `-NOSCREEN` : Pas d'affichage écran

## Fonctions spécialisées de LOGS

| Fonction | Description | Couleur | Destination |
|----------|-------------|---------|-------------|
| `DBG()` | Messages de débogage | Gris | LOG (si debug=yes) |
| `LOG()` | Messages informatifs | Cyan | LOG + Écran |
| `ERR()` | Messages d'erreur | Rouge | ERR + LOG + Email |
| `WRN()` | Avertissements | Magenta | LOG + Email |
| `MOD()` | Modifications | Jaune | MOD + LOG + Email |
| `DLT()` | Deltas | Gris | DLT (pas d'écran) |
| `INA()` | Matricules inactifs | Gris | INA (pas d'écran) |

## Fonctions de fin de script

- `QUIT()` : Fin normale avec statistiques
- `QUITEX()` : Fin avec erreur et stack trace

## Fichiers de log générés

1. **pathfilelog** : Log général
2. **pathfileerr** : Erreurs uniquement
3. **pathfilemod** : Modifications
4. **pathfiledlt** : Deltas
5. **pathfileina** : Matricules inactifs

# Maintenance

## Ajout de nouvelles sources/traitements/tables

1. Créer la fonction `Query_XXX`
2. Ajouter la fonction `Compute_XXX`
3. Créer les fonctions `QUERY/UPDATE_BDD_XXX`
4. Intégrer dans le flux principal (main)

## Personnalisation du logging

- Utiliser les fonctions `LOG()`, `ERR()`, `WRN()`, etc.
- Ajouter des switches spécifiques si nécessaire

## Évolution de la configuration

- Ajouter des sections dans le fichier INI
- Mettre à jour la fonction `LoadIni()`
- Utiliser `GetFilePath()` pour les nouveaux chemins

# Dépannage

## Problèmes courants

1. **Fichier de configuration manquant** : Vérifier `AnnualReview.ini`
2. **Erreurs de connexion BDD** : Contrôler la configuration PostgreSQL
3. **Fichiers sources introuvables** : Vérifier les chemins et wildcards
4. **Erreurs d'encodage** : Vérifier la configuration de culture

## Logs de débogage

- Activer `debug = yes` dans la configuration
- Consulter le fichier de log général
- Examiner les fichiers de log spécialisés
