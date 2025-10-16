# AnnualReview.ps1

function Log_Deltas {
	DLT "Log_Deltas" "Correspondance des matricules USER >> ListeBM" -CRLF
	$cptadmin = 0
	$cptlistebm = 0
	foreach ( $zeromatricule in $script:USER.keys ) {
		$cptadmin++
		$matricule = $zeromatricule  -replace '^0+', ''
		if ( -not ($script:ListeBM.ContainsKey($matricule)) ) {
			if ( $script:EXCLUS.Contains($matricule) ) { 
				DLT "Log_Deltas" "Matricule USER [$zeromatricule] absent dans ListeBM : $($script:EXCLUS[$matricule])"
			} else {
				DLT "Log_Deltas" "Matricule USER [$zeromatricule] absent dans ListeBM"
			}
		} else {
			$cptlistebm++
		}
	}
	DLT "Log_Deltas" "$cptlistebm (ListeBM) /$cptadmin (USER) Matricules sont communs : $($cptadmin-$cptlistebm) sont absent de ListeBM"

	DLT "Log_Deltas" "Correspondance des matricules ListeBM >> Review" -CRLF
	$cptlistebm = 0
	$cptReview = 0
	foreach ( $matricule in $script:ListeBM.keys ) {
		$zeromatricule = $matricule.PadLeft(8, '0')
		$cptlistebm++
		if ( ($script:BDDuser.ContainsKey($matricule)) ) {
			$societe = $script:BDDuser[$matricule]["entite"]
		} else {
			DLT "Log_Deltas" "Matricule ListeBM [$zeromatricule] absent dans BDDuser"
			continue
		}
		if ( ($script:USER.ContainsKey($zeromatricule)) ) {
			$dateentree = $script:USER[$zeromatricule]["Date d'entrée"]
		} else {
			DLT "Log_Deltas" "Matricule ListeBM [$zeromatricule] absent dans USER"
			continue
		}
		if ( -not ($script:Review.ContainsKey($zeromatricule)) ) {
			if ( $script:EXCLUS.Contains($matricule) ) { 
				DLT "Log_Deltas" "Matricule ListeBM [$zeromatricule] absent dans Review : $($script:EXCLUS[$matricule])  $societe"
			} else {
				DLT "Log_Deltas" "Matricule ListeBM [$zeromatricule] absent dans Review (Date entrée = $($dateentree))  $societe"
			}
		} else {
			$cptReview++
		}
	}
	DLT "Log_Deltas" "$cptReview (Review) /$cptlistebm (ListeBM) Matricules sont communs : $($cptlistebm - $cptReview) sont absent de ListeBM"


	DLT "Log_Deltas" "Correspondance des matricules Review >> ListeBM" -CRLF
	$cptlistebm = 0
	$cptReview = 0
	foreach ( $zeromatricule in $script:Review.keys ) {
		$matricule = $zeromatricule  -replace '^0+', ''
		$cptReview++
		if ( -not ($script:ListeBM.ContainsKey($matricule)) ) {
			
			if ( $script:EXCLUS.Contains($matricule) ) { 
				DLT "Log_Deltas" "Matricule Review [$zeromatricule] absent dans ListeBM : $($script:EXCLUS[$matricule])"
			} else {
				
				DLT "Log_Deltas" "Matricule Review [$zeromatricule] absent dans ListeBM : Date entrée = $($script:USER[$zeromatricule]["Date d'entrée"])"
			}
		} else {
			$cptlistebm++
		}
	}
	DLT "Log_Deltas" "$cptlistebm (ListeBM) /$cptReview (Review) Matricules sont communs : $($cptReview - $cptlistebm) sont absent de ListeBM"

}

# --------------------------------------------------------
#               Chargement fichier .ini
# --------------------------------------------------------

function LoadIni {
	# initialisation variables liste des logs
	$script:pathfilelog = @()
	$script:pathfileerr = @()
	$script:pathfileina = @()
	$script:pathfiledlt = @()
	$script:pathfilemod = @()

	# definition des sections du fichier .ini
	$script:cfg = @{
        "start"                   = @{}
        "intf"                    = @{}
        "CSV_Salaries"            = @{}
        "XLS_Salaries-complement" = @{}
        "CSV_Salaires-primes"     = @{}
        "SQL_Review"              = @{}
        "email"                   = @{}
    }
    # Recuperation des parametres passes au script 
    $script:execok  = $false

    if (-not(Test-Path $($script:cfgFile) -PathType Leaf)) { Write-Host "Fichier de parametrage $script:cfgFile innexistant"; exit 1 }
    Write-Host "Fichier de parametrage $script:cfgFile"

    # Initialisation des sections parametres.
    $script:start    = [System.Diagnostics.Stopwatch]::startNew()
    $script:MailErr  = $false
    $script:WARNING  = 0
    $script:ERREUR   = 0
	
	$script:emailtxt = New-Object 'System.Collections.Generic.List[string]'

	$script:cfg = Add-IniFiles $script:cfg $script:cfgFile

	# Recherche des chemins de tous les fichiers et verification de leur existence
	if (-not ($script:cfg["intf"].ContainsKey("rootpath")) ) {
		$script:cfg["intf"]["rootpath"]                = $PSScriptRoot
	}
	$script:cfg["intf"]["pathfilelog"]                 				= GetFilePath $script:cfg["intf"]["pathfilelog"]
	$script:cfg["intf"]["pathfileina"]                 				= GetFilePath $script:cfg["intf"]["pathfileina"]
	$script:cfg["intf"]["pathfiledlt"]                 				= GetFilePath $script:cfg["intf"]["pathfiledlt"]
	$script:cfg["intf"]["pathfileerr"]                 				= GetFilePath $script:cfg["intf"]["pathfileerr"]
	$script:cfg["intf"]["pathfilemod"]                 				= GetFilePath $script:cfg["intf"]["pathfilemod"]

	$script:cfg["sql_postgre_review"]["microsoftext"]  				= GetFilePath $script:cfg["sql_postgre_review"]["microsoftext"] -Needed
	$script:cfg["sql_postgre_review"]["pathdll"]       				= GetFilePath $script:cfg["sql_postgre_review"]["pathdll"] -Needed

	$script:cfg["csv_salaries"]["fichiercsv"]             			= GetFilePath $script:cfg["csv_salaries"]["fichiercsv"] -Needed
	$script:cfg["xls_salaries-complement"]["fichierxls"]        	= GetFilePath $script:cfg["xls_salaries-complement"]["fichierxls"] -Needed
	$script:cfg["csv_salaires-primes"]["fichiercsv_histo"]      	= GetFilePath $script:cfg["csv_salaires-primes"]["fichiercsv_histo"] -Needed
	$script:cfg["csv_salaires-primes"]["fichiercsv_current"]		= GetFilePath $script:cfg["csv_salaires-primes"]["fichiercsv_current"] -Needed

	$script:cfg["xls_csv_frais"]["fichierxls_fortil_histo"]  		= GetFilePath $script:cfg["xls_csv_frais"]["fichierxls_fortil_histo"] -Needed
	$script:cfg["xls_csv_frais"]["fichierxls_fortil_current"]  		= GetFilePath $script:cfg["xls_csv_frais"]["fichierxls_fortil_current"] -Needed
	$script:cfg["xls_csv_frais"]["fichierxls_fusion_histo"] 		= GetFilePath $script:cfg["xls_csv_frais"]["fichierxls_fusion_histo"] -Needed
	$script:cfg["xls_csv_frais"]["fichierxls_fusion_current"] 		= GetFilePath $script:cfg["xls_csv_frais"]["fichierxls_fusion_current"] -Needed
	$script:cfg["xls_csv_frais"]["fichiercsv_famille"] 				= GetFilePath $script:cfg["xls_csv_frais"]["fichiercsv_famille"] -Needed
	$script:cfg['xls_csv_frais']['fichierxls_reservation-berceau'] 	= GetFilePath $script:cfg['xls_csv_frais']['fichierxls_reservation-berceau'] -Needed

	$script:cfg["csv_entrepreneuriat"]["fichiercsv"]      			= GetFilePath $script:cfg["csv_entrepreneuriat"]["fichiercsv"] -Needed
	$script:cfg["xls_projets-liste"]["fichierxls"]         			= GetFilePath $script:cfg["xls_projets-liste"]["fichierxls"] -Needed
	$script:cfg["xls_projets-affectation"]["fichierxls"]   			= GetFilePath $script:cfg["xls_projets-affectation"]["fichierxls"] -Needed
	$script:cfg["xlsx_historiques_salaires"]["fichierxlsx"] 			= GetFilePath $script:cfg["xlsx_historiques_salaires"]["fichierxlsx"] -Needed

	# Suppression des fichiers One_Shot et création des fichiers innexistants
	if ((Test-Path $($script:cfg["intf"]["pathfilelog"]) -PathType Leaf)) { Remove-Item -Path $script:cfg["intf"]["pathfilelog"]}    
	if ((Test-Path $($script:cfg["intf"]["pathfileina"]) -PathType Leaf)) { Remove-Item -Path $script:cfg["intf"]["pathfileina"]} 
	if ((Test-Path $($script:cfg["intf"]["pathfiledlt"]) -PathType Leaf)) { Remove-Item -Path $script:cfg["intf"]["pathfiledlt"]} 

	$null = New-Item -type file $($script:cfg["intf"]["pathfilelog"]) -Force;
	$null = New-Item -type file $($script:cfg["intf"]["pathfileina"]) -Force;
	$null = New-Item -type file $($script:cfg["intf"]["pathfiledlt"]) -Force;
	if (-not(Test-Path $($script:cfg["intf"]["pathfileerr"]) -PathType Leaf)) { $null = New-Item -type file $($script:cfg["intf"]["pathfileerr"]) -Force; }
	if (-not(Test-Path $($script:cfg["intf"]["pathfilemod"]) -PathType Leaf)) { $null = New-Item -type file $($script:cfg["intf"]["pathfilemod"]) -Force; }

    # Définir la variable $ValidEntities
    $script:ValidEntities = $script:cfg["Exclude"]["comite_entreprise_30_euros"].Split(',') | ForEach-Object { $_.Trim() }
    # Définir la variable $ValidFrais
    $script:ValidFrais = $script:cfg["Include Only"]["Frais"].Split(',') | ForEach-Object { $_.Trim() }
    # Définir la variable $ValidTransport
    $script:ValidTransport = $script:cfg["Include Only"]["Categorie_transport"].Split(',') | ForEach-Object { $_.Trim() }
}

# --------------------------------------------------------
#   Fichier ["CSV_Salaries"]["fichierCSV"]					[req22_Revue_annuelle_Admin-*.csv]
#	$script:USER    = @{}
#	$script:USERAll = @{}
#	$script:EXCLUS   = @{}
# --------------------------------------------------------

function Query_CSV_Salaries {
	$csvfile         = $script:cfg["csv_salaries"]["fichiercsv"]
	$headerstartline = $script:cfg["csv_salaries"]["headerstartline"] - 1
	$datecol 		 = @("Date Naissance","Date d'entrée","D Entrée groupe","Date de sortie dans la société")

	LOG "Query_CSV_Salaries" "Chargement du fichier $csvfile"
	$script:USER    = @{}
	$script:USER    = Invoke-CSVQuery -csvfile $csvfile -key "Matricule" -separator "," -row $headerstartline -frmtdateOUT $script:cfg["sql_postgre_review"]["frmtdateout"] -datecol $datecol
	$script:USERAll = $script:USER.Clone()
}
function Compute_USER {
	$script:EXCLUS = @{}
	$exclu 		   = 0
	$lim           = [datetime](Get-Date).Date
	$format        = $script:cfg["sql_postgre_review"]["frmtdateout"]

	foreach ($zeromatricule in @($script:USER.Keys)) { 
		$matricule = $zeromatricule  -replace '^0+', ''
		$de        = $script:USER[$zeromatricule]["Date d'entrée"]
		$ds        = $script:USER[$zeromatricule]["Date de sortie dans la société"]
		$obs       = $script:USER[$zeromatricule]["InfosG Observations"]

		# InfosG Observations" = PROFIL SECONDAIRE
		if ( $obs -eq "PROFIL SECONDAIRE" ) {
			Add-Exclusion $matricule "CSV_Salaries" "InfosG Observations = $obs" -USER
			$script:USER.Remove($zeromatricule)
			$exclu++
			continue
		}
		
		# "date d'entrée" > Today
		if ( -not ([string]::IsNullOrWhiteSpace($de)) ) {
			$dateResult = ConvertTo-SafeDate -dateString $de -format $format -matricule $zeromatricule -fieldName "date d'entrée" -functionName "QueryCSVUSER"
			if (-not $dateResult.Success) {
				Add-Exclusion $matricule "CSV_Salaries" "Erreur parsing date d'entrée = $de" -USER
				$script:USER.Remove($zeromatricule)
				$exclu++
				continue
			}
			if ($dateResult.Date -gt $lim) {
				Add-Exclusion $matricule "CSV_Salaries" "Date d'entrée = $de" -USER
				$script:USER.Remove($zeromatricule)
				$exclu++
				continue
			}
		}

		# "date de sortie" < Today
		if ( -not ([string]::IsNullOrWhiteSpace($ds)) ) {
			$dateResult = ConvertTo-SafeDate -dateString $ds -format $format -matricule $zeromatricule -fieldName "date de sortie" -functionName "QueryCSVUSER"
			if (-not $dateResult.Success) {
				Add-Exclusion $matricule "CSV_Salaries" "Erreur parsing date de sortie = $ds" -USER
				$script:USER.Remove($zeromatricule)
				$exclu++
				continue
			}
			if ($dateResult.Date -lt $lim) {
				Add-Exclusion $matricule "CSV_Salaries" "Date de sortie dans la société = $ds" -USER
				$script:USER.Remove($zeromatricule)
				$exclu++
				continue
			}
		}
	}
	LOG "Compute_USER" "$($script:USER.Count) Matricules valides, $($exclu) Matricules exclues du fichier $csvfile"
}

# --------------------------------------------------------
#   Fichier ["XLS_Salaries-complement"]["fichierXLS"]		[Progessi - Liste BM_*.xls]
#	$script:ListeBM     = @{}
#	$script:ListeBMAll  = @{}
#	$script:ListeBMo365 = @{}
# --------------------------------------------------------

function Query_XLS_Salaries-complement {
    $xlsfile  =  $script:cfg["XLS_Salaries-complement"]["fichierXLS"]
	$sqlquery = "SELECT * FROM [Sheet0$]"
	$header   = "Dénomination sociale,statusId,Matricule,Prénom_act,Nom_act,UPN O365_act,Département,Prénom_manager,Nom_manager,UPN O365_manager,Date d'entrée,Date de sortie,Emploi,Interne/externe,Courriel,TA source,TA finale,Etablissement secondaire,Date d'activation,Date de désactivation,Dénomination sociale2,Id Acteur,Classe horaire de présence,ID profil STT source,Statut de validation,Date de création,Commentaires,Matricule INT ou STT,Statut,Adresse n°1,Adresse n°2,Ville,Code postal,Réf. géographique de pays,Type de contrat,Dernier diplôme obtenu,Métier résolu,Effectif"
	$datecol  = @("Date d'entrée","Date de sortie","Date d'activation","Date de désactivation")
	$headers  = $header -split ","

    LOG "Query_XLS_Salaries-complement" "Chargement du fichier Excel : $xlsfile" -CRLF
    $result = Invoke-ExcelQuery -filePath $xlsfile -sqlQuery $sqlquery -functionName "Query_XLS_Salaries-complement" -header $headers -frmtdateOUT $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -datecol $datecol -dateLocale "FR"
    $script:table_XLS_Salaries_complement = $result.Table
	LOG "Query_XLS_Salaries-complement" "Nombre de lignes chargées : $($script:table_XLS_Salaries_complement.Rows.Count)"
}
function Compute_ListeBM {
	$script:ListeBM     = @{}
	$script:ListeBMAll  = @{}
	$script:ListeBMo365 = @{}

	# Charger une hashtable
	$cptvalidmatricule = 0
	$cptvalidupn       = 0
	$cptexclu          = 0
	$cptnomatricule    = 0
	$cptnoupn          = 0
	$cptdoublonmat     = 0
	$cptdoublonupn     = 0

	$ExcludeContrat = $script:cfg["Exclude"]["contrat"].Split(',')              | ForEach-Object { $_.Trim() }
	$ExcludeSocial  = $script:cfg["Exclude"]["denomination_sociale"].Split(',') | ForEach-Object { $_.Trim() }
	$ExcludeHoraire = $script:cfg["Exclude"]["Classe horaire"].Split(',')       | ForEach-Object { $_.Trim() }
	$IncludeStatut  = $script:cfg["Include Only"]["statusId"].Split(',')        | ForEach-Object { $_.Trim() }
	$IncludeProfil  = $script:cfg["Include Only"]["Interne/externe"].Split(',') | ForEach-Object { $_.Trim() }

	$keycol = "Matricule"
    foreach ($row in $script:table_XLS_Salaries_complement.Rows) {
		[string]$matricule = $row[$keycol]
		$zeromatricule = $matricule.PadLeft(8, '0')

		$nom     = $row["nom_act"];
		$prenom  = $row["prénom_act"];
		$party   = $row["statusId"];
		$profil  = $row["Interne/externe"];
		$o365    = $row["UPN O365_act"];
		$horaire = $row["Classe horaire de présence"];
		$contrat = $row["Type de contrat"];
		$social  = $row["Dénomination sociale"];

		if ( [string]::IsNullOrWhiteSpace($matricule) ) { 
			$OKmatricule = $false
			$OKadmin = $false
		} else {
			$OKmatricule = $true
			if ( ($script:USER.ContainsKey($zeromatricule)) ) {
				$OKadmin = $true
			} else {
				$OKadmin = $false
			}
		}

		$OKparty     = $false
		$OKprofil    = $false
		$OKhoraire   = $false
		$OK          = $false

		if ( $zeromatricule -ne "00000000" ) {
			if (-not $script:ListeBMAll.ContainsKey($matricule)) {
				$script:ListeBMAll[$matricule] = @{}
			} else {
				WRN "Compute_ListeBM" "Matricule : [$zeromatricule] Doublon dans ListeBM"
				$cptdoublonmat++
			}
			# Choix de recopier les données du doublon si le matricule est présent dans listeBM
			foreach ($col in $table_XLS_Salaries_complement.Columns) {
				$script:ListeBMAll[$matricule][$col.ColumnName] = $row[$col.ColumnName]
			}
		}

		# Traitement des exclusions
		if ( $OKadmin ) { 
			$OK = $true

			# Exclusion si Date d'entrée > Today
			$dateentree = ConvertTo-SafeDate -dateString $script:USER[$zeromatricule]["Date d'entrée"] -format $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -matricule $zeromatricule -fieldName "Date d'entrée" -functionName "Query_XLS_Salaries-complement"

			# Exclusion des contrats invalides
			foreach ( $exclude in $ExcludeContrat ) {
				if ( $contrat.startswith($exclude) ) {
					$OK = $false;
					Add-Exclusion $matricule "XLS_Salaries-complement" "Exclu pour Type de contrat = $exclude   $nom $prenom" -USER
				}
			}
			# Exclusion des Dénomination sociale invalides
			foreach ( $exclude in $ExcludeSocial ) {
				if ( $social.startswith($exclude) ) {
					$OK = $false;
					Add-Exclusion $matricule "XLS_Salaries-complement" "Exclu pour Dénomination sociale = $exclude   $nom $prenom" -USER
				}
			}

			# Exclusion des classe horaire de présence invalides
			foreach ( $exclude in $ExcludeHoraire ) {
				if ( $horaire.startswith($exclude) ) {
					$OK = $false; 
					Add-Exclusion $matricule "XLS_Salaries-complement" "Exclu pour Classe horaire de présence = $exclude   $nom $prenom" -USER
				} else {
					$OKhoraire = $true
				}
			}

			# INCLUSION STRICTE

			# Inclusion stricte sur des statusId valide
			foreach ( $include in $IncludeStatut ) {
				if ( $party.startswith($include) ) {
					$OKparty = $true
					break
				}
			}
			if ( -not ($OKparty) ) {
				$OK = $false; 
				Add-Exclusion $matricule "XLS_Salaries-complement" "Exclu pour statusId = $party   $nom $prenom" -USER
			}

			# Inclusion stricte sur des Interne/externe valide
			$OKprofil = $false
			foreach ( $include in $IncludeProfil) {
				if ( $profil.startswith($include) ) {
					$OKprofil = $true
					break
				}
			}
			if ( -not ($OKprofil) ) {
				$OK = $false; 
				Add-Exclusion $matricule "XLS_Salaries-complement" "Exclu pour Interne/externe = $profil   $nom $prenom" -USER
			}

		}

		# Exclusion si UPN O365 non present
		if ( $OKparty -and $OKprofil -and $OKhoraire -and $OKadmin ) {  # -and $OKactif
			if ( [string]::IsNullOrWhiteSpace($o365) -or $o365 -eq "-" ) {
				$OK = $false
				ERR "Compute_ListeBM" "Matricule : [$zeromatricule] Pas de UPN O365 : $nom $prenom"
				$cptnoupn++
			}
		}

		if ( -not ($OK) ) { 
			if ( -not ($OKadmin) ) {
				if ( -not ($OKmatricule) ) {
					Add-Exclusion $matricule "XLS_Salaries-complement" "$nom $prenom" -USER
					$cptnomatricule++
				} else {
					Add-Exclusion $matricule "XLS_Salaries-complement" "présent dans Liste_BM.xls n'existe pas dans les non exclus de USER.csv $nom $prenom" -USER
				}
			} else {
				Add-Exclusion $matricule "XLS_Salaries-complement" "statusId = $party   Interne/externe = $profil    Date entree = $($dateentree.Date)   Classe horaire de présence = $horaire  UPN O365 = $o365  ($nom $prenom)" -USER
			}
			
			$cptexclu++			
			continue 
		}

		# Crée la clé Matricule si valid
        if (-not $script:ListeBM.ContainsKey($matricule)) {
            $script:ListeBM[$matricule] = @{}
			$cptvalidmatricule++
        } else {
			ERR "Compute_ListeBM" "Doublon sur Matricule : [$zeromatricule] : $nom $prenom"
			continue
		}
        foreach ($col in $table_XLS_Salaries_complement.Columns) {
            $script:ListeBM[$matricule][$col.ColumnName] = $row[$col.ColumnName]
        }

		# Crée la clé ListeBMo365[$o365] si o365 valid
		if ( (IsEmail $o365 $zeromatricule)  ) {
			if (-not $script:ListeBMo365.ContainsKey($o365)) {
				$script:ListeBMo365[$o365] = @{}
				$cptvalidupn++
			} else {
				ERR "Fill_ListeBM_without_exclued" "Doublon sur UPN O365 (acteur) : $o365 : $nom $prenom"
				$cptdoublonupn++
				continue
			}
			foreach ($col in $table_XLS_Salaries_complement.Columns) {
				$script:ListeBMo365[$o365][$col.ColumnName] = "$($row[$col.ColumnName])"
			}
		} else {
			ERR "Fill_ListeBM_without_exclued" "Matricule [$zeromatricule] UPN O365 [$o365] non valide : $nom $prenom"	
		}
	}  

	foreach ($zeromatricule in $script:USER.Keys) { 
		$matricule = $zeromatricule  -replace '^0+', ''
		if ( -not ($script:ListeBM.ContainsKey($matricule)) ) {
			Add-Exclusion $matricule "XLS_Salaries-complement" "présent dans USER, sans correspondance de Matricule dans les non exclus de Liste_BM" -USER
		}
	}

	$c = $cptvalidmatricule - $cptvalidupn
	LOG "Compute_ListeBM" "Valid(Matricule) : $cptvalidmatricule   Valid(UPN) : $cptvalidupn   Exclu : $cptexclu   Sans Matricule : $cptnomatricule   Sans UPN O365 : $cptnoupn   Doublon Matricule : $cptdoublonmat   Doublon UPN O365 : $cptdoublonupn   Avec Matricule, mais sans UPN O365 : $c"
	LOG "Compute_ListeBM" "Total : $($script:ListeBM.Count) Matricules valides, $($script:ListeBMAll.Count) dans All"
}

# --------------------------------------------------------
#    Fichier ["CSV_Salaires-primes"]["fichierCSV_Histo"]	[1 - Revue annuelle - Primes et charges 2018 -2024.csv]
#    Fichier ["CSV_Salaires-primes"]["fichierCSV_Current"]	[req23_Revue_annuelle_primes_charges-*.csv            ]
#    $script:Review    = @{}    
#    $script:ReviewAll = @{}
#    $script:PRIMAIRE  = @{}
# --------------------------------------------------------

function Query_CSV_Salaires-primes {
    $script:Review = @{}    
    $script:ReviewAll = @{}    
    $header = "Année,Matricule,Code,Cat,Cat RUB,Montant,Montant rectifié"

	$csvfile       = $script:cfg["CSV_Salaires-primes"]["fichierCSV_Histo"]
	$datastartline = $script:cfg["CSV_Salaires-primes"]["DATAstartline_Histo"] - 1
	Query_CSV_Salaires-primes_by_file $csvfile $datastartline $header

	$csvfile       = $script:cfg["CSV_Salaires-primes"]["fichierCSV_Current"]
	$datastartline = $script:cfg["CSV_Salaires-primes"]["DATAstartline_Current"] - 1
	Query_CSV_Salaires-primes_by_file $csvfile $datastartline $header
}
function Query_CSV_Salaires-primes_by_file {
	Param ( $csvfile, $datastartline ,$header)

    LOG "Query_CSV_Salaires-primes" "Chargement du fichier $csvfile" -CRLF
	$PC = @{}
	$PC = Invoke-CSVQuery -csvfile $csvfile -key "_index" -separator "," -row $datastartline -header $header -frmtdateOUT $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"]

    LOG "Query_CSV_Salaires-primes" "Verification des Matricules"

	$ex = @()
	$exclu   = 0

    foreach ($key in $PC.Keys) {
        $zeromatricule = $PC[$key]['Matricule']
		$matricule = $zeromatricule  -replace '^0+', ''

		$annee     = $PC[$key]['Année']
		$CatRub    = $PC[$key]['Cat RUB']
		$montantStr = $PC[$key]['Montant rectifié']
		$amountResult = ConvertTo-SafeAmount -amountString $montantStr -matricule $zeromatricule -annee $annee -catRub $CatRub -functionName "Query_CSV_Salaires-primes"
		$montant = $amountResult.Amount

		# ------------------------
		# remplissage de ReviewAll
		# ------------------------
		if (-not $script:ReviewAll.ContainsKey($zeromatricule)) {
			$script:ReviewAll[$zeromatricule] = @{}
		}
		if (-not $script:ReviewAll[$zeromatricule].ContainsKey($annee)) {
			$script:ReviewAll[$zeromatricule][$annee] = @{}
		}
		# Stocke la valeur du champ (Somme pour les champs identiques)
		if ( [string]::IsNullOrWhiteSpace($CatRub) ) {
			DBG "Query_CSV_Salaires-primes" "Matricule [$zeromatricule] [$annee] : Colonne [Cat RUB] vide pour un montant de $montant"
		} else {
			if ( $script:ReviewAll[$zeromatricule][$annee].ContainsKey($CatRub) ) {
				$script:ReviewAll[$zeromatricule][$annee][$CatRub] += $montant
			} else {
				$script:ReviewAll[$zeromatricule][$annee][$CatRub] = $montant
			}
		}

		# ------------------------
		# remplissage de Review
		# ------------------------
		if ( $script:USER.ContainsKey($zeromatricule) ) {
			
			# Vérifie si le matricule existe déjà
			if (-not $script:Review.ContainsKey($zeromatricule)) {
				$script:Review[$zeromatricule] = @{}
			}

			# Vérifie si l’année existe pour ce matricule
			if (-not $script:Review[$zeromatricule].ContainsKey($annee)) {
				$script:Review[$zeromatricule][$annee] = @{}
			}

			# Stocke la valeur du champ (Somme pour les champs identiques)
			if ( [string]::IsNullOrWhiteSpace($CatRub) ) {
				DBG "Query_CSV_Salaires-primes" "Matricule [$zeromatricule] [$annee] : Colonne [Cat RUB] vide pour un montant de $montant"
			} else {
				if ( $script:Review[$zeromatricule][$annee].ContainsKey($CatRub) ) {
					$script:Review[$zeromatricule][$annee][$CatRub] += $montant
				} else {
					$script:Review[$zeromatricule][$annee][$CatRub] = $montant
				}
			}
		} else {
			# Ne compter qu'une seule fois le matricule
			if ( -not ($ex.Contains($matricule)) ) {
				$ex += $matricule
				$exclu++
				Add-Exclusion $matricule "CSV_Salaires-primes" "non trouvé dans les Matricules USER valide" -USER
			}
		}
    }
	LOG "Query_CSV_Salaires-primes" "Nombre de matricules valide distinct : $($script:Review.Count) ($exclu Matricules ont été exclus du fichier $csvfile)"
}
function Compute_Profil_secondaire {
	# $script:PRIMAIRE[Primaire][Secondaire]
	$script:PRIMAIRE = @{}

	LOG "Compute_Profil_secondaire" "Analyse des profils multiples" -CRLF
	
	# Créer un index inverse pour optimiser les recherches : EmailBase -> [liste des matricules]
	$emailBaseIndex = @{}
	foreach ($mat in $script:ListeBMAll.keys) {
		$upn = $script:ListeBMAll[$mat]['UPN O365_act']
		if (-not [string]::IsNullOrWhiteSpace($upn)) {
			# Extraire la base de l'email (avant un chiffre ou @)
			if ($upn -match '^([^0-9@]+)') {
				$emailBase = $matches[1]
				if (-not $emailBaseIndex.ContainsKey($emailBase)) {
					$emailBaseIndex[$emailBase] = @()
				}
				$emailBaseIndex[$emailBase] += $mat
			}
		}
	}

	# Traiter les profils primaires
	foreach ($zeromatricule in $script:USER.keys) {
		$matricule = $zeromatricule -replace '^0+', ''
		$obs = $script:USER[$zeromatricule]["InfosG Observations"]

		if ($obs -like "*PRIMAIRE*") { 
			# Vérifier que le matricule existe dans ListeBM
			if (-not $script:ListeBM.ContainsKey($matricule)) {
				ERR "Compute_Profil_secondaire" "Matricule [$zeromatricule] absent de ListeBM. Impossible de déterminer son Email."
				continue
			}

			$emailUPN = $script:ListeBM[$matricule]['UPN O365_act']
			$emailBase = ""
			
			# Extraire la base de l'email
			if ($emailUPN -match '^([^0-9@]+)') {
				$emailBase = $matches[1]
			}

			if ([string]::IsNullOrWhiteSpace($emailBase)) {
				ERR "Compute_Profil_secondaire" "Matricule [$zeromatricule] : Email Base vide"
				continue
			}

			# Initialiser le profil primaire
			$script:PRIMAIRE[$matricule] = @{}
			LOG "Compute_Profil_secondaire" "[$emailBase] -> Profil primaire [$zeromatricule]"

			# Rechercher les profils secondaires associés à cette base d'email
			if ($emailBaseIndex.ContainsKey($emailBase)) {
				foreach ($mat in $emailBaseIndex[$emailBase]) {
					# Exclure le matricule primaire lui-même
					if ($mat -ne $matricule) {
						$zeromat = $mat.PadLeft(8, '0')

						# Vérifications de présence dans USERAll et ReviewAll
						if (-not $script:USERAll.ContainsKey($zeromat)) {
							ERR "Compute_Profil_secondaire" "[$emailBase] -> Profil secondaire [$zeromat] inexistant dans USERAll"
							continue
						}

						if (-not $script:ReviewAll.ContainsKey($zeromat)) {
							ERR "Compute_Profil_secondaire" "[$emailBase] -> Profil secondaire [$zeromat] inexistant dans Review"
							continue
						}

						# Ajouter le profil secondaire
						$script:PRIMAIRE[$matricule][$mat] = "Secondaire"
						LOG "Compute_Profil_secondaire" "[$emailBase] -> Profil secondaire [$zeromat]"
					}
				}
			}
		}
	}
}

# --------------------------------------------------------
#    Table utilisateurs
#	 $script:BDDuser = @{}
# --------------------------------------------------------

function Query_BDD_utilisateurs {
	$script:BDDuser = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableUtilisateurs"] -functionName "QueryBDDutilisateurs" -keyColumns @("id") -targetVariable $script:BDDuser -UseFrmtDateOUT
}
function Compute_Managers {
	foreach ($matricule in $script:ListeBM.Keys) {
		# Collaborateur, par defaut
		$script:ListeBM[$matricule]['role'] = "collaborateur"
		$script:ListeBM[$matricule]['id_manager'] = ""
	}

	foreach ($matricule in $script:ListeBM.Keys) {
		$zeromatricule = $matricule.PadLeft(8, '0')
		# pour chaque utilisateur, on cherche le manager
		$upnact = $script:ListeBM[$matricule]['UPN O365_act']
		$upnmanager = $script:ListeBM[$matricule]['UPN O365_manager']

		# si manager
		if ( IsNotNullPGS $upnmanager ) {
			if ($script:ListeBMo365.ContainsKey($upnmanager) -and $script:ListeBMo365[$upnmanager].ContainsKey('Matricule')) {
				# Rechercher le matricule du manager
				[string]$matriculemanager = $script:ListeBMo365[$upnmanager]['Matricule']
				$script:ListeBM[$matricule]['UPN O365_manager'] = $upnmanager

				# Affecter le rôle Manager au compte du manager
				$script:ListeBMo365[$upnmanager]['role'] = "manager"
				$script:ListeBM[$matriculemanager]['role']= "manager"

				# Affecter id_manager au compte du user
				if ( $script:ListeBMo365.ContainsKey($upnact) ) {  
					$script:ListeBMo365[$upnact]['id_manager'] = $matriculemanager
				} else {
					ERR "Compute-Managers" "Matricule [$zeromatricule] UPN O365 [$upnact] non résolu dans ListeBM."
				}
				
				$script:ListeBM[$matricule]['id_manager'] = $matriculemanager
			} else {
				ERR "Compute-Managers" "Matricule [$zeromatricule] UPN O365 [$upnact] a pour manager [$upnmanager] non résolu dans la liste des utilisateurs non exclus."
			}
		}
	}
}
function Compute_Administrator {
	# Traitement des exceptions "admin" (.ini, section [users_ADMIN])
	foreach ($lst in $script:cfg["users_ADMIN"].Keys) {
		$upnadmin = $script:cfg["users_ADMIN"][$lst]
		if ( $script:ListeBMo365.ContainsKey( $upnadmin ) ) {
			$matricule_admin = $script:ListeBMo365[$upnadmin]['Matricule']
			$script:ListeBM[$matricule_admin]['role'] = "admin"
			$script:ListeBMo365[$upnadmin]['role'] = "admin"
		} else {
			ERR "Compute_Administrator" "Exception USER [$lst] UPN O365 [$upnadmin] non résolu dans ListeBM."
		}
	}

	# Traitement des exceptions 1@ (.ini, section [users_desactive_exception])
	$script:ex = @()
	foreach ($lst in $script:cfg['users_desactive_exception'].Keys) {
		$upnadmin = $script:cfg["users_desactive_exception"][$lst]
		if ( $script:ListeBMo365.ContainsKey( $upnadmin ) ) {
			$script:ex += $script:cfg['users_desactive_exception'][$lst]
		} else {
			ERR "Compute_Administrator" "Exception xxxx.xxxx@xxx.xxx [$lst] UPN O365 non résolu dans ListeBM."
		}
	}
}
function Prepare_two_pass {
	$script:pass1 = @{}
	$script:pass2 = @{}
	
	foreach ($matricule in $script:ListeBM.Keys) {
		$zeromatricule = $matricule.PadLeft(8, '0')
		if ( $script:USER.ContainsKey($zeromatricule)) {
			$script:pass1[$matricule] = @{}
			$script:pass1[$matricule]["id"]  		= $matricule  
			$script:pass1[$matricule]["entite"]  	= $script:USER[$zeromatricule]["L Etablissement St. Jur."]  
			$script:pass1[$matricule]["checkin"]  	= $script:USER[$zeromatricule]["Date d'entrée"]
			$script:pass1[$matricule]["poste"]      = $script:USER[$zeromatricule]["Emploi"]
			$script:pass1[$matricule]["prenom"]     = $script:ListeBM[$matricule]["Prénom_act"]
			$script:pass1[$matricule]["nom"]        = $script:ListeBM[$matricule]["Nom_act"]
			if ( $script:cfg["start"]["Desactive_login"] -eq "yes" -and -not ( $script:ListeBM[$matricule]["UPN O365_act"] -in $script:ex ) ) {
				$script:pass1[$matricule]["email"] = $script:ListeBM[$matricule]["UPN O365_act"].Replace("@", "9@")
			} else {
				$script:pass1[$matricule]["email"] = $script:ListeBM[$matricule]["UPN O365_act"]
			}
			$script:pass1[$matricule]["role"]		= $script:ListeBM[$matricule]["role"]

			$script:pass2[$matricule] = @{}
			$script:pass2[$matricule]["id"]  		= $matricule  
			$script:pass2[$matricule]["manager_id"]	= $script:ListeBM[$matricule]["id_manager"]
		}
	}
}
function Update_BDD_utilisateurs {
	$table = $script:cfg["SQL_Postgre_Review"]["tableUtilisateurs"]
	$keycolumns  = @("id")	

	# Pass 1
	LOG "Update_BDD_utilisateurs" "Passe 1 : Modif table [$table] (Tous attributs, sauf [manager_id])" -CRLF
	Update_BDDTable $script:pass1 $script:BDDuser $keycolumns $table "Update_BDD_utilisateurs" { Query_BDD_utilisateurs }

	# Pass 2
	LOG "Update_BDD_utilisateurs" "Passe 2 : Modif table [$table] ([manager_id] seulement)" -CRLF
	Update_BDDTable $script:pass2 $script:BDDuser $keycolumns $table "Update_BDD_utilisateurs" { Query_BDD_utilisateurs }
}

# --------------------------------------------------------
#   Table remuneration
#	$script:REMUN    = @{}
#	$script:BDDRemun = @{}
# --------------------------------------------------------

function Query_BDD_remuneration {
	$script:BDDRemun = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableRemuneration"] -functionName "Query_BDD_remuneration" -keyColumns @("utilisateur_id","annee") -targetVariable $script:BDDRemun
}
function Compute_Remuneration {
	LOG "Compute_Remuneration" "Calcul des remunerations par annee"	
	$first = $true
	$script:REMUN = @{}
    foreach ($zeromatricule in $script:Review.Keys) {
		$matricule = $zeromatricule  -replace '^0+', ''

		# Verifier si l'utilisateur existe dans la BDD utilisateur 
		if (-not (Test-UserExists $matricule $zeromatricule "Update_BDD_remuneration")) {
			continue
		}
		$script:REMUN[$matricule] = @{}
		foreach ($annee in $script:Review[$zeromatricule].Keys) {
			$script:REMUN[$matricule][$annee] = @{}
			$salaire_base       = 0
			$primes             = 0
			$heures_supp        = 0
			$cot_salariales     = 0
			$cot_patronale      = 0
			$script:REMUN[$matricule][$annee]["utilisateur_id"]     = $matricule
			$script:REMUN[$matricule][$annee]["annee"]              = $annee
			foreach ($CatRub in $script:Review[$zeromatricule][$annee].Keys) {
				$montant =  [Math]::Ceiling($script:Review[$zeromatricule][$annee][$CatRub])
				if ( $CatRub -eq "SALAIRE DE BASE" )   { 
					$salaire_base   += $montant }
				if ( $CatRub -eq "PRIMES" )            { $primes         += $montant }
				if ( $CatRub -eq "HS" )                { $heures_supp    += $montant }
				if ( $CatRub -eq "TOTAL CHARGES SAL" ) { $cot_salariales += $montant }
				if ( $CatRub -eq "TOTAL CHARGES PAT" ) { $cot_patronale  += $montant }
			}
			$script:REMUN[$matricule][$annee]["salaire_base"]      = $salaire_base
			$script:REMUN[$matricule][$annee]["primes"]            = $primes
			$script:REMUN[$matricule][$annee]["heures_supp"]       = $heures_supp
			$script:REMUN[$matricule][$annee]["cot_salariales"]    = $cot_salariales
			$script:REMUN[$matricule][$annee]["cot_patronale"]     = $cot_patronale
			# total_remuneration est calculé par un trigger en BDD
		}
    }
}
function Compute_multi_remuneration {
	LOG "Compute_multi_remuneration" "Additionne au profil primaire les remuneration des profils secondaires"
	foreach ($primaire in $script:PRIMAIRE.keys) {
		$zeroprimaire = $primaire.PadLeft(8, '0')
		foreach ($secondaire in $script:PRIMAIRE[$primaire].keys) {
			$zerosecondaire = $secondaire.PadLeft(8, '0')

			# Pour chaque année, ajouter les rémunerations des profils secondaires 
			foreach ($annee in $script:ReviewAll[$zerosecondaire].Keys) {
				$salaire_base       = 0
				$primes             = 0
				$heures_supp        = 0
				$cot_salariales     = 0
				$cot_patronale      = 0
				foreach ($CatRub in $script:ReviewAll[$zerosecondaire][$annee].Keys) {
					$montant =  [Math]::Ceiling($script:ReviewAll[$zerosecondaire][$annee][$CatRub])
					if ( $CatRub -eq "SALAIRE DE BASE" )   { $salaire_base   += $montant }
					if ( $CatRub -eq "PRIMES" )            { $primes         += $montant }
					if ( $CatRub -eq "HS" )                { $heures_supp    += $montant }
					if ( $CatRub -eq "TOTAL CHARGES SAL" ) { $cot_salariales += $montant }
					if ( $CatRub -eq "TOTAL CHARGES PAT" ) { $cot_patronale  += $montant }
				}

				if ( -not ($script:REMUN.ContainsKey($primaire)) ) {
					$script:REMUN[$primaire] = @{}
				}

				# gere l'absence de remuneration pour une année dans le profil primaire
				if ( -not ($script:REMUN[$primaire].ContainsKey($annee)) ) {
					$script:REMUN[$primaire][$annee] = @{}
					$script:REMUN[$primaire][$annee]["utilisateur_id"]    = $primaire
					$script:REMUN[$primaire][$annee]["annee"]             = $annee
					$script:REMUN[$primaire][$annee]["salaire_base"]      = $salaire_base
					$script:REMUN[$primaire][$annee]["primes"]            = $primes
					$script:REMUN[$primaire][$annee]["heures_supp"]       = $heures_supp
					$script:REMUN[$primaire][$annee]["cot_salariales"]    = $cot_salariales
					$script:REMUN[$primaire][$annee]["cot_patronale"]     = $cot_patronale
					DBG "Compute_multi_remuneration" "Matricule [$zeroprimaire] n'avait pas de remuneration pour l'année $annee >> presence de remuneration dans profil secondaire [$zerosecondaire]" 
				} else {
					$script:REMUN[$primaire][$annee]["salaire_base"]      += $salaire_base
					$script:REMUN[$primaire][$annee]["primes"]            += $primes
					$script:REMUN[$primaire][$annee]["heures_supp"]       += $heures_supp
					$script:REMUN[$primaire][$annee]["cot_salariales"]    += $cot_salariales
					$script:REMUN[$primaire][$annee]["cot_patronale"]     += $cot_patronale
				}
				DBG "Compute_multi_remuneration" "$primaire a pour profil secondaire $secondaire  $annee >> salaire_base $salaire_base, primes $primes, heures_supp $heures_supp, cot_salariales $cot_salariales,  cot_patronale $cot_patronale"
			}
		}
	}
}
function Update_BDD_remuneration {
	$table = $script:cfg["SQL_Postgre_Review"]["tableRemuneration"]
	$keycolumns = @("utilisateur_id","annee")
	
	#LOG "Update_BDD_remuneration" "Update de la table $table" -CRLF
	Update_BDDTable $script:REMUN $script:BDDRemun $keycolumns $table "Update_BDD_remuneration" { Query_BDD_remuneration }
}

# --------------------------------------------------------
#   Table performance_co
#	$script:PERF    = @{}
#	$script:BDDPerf = @{}
# --------------------------------------------------------

function Query_BDD_performance_co {
	$script:BDDPerf = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableutilisateur_performance_co"] -functionName "Query_BDD_performance_co" -keyColumns @("utilisateur_id","annee") -targetVariable $script:BDDPerf
}
function Compute_performance_co {
	$script:PERF = @{}
    foreach ($zeromatricule in $script:Review.Keys) {
		$matricule = $zeromatricule  -replace '^0+', ''

		# Verifier si l'utilisateur existe dans la BDD utilisateur 
		if (-not (Test-UserExists $matricule $zeromatricule "Compute_performance_co")) {
			continue
		}

		$script:PERF[$matricule] = @{}
		foreach ($annee in $script:Review[$zeromatricule].Keys) {
			$script:PERF[$matricule][$annee] = @{}
			$epargne_salariale = 0
			$script:PERF[$matricule][$annee]["utilisateur_id"]     = $matricule
			$script:PERF[$matricule][$annee]["annee"]              = $annee
			foreach ($CatRub in $script:Review[$zeromatricule][$annee].Keys) {
				$montant =  [Math]::Ceiling($script:Review[$zeromatricule][$annee][$CatRub])
				if ( $CatRub -eq "EPARGNE SAL" )   { $epargne_salariale  += $montant }
			}
			$script:PERF[$matricule][$annee]["epargne_salariale"] += $epargne_salariale
		}
    }
}
function Compute_multi_performance_co {
	LOG "Compute_multi_performance_co" "Additionne au profil primaire la performance_co des profils secondaires" -CRLF
	foreach ($primaire in $script:PRIMAIRE.keys) {
		$zeroprimaire = $primaire.PadLeft(8, '0')
		foreach ($secondaire in $script:PRIMAIRE[$primaire].keys) {
			$zerosecondaire = $secondaire.PadLeft(8, '0')

			# Pour chaque année, ajouter les performance_co des profils secondaires 
			foreach ($annee in $script:ReviewAll[$zerosecondaire].Keys) {
				$epargne_salariale = 0
				foreach ($CatRub in $script:ReviewAll[$zerosecondaire][$annee].Keys) {
					$montant =  [Math]::Ceiling($script:ReviewAll[$zerosecondaire][$annee][$CatRub])
					if ( $CatRub -eq "EPARGNE SAL" )   { $epargne_salariale  = $montant }
				}

				if ( -not ($script:BDDRemun.ContainsKey($primaire)) ) {
					$script:PERF[$primaire] = @{}
				}

				# gere l'absence de performance_co pour une année dans le profil primaire
				if ( -not ($script:PERF[$primaire].ContainsKey($annee)) ) {
					$script:PERF[$primaire][$annee] = @{}
					$script:PERF[$primaire][$annee]["utilisateur_id"]    = $primaire
					$script:PERF[$primaire][$annee]["annee"]             = $annee
					$script:PERF[$primaire][$annee]["epargne_salariale"] = $epargne_salariale
					DBG "Compute_multi_performance_co" "Matricule [$zeroprimaire] n'avait pas de performance_co pour l'année $annee >> presence de performance_co dans profil secondaire [$zerosecondaire]" 
				} else {
					$script:PERF[$primaire][$annee]["epargne_salariale"] += $epargne_salariale
				}
				DBG "Compute_multi_performance_co" "$primaire a pour profil secondaire $secondaire  $annee >> performance_co $epargne_salariale"
			}
			
		}
	}
}
function Update_BDD_performance_co {
	$table = $script:cfg["SQL_Postgre_Review"]["tableutilisateur_performance_co"]
	$keycolumns = @("utilisateur_id","annee")

	#LOG "Update_BDD_performance_co" "Update de la table $table" -CRLF
	Update_BDDTable $script:PERF $script:BDDPerf $keycolumns $table "Update_BDD_performance_co" { Query_BDD_performance_co }
}

# --------------------------------------------------------
#   Table protection sociale
#	$script:PROT       = @{}
#	$script:BDDProt = @{}
# --------------------------------------------------------

function Query_BDD_protect_sociale {
	$script:BDDProt = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableprotection_sociale_utilisateur"] -functionName "Query_BDD_protect_sociale" -keyColumns @("utilisateur_id","annee") -targetVariable $script:BDDProt
}
function Compute_protect_sociale {
	$script:PROT = @{}
	$currentyear = (Get-Date).Year
    foreach ($zeromatricule in $script:Review.Keys) {
		$matricule = $zeromatricule  -replace '^0+', ''

		# Verifier si l'utilisateur existe dans la BDD utilisateur 
		if ( $script:BDDuser.ContainsKey($matricule) -eq $false ) {
			Add-Exclusion $matricule "BDDuser" "Matricule [$zeromatricule] n'existe pas dans la table [$($script:cfg["SQL_Postgre_Review"]["tableUtilisateurs"])]" -USER
			continue
		}

		$script:PROT[$matricule] = @{}
		foreach ($annee in $script:Review[$zeromatricule].Keys) {
			if ( $script:cfg["start"]["BypassProtSocialeCurrentYear"] -eq "yes") {
				if ($annee -eq $currentyear) {
					continue
				}
			}
			$script:PROT[$matricule][$annee] = @{}
			$script:PROT[$matricule][$annee]["utilisateur_id"] = $matricule
			$script:PROT[$matricule][$annee]["annee"]          = $annee
			$prevoyance_salariale = 0
			$prevoyance_patronale = 0
			$mutuelle_salariale   = 0
			$mutuelle_patronale   = 0
			$retraite_salariale   = 0
			$retraite_patronale   = 0
			foreach ($CatRub in $script:Review[$zeromatricule][$annee].Keys) {
				$montant =  [Math]::Ceiling($script:Review[$zeromatricule][$annee][$CatRub])
				if ( $CatRub -eq "PREVOYANCE SAL" ) { $prevoyance_salariale = $montant }
				if ( $CatRub -eq "PREVOYANCE PAT" ) { $prevoyance_patronale = $montant }
				if ( $CatRub -eq "MUTUELLE SAL" )   { $mutuelle_salariale   = $montant }
				if ( $CatRub -eq "MUTUELLE PAT" )   { $mutuelle_patronale   = $montant }
				if ( $CatRub -eq "RETRAITE SAL" )   { $retraite_salariale   = $montant }
				if ( $CatRub -eq "RETRAITE PAT" )   { $retraite_patronale   = $montant }
			}
			$script:PROT[$matricule][$annee]["prevoyance_salariale"] = $prevoyance_salariale
			$script:PROT[$matricule][$annee]["prevoyance_patronale"] = $prevoyance_patronale
			$script:PROT[$matricule][$annee]["mutuelle_salariale"]   = $mutuelle_salariale
			$script:PROT[$matricule][$annee]["mutuelle_patronale"]   = $mutuelle_patronale
			$script:PROT[$matricule][$annee]["retraite_salariale"]   = $retraite_salariale
			$script:PROT[$matricule][$annee]["retraite_patronale"]   = $retraite_patronale
		}
    }
}
function Compute_multi_protect_sociale {
	LOG "Compute_multi_protect_sociale" "Additionne au profil primaire la protection sociale des profils secondaires" -CRLF
	foreach ($primaire in $script:PRIMAIRE.keys) {
		$zeroprimaire = $primaire.PadLeft(8, '0')
		foreach ($secondaire in $script:PRIMAIRE[$primaire].keys) {
			$zerosecondaire = $secondaire.PadLeft(8, '0')

			# Pour chaque année, ajouter les performance_co des profils secondaires 
			foreach ($annee in $script:ReviewAll[$zerosecondaire].Keys) {
				$prevoyance_salariale = 0
				$prevoyance_patronale = 0
				$mutuelle_salariale   = 0
				$mutuelle_patronale   = 0
				$retraite_salariale   = 0
				$retraite_patronale   = 0
				foreach ($CatRub in $script:ReviewAll[$zerosecondaire][$annee].Keys) {
					$montant =  [Math]::Ceiling($script:ReviewAll[$zerosecondaire][$annee][$CatRub])
					if ( $CatRub -eq "PREVOYANCE SAL" ) { $prevoyance_salariale = $montant }
					if ( $CatRub -eq "PREVOYANCE PAT" ) { $prevoyance_patronale = $montant }
					if ( $CatRub -eq "MUTUELLE SAL" )   { $mutuelle_salariale   = $montant }
					if ( $CatRub -eq "MUTUELLE PAT" )   { $mutuelle_patronale   = $montant }
					if ( $CatRub -eq "RETRAITE SAL" )   { $retraite_salariale   = $montant }
					if ( $CatRub -eq "RETRAITE PAT" )   { $retraite_patronale   = $montant }
				}

				if ( -not ($script:PROT.ContainsKey($primaire)) ) {
					$script:PROT[$primaire] = @{}
				}

				# gere l'absence de protection sociale pour une année dans le profil primaire
				if ( -not ($script:PROT[$primaire].ContainsKey($annee)) ) {
					$script:PROT[$primaire][$annee] = @{}
					$script:PROT[$primaire][$annee]["utilisateur_id"]       = $primaire
					$script:PROT[$primaire][$annee]["annee"]                = $annee
					$script:PROT[$primaire][$annee]["prevoyance_salariale"] = $prevoyance_salariale
					$script:PROT[$primaire][$annee]["prevoyance_patronale"] = $prevoyance_patronale
					$script:PROT[$primaire][$annee]["mutuelle_salariale"]   = $mutuelle_salariale
					$script:PROT[$primaire][$annee]["mutuelle_patronale"]   = $mutuelle_patronale
					$script:PROT[$primaire][$annee]["retraite_salariale"]   = $retraite_salariale
					$script:PROT[$primaire][$annee]["retraite_patronale"]   = $retraite_patronale
					DBG "Compute_multi_protect_sociale" "Matricule [$zeroprimaire] n'avait pas de protection sociale pour l'année $annee >> presence de protection sociale dans profil secondaire [$zerosecondaire]" 
				} else {
					$script:PROT[$primaire][$annee]["prevoyance_salariale"] += $prevoyance_salariale
					$script:PROT[$primaire][$annee]["prevoyance_patronale"] += $prevoyance_patronale
					$script:PROT[$primaire][$annee]["mutuelle_salariale"]   += $mutuelle_salariale
					$script:PROT[$primaire][$annee]["mutuelle_patronale"]   += $mutuelle_patronale
					$script:PROT[$primaire][$annee]["retraite_salariale"]   += $retraite_salariale
					$script:PROT[$primaire][$annee]["retraite_patronale"]   += $retraite_patronale
				}
				DBG "Compute_multi_protect_sociale" "$primaire a pour profil secondaire $secondaire  $annee >> Valeurs : $prevoyance_salariale, $prevoyance_patronale, $mutuelle_salariale, $mutuelle_patronale, $retraite_salariale, $retraite_patronale"
			}
		}
	}
}
function Update_BDD_protect_sociale {
	$table = $script:cfg["SQL_Postgre_Review"]["tableprotection_sociale_utilisateur"]
	$keycolumns  = @("utilisateur_id","annee")

	#LOG "Update_BDD_protect_sociale" "Update de la table $table" -CRLF
	Update_BDDTable $script:PROT $script:BDDProt $keycolumns $table "Update_BDD_protect_sociale" { Query_BDD_protect_sociale }
}

# --------------------------------------------------------
#   Table conges_utilisateur
#	$script:CONGES    = @{}
#	$script:BDDConges = @{}
# --------------------------------------------------------

function Query_BDD_conges_utilisateur {
	$script:BDDConges = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableconges_utilisateur"] -functionName "Query_BDD_conges_utilisateur" -keyColumns @("utilisateur_id","annee") -targetVariable $script:BDDConges
}
function Compute_conges_utilisateur {
	$invalid = $script:cfg["Exclude"]["zero_conge_jours_fortil"].Split(',') | ForEach-Object { $_.Trim() }

	$script:CONGES = @{}
    foreach ($matricule in $script:BDDRemun.Keys) {
		$zeromatricule = $matricule.PadLeft(8, '0')
		# Verifier si l'utilisateur existe dans la BDD utilisateur 
		if ( $script:BDDuser.ContainsKey($matricule) -eq $false ) {
			Add-Exclusion $matricule "BDDRemun" "n'existe pas dans la table [$($script:cfg["SQL_Postgre_Review"]["tableUtilisateurs"])]" -USER
			continue
		}
		$script:CONGES[$matricule] = @{}
		$date_entree = [datetime]::ParseExact($script:BDDuser[$matricule]["checkin"],$script:cfg["SQL_Postgre_Review"]["frmtdateOUT"],$null)
		$entite = $script:BDDuser[$matricule]["entite"]

		foreach ($annee in $script:BDDRemun[$matricule].Keys) {
			$script:CONGES[$matricule][$annee] = @{}
			$script:CONGES[$matricule][$annee]["utilisateur_id"] = $matricule
			$script:CONGES[$matricule][$annee]["annee"]          = $annee

			$salaire_base_annee = $script:BDDRemun[$matricule][$annee]["salaire_base"]

			# Date de base de l'année courante : 01/01/$annee
			try {
				$date_debut_annee  = [datetime]::ParseExact("$annee/01/01","yyyy/MM/dd",$null)
				$date_middle_annee = [datetime]::ParseExact("$annee/06/01","yyyy/MM/dd",$null)
			} catch {
				ERR "Compute_conges_utilisateur" "Matricule [$zeromatricule] : Erreur lors de la conversion des dates pour l'annee $annee" -CRLF
				continue
			}

			[string]$annee_actuelle = (Get-Date).Year
			if ( $annee -eq $annee_actuelle ) {
				# Cas particulier pour l'annee actuelle : La date_fin_annee est la date du jour
				$date_fin_annee = Get-Date
			} else {
				$date_fin_annee = [datetime]::ParseExact("$annee/12/31","yyyy/MM/dd",$null)
			}
			
			# calcul ancienneté en année (>= 1 an) par rapport au 01/06/$annee (en années pleines avec comme base de debut d'annee 01/06)
			$anciennete_annee_middle = Annee_pleine $date_middle_annee $date_entree

			# calcul ancienneté en année (>= 1 an) par rapport au 01/01/$annee (en années pleines avec comme base de debut d'annee 01/01)
			$anciennete_annee_debut  = Annee_pleine $date_debut_annee $date_entree

			# Nombre de jours de presence dans $annee
			$nb_jours_presence_dans_annee = [Math]::Max(0,[Math]::Min(365, ($date_fin_annee - $date_entree).Days))

			# Verification de la date de checkin
			if ( $date_entree -gt $date_fin_annee ) {
				if ( $salaire_base_annee -gt 0 ) {
					ERR "Compute_conges_utilisateur" "Matricule [$zeromatricule] a une date de checkin [$($date_entree.ToString("yyyy/MM/dd"))] supérieure à l'année $annee (Salaire $annee : $salaire_base_annee)"
				} else {
					WRN "Compute_conges_utilisateur" "Matricule [$zeromatricule] a une date de checkin [$($date_entree.ToString("yyyy/MM/dd"))] supérieure à l'année $annee (Salaire $annee : 0)"
				}
			}

			# jours total : 25/365*(nbre de jour de presence).
			$jours_total = 25/365 * $nb_jours_presence_dans_annee

			$jours_syntec = Calcul_Jours_Syntec $anciennete_annee_middle

			# Calcul monetisation_syntec
			$monetisation_syntec = Calcul_Monetisation $salaire_base_annee $annee $jours_syntec

			# Calcul_Jours_Fortil
			$jours_fortil = Calcul_Jours_Fortil $annee $anciennete_annee_debut $entite $invalid

			# Calcul monetisation fortil
			$monetisation_fortil = Calcul_Monetisation $salaire_base_annee $annee $jours_fortil

			$script:CONGES[$matricule][$annee]["jours_total"]         = [decimal]::Round($jours_total,2)
			$script:CONGES[$matricule][$annee]["jours_syntec"]        = $jours_syntec
			$script:CONGES[$matricule][$annee]["monetisation_syntec"] = [math]::Ceiling($monetisation_syntec)
			$script:CONGES[$matricule][$annee]["jours_fortil"]        = $jours_fortil
			$script:CONGES[$matricule][$annee]["monetisation_fortil"] = [math]::Ceiling($monetisation_fortil)
		}
    }
}
function Update_BDD_conges_utilisateur {
	$table = $script:cfg["SQL_Postgre_Review"]["tableconges_utilisateur"]
	$keycolumns  = @("utilisateur_id","annee")

	#LOG "Update_BDD_conges_utilisateur" "Update de la table $table" -CRLF
	Update_BDDTable $script:CONGES $script:BDDConges $keycolumns $table "Update_BDD_conges_utilisateur" { Query_BDD_conges_utilisateur }
}
function Calcul_Jours_Syntec {
	param ($anciennete_annee_middle)

	$jours_syntec = 0
	if     ( $anciennete_annee_middle -lt  5 ) 	{ $jours_syntec = 0 } # de 0 à 4 ans inclus
	elseif ( $anciennete_annee_middle -lt 10 ) 	{ $jours_syntec = 1 } # de 5 à 9 ans inclus
	elseif ( $anciennete_annee_middle -lt 15 ) 	{ $jours_syntec = 2 } # de 10 à 14 ans inclus
	elseif ( $anciennete_annee_middle -lt 20 ) 	{ $jours_syntec = 3 } # de 15 à 19 ans inclus
	else                                		{ $jours_syntec = 4 } # >= 20 ans

	return $jours_syntec
}
function Calcul_Jours_Fortil {
	param ($annee, $anciennete_annee_debut, $entite, $invalid)

	if ( -not ([string]::IsNullOrWhiteSpace($entite)) )   {
		if ($invalid -ne $null -and $invalid.Count -gt 0) {
			foreach ($v in $invalid) {
				if ( $entite.startswith($v)  ) {
					return 0
				}
			}
		}
	}

	$jours_fortil = 0
	if       ( $annee -eq "2023" -or $annee -eq "2024") { 
		# traitement cas 2023 - 2024
		if     ( $anciennete_annee_debut -ge 4 ) { $jours_fortil = 5 } # si >= 48 mois = 4 ans
		elseif ( $anciennete_annee_debut -ge 2 ) { $jours_fortil = 4 } # si >= 24 mois = 2 ans
	} elseif ( $annee -ge "2025" ) {													
		# traitement 2025 et +
		if     ( $anciennete_annee_debut -ge 4 ) { $jours_fortil = 8 } # si >= 48 mois = 4 ans
		elseif ( $anciennete_annee_debut -ge 2 ) { $jours_fortil = 6 } # si >= 24 mois = 2 ans
		elseif ( $anciennete_annee_debut -ge 1 ) { $jours_fortil = 4 } # si >= 12 mois = 1 ans
	}

	return $jours_fortil
}
function Calcul_Monetisation {
	param ($SB, $annee, $nbjours)

	$nbr_jours_ouvrables_moyen_par_mois = 21.667
	$currentyear = "$((Get-Date).Year)"
	$currentmonth = (Get-Date).Month

	if ( $annee -eq $currentyear) { 
		if ( $currentmonth -eq 1 ) { return 0 }
		$val = (( $SB / ( $currentmonth -1 ) ) * 12) / 12 / $nbr_jours_ouvrables_moyen_par_mois * $nbjours
	} else {
		$val = $SB / 12 / $nbr_jours_ouvrables_moyen_par_mois * $nbjours
	}
	return $val
}

# --------------------------------------------------------
#	Fichier ['XLS_CSV_Frais']['fichierXLS_Fusion_histo']			[DRH - Frais Policy_Fusion_Group_YTD_2024*.xls]
#	Fichier ['XLS_CSV_Frais']['fichierXLS_Fusion_current']			[DRH - Frais Policy_Fusion_Group_YTD_2025*.xls]
#	Fichier ['XLS_CSV_Frais']['fichierXLS_Fortil_histo']			[DRH - Frais Policy_FORTIL GROUP_YTD_2024*.xls]
#	Fichier ['XLS_CSV_Frais']['fichierXLS_Fortil_current']			[DRH - Frais Policy_FORTIL GROUP_YTD_2025*.xls]
#	Fichier ['XLS_CSV_Frais']['fichierXLS_Reservation-berceau']		[Politique famille - réservation berceau.xlsx ]
#	Fichier ['XLS_CSV_Frais']['fichierCSV_famille']					[req21_Revue_annuelle__Absences-*.csv         ]
#   Table Utilisateurs_frais
#	$script:CSVFam   	= @{}
#	$script:XLSBerceau	= @{}
#	$script:FRAIS    	= @{}
#	$script:PMSS     	= @{}
#	$script:FAM      	= @{}
#	$script:BDDFrais 	= @{}
# --------------------------------------------------------

function Query_XLS_fusion {
	$script:FRAIS = @{}
    # ---------------------------------------------------------
    # Chargement du fichier fichierXLS_Fusion_histo en memoire
    # ---------------------------------------------------------
    $xlsfusionhisto = $script:cfg['XLS_CSV_Frais']['fichierXLS_Fusion_histo']
    Cumul_XLS_Frais -FilePath $xlsfusionhisto -FileType "Fusion_histo" -FunctionName "Query_XLS_fusion" -CountNewEntries $false
    # ---------------------------------------------------------
    # Chargement du fichier fichierXLS_Fusion_current en memoire
    # ---------------------------------------------------------
    $xlsfusioncurrent = $script:cfg['XLS_CSV_Frais']['fichierXLS_Fusion_current']
    Cumul_XLS_Frais -FilePath $xlsfusioncurrent -FileType "Fusion_current" -FunctionName "Query_XLS_fusion" -CountNewEntries $false
}
function Query_XLS_fortil {
    # ---------------------------------------------------------
    # Chargement du fichier fichierXLS_Fortil_histo en memoire
    # ---------------------------------------------------------
    $xlsfortilhisto = $script:cfg['XLS_CSV_Frais']['fichierXLS_Fortil_histo']
    Cumul_XLS_Frais -FilePath $xlsfortilhisto -FileType "Fortil_histo" -FunctionName "Query_XLS_fortil" -CountNewEntries $true
    # ---------------------------------------------------------
    # Chargement du fichier fichierXLS_Fortil_current en memoire
    # ---------------------------------------------------------
    $xlsfortilcurrent = $script:cfg['XLS_CSV_Frais']['fichierXLS_Fortil_current']
    Cumul_XLS_Frais -FilePath $xlsfortilcurrent -FileType "Fortil_current" -FunctionName "Query_XLS_fortil" -CountNewEntries $true
}
function Cumul_XLS_Frais {
    param(
        [string]$FilePath,
        [string]$FileType,  # "Fusion" ou "Fortil"
        [string]$FunctionName,
        [bool]$CountNewEntries = $false
    )
    
    $keycol = "Matricule"
    $sqlquery = "SELECT [Matricule], [Dénomination sociale], [Catégorie (Description)], [Date du frais], [Montant] FROM [Sheet0$] WHERE [Statut] = 'PMNT_NOT_PAID'"
	$datecol = @("Date du frais")

    $result = Invoke-ExcelQuery -filePath $FilePath -sqlQuery $sqlquery -functionName $FunctionName -frmtdateOUT $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -datecol $datecol -dateLocale "FR"
    $table = $result.Table
    $n = $result.RowCount

    LOG $FunctionName "Chargement du fichier $FilePath en memoire ($n lignes)" -CRLF

    # Traitement des données
    $oldmat = 0
    $exclu = 0
    $cptNew = 0
    
    foreach ($row in $table.Rows) {
        [string]$matricule = $row[$keycol]
        $zeromatricule = $matricule.PadLeft(8, '0')
        $categorie = ""
        [datetime]$datefrais = 0
        $montant = 0
        $entite = ""

        if (-not $script:ListeBM.ContainsKey($matricule)) {
            if ($matricule -ne $oldmat) { 
                $oldmat = $matricule
                Add-Exclusion $matricule "XLS_Utilisateurs_frais" "($FileType) non trouvé dans Liste BM" -USER
                $exclu++
            }
            continue
        }

        foreach ($col in $table.Columns) {
            if ($col.Caption -eq "Catégorie (Description)") { $categorie = $row[$col] }
            if ($col.Caption -eq "Date du frais") { $datefrais = $row[$col] }
            if ($col.Caption -eq "Montant") { $montant = [math]::Ceiling($row[$col]) }
            if ($col.Caption -eq "Dénomination sociale") { $entite = $row[$col] }
        }

        if (-not $script:FRAIS.ContainsKey($matricule)) { 
            $script:FRAIS[$matricule] = @{} 
            if ($CountNewEntries) { $cptNew++ }
        } 
        
        $annee = $datefrais.Year
        if (-not $script:FRAIS[$matricule].ContainsKey($annee)) { 
            $script:FRAIS[$matricule][$annee] = @{} 
            $script:FRAIS[$matricule][$annee]["utilisateur_id"] = $matricule
            $script:FRAIS[$matricule][$annee]["annee"] = $annee
            $script:FRAIS[$matricule][$annee]["repas"] = 0
            $script:FRAIS[$matricule][$annee]["transport"] = 0
            $script:FRAIS[$matricule][$annee]["comite_entreprise"] = 0
            $script:FRAIS[$matricule][$annee]["formation"] = 0
            $script:FRAIS[$matricule][$annee]["prime_naissance"] = 0
        }

        # Repas
        if ($categorie -in $script:ValidFrais) {
            $script:FRAIS[$matricule][$annee]["repas"] += $montant    
        }
        # Transport
        if ($categorie -in $script:ValidTransport) {
            $script:FRAIS[$matricule][$annee]["transport"] += $montant    
        }
        
        $script:FRAIS[$matricule][$annee]["comite_entreprise"] = comite_entreprise $entite $script:ValidEntities $annee    
    }
    
    if ($CountNewEntries) {
        LOG $FunctionName "$cptNew Matricules supplementaires($FileType) ont eu des frais. ($exclu Matricules ont été exclus car innexistant dans Liste BM)"
        LOG $FunctionName "$($script:FRAIS.Count) Matricules (Fusion+Fortil) ont eu des frais."
    } else {
        LOG $FunctionName "$($script:FRAIS.Count) Matricules ($FileType) ont eu des frais. ($exclu Matricules ont été exclus car innexistant dans Liste BM)"
    }
}
function Fill_Empty_Frais {
	# ---------------------------------------------------------
	# Completer avec tous les users de Liste BM qui n'ont pas de frais depuis 2022 inclu, si la date d'entree est plus ancienne que l'année à remplir
	# Mettre à 0 toutes les valeurs, sauf $fortil[$matricule][$annee]["comite_entreprise"] à 30 ou 0
	# ---------------------------------------------------------

	LOG "Fill_Empty_Frais" "Remplissage des frais à 0 des Matricules de Liste BM sans frais (pour les années non remplies) et 0 ou 30 dans [comite_entreprise] depuis 2019" -CRLF
   	foreach ($matricule in $script:ListeBM.Keys) {
		$zeromatricule = $matricule.PadLeft(8, '0')
		$anneestart  = 2019
		$anneeend    = (Get-Date).Year
		if (-not $script:BDDuser.ContainsKey($matricule)) { 
			ERR "Fill_Empty_Frais" "Matricule [$zeromatricule] innexistant dans la table Utilisateur"
			continue
		}
		[datetime]$dateentree = $script:BDDuser[$matricule]["checkin"]
		$anneeentree = $dateentree.Year 
		$entite      = $script:ListeBM[$matricule]["Dénomination sociale"]

		for ($annee = $anneestart; $annee -le $anneeend; $annee++) {
			if ( $anneeentree -le $annee ) {
				if (-not $script:FRAIS.ContainsKey($matricule)) { 
					$script:FRAIS[$matricule] = @{} 
				} 

				if (-not $script:FRAIS[$matricule].ContainsKey($annee)) { 
					$script:FRAIS[$matricule][$annee] = @{} 
					$script:FRAIS[$matricule][$annee]["utilisateur_id"]    = $matricule
					$script:FRAIS[$matricule][$annee]["annee"]             = $annee
					$script:FRAIS[$matricule][$annee]["repas"]             = 0
					$script:FRAIS[$matricule][$annee]["transport"]         = 0
					$script:FRAIS[$matricule][$annee]["comite_entreprise"] = comite_entreprise $entite $script:ValidEntities $annee # >>>>
					$script:FRAIS[$matricule][$annee]["formation"]         = 0
					$script:FRAIS[$matricule][$annee]["prime_naissance"] = 0
				}
			}
		}
    }
}
function Compute_FLOTTE_AUTO {
	# Decrementer la somme des frais transport avec CatRUB = "FLOTTE AUTO" de Review
	LOG "Compute_FLOTTE_AUTO" "Completer la somme transport avec CatRUB = [FLOTTE AUTO] de Review" -CRLF
    foreach ($matricule in $script:ListeBM.Keys) {
		$zeromatricule = $matricule.PadLeft(8, '0')
		foreach ($annee in $script:Review[$zeromatricule].Keys) {
			$an = [int]$annee
			if ( $script:Review[$zeromatricule][$annee].ContainsKey("FLOTTE AUTO") ) {
				try {
					$val = [math]::Ceiling($script:Review[$zeromatricule][$annee]["FLOTTE AUTO"])
					try {
						$script:FRAIS[$matricule][$an]["transport"] -= $val
					} catch {
						ERR "Compute_FLOTTE_AUTO" "Matricule [$zeromatricule] [$annee] : Erreur ajout FLOTTE AUTO : valeur [$val] dans [transport]"
					}
				} catch {
					ERR "Compute_FLOTTE_AUTO" "Matricule [$zeromatricule] [$annee] : Erreur valeur FLOTTE AUTO"
				}
			}
		}
	}
}
function Query_CSV_prime_naissance {
	# ---------------------------------------------------------
	# Chargement du fichier fichierCSV_famille en memoire pour champ "prime_naissance"
	# ---------------------------------------------------------
	LOG "Query_CSV_prime_naissance" "Chargement du fichier $csvfamille" -CRLF

	$csvfamille = $script:cfg['XLS_CSV_Frais']['fichierCSV_famille']
	$headerstartline = $script:cfg['XLS_CSV_Frais']['HEADERstartline_famille'] - 1
	$datecol  = @("D Naissance fam.")
	$script:CSVFam = @{}
	$script:CSVFam = Invoke-CSVQuery -csvfile $csvfamille -key "_index" -separator "," -row $headerstartline -frmtdateOUT $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -datecol $datecol
	LOG "Query_CSV_prime_naissance" "Chargement du fichier $csvfamille en memoire ($($script:FAM.Count) lignes)" 
}
function Compute_prime_naissance {
	# Calculer du pourcentage par année pour prime_naissance
	$script:PMSS = @{}
	foreach ($annee in $script:cfg["PMSS"].Keys) { 
		$script:PMSS[$annee] = @{}
		foreach ($str in $script:cfg["PMSS"][$annee]) {
			$amount  = [decimal]$str.split(",")[0]
			$percent = [decimal]$str.split(",")[1]
			$script:PMSS[$annee]["amount"]  = $amount
			$script:PMSS[$annee]["percent"] = $percent
			$script:PMSS[$annee]["prime_naissance"]  = [math]::Ceiling($amount * $percent / 100)
		}
	}

	$script:FAM = @{}

	foreach ($key in $script:CSVFam.keys) {
		# parcourrir toutes les lignes de $script:CSVFam, 
		# inclure dans $script:FAM tous les [matricules][annee] dont [Matricule] présent (et encore actif) avant [D Naissance fam.] de l'année
		$zeromatricule = $script:CSVFam[$key]['Matricule']
		if ( $script:USER.ContainsKey($zeromatricule)) {
			$matricule = $zeromatricule  -replace '^0+', ''
			$entree    = ConvertTo-SafeDate -dateString $script:USER[$zeromatricule]["Date d'entrée"] -format  $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -matricule $zeromatricule -fieldName "Date d'entrée" -functionName "Query_XLS_CSV_Utilisateurs_frais"
			$naissance = ConvertTo-SafeDate -dateString $script:CSVFam[$key]['D Naissance fam.'] -format  $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -matricule $zeromatricule -fieldName "D Naissance fam." -functionName "Query_XLS_CSV_Utilisateurs_frais"
			if ( $naissance.Date -ne $null ) {
				if ( $naissance.Date -ge $entree.Date ) {
					$annee = $naissance.Date.Year
					if ( $script:PMSS.ContainsKey("$annee") ) {
						if ( $script:FAM.ContainsKey($zeromatricule) -eq $false ) {
							$script:FAM[$zeromatricule] = @{}
						}
						if ( $script:FAM[$zeromatricule].ContainsKey($annee) -eq $false ) {
							$script:FAM[$zeromatricule][$annee] = $script:PMSS["$annee"]["prime_naissance"]
						} else {
							DBG "Compute_prime_naissance" "Matricule [$zeromatricule] [$annee]: Deux enfants dans la même année"
							$script:FAM[$zeromatricule][$annee] += $script:PMSS["$annee"]["prime_naissance"]
						}
					}
				} else {
					Add-Exclusion $matricule "CSV_famille" "Date naissance [$($naissance.Date)] inférieure Date Entrée [$($dateentree.Date)]"
				}
			}
		} else {
			Add-Exclusion $matricule "CSV_famille" "n'existe pas dans les utilisateurs USER non exclus"
		}
	}

	# ---------------------------------------------------------
	# Remplacer les valeurs de $script:FRAIS[$matricule][$an]['prime_naissance'] par celles de $script:FAM quand elles existent
	# ---------------------------------------------------------
	foreach ($zeromatricule in $script:FAM.keys) {
		$matricule = $zeromatricule  -replace '^0+', ''
		if ( $script:FRAIS.ContainsKey($matricule) ) {
			foreach ($annee in $script:FAM[$zeromatricule].keys) {
				$an = [int]$annee
				if ( $script:FRAIS[$matricule].ContainsKey($an) ) {
					$script:FRAIS[$matricule][$an]['prime_naissance'] = $script:FAM[$zeromatricule][$annee]
				} else {
					ERR "Compute_prime_naissance" "Matricule [$zeromatricule] [$annee]: Pas de données FRAIS disponible pour l'année [$annee]"
				}
			}
		}
	}
}
function Query_XLS_Reservation_berceau {
	# ---------------------------------------------------------
	# Chargement du fichier XLS_Reservation-berceau en memoire
	# Hypothèse : Chaque mois entamé est un mois due.
	# ---------------------------------------------------------
	$fichierxls  = $script:cfg['XLS_CSV_Frais']['fichierXLS_Reservation-berceau']
	LOG "Query_XLS_Reservation-berceau" "Chargement du fichier $fichierxls" -CRLF

	$columnMapping = @{
		"matricule"    = 2
		"date_debut"   = 4  
		"date_fin"     = 5
		"prix_mensuel" = 9
	}
	$datecol = @("date_debut", "date_fin")
	$sqlquery  = "SELECT [matricule], [date_debut], [date_fin], [prix_mensuel] FROM [Liste berceau$]"

	$result = Invoke-ExcelQuery -filePath $fichierxls -sqlQuery $sqlquery -functionName "Query_XLS_Reservation-berceau" -columnMapping $columnMapping -frmtdateOUT $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -datecol $datecol -dateLocale "FR"
    $script:XLSBerceau = $result.Table

}
function Compute_Reservation_berceau {
	$today = (Get-Date)
	$currentyear = $today.Year

	$prevmonth   = (Get-Date -Year $currentyear -Month $today.Month -Day 1).AddDays(-1)
	$PrevMonth = (Get-Date -Year $today.Year -Month $today.Month -Day 1).AddDays(-1)
    foreach ($row in $script:XLSBerceau.Rows) {
        [string]$zeromatricule = "$($row['matricule'])".Trim()
		if ( $zeromatricule -match '\d+' ) { 
			$matricule = $zeromatricule  -replace '^0+', ''
			if ( $script:ListeBM.ContainsKey($matricule)) {
				$date_debut = $row['date_debut']
				$date_fin   = $row["date_fin"]
				
				# Convertir les chaînes de dates en objets DateTime pour les calculs
				if ($date_debut -is [string]) {
					try {
						$date_debut = [datetime]::Parse($date_debut)
					} catch {
						ERR "Compute_Reservation_berceau" "Matricule [$zeromatricule] : [Date debut] non valide - $($_.Exception.Message)"
						continue
					}
				}
				if ($date_fin -is [string]) {
					try {
						$date_fin = [datetime]::Parse($date_fin)
					} catch {
						ERR "Compute_Reservation_berceau" "Matricule [$zeromatricule] : [Date fin] non valide - $($_.Exception.Message)"
						continue
					}
				}
				
				if (-not ($date_debut -is [datetime])) { ERR "Query_XLS_Reservation-berceau" "Matricule [$zeromatricule] : [Date debut] non valide"; continue}
				if (-not ($date_fin -is [datetime]))   { ERR "Query_XLS_Reservation-berceau" "Matricule [$zeromatricule] : [Date fin]  non valide"; continue}

				# Exclure si $date_debut > today
				if ( $date_debut -gt $today ) {
					Add-Exclusion $matricule "XLS_Reservation-berceau" "Date début [$date_debut] supérieure à aujourd'hui"
					continue
				}
				
				# prix_mensuel
				$maxyear = [math]::Min($date_fin.Year, $currentyear)
				# Nettoyer la valeur monétaire pour extraire seulement les chiffres
				$prix_mensuel_raw = $row["prix_mensuel"].ToString()
				# Supprimer les espaces, le symbole €, et remplacer la virgule par un point
				$prix_mensuel_clean = $prix_mensuel_raw -replace '[€\s]', '' -replace ',', '.'
				
				try {
					$prix_mensuel = [math]::Ceiling([double]$prix_mensuel_clean)
				} catch {
					ERR "Compute_Reservation_berceau" "Matricule [$zeromatricule] : Prix mensuel non valide [$prix_mensuel_raw] - $($_.Exception.Message)"
					continue
				}
				for ( $an=$date_debut.Year; $an -le $maxyear; $an++) {
					$montant = Calcul_berceau $date_debut $date_fin $prix_mensuel $an
					$script:FRAIS[$matricule][$an]['reservation_berceau'] += $montant
				}
				
			} else {
				Add-Exclusion $matricule "XLS_Reservation-berceau" "n'existe pas dans les utilisateurs valide de ListeBM"
			}
		} else {
			INA "Compute_Reservation_berceau" "Exclu : Matricule [$zeromatricule] n'est pas un INTEGER"
		}
	}
}
function Query_BDD_Utilisateurs_frais {
	$script:BDDFrais = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableutilisateur_frais"] -functionName "Query_BDD_Utilisateurs_frais" -keyColumns @("utilisateur_id","annee") -targetVariable $script:BDDFrais
}
function Update_BDD_Utilisateurs_frais {
	$table = $script:cfg["SQL_Postgre_Review"]["tableutilisateur_frais"]
	$keycolumns = @("utilisateur_id","annee")
	
	Update_BDDTable $script:FRAIS $script:BDDFrais $keycolumns $table "Update_BDD_Utilisateurs_frais" { Query_BDD_Utilisateurs_frais }
}
function comite_entreprise {
	Param ( [string]$entite, [string[]]$valid, $annee )

	if ([string]::IsNullOrWhiteSpace($entite))   { return 0 }
	if ($valid -eq $null -or $valid.Count -eq 0) { return 0 }

	# Vérifier si l'entité commence par l'une des valeurs dans le tableau $valid et $annee >= 2024
	if ( $annee -ge 2024 ) {
		foreach ($validValue in $valid) {
			if (-not [string]::IsNullOrWhiteSpace($validValue) -and $entite.StartsWith($validValue)) {
				return 30	
			}
		}
	}
	return 0
}
function Calcul_berceau {
    param (
        [datetime]$date_entree,
        [datetime]$date_sortie,
        [int]$prix_mensuel,
        [int]$annee
    )

    $montant = 0

    # Sécurité : vérifier que les dates sont valides
    if (-not $date_entree -or -not $date_sortie) {
        DBG "Calcul_berceau" "date_entree ou date_sortie non fournie ou invalide."
        return 0
    }

    # Définir les bornes de l’année
    $debut_annee = Get-Date -Year $annee -Month 1 -Day 1
    $fin_annee   = Get-Date -Year $annee -Month 12 -Day 31
    
    # Pour l'année courante, la fin ne peut pas dépasser la fin du mois précédent
    $annee_courante = (Get-Date).Year
    if ($annee -eq $annee_courante) {
        $aujourd_hui = Get-Date
        $fin_mois_precedent = Get-Date -Year $aujourd_hui.Year -Month $aujourd_hui.Month -Day 1
        $fin_mois_precedent = $fin_mois_precedent.AddDays(-1)
        $fin_annee = if ($fin_mois_precedent -lt $fin_annee) { $fin_mois_precedent } else { $fin_annee }
    }

    # Calculer la période effective à prendre en compte
    $periode_debut = if ($date_entree -gt $debut_annee) { $date_entree } else { $debut_annee }
    $periode_fin   = if ($date_sortie -lt $fin_annee)   { $date_sortie } else { $fin_annee }

    # Si la période est invalide, on arrête ici
    if ($periode_debut -gt $periode_fin) {
        return 0
    }

    try {
        $mois_courant = Get-Date -Year $periode_debut.Year -Month $periode_debut.Month -Day 1
    } catch {
        DBG "Calcul_berceau" "Erreur lors de la génération de la date de départ."
        return 0
    }
	$cpt = 0
    while ($mois_courant -le $periode_fin) {
        $debut_mois = $mois_courant
        $fin_mois = $mois_courant.AddMonths(1).AddDays(-1)

        # Ne pas compter le mois si la période se termine avant la fin du mois
        if (($periode_fin.Date -ge $debut_mois.Date) -and ($periode_fin.Date -lt $fin_mois.Date)) {
            break
        }

        $montant += $prix_mensuel
        $mois_courant = $mois_courant.AddMonths(1)
		$cpt++
    }

    if ($cpt -gt 0) {
    	DBG "Calcul_berceau" "Matricule [$zeromatricule] Période du $($periode_debut.ToString('dd/MM/yyyy')) au $($periode_fin.ToString('dd/MM/yyyy')), ($cpt mois) : Montant = $montant €"
    }
	
    return $montant
}

# --------------------------------------------------------
#	Fichier ["CSV_entrepreneuriat"]["fichierCSV"]		[import_valorisation_entrepreneuriat*.csv]
#   Table vallorisation_entrepreneuriat
#	$script:CSVEntrep = @{}
#	$script:VALOR     = @{}
#	$script:BDDValor  = @{}
# --------------------------------------------------------

function Query_CSV_val_entrepreneuriat {
	$csvfile = $script:cfg["CSV_entrepreneuriat"]["fichierCSV"]
    $header  = "utilisateur_id;annee;cashout_dividendes"

	$script:CSVEntrep = @{}
	$script:CSVEntrep = Invoke-CSVQuery -csvfile $csvfile -key "_index" -separator ";" -row 2 -header $header
 	LOG "Query_CSV_val_entrepreneuriat" "Chargement du fichier $csvfile en memoire ($($script:CSVEntrep.count) lignes)" -CRLF
}
function Compute_val_entrepreneuriat {
	$script:VALOR = @{}
	foreach ( $row in $script:CSVEntrep.Keys ) {
		[string]$matricule = $script:CSVEntrep[$row]["utilisateur_id"]
		$zeromatricule = $matricule.PadLeft(8, '0')
		if ( $script:BDDuser.ContainsKey($matricule) -eq $false ) {
			Add-Exclusion $matricule "CSV_entrepreneuriat" "n'existe pas dans la table [$($script:cfg["SQL_Postgre_Review"]["tableUtilisateurs"])]" -USER
			continue
		}
		[string]$annee = ([int]$script:CSVEntrep[$row]["annee"]).ToString()
		$dividendes    =  ConvertTo-SafeAmount -amountString $script:CSVEntrep[$row]["cashout_dividendes"] -matricule $zeromatricule -annee $annee -catRub "cashout_dividendes" -functionName "Query_CSV_val_entrepreneuriat"

		if ( -not ($script:VALOR.ContainsKey($matricule)) ) {
			$script:VALOR[$matricule] = @{}
		}
		if ( -not ($script:VALOR[$matricule].ContainsKey($annee)) ) {
			$script:VALOR[$matricule][$annee] = @{}
		}
		$script:VALOR[$matricule][$annee]["utilisateur_id"]     = $matricule
		$script:VALOR[$matricule][$annee]["annee"]              = $annee
		$script:VALOR[$matricule][$annee]["cashout_dividendes"] = [math]::Ceiling($dividendes.Amount)
	}
}
function Query_BDD_val_entrepreneuriat {
	$script:BDDValor = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tablevalorisation_entrepreneuriat"] -functionName "Query_BDD_val_entrepreneuriat" -keyColumns @("utilisateur_id","annee") -targetVariable $script:BDDValor
}
function Update_BDD_val_entrepreneuriat {
	$table = $script:cfg["SQL_Postgre_Review"]["tablevalorisation_entrepreneuriat"]
	$keycolumns = @("utilisateur_id","annee")
	
	Update_BDDTable $script:VALOR $script:BDDValor $keycolumns $table "Update_BDD_val_entrepreneuriat" { Query_BDD_val_entrepreneuriat }
}

# --------------------------------------------------------
#	Fichier ["XLS_Projets-liste"]["fichierXLS"]			[OLAP-Projets_liste_*.xls]
#   Table projets
#	$script:XLSProj = @{}
#	$script:PROJ    = @{}
#	$script:BDDProj = @{}
# --------------------------------------------------------

function Query_XLS_Projets_liste {
	$fichierxls  = $script:cfg["XLS_Projets-liste"]["fichierXLS"]
	$keycol    = "id"
	$columnMapping = @{
		"id" = 4
		"nom" = 5  
		"client" = 12
		"statut" = 3
	}
	$sqlquery  = "SELECT [id], [nom], [client], [statut] FROM [Sheet0$]"

    $result = Invoke-ExcelQuery -filePath $fichierxls -sqlQuery $sqlquery -functionName "Query_XLS_Projets-liste" -columnMapping $columnMapping
    $script:XLSProj = $result.Table
    $n = $result.RowCount
	LOG "Query_XLS_Projets_liste" "Chargement du fichier $fichierxls en memoire ($n lignes)" -CRLF
}
function Compute_Projets_liste {
	$script:PROJ = @{}
	[string]$id = ""

	$keycol    = "id"
    foreach ($row in $script:XLSProj.Rows) {
        [string]$id = "$($row[$keycol])".Trim()
		if ( $id -match '\d+' ) { 
			foreach ($col in $script:XLSProj.Columns) {
				if ( $col.Caption -eq "id" )        { [string]$id     = $row[$col] }
				if ( $col.Caption -eq "nom" ) 		{ [string]$nom    = $row[$col] }
				if ( $col.Caption -eq "client" )    { [string]$client = $row[$col] }
				if ( $col.Caption -eq "statut" )    { [string]$Statut = $row[$col] }
			}

			if (-not $script:PROJ.ContainsKey($id)) { 
				$script:PROJ[$id] = @{} 
			} else {
				ERR "Compute_Projets_liste" "Doublon sur Projet ID [$id]"
			}

			$script:PROJ[$id]['id']          = $id
			$script:PROJ[$id]['nom']         = $nom
			$script:PROJ[$id]['client']      = $client

			# Remplacer: PRJ_ACTIVE >> En cours, PRJ_CLOSED >> Terminé
			if ( $Statut -eq "PRJ_ACTIVE" ) {
				$script:PROJ[$id]['statut']      = "En cours"
			} elseif ( $Statut -eq "PRJ_CLOSED" ) {
				$script:PROJ[$id]['statut']      = "Terminé"
			} else {
				$script:PROJ[$id]['statut']      = "Unknown"
				ERR "Compute_Projets_liste" "Statut [$Statut] non reconnu"
			}
			$script:PROJ[$id]['description'] = ""

		} else {
			ERR "Compute_Projets_liste" "Exclu : Project ID [$id] n'est pas un INTEGER"
		}
	}
}
function Compute_Fake_Projets {
	LOG "Compute_Fake_Projets" "Ajout des projets factice > 1000000" -CRLF
	foreach ($id in $script:cfg['Projets_factice'].keys) {
		$script:PROJ[$id]           = @{}
		$script:PROJ[$id]['id']     = $id
		$script:PROJ[$id]['nom']    = $script:cfg['Projets_factice'][$id]
		$script:PROJ[$id]['statut'] = "En cours"
	}
}
function Query_BDD_projets {
	$script:BDDProj = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableprojets"] -functionName "Query_BDD_projets" -keyColumns @("id") -targetVariable $script:BDDProj
}
function Update_BDD_projets {
	$table = $script:cfg["SQL_Postgre_Review"]["tableprojets"]
	$keycolumns = @("id")
	
	Update_BDDTable $script:PROJ $script:BDDProj $keycolumns $table "Update_BDD_projets" { Query_BDD_projets }
}

# --------------------------------------------------------
#	Fichier ["XLS_Projets-affectation"]["fichierXLS"]	[OLAP-affectations_*.xlsx]
#   Table affectations
#	$script:XLSAffect = @{}
#	$script:AFFECT    = @{}
#	$script:BDDAffect = @{}
# --------------------------------------------------------

function Query_XLS_Projets_affectation {
	$xlsfile  = $script:cfg["XLS_Projets-affectation"]["fichierXLS"]
	$keycol    = "Matricule"
	$sqlquery  = "SELECT [Matricule], [Date de la FDT (Année)], [Tâche], [Quantité] FROM [Sheet0$]"

    $result = Invoke-ExcelQuery -filePath $xlsfile -sqlQuery $sqlquery -functionName "Query_XLS_Projets-affectation"
    $script:XLSAffect = $result.Table
    $n = $result.RowCount
	LOG "Query_XLS_Projets-affectation" "Chargement du fichier $xlsfile en memoire ($n lignes)" -CRLF
}
function Compute_Projets_affectation {
   # Charger la hashtable AFFECT
	[string]$projet_id      = ""
	[string]$utilisateur_id = ""
	$keycol                 = "Matricule"

    $script:AFFECT = @{}
    foreach ($row in $script:XLSAffect.Rows) {
        [string]$matricule = $row[$keycol]
		if ( $matricule -match '\d+' ) { 
			$zeromatricule = $matricule.PadLeft(8, '0')
			if ($script:ListeBM.ContainsKey($matricule) ) {
				[string]$utilisateur_id = $row["Matricule"]
				[string]$annee          = $row["Date de la FDT (Année)"]
				[string]$projet_id      = if ($row["Tâche"] -eq $null -or $row["Tâche"] -eq [DBNull]::Value) { "" } else { $row["Tâche"].ToString() }
				[string]$jours_passes   = $row["Quantité"]

				$projet_id = Replace_Fake_affectation $projet_id

				if ( $projet_id -match '\d+' ) {
					if ( $script:BDDProj.ContainsKey($projet_id) ) {
						if (-not $script:AFFECT.ContainsKey($utilisateur_id)) { 
							$script:AFFECT[$utilisateur_id] = @{} 
						} 
						if (-not $script:AFFECT[$utilisateur_id].ContainsKey($projet_id)) { 
							$script:AFFECT[$utilisateur_id][$projet_id] = @{} 
						}
						if (-not $script:AFFECT[$utilisateur_id][$projet_id].ContainsKey($annee)) { 
							$script:AFFECT[$utilisateur_id][$projet_id][$annee] = @{} 
						}
						$script:AFFECT[$utilisateur_id][$projet_id][$annee]['utilisateur_id'] = $utilisateur_id
						$script:AFFECT[$utilisateur_id][$projet_id][$annee]['projet_id']      = $projet_id
						$script:AFFECT[$utilisateur_id][$projet_id][$annee]['annee']          = $annee
						$script:AFFECT[$utilisateur_id][$projet_id][$annee]['jours_passes']   = ([decimal]$jours_passes).ToString("F2")
					} else {
						ERR "Compute_Projets_affectation" "Exclu : Matricule [$zeromatricule][$annee] : projet_id [$projet_id] innexistant dans la table projets"
						Add-Exclusion $matricule "XLS_Projets-affectation" "[$annee] : projet_id [$projet_id] innexistant dans la table projets"
					}
				} else {
					Add-Exclusion $matricule "XLS_Projets-affectation" "[$annee] : contient une Tâche qui n'est pas un INTEGER [$projet_id]"
				}
			} else {
				Add-Exclusion $matricule "XLS_Projets-affectation" "sans correspondance de Matricule dans les non exclus de Liste_BM" -USER
			}
		} else {
			Add-Exclusion $matricule "XLS_Projets-affectation" "Le matricule n'est pas un INTEGER" -USER
		}
	}

}
function Replace_Fake_affectation {
	Param ( [string]$id )

	if ( $id -match '\d+' ) { return $id }
	foreach ($key in $script:cfg['Projets_factice'].Keys) { 
		if ( $id -eq $script:cfg['Projets_factice'][$key] ) {
			$id = $key
			break
		}
	}
	return $id
}
function Query_BDD_affectations {
	$script:BDDAffect = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tableutilisateur_projet"] -functionName "Query_BDD_affectations" -keyColumns @("utilisateur_id","projet_id","annee") -targetVariable $script:BDDAffect
}
function Update_BDD_affectations {
	$table = $script:cfg["SQL_Postgre_Review"]["tableutilisateur_projet"]
	$keycolumns = @("utilisateur_id","projet_id","annee")
	
	Update_BDDTable $script:AFFECT $script:BDDAffect $keycolumns $table "Update_BDD_affectations" { Query_BDD_affectations }
}

# --------------------------------------------------------
#	Fichier ["XLSX_historiques_salaires"]["fichierXLSX"]	[1 - Revue annuelle - Salaires contractuels.csv]
#   Table historiques_salaires
#	$script:HISTO    = @{}
#	$script:BDDHisto = @{}
# --------------------------------------------------------

function Query_XLSX_historiques_salaires {
	$XLSXfile        = $script:cfg["XLSX_historiques_salaires"]["fichierXLSX"]
	$headerstartline = $script:cfg["XLSX_historiques_salaires"]["HEADERstartline"] - 1
	$sheetname       = $script:cfg["XLSX_historiques_salaires"]["SheetName"]
	$datecol         = @("DATE EFFET")

	LOG "Query_XLSX_historiques_salaires" "Chargement du fichier $XLSXfile" -CRLF
	$script:XLSXHisto = Invoke-ExcelQuery -filePath $XLSXfile -sqlQuery "SELECT * FROM [$sheetname$]" -functionName "Query_XLSX_historiques_salaires" -frmtdateOUT $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"] -HEADERstartline 3 -key "_index" -datecol $datecol -dateLocale "FR" -ConvertToHashtable
}
function Compute_historiques_salaires {
	# Mappage des colonnes CSV vers les colonnes SQL. clé : utilisateur_id, date
	$script:HISTO = @{}
	foreach ( $row in $script:XLSXHisto.Keys ) {
		$zeromatricule = $script:XLSXHisto[$row]["Matricule"]
		$matricule = $zeromatricule  -replace '^0+', ''

		$date = $script:XLSXHisto[$row]["DATE EFFET"]
		if ( $script:BDDuser.ContainsKey($matricule) ) {
			if ( -not $script:HISTO.ContainsKey($matricule) ) {
				$script:HISTO[$matricule] = @{}
			}
			if ( -not $script:HISTO[$matricule].ContainsKey($date) ) {
				$script:HISTO[$matricule][$date] = @{}
			}
			$script:HISTO[$matricule][$date]["utilisateur_id"] = $matricule
			$script:HISTO[$matricule][$date]["date"]           = $date
			$script:HISTO[$matricule][$date]["montant"]        = $script:XLSXHisto[$row]["SALAIRE ANNUEL"]

		} else {
			Add-Exclusion $matricule "XLSX_historiques_salaires" "non trouvé dans la table utilisateurs" -USER
		}
	}
}
function Query_BDD_historiques_salaires {
	$script:BDDHisto = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Postgre_Review"]["tablehistoriques_salaires"] -functionName "Query_BDD_historiques_salaires" -keyColumns @("utilisateur_id","date") -targetVariable $script:BDDHisto -UseFrmtDateOUT
}
function Update_BDD_historiques_salaires {
	$table       = $script:cfg["SQL_Postgre_Review"]["tablehistoriques_salaires"]
	$keycolumns  = @("utilisateur_id","date")

	Update_BDDTable  $script:HISTO $script:BDDHisto $keycolumns $table "Update_BDD_historiques_salaires" { Query_BDD_historiques_salaires }
}

# --------------------------------------------------------
#               Utilitaires BDD
# --------------------------------------------------------

# Fonction utilitaire pour obtenir les paramètres de connexion BDD
function Get-BDDConnectionParams {
    return @{
        server      = $script:cfg["SQL_Postgre_Review"]["server"]
        database    = $script:cfg["SQL_Postgre_Review"]["database"]
        login       = $script:cfg["SQL_Postgre_Review"]["login"]
        password    = Encode $script:cfg["SQL_Postgre_Review"]["password"]
        datefrmtout = $script:cfg["SQL_Postgre_Review"]["frmtdateIN"]
    }
}
# Fonction utilitaire pour valider si un utilisateur existe
function Test-UserExists {
    param( [string]$matricule, [string]$zeromatricule, [string]$functionName )
    
    if ( $script:BDDuser.ContainsKey($matricule) -eq $false ) {
        Add-Exclusion $matricule "BDDuser" "N'existe pas dans la table [$($script:cfg['SQL_Postgre_Review']['tableUtilisateurs'])]" -USER
        return $false
    }
    return $true
}
# Fonction utilitaire pour effectuer une requête BDD standard
function Query_BDDTable {
    param(
        [string]$tableName,
        [string]$functionName,
        [array]$keyColumns,
        [hashtable]$targetVariable,
        [switch]$UseFrmtDateOUT
    )
    
    $params = Get-BDDConnectionParams
    
    LOG $functionName "Chargement de la table [$tableName] en memoire" -CRLF
    
    # Vider la hashtable cible
    $targetVariable.Clear()
    
    # Paramètres pour QueryTable
    $queryParams = @{
        server = $params.server
        database = $params.database
        table = $tableName
        login = $params.login
        password = $params.password
        keycolumns = $keyColumns
    }
    
    # Ajouter le format de date si demandé
    if ($UseFrmtDateOUT) {
        $queryParams.frmtdateOUT = $script:cfg["SQL_Postgre_Review"]["frmtdateOUT"]
    }
    
    # Exécuter la requête et affecter le résultat
    $result = QueryTable @queryParams
    
    # Copier le résultat dans la variable cible
    foreach ($key in $result.Keys) {
        $targetVariable[$key] = $result[$key]
    }
}
# Fonction utilitaire pour effectuer une mise à jour BDD standard
function Update_BDDTable {
    param(
        [hashtable]$sourceData,
        [hashtable]$targetData,
        [array]$keyColumns,
        [string]$tableName,
        [string]$functionName,
        [scriptblock]$reloadFunction
    )
    
    $params = Get-BDDConnectionParams
    
    LOG $functionName "Update de la table $tableName" -CRLF
    
    UpdateTable $sourceData $targetData $keyColumns $params.server $params.database $tableName $params.login $params.password $script:cfg["start"]["ApplyUpdate"]
    
    # Recharger les modifs en memoire
    if ($reloadFunction) {
        & $reloadFunction
    }
}

# --------------------------------------------------------
#               Utilitaires
# --------------------------------------------------------

function IsNullPGS {
	Param ( [string]$str )

	if ( $str.Trim() -eq "-" ) { return $true }
	if ( [string]::IsNullOrWhiteSpace($str) ) { return $true }
	return $false
}
function IsNotNullPGS {
	Param ( [string]$str )
	return -not (IsNullPGS $str)
}
function Annee_pleine {
	param ( [datetime]$base, [datetime]$entree )
	# Calcul de la différence brute en années
	$year = $base.Year - $entree.Year

	# Ajustement si le mois/jour d'entrée n'est pas encore atteint dans l'année de référence
	if ($base.Month -lt $entree.Month -or ($base.Month -eq $entree.Month -and $base.Day -lt $entree.Day) ) { $year-- }
	return [math]::Max(0, $year)
}
function Add-Exclusion {
	param ( [string]$matricule, [string]$source, [string]$raison, [switch]$USER )

	$m = $matricule.Trim()
	$src = $source.PadRight(23, ' ')
	if ( [string]::IsNullOrWhiteSpace($m)) { 
		$fn = (Get-PSCallStack)[1].FunctionName
		$m = "Vide    "
		INA $fn "xxxx3 : Matricule [$m] : $src : $raison"
	} else {
		if ([int]::TryParse($m, [ref]$null)) {
			if ( -not ($script:EXCLUS.ContainsKey($m)) ) {
				$fn = (Get-PSCallStack)[1].FunctionName
				$zeromatricule = $m.PadLeft(8, '0')
				INA $fn "xxxx2 : Matricule [$zeromatricule] : $src : $raison"
				if ( $USER ) { $script:EXCLUS[$m] = $raison }
			}
		} else {
			$fn = (Get-PSCallStack)[1].FunctionName
			$m = if ($m.Length -gt 8) { $m.Substring(0,8) } else { $m.PadRight(8) }
			INA $fn "xxxx1 : Matricule [$m] : $src : $raison"
		}
	}
}

# --------------------------------------------------------
#               Main
# --------------------------------------------------------

$script:cfgFile = "$PSScriptRoot\AnnualReview.ini"

# Chargement des modules
. "$PSScriptRoot\Modules\Ini.ps1"        > $null 
. "$PSScriptRoot\Modules\Log.ps1"        > $null 
. "$PSScriptRoot\Modules\Encode.ps1"     > $null 
. "$PSScriptRoot\Modules\Csv.ps1"        > $null 
. "$PSScriptRoot\Modules\XLSX.ps1"       > $null 
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null
. "$PSScriptRoot\Modules\SendEmail.ps1"  > $null 

LoadIni

SetConsoleToUFT8

Add-Type -AssemblyName System.Web
Add-Type -Path $script:cfg["SQL_Postgre_Review"]["microsoftExt"]
Add-Type -Path $script:cfg["SQL_Postgre_Review"]["pathdll"]

. "$PSScriptRoot\Modules\SQL - Transaction.ps1" > $null
if ($script:cfg["start"]["TransacSQL"] -eq "AllInOne" ) {
	. "$PSScriptRoot\Modules\PostgreSQL - TransactionAllInOne.ps1" > $null
} else {
	. "$PSScriptRoot\Modules\PostgreSQL - TransactionOneByOne.ps1" > $null
}

Query_CSV_Salaries
	Compute_USER

Query_XLS_Salaries-complement
	Compute_ListeBM

Query_CSV_Salaires-primes
	Compute_Profil_secondaire

# Utilisateur
Query_BDD_Utilisateurs
	Compute_Managers
	Compute_Administrator
	Prepare_two_pass
Update_BDD_utilisateurs

# Remuneration
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
	if ( $script:cfg["Options"]["Compute_FLOTTE_AUTO"] -eq "yes" ) {
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

# Projets-liste
Query_XLS_Projets_liste
	Compute_Projets_liste
	Compute_Fake_Projets
Query_BDD_projets
Update_BDD_projets

# Projets_affectation
Query_XLS_Projets_affectation
	Compute_Projets_affectation
Query_BDD_affectations
Update_BDD_affectations

Query_XLSX_historiques_salaires
	Compute_historiques_salaires
Query_BDD_historiques_salaires
Update_BDD_historiques_salaires

Log_Deltas

QUIT "Main" "Fin du process"

