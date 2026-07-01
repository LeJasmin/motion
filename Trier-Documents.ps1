<#
================================================================================
  Trier-Documents.ps1
  --------------------------------------------------------------------------
  Scanne ton dossier Documents, regroupe / trie / RENOMME tes fichiers
  par affaire judiciaire et par dossier :

      Affaire SSDH | Affaire Azzouz | These | Sante | Famille | Administratif

  CONVENTION DE NOMMAGE (d'apres ton systeme "Second Cerveau") :
      Affaire/Dossier - contenu du document - date (si necessaire)
      ex :  Affaire SSDH - Conclusions - 2024-06-20.pdf
            These - Chapitre 2 redaction.docx

  STRUCTURE DE L'AFFAIRE SSDH (d'apres ta fiche master, 3 volets) :
      Affaire SSDH\
          1. Volet juridique
          2. Volet logement social
          3. Volet sinistres
          0. A classer            (pieces non reconnues)

  >>> PAR DEFAUT : MODE SIMULATION. Le script NE DEPLACE RIEN. <<<
      Il affiche le plan de classement + les nouveaux noms, et cree un
      fichier Excel recapitulatif (Plan-de-classement.csv) a verifier.

  --------------------------------------------------------------------------
  UTILISATION (clic droit > Executer avec PowerShell, ou en console) :

    # 1) Voir le plan + les nouveaux noms, sans rien toucher :
    powershell -ExecutionPolicy Bypass -File .\Trier-Documents.ps1

    # 2) Une fois verifie, faire le vrai tri (deplace + renomme) :
    powershell -ExecutionPolicy Bypass -File .\Trier-Documents.ps1 -Executer

  OPTIONS :
    -Copier        Copie au lieu de deplacer (les originaux restent en place)
    -GarderNoms    Range les fichiers SANS les renommer (garde le nom d'origine)
    -AvecDate      Ajoute toujours la date AAAA-MM-JJ (sinon seulement si utile)
    -SansSousDossiers   Met l'affaire SSDH a plat (pas de 3 volets)

  Apres un vrai tri, un script  Annuler-le-tri.ps1  est cree pour tout
  remettre comme avant si besoin.
================================================================================
#>

param(
    [string]$Source = "$env:USERPROFILE\Documents",
    [string]$Destination = "$env:USERPROFILE\Documents\_Classement",
    [switch]$Executer,
    [switch]$Copier,
    [switch]$GarderNoms,
    [switch]$AvecDate,
    [switch]$SansSousDossiers
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$ErrorActionPreference = "Stop"

# ============================================================================
#  1) DOSSIERS, MOTS-CLES, ET SOUS-DOSSIERS
#     -> Tu peux modifier librement les listes de mots-cles ci-dessous.
#     L'ordre compte : les affaires nominatives passent en premier.
# ============================================================================
$Categories = @(

    [pscustomobject]@{
        Nom  = "Affaire SSDH"
        Mots = @(
            "ssdh","seine saint denis habitat","seine-saint-denis habitat",
            "oph","lebcir vs","jcp ssdh","lr ssdh","insalubrite","amiante",
            "degat des eaux","degats des eaux","fsl","plan d apurement",
            "apurement","pacifica","espacil","tribunal de proximite",
            "rapport d expertise","reconnaissance","dde","mage")
        # Sous-dossiers = les 3 volets de ta fiche master (+ "0. A classer")
        SousDossiers = @(
            [pscustomobject]@{ Nom = "1. Volet juridique"; Mots = @(
                "juridique","conclusion","conclusions","tribunal","jcp","juge",
                "assignation","audience","plaidoirie","requete","avocat",
                "mise en demeure","lettre recommandee","lr ","jugement",
                "citation","huissier","signification","proximite","plainte","plaintes") },
            [pscustomobject]@{ Nom = "3. Volet sinistres"; Mots = @(
                "sinistre","degat des eaux","degats des eaux","expertise","expert",
                "assurance","pacifica","insalubrite","amiante","dde",
                "reconnaissance","catastrophe","gare") },
            [pscustomobject]@{ Nom = "2. Volet logement social"; Mots = @(
                "logement social","relogement","reloge","bail","dalo",
                "demande de logement","attribution","val d oise","val-d'oise",
                "commission","syplo","fsl","apurement","quittance","loyer") }
        )
    },

    [pscustomobject]@{
        Nom  = "Affaire Azzouz"
        # >>> A COMPLETER : ajoute ici les mots qui identifient cette affaire
        #     (nom de la partie adverse, n0 de RG, bailleur, etc.)
        Mots = @("azzouz")
        SousDossiers = @()
    },

    [pscustomobject]@{
        Nom  = "These"
        Mots = @(
            "these","thesis","doctorat","doctorant","doctorante","manuscrit",
            "soutenance","chapitre","bibliographie","biblio","zotero",
            "auto-ethno","autoethno","ethnographie","corpus","jury",
            "directrice de these","directeur de these","sic","colloque",
            "publication","article scientifique","cadre theorique","methodo")
        SousDossiers = @()
    },

    [pscustomobject]@{
        Nom  = "Sante"
        Mots = @(
            "sante","medical","medecin","docteur","ordonnance","analyse",
            "radio","irm","scanner","labo","mutuelle","cpam","ameli","secu",
            "remboursement","hopital","clinique","dentiste","ophtalmo","kine",
            "vaccin","carnet de sante","arret de travail","expertise medicale",
            "psychologue","therapie","nutrition")
        SousDossiers = @()
    },

    [pscustomobject]@{
        Nom  = "Famille"
        Mots = @(
            "famille","carmen","montessori","luna","enfant","bebe","mariage",
            "naissance","bapteme","photo","vacances","noel","anniversaire",
            "livret de famille","scolaire","ecole","creche","cantine","garde",
            "pension","maman","couple")
        SousDossiers = @()
    },

    [pscustomobject]@{
        Nom  = "Administratif"
        Mots = @(
            "impot","impots","fisc","facture","edf","engie","eau","internet",
            "assurance","banque","releve","rib","contrat","caf","pole emploi",
            "france travail","urssaf","attestation","carte d identite",
            "passeport","permis","amende","contravention","courrier","lettre",
            "devis","abonnement","cotisation","retraite","numerise")
        SousDossiers = @()
    }
)

# Dossier pour les fichiers qui ne correspondent a AUCUNE categorie.
$DossierNonClasse = "A trier manuellement"
# Sous-dossier par defaut quand une affaire a des volets mais que la piece
# n'en touche aucun.
$SousDossierDefaut = "0. A classer"

# Extensions considerees comme "documents". $true => inclure TOUS les fichiers.
$ToutInclure = $false
$ExtensionsDocs = @(
    ".pdf",".doc",".docx",".odt",".rtf",".txt",".md",
    ".xls",".xlsx",".ods",".csv",
    ".ppt",".pptx",".odp",
    ".jpg",".jpeg",".png",".gif",".tif",".tiff",".heic",
    ".eml",".msg",".zip"
)

# ============================================================================
#  2) FONCTIONS
# ============================================================================
function Remove-Diacritics {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $norm = $Text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $norm.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString()
}

function Normaliser {
    param([string]$Text)
    $t = Remove-Diacritics $Text
    $t = $t.ToLower()
    $t = [Regex]::Replace($t, "[^a-z0-9]+", " ")
    return " " + $t.Trim() + " "
}

# Renvoie $true si l'un des mots-cles est present dans le texte normalise.
function Contient-MotCle {
    param([string]$HayNormalise, [string[]]$Mots)
    foreach ($mot in $Mots) {
        if ($HayNormalise.Contains((Normaliser $mot))) { return $true }
    }
    return $false
}

# Trouve la categorie (et le sous-dossier eventuel) d'un fichier.
function Trouver-Classement {
    param([string]$CheminRelatif)
    $hay = Normaliser $CheminRelatif
    foreach ($cat in $Categories) {
        if (Contient-MotCle $hay $cat.Mots) {
            $sous = ""
            if (-not $SansSousDossiers -and $cat.SousDossiers -and $cat.SousDossiers.Count -gt 0) {
                $sous = $SousDossierDefaut
                foreach ($sd in $cat.SousDossiers) {
                    if (Contient-MotCle $hay $sd.Mots) { $sous = $sd.Nom; break }
                }
            }
            return [pscustomobject]@{ Categorie = $cat.Nom; SousDossier = $sous }
        }
    }
    return [pscustomobject]@{ Categorie = $DossierNonClasse; SousDossier = "" }
}

# Nettoie le nom d'origine pour en faire le "contenu" lisible.
function Nettoyer-Contenu {
    param([string]$NomSansExt)
    $t = $NomSansExt
    $t = $t -replace '_', ' '                      # underscores -> espaces
    $t = $t -replace '^\s*\d+[\.\)\-]\s*', ''      # enleve "5." "4) " en debut
    $t = $t -replace '\s*\(\d+\)\s*$', ''          # enleve " (1)" copie en fin
    $t = [Regex]::Replace($t, '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { $t = $NomSansExt.Trim() }
    return $t
}

# Detecte si le nom contient deja une date (annee 19xx/20xx).
function Contient-Date {
    param([string]$Texte)
    return ($Texte -match '(19|20)\d{2}')
}

function Chemin-Unique {
    param([string]$Dossier, [string]$NomFichier)
    $base = [IO.Path]::GetFileNameWithoutExtension($NomFichier)
    $ext  = [IO.Path]::GetExtension($NomFichier)
    $cible = Join-Path $Dossier $NomFichier
    $i = 2
    while ((Test-Path -LiteralPath $cible) -or ($script:CiblesPrevues -contains $cible)) {
        $cible = Join-Path $Dossier ("{0} ({1}){2}" -f $base, $i, $ext)
        $i++
    }
    return $cible
}

# Enleve les caracteres interdits dans un nom de fichier Windows.
function Nom-Valide {
    param([string]$Nom)
    $invalides = [IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = "[" + [Regex]::Escape($invalides) + "]"
    return ([Regex]::Replace($Nom, $pattern, ' ') -replace '\s+', ' ').Trim()
}

# ============================================================================
#  3) PREPARATION
# ============================================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
if ($Executer) {
    $verbe = if ($Copier) { "COPIES" } else { "DEPLACES" }
    Write-Host "  MODE REEL : les fichiers vont etre $verbe + renommes." -ForegroundColor Yellow
} else {
    Write-Host "  MODE SIMULATION : aucun fichier ne sera deplace." -ForegroundColor Green
}
Write-Host "  Source      : $Source"
Write-Host "  Destination : $Destination"
if ($GarderNoms) { Write-Host "  Renommage   : NON (-GarderNoms)" } else { Write-Host "  Renommage   : OUI  (Affaire - contenu - date si utile)" }
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path -LiteralPath $Source)) {
    Write-Host "ERREUR : le dossier source n'existe pas : $Source" -ForegroundColor Red
    Write-Host 'Relance en precisant le chemin, ex : .\Trier-Documents.ps1 -Source "C:\Users\TonNom\Documents"' -ForegroundColor Red
    return
}

$DestFull = [IO.Path]::GetFullPath($Destination)

Write-Host "Scan en cours..." -ForegroundColor Gray
$fichiers = Get-ChildItem -LiteralPath $Source -File -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object {
        (-not $_.FullName.StartsWith($DestFull, [StringComparison]::OrdinalIgnoreCase)) -and
        (-not ($_.Attributes -band [IO.FileAttributes]::Hidden)) -and
        (-not ($_.Attributes -band [IO.FileAttributes]::System)) -and
        ($ToutInclure -or ($ExtensionsDocs -contains $_.Extension.ToLower()))
    }

if (-not $fichiers -or $fichiers.Count -eq 0) {
    Write-Host "Aucun document trouve a trier dans : $Source" -ForegroundColor Yellow
    return
}
Write-Host ("$($fichiers.Count) document(s) trouve(s).") -ForegroundColor Gray
Write-Host ""

# ============================================================================
#  4) CALCUL DU PLAN
# ============================================================================
$script:CiblesPrevues = @()
$plan = New-Object System.Collections.Generic.List[object]

foreach ($f in $fichiers) {
    $rel = $f.FullName.Substring($Source.Length).TrimStart('\','/')
    $cl  = Trouver-Classement $rel

    # Dossier cible (categorie + sous-dossier eventuel)
    $dossierCible = Join-Path $Destination $cl.Categorie
    if ($cl.SousDossier) { $dossierCible = Join-Path $dossierCible $cl.SousDossier }

    # Nouveau nom
    if ($GarderNoms) {
        $nouveauNom = $f.Name
    } else {
        $ext      = $f.Extension
        $contenu  = Nettoyer-Contenu ([IO.Path]::GetFileNameWithoutExtension($f.Name))
        $prefixe  = $cl.Categorie
        $nom      = "$prefixe - $contenu"
        # Date : ajoutee si demandee (-AvecDate) et absente du nom
        if ($AvecDate -and -not (Contient-Date $contenu)) {
            $nom = "$nom - " + $f.LastWriteTime.ToString("yyyy-MM-dd")
        }
        $nouveauNom = (Nom-Valide $nom) + $ext
    }

    $cible = Chemin-Unique -Dossier $dossierCible -NomFichier $nouveauNom
    # Si collision et qu'on a renomme sans date : la date sert a desambiguer
    if (-not $GarderNoms -and -not $AvecDate -and ([IO.Path]::GetFileName($cible) -ne $nouveauNom)) {
        $ext     = $f.Extension
        $sansExt = [IO.Path]::GetFileNameWithoutExtension($nouveauNom)
        $avecDate = (Nom-Valide ($sansExt + " - " + $f.LastWriteTime.ToString("yyyy-MM-dd"))) + $ext
        $cible = Chemin-Unique -Dossier $dossierCible -NomFichier $avecDate
    }
    $script:CiblesPrevues += $cible

    $plan.Add([pscustomobject]@{
        Categorie   = $cl.Categorie
        SousDossier = $cl.SousDossier
        AncienNom   = $f.Name
        NouveauNom  = [IO.Path]::GetFileName($cible)
        Source      = $f.FullName
        Destination = $cible
        TailleKo    = [math]::Round($f.Length / 1KB, 1)
    })
}

# ============================================================================
#  5) RECAPITULATIF
# ============================================================================
Write-Host "RECAPITULATIF PAR DOSSIER :" -ForegroundColor Cyan
Write-Host "----------------------------------------------------------------"
$ordre = ($Categories.Nom + $DossierNonClasse)
foreach ($nom in $ordre) {
    $sousPlan = $plan | Where-Object { $_.Categorie -eq $nom }
    $n = ($sousPlan | Measure-Object).Count
    if ($n -gt 0) {
        $couleur = if ($nom -eq $DossierNonClasse) { "Yellow" } else { "White" }
        Write-Host ("  {0,-22} : {1,4} fichier(s)" -f $nom, $n) -ForegroundColor $couleur
        # Detail des sous-dossiers
        $sousNoms = $sousPlan | Where-Object { $_.SousDossier } | Select-Object -ExpandProperty SousDossier -Unique | Sort-Object
        foreach ($sn in $sousNoms) {
            $ns = ($sousPlan | Where-Object { $_.SousDossier -eq $sn } | Measure-Object).Count
            Write-Host ("      - {0,-18} : {1,4}" -f $sn, $ns) -ForegroundColor DarkGray
        }
    }
}
Write-Host "----------------------------------------------------------------"
Write-Host ("  {0,-22} : {1,4} fichier(s)" -f "TOTAL", $plan.Count) -ForegroundColor Cyan
Write-Host ""

# Apercu de quelques renommages
if (-not $GarderNoms) {
    Write-Host "EXEMPLES DE RENOMMAGE (avant  ->  apres) :" -ForegroundColor Cyan
    $plan | Where-Object { $_.Categorie -ne $DossierNonClasse } | Select-Object -First 8 | ForEach-Object {
        Write-Host ("  {0}" -f $_.AncienNom) -ForegroundColor DarkGray
        Write-Host ("    -> {0}\{1}\{2}" -f $_.Categorie, $_.SousDossier, $_.NouveauNom) -ForegroundColor Gray
    }
    Write-Host ""
}

# CSV detaille
$dossierRapport = if (Test-Path -LiteralPath $Destination) { $Destination } else { $Source }
$csvPath = Join-Path $dossierRapport "Plan-de-classement.csv"
try {
    $plan | Sort-Object Categorie, SousDossier, NouveauNom |
        Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Plan detaille (ouvre-le dans Excel pour tout verifier) :" -ForegroundColor Green
    Write-Host "  $csvPath" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "(Impossible d'ecrire le CSV : $($_.Exception.Message))" -ForegroundColor Yellow
}

# ============================================================================
#  6) EXECUTION (uniquement avec -Executer)
# ============================================================================
if (-not $Executer) {
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "  SIMULATION TERMINEE - aucun fichier n'a ete deplace." -ForegroundColor Green
    Write-Host "  Verifie le fichier Plan-de-classement.csv ci-dessus." -ForegroundColor Green
    Write-Host "  Quand tout te convient, relance avec -Executer :" -ForegroundColor Green
    Write-Host "    .\Trier-Documents.ps1 -Executer" -ForegroundColor White
    Write-Host "================================================================" -ForegroundColor Green
    return
}

Write-Host "Tri en cours..." -ForegroundColor Gray
$journal = New-Object System.Collections.Generic.List[object]
$ok = 0; $erreurs = 0
foreach ($item in $plan) {
    try {
        $dossierCible = Split-Path -Parent $item.Destination
        if (-not (Test-Path -LiteralPath $dossierCible)) {
            New-Item -ItemType Directory -Path $dossierCible -Force | Out-Null
        }
        if ($Copier) {
            Copy-Item -LiteralPath $item.Source -Destination $item.Destination -Force
        } else {
            Move-Item -LiteralPath $item.Source -Destination $item.Destination -Force
        }
        $journal.Add([pscustomobject]@{ De = $item.Source; Vers = $item.Destination })
        $ok++
    } catch {
        Write-Host ("  ECHEC : {0}  ->  {1}" -f $item.Source, $_.Exception.Message) -ForegroundColor Red
        $erreurs++
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ("  TERMINE : {0} fichier(s) traite(s), {1} erreur(s)." -f $ok, $erreurs) -ForegroundColor Cyan

if (-not $Copier -and $journal.Count -gt 0) {
    $undoPath = Join-Path $Destination "Annuler-le-tri.ps1"
    $lignes = New-Object System.Collections.Generic.List[string]
    $lignes.Add('# Annule le tri : remet chaque fichier a son emplacement et nom d''origine.')
    $lignes.Add('$ErrorActionPreference = "Continue"')
    foreach ($j in $journal) {
        $src = $j.Vers.Replace("'", "''")
        $dst = $j.De.Replace("'", "''")
        $lignes.Add("`$d = Split-Path -Parent '$dst'; if(-not(Test-Path -LiteralPath `$d)){New-Item -ItemType Directory -Path `$d -Force | Out-Null}")
        $lignes.Add("Move-Item -LiteralPath '$src' -Destination '$dst' -Force")
    }
    Set-Content -LiteralPath $undoPath -Value $lignes -Encoding UTF8
    Write-Host "  Pour TOUT annuler si besoin, lance :" -ForegroundColor Yellow
    Write-Host "    $undoPath" -ForegroundColor Yellow
}
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
