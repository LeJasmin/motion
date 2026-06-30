<#
================================================================================
  Trier-Documents.ps1
  --------------------------------------------------------------------------
  Scanne ton dossier Documents, regroupe / trie / renomme tes fichiers
  par affaire judiciaire et par dossier :

      Affaire SSDH | Affaire Azzouz | These | Sante | Famille | Administratif

  >>> PAR DEFAUT : MODE SIMULATION. Le script NE DEPLACE RIEN. <<<
      Il affiche seulement le plan de classement et cree un fichier Excel
      recapitulatif (Plan-de-classement.csv) que tu peux ouvrir et verifier.

  Quand le plan te convient, relance avec l'option  -Executer  pour
  effectuer reellement le tri.

  --------------------------------------------------------------------------
  UTILISATION (clic droit > Executer avec PowerShell, ou dans une console) :

    # 1) Voir le plan, sans rien toucher (recommande en premier) :
    powershell -ExecutionPolicy Bypass -File .\Trier-Documents.ps1

    # 2) Une fois le plan verifie, faire le vrai tri :
    powershell -ExecutionPolicy Bypass -File .\Trier-Documents.ps1 -Executer

    # Variante encore plus prudente : COPIER au lieu de DEPLACER
    powershell -ExecutionPolicy Bypass -File .\Trier-Documents.ps1 -Executer -Copier

    # Ajouter la date (AAAA-MM-JJ) devant chaque nom de fichier :
    powershell -ExecutionPolicy Bypass -File .\Trier-Documents.ps1 -Executer -Renommer

  Apres un vrai tri, un script  Annuler-le-tri.ps1  est cree : il permet
  de TOUT remettre comme avant si besoin.
================================================================================
#>

param(
    # Dossier a scanner. Par defaut : ton dossier Documents.
    [string]$Source = "$env:USERPROFILE\Documents",

    # Ou ranger les dossiers tries. Par defaut : Documents\_Classement.
    [string]$Destination = "$env:USERPROFILE\Documents\_Classement",

    # Sans ce flag => SIMULATION (rien n'est deplace). Avec => vrai tri.
    [switch]$Executer,

    # Copie les fichiers au lieu de les deplacer (les originaux restent).
    [switch]$Copier,

    # Ajoute la date de modif (AAAA-MM-JJ_) devant chaque nom de fichier.
    [switch]$Renommer
)

# --- Reglages d'affichage (accents corrects dans la console) -----------------
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$ErrorActionPreference = "Stop"

# ============================================================================
#  1) DEFINITION DES DOSSIERS ET DES MOTS-CLES
#     -> Tu peux ajouter / retirer des mots-cles librement ci-dessous.
#     L'ordre compte : les affaires nominatives passent en premier.
# ============================================================================
$Categories = @(
    [pscustomobject]@{ Nom = "Affaire SSDH";    Mots = @("ssdh") },
    [pscustomobject]@{ Nom = "Affaire Azzouz";  Mots = @("azzouz") },
    [pscustomobject]@{ Nom = "These";           Mots = @(
        "these","thesis","doctorat","doctorant","manuscrit","soutenance",
        "chapitre","bibliographie","biblio","recherche","laboratoire",
        "article","publication","corpus","directeur de these","jury") },
    [pscustomobject]@{ Nom = "Sante";           Mots = @(
        "sante","medical","medecin","docteur","ordonnance","analyse",
        "radio","irm","scanner","labo","laboratoire d analyse","mutuelle",
        "cpam","ameli","secu","remboursement","hopital","clinique",
        "dentiste","ophtalmo","kine","vaccin","carnet de sante","arret de travail") },
    [pscustomobject]@{ Nom = "Famille";         Mots = @(
        "famille","enfant","bebe","mariage","naissance","bapteme","photo",
        "vacances","noel","anniversaire","livret de famille","scolaire",
        "ecole","college","lycee","bulletin","cantine","garde","pension") },
    [pscustomobject]@{ Nom = "Administratif";   Mots = @(
        "impot","impots","fisc","facture","edf","engie","eau","internet",
        "assurance","banque","releve","rib","contrat","bail","loyer",
        "quittance","caf","pole emploi","france travail","urssaf","attestation",
        "carte d identite","passeport","permis","amende","contravention",
        "courrier","lettre","devis","abonnement","cotisation","retraite") }
)

# Dossier ou vont les fichiers qui ne correspondent a AUCUN mot-cle
# (rien n'est jamais perdu).
$DossierNonClasse = "A trier manuellement"

# Extensions considerees comme "documents". Mets $true pour TOUT inclure.
$ToutInclure = $false
$ExtensionsDocs = @(
    ".pdf",".doc",".docx",".odt",".rtf",".txt",".md",
    ".xls",".xlsx",".ods",".csv",
    ".ppt",".pptx",".odp",
    ".jpg",".jpeg",".png",".gif",".tif",".tiff",".heic",
    ".eml",".msg",".zip"
)

# ============================================================================
#  2) FONCTIONS UTILITAIRES
# ============================================================================

# Enleve les accents pour comparer (e = e accent, etc.)
function Remove-Diacritics {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $norm = $Text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $norm.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString()
}

# Normalise un texte pour la recherche de mots-cles.
function Normaliser {
    param([string]$Text)
    $t = Remove-Diacritics $Text
    $t = $t.ToLower()
    # Remplace tout ce qui n'est pas lettre/chiffre par un espace
    $t = [Regex]::Replace($t, "[^a-z0-9]+", " ")
    return " " + $t.Trim() + " "
}

# Trouve la categorie d'un fichier d'apres son chemin (dossiers + nom).
function Trouver-Categorie {
    param([string]$CheminRelatif)
    $hay = Normaliser $CheminRelatif
    foreach ($cat in $Categories) {
        foreach ($mot in $cat.Mots) {
            $m = Normaliser $mot
            if ($hay.Contains($m)) {
                return $cat.Nom
            }
        }
    }
    return $DossierNonClasse
}

# Donne un chemin de destination unique (gere les doublons de nom).
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

# ============================================================================
#  3) PREPARATION
# ============================================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
if ($Executer) {
    Write-Host "  MODE REEL : les fichiers vont etre $([string]::Format('{0}', $(if($Copier){'COPIES'}else{'DEPLACES'})))." -ForegroundColor Yellow
} else {
    Write-Host "  MODE SIMULATION : aucun fichier ne sera deplace." -ForegroundColor Green
}
Write-Host "  Source      : $Source"
Write-Host "  Destination : $Destination"
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path -LiteralPath $Source)) {
    Write-Host "ERREUR : le dossier source n'existe pas : $Source" -ForegroundColor Red
    Write-Host "Relance le script en precisant le bon chemin, par ex. :" -ForegroundColor Red
    Write-Host '  .\Trier-Documents.ps1 -Source "C:\Users\TonNom\Documents"' -ForegroundColor Red
    return
}

# On evite de re-scanner le dossier de classement lui-meme.
$DestFull = [IO.Path]::GetFullPath($Destination)

Write-Host "Scan en cours..." -ForegroundColor Gray
$fichiers = Get-ChildItem -LiteralPath $Source -File -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object {
        # Ignore le dossier de destination
        (-not $_.FullName.StartsWith($DestFull, [StringComparison]::OrdinalIgnoreCase)) -and
        # Ignore les fichiers caches/systeme
        (-not ($_.Attributes -band [IO.FileAttributes]::Hidden)) -and
        (-not ($_.Attributes -band [IO.FileAttributes]::System)) -and
        # Filtre par extension (sauf si $ToutInclure)
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
    $categorie = Trouver-Categorie $rel

    $nouveauNom = $f.Name
    if ($Renommer) {
        $date = $f.LastWriteTime.ToString("yyyy-MM-dd")
        # Evite de doubler la date si elle est deja en debut de nom
        if ($f.Name -notmatch '^\d{4}-\d{2}-\d{2}') {
            $nouveauNom = "{0}_{1}" -f $date, $f.Name
        }
    }

    $dossierCible = Join-Path $Destination $categorie
    $cible = Chemin-Unique -Dossier $dossierCible -NomFichier $nouveauNom
    $script:CiblesPrevues += $cible

    $plan.Add([pscustomobject]@{
        Categorie   = $categorie
        Fichier     = $f.Name
        NouveauNom  = [IO.Path]::GetFileName($cible)
        Source      = $f.FullName
        Destination = $cible
        TailleKo    = [math]::Round($f.Length / 1KB, 1)
    })
}

# ============================================================================
#  5) AFFICHAGE DU RECAPITULATIF
# ============================================================================
Write-Host "RECAPITULATIF PAR DOSSIER :" -ForegroundColor Cyan
Write-Host "----------------------------------------------------------------"
$ordre = ($Categories.Nom + $DossierNonClasse)
foreach ($nom in $ordre) {
    $n = ($plan | Where-Object { $_.Categorie -eq $nom }).Count
    if ($n -gt 0) {
        $couleur = if ($nom -eq $DossierNonClasse) { "Yellow" } else { "White" }
        Write-Host ("  {0,-22} : {1,4} fichier(s)" -f $nom, $n) -ForegroundColor $couleur
    }
}
Write-Host "----------------------------------------------------------------"
Write-Host ("  {0,-22} : {1,4} fichier(s)" -f "TOTAL", $plan.Count) -ForegroundColor Cyan
Write-Host ""

# Toujours ecrire le plan detaille dans un CSV (ouvrable dans Excel).
if (-not (Test-Path -LiteralPath $Destination)) {
    if ($Executer) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
}
$dossierRapport = if (Test-Path -LiteralPath $Destination) { $Destination } else { $Source }
$csvPath = Join-Path $dossierRapport "Plan-de-classement.csv"
try {
    $plan | Sort-Object Categorie, Fichier |
        Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Plan detaille enregistre ici (ouvre-le dans Excel pour verifier) :" -ForegroundColor Green
    Write-Host "  $csvPath" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "(Impossible d'ecrire le CSV : $($_.Exception.Message))" -ForegroundColor Yellow
}

# ============================================================================
#  6) EXECUTION (uniquement si -Executer)
# ============================================================================
if (-not $Executer) {
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "  SIMULATION TERMINEE - aucun fichier n'a ete deplace." -ForegroundColor Green
    Write-Host "  Verifie le fichier Plan-de-classement.csv ci-dessus." -ForegroundColor Green
    Write-Host ""
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

# Script d'annulation (remet tout comme avant) - seulement si on a DEPLACE.
if (-not $Copier -and $journal.Count -gt 0) {
    $undoPath = Join-Path $Destination "Annuler-le-tri.ps1"
    $lignes = New-Object System.Collections.Generic.List[string]
    $lignes.Add('# Annule le tri : remet chaque fichier a son emplacement d''origine.')
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
