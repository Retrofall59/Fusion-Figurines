# ============================================================================
# Fusion-Figurines-GUI.ps1
# Detecte des groupes de pieces .stl (un sous-dossier = une figurine a plusieurs
# parties), les fusionne + oriente via la ligne de commande de Bambu Studio
# (--orient --export-3mf), puis reexporte le resultat en .stl pret pour
# Coloriser-STL. Apercu isometrique statique (pas interactif) des pieces
# d'entree et du resultat de sortie.
#
# 26/10/2026 : premiere version. Le tout premier vrai lancement sert AUSSI de
# test pour voir si --orient fait reellement le travail - a verifier avec
# Tomyn sur ses vrais fichiers avant de considerer ce pipeline fiable.
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
# 29/10/2026 : pour [Microsoft.VisualBasic.Interaction]::InputBox (demander un nouveau nom a
# Tomyn avant de fusionner - voir Lancer-Traitement, cas "plusieurs figurines dans le meme
# dossier source").
Add-Type -AssemblyName Microsoft.VisualBasic

# 27/10/2026 : demande de Tomyn - un mode clair en plus du mode sombre existant (bascule au
# clic, voir le bouton pres du haut de la fenetre). $THEME reste l'objet UNIQUE utilise partout
# dans le script (controles, apercus...) - on ne le remplace jamais par un AUTRE objet, on
# MUTE ses valeurs (meme pattern que $script:EtatTraitement plus bas : muter, pas reassigner)
# a partir de l'un des deux palettes ci-dessous, pour que tout ce qui l'a deja capture (y
# compris les closures des apercus asynchrones) voie la bascule sans avoir besoin d'etre
# recree. Rafraichir-ControlesTheme (definie plus bas, apres la creation de $form) repeint
# ensuite les controles existants avec les nouvelles valeurs.
$script:PaletteSombre = @{
    FondApp            = [System.Drawing.Color]::FromArgb(30, 30, 34)
    FondPanneau        = [System.Drawing.Color]::FromArgb(40, 40, 46)
    FondChamp          = [System.Drawing.Color]::FromArgb(50, 50, 58)
    Texte              = [System.Drawing.Color]::FromArgb(225, 225, 230)
    TexteAttenue       = [System.Drawing.Color]::FromArgb(160, 160, 168)
    AccentPrincipal    = [System.Drawing.Color]::FromArgb(90, 140, 230)
    AccentPrincipalTxt = [System.Drawing.Color]::White
    Bordure            = [System.Drawing.Color]::FromArgb(60, 60, 68)
}
# 27/10/2026 : mode clair "ivoire" demande par Tomyn - pas blanc pur (trop clinique/dur pour
# les yeux), un blanc casse/ivoire chaud avec un texte presque noir (pas un vrai noir, un peu
# plus doux) et un accent bleu legerement fonce pour garder un bon contraste sur fond clair.
$script:PaletteClaire = @{
    # 27/10/2026 : revenu pres de la valeur d'origine (238,231,216 -> 244,239,228) - assombrir
    # le fond pour faire ressortir les champs finissait par rapprocher le mode clair du mode
    # sombre (remarque de Tomyn, a raison). Le contraste des champs/listes/apercus est
    # maintenant regle par une VRAIE bordure (voir plus bas, BorderStyle restaure sur ces
    # controles) plutot qu'en assombrissant tout le fond.
    FondApp            = [System.Drawing.Color]::FromArgb(244, 239, 228)
    # 27/10/2026 : remonte (240,235,224 -> 223,215,198) - Tomyn + moi avons remarque que les
    # boutons neutres ("Tout reinclure", "Nettoyer dossiers temp", "Parcourir...") se
    # distinguaient a peine du fond general en mode clair (l'ecart etait bien plus marque en
    # sombre). Un ecart plus franc ici regle ca sans toucher au mode sombre.
    FondPanneau        = [System.Drawing.Color]::FromArgb(223, 215, 198)
    FondChamp          = [System.Drawing.Color]::FromArgb(255, 255, 255)
    Texte              = [System.Drawing.Color]::FromArgb(40, 38, 34)
    TexteAttenue       = [System.Drawing.Color]::FromArgb(120, 113, 101)
    AccentPrincipal    = [System.Drawing.Color]::FromArgb(55, 105, 195)
    AccentPrincipalTxt = [System.Drawing.Color]::White
    Bordure            = [System.Drawing.Color]::FromArgb(190, 178, 155)
}
$THEME = @{}
foreach ($cle in $script:PaletteSombre.Keys) { $THEME[$cle] = $script:PaletteSombre[$cle] }

function Formater-DureeMinSec([int]$secondesTotal) {
    # 26/10/2026 : demande de Tomyn - au-dela de 60s, afficher "X min YY s" plutot qu'un
    # nombre de secondes brut qui grossit sans fin (ex: "61 s" au lieu de "1 min 01 s").
    if ($secondesTotal -lt 60) { return "$secondesTotal s" }
    $min = [int]($secondesTotal / 60)
    $sec = $secondesTotal % 60
    return "{0} min {1:D2} s" -f $min, $sec
}

function Arrondir-Bouton([System.Windows.Forms.Control]$bouton, [int]$rayon = 8) {
    $largeur = $bouton.Width; $hauteur = $bouton.Height
    if ($largeur -le 0 -or $hauteur -le 0) { return }
    $chemin = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $rayon * 2
    $chemin.AddArc(0, 0, $d, $d, 180, 90)
    $chemin.AddArc($largeur - $d, 0, $d, $d, 270, 90)
    $chemin.AddArc($largeur - $d, $hauteur - $d, $d, $d, 0, 90)
    $chemin.AddArc(0, $hauteur - $d, $d, $d, 90, 90)
    $chemin.CloseFigure()
    $bouton.Region = New-Object System.Drawing.Region($chemin)
}

# ============================================================================
# CONFIG
# ============================================================================
$script:CheminConfig = Join-Path $PSScriptRoot "config-fusion-figurines.json"

# 29/10/2026 : demande de Tomyn ("idee d'amelioration") - jusqu'ici Ajouter-Journal n'ecrivait
# QUE dans la zone de texte de l'interface (RichTextBox en memoire) : un crash/gel de l'appli
# pendant un gros lot (antivirus, Bambu Studio qui part en vrille, Windows qui redemarre tout
# seul...) effacait TOUT l'historique avec elle, et le rapport CSV de fin de lot n'existe que
# s'il reste de bout en bout. Un fichier .log sur disque survit a un crash et evite en plus de
# devoir copier-coller la zone de texte a la main pour diagnostiquer un bug. Un fichier par
# lancement de l'appli (horodate), dans un sous-dossier "logs" a cote du script.
$script:DossierLogs = Join-Path $PSScriptRoot "logs"
try {
    if (-not (Test-Path $script:DossierLogs)) { New-Item -ItemType Directory -Path $script:DossierLogs -Force | Out-Null }
    $script:CheminJournalFichier = Join-Path $script:DossierLogs "journal-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
} catch {
    # Dossier logs inaccessible (droits, disque reseau en lecture seule...) : on degrade
    # silencieusement vers la zone de texte seule plutot que de faire planter l'appli au
    # demarrage pour ca.
    $script:CheminJournalFichier = $null
}

function Obtenir-ConfigParDefaut {
    return [PSCustomObject]@{
        CheminBambuExe = "C:\Program Files\Bambu Studio\bambu-studio.exe"
        DossierRacine  = ""
        DossierSortie  = ""
        # 27/10/2026 : "Sombre" ou "Claire" - memorise le choix de theme de Tomyn d'une
        # session a l'autre.
        Theme          = "Sombre"
        # 27/10/2026 : demande de Tomyn - le seuil de detection "piece pas vraiment en
        # contact" (voir Test-PiecesEnContact) etait fige a 0.3mm en dur. Certaines figurines
        # ont des pieces volontairement tres rapprochees (aimants, charnieres, encoches) qui
        # declenchaient l'avertissement a tort. Reglable depuis l'interface desormais, 0.3mm
        # reste la valeur par defaut (largeur de buse habituelle ~0.4mm, voir le commentaire
        # original a Test-PiecesEnContact).
        SeuilContactMm = 0.3
        # 27/10/2026 : demande de Tomyn - "reprise intelligente" (voir Lancer-Traitement),
        # activee par defaut.
        SauterDejaTraitees = $true
    }
}

function Charger-Config {
    $defaut = Obtenir-ConfigParDefaut
    if (Test-Path $script:CheminConfig) {
        try {
            $charge = Get-Content $script:CheminConfig -Raw | ConvertFrom-Json
            foreach ($prop in $defaut.PSObject.Properties.Name) {
                if (-not ($charge.PSObject.Properties.Name -contains $prop)) {
                    $charge | Add-Member -MemberType NoteProperty -Name $prop -Value $defaut.$prop
                }
            }
            return $charge
        } catch { return $defaut }
    }
    return $defaut
}

function Sauver-Config {
    $script:Config | ConvertTo-Json | Set-Content -Path $script:CheminConfig -Encoding UTF8
}

$script:Config = Charger-Config

# 27/10/2026 : applique le theme sauvegarde AVANT la creation du moindre controle (tous les
# controles lisent $THEME.XXX au moment de leur creation, plus bas dans le script) - pas besoin
# de repeindre quoi que ce soit ici, juste avoir les bonnes valeurs des le depart.
$script:PaletteActive = if ($script:Config.Theme -eq "Claire") { $script:PaletteClaire } else { $script:PaletteSombre }
foreach ($cle in $script:PaletteActive.Keys) { $THEME[$cle] = $script:PaletteActive[$cle] }

# ============================================================================
# LECTURE STL (binaire et ASCII) -> liste de triangles @{A=(x,y,z); B=...; C=...}
# ============================================================================
function Lire-TrianglesSTL([string]$cheminFichier) {
    $octets = [System.IO.File]::ReadAllBytes($cheminFichier)
    $triangles = [System.Collections.Generic.List[object]]::new()

    $estBinaire = $false
    if ($octets.Length -ge 84) {
        $nbTrianglesAnnonce = [System.BitConverter]::ToUInt32($octets, 80)
        $tailleAttendue = 84 + ($nbTrianglesAnnonce * 50)
        if ($tailleAttendue -eq $octets.Length) { $estBinaire = $true }
    }

    if ($estBinaire) {
        $pos = 84
        for ($i = 0; $i -lt $nbTrianglesAnnonce; $i++) {
            $pos += 12  # normale (ignoree, on la recalcule)
            $ax = [System.BitConverter]::ToSingle($octets, $pos);      $ay = [System.BitConverter]::ToSingle($octets, $pos + 4);  $az = [System.BitConverter]::ToSingle($octets, $pos + 8)
            $bx = [System.BitConverter]::ToSingle($octets, $pos + 12); $by = [System.BitConverter]::ToSingle($octets, $pos + 16); $bz = [System.BitConverter]::ToSingle($octets, $pos + 20)
            $cx = [System.BitConverter]::ToSingle($octets, $pos + 24); $cy = [System.BitConverter]::ToSingle($octets, $pos + 28); $cz = [System.BitConverter]::ToSingle($octets, $pos + 32)
            $pos += 36 + 2
            # 27/10/2026 : CORRECTIF PERF - [PSCustomObject] remplace par une simple Hashtable
            # (acces .A/.B/.C identique, voir benchmark sandbox : ~3x plus rapide a construire
            # en boucle sur de gros volumes - PSCustomObject passe par l'adaptateur de type
            # PowerShell/PSObject, une Hashtable non). Mesure reelle qui a motive ce correctif :
            # Lire-Maillage3MF (meme principe de construction par triangle) prenait 130.3 s sur
            # un maillage de 565366 triangles - voir le commentaire a Resoudre-Objet3MF.
            $triangles.Add(@{
                A = [double[]]@($ax, $ay, $az); B = [double[]]@($bx, $by, $bz); C = [double[]]@($cx, $cy, $cz)
            })
        }
    } else {
        # ASCII : on parse ligne par ligne, on cherche "vertex x y z"
        $lignes = [System.IO.File]::ReadAllLines($cheminFichier)
        $sommetsCourants = [System.Collections.Generic.List[double[]]]::new()
        foreach ($ligne in $lignes) {
            $l = $ligne.Trim()
            if ($l.StartsWith("vertex")) {
                $parties = $l -split '\s+'
                if ($parties.Count -ge 4) {
                    $sommetsCourants.Add([double[]]@([double]$parties[1], [double]$parties[2], [double]$parties[3]))
                }
            } elseif ($l.StartsWith("endfacet")) {
                if ($sommetsCourants.Count -eq 3) {
                    $triangles.Add(@{ A = $sommetsCourants[0]; B = $sommetsCourants[1]; C = $sommetsCourants[2] })
                }
                $sommetsCourants.Clear()
            }
        }
    }
    return $triangles
}

# ============================================================================
# LECTURE STL "ECHANTILLONNEE" - pour l'APERCU UNIQUEMENT (pas pour la fusion reelle).
# 26/10/2026 : meme en arriere-plan (hors thread UI), lire + afficher TOUS les
# triangles d'un maillage tres detaille (plusieurs centaines de milliers, voire
# millions) est beaucoup trop lent en PowerShell (langage interprete, pas fait pour
# ca). Comme l'apercu n'est qu'un controle visuel rapide (forme/orientation), on ne
# lit qu'un SOUS-ECHANTILLON regulier des triangles - directement au niveau des
# octets pour le format binaire (on ne decode meme pas les triangles sautes, donc le
# gain de vitesse est reel, pas juste au niveau de l'affichage).
# ============================================================================
function Lire-TrianglesSTLEchantillon([string]$cheminFichier, [int]$maxTriangles) {
    # Renvoie un objet { Triangles = <liste echantillonnee>; NombreOriginal = <total reel
    # dans le fichier> } - le nombre original sert ensuite a calculer de combien on doit
    # "gonfler" chaque triangle affiche pour compenser ceux qu'on a sautes (voir
    # Rendre-ApercuIsometrique / FacteurGonflage), sinon la surface semble trouee.
    #
    # 26/10/2026 : CORRECTIF MAJEUR DE VITESSE - cette fonction utilisait avant
    # [System.IO.File]::ReadAllBytes, qui charge le fichier ENTIER en memoire avant de
    # choisir quels triangles garder. Sur un gros fichier (plusieurs centaines de Mo), c'est
    # CETTE lecture complete du disque qui dominait le temps total (plusieurs minutes),
    # pas le traitement des triangles eux-memes - l'echantillonnage ne servait donc a rien
    # cote vitesse. Fix : on ouvre le fichier en FileStream et on se positionne (Seek)
    # directement sur l'octet de CHAQUE triangle garde, en ne lisant que les 50 octets de
    # ce triangle - le reste du fichier n'est JAMAIS lu, ni decode, ni meme touche.
    $triangles = [System.Collections.Generic.List[object]]::new()
    $nombreOriginal = 0

    $flux = [System.IO.File]::OpenRead($cheminFichier)
    try {
        $tailleFichier = $flux.Length

        # 26/10/2026 : CORRECTIF CRASH/BLOCAGE - l'ancienne detection binaire/ASCII reposait
        # UNIQUEMENT sur "la taille du fichier correspond EXACTEMENT a 84 + n*50 octets". Un
        # fichier binaire avec le moindre octet de trop ou de padding (ce qui arrive avec
        # certains exports/scans) faisait echouer ce test et le faisait tomber dans la
        # branche ASCII - qui lit alors du binaire arbitraire comme si c'etait du texte.
        # Comme un vrai triangle ASCII ("vertex ...", "endfacet") n'apparait jamais dans du
        # binaire, la boucle ASCII ne trouvait JAMAIS assez de triangles pour s'arreter au
        # budget demande, et scannait alors le fichier ENTIER ligne par ligne (potentiellement
        # des centaines de milliers/millions de "lignes" sur un gros fichier) - ca explique
        # exactement les symptomes observes : CPU moderee (lecture/allocation de chaines en
        # boucle, pas un vrai calcul), duree qui explose avec la taille du fichier source (et
        # non avec le budget de triangles demande), et un crash eventuel (pression memoire).
        #
        # Detection plus fiable : un fichier STL ASCII commence TOUJOURS par le mot "solid"
        # (c'est la seule signature fiable du format texte). Si ce n'est pas le cas, on le
        # traite TOUJOURS comme binaire (meme si la taille ne correspond pas exactement a la
        # formule), car scanner un gros fichier comme "texte" sans jamais trouver de triangle
        # est bien pire que de perdre quelques triangles en fin de fichier si le nombre
        # annonce dans l'en-tete est legerement incorrect.
        $debut5 = New-Object byte[] ([Math]::Min(5, $tailleFichier))
        if ($debut5.Length -gt 0) { [void]$flux.Read($debut5, 0, $debut5.Length) }
        $commenceParSolid = ([System.Text.Encoding]::ASCII.GetString($debut5).ToLowerInvariant() -eq "solid")

        $estBinaire = $true
        $nbTrianglesAnnonce = 0
        if ($tailleFichier -ge 84) {
            [void]$flux.Seek(0, [System.IO.SeekOrigin]::Begin)
            $entete = New-Object byte[] 84
            [void]$flux.Read($entete, 0, 84)
            $nbTrianglesAnnonce = [System.BitConverter]::ToUInt32($entete, 80)
        }
        if ($commenceParSolid) {
            # Signature ASCII sans ambiguite - mais on verifie quand meme la formule binaire
            # au cas improbable ou un fichier binaire aurait "solid" comme debut d'en-tete
            # (autorise par le format, meme si rare en pratique).
            $tailleAttendue = 84 + ([int64]$nbTrianglesAnnonce * 50)
            if ($tailleAttendue -ne $tailleFichier) { $estBinaire = $false }
        }

        if ($estBinaire) {
            $nombreOriginal = $nbTrianglesAnnonce
            $pas = [Math]::Max(1, [Math]::Floor($nbTrianglesAnnonce / [Math]::Max(1, $maxTriangles)))
            if ($pas -eq 1) {
                # 27/10/2026 : CORRECTIF PERF - quand on garde TOUS les triangles (cas de la
                # fusion locale : budget = nombre reel, donc $pas=1 toujours), faire un
                # Seek()+Read(50 octets) PAR TRIANGLE revient a des MILLIONS d'appels systeme/
                # methode individuels sur un gros maillage (1.46M triangles = 1.46M allers-
                # retours) - meme categorie de probleme que le correctif applique a
                # Ecrire-STLBinaire (voir son commentaire). Comme on lit de toute facon TOUT le
                # bloc de triangles de maniere contigue dans ce cas, un seul GROS Read() en
                # memoire est strictement equivalent et bien plus rapide.
                $tailleBloc = [int64]$nbTrianglesAnnonce * 50
                if (84 + $tailleBloc -gt $tailleFichier) { $tailleBloc = $tailleFichier - 84 }
                $tamponComplet = New-Object byte[] $tailleBloc
                [void]$flux.Seek(84, [System.IO.SeekOrigin]::Begin)
                $luTotal = 0
                while ($luTotal -lt $tailleBloc) {
                    $n = $flux.Read($tamponComplet, $luTotal, $tailleBloc - $luTotal)
                    if ($n -le 0) { break }
                    $luTotal += $n
                }
                $nbCompletsLus = [Math]::Floor($luTotal / 50)
                for ($i = 0; $i -lt $nbCompletsLus; $i++) {
                    $o = $i * 50 + 12
                    $ax = [System.BitConverter]::ToSingle($tamponComplet, $o);      $ay = [System.BitConverter]::ToSingle($tamponComplet, $o + 4);  $az = [System.BitConverter]::ToSingle($tamponComplet, $o + 8)
                    $bx = [System.BitConverter]::ToSingle($tamponComplet, $o + 12); $by = [System.BitConverter]::ToSingle($tamponComplet, $o + 16); $bz = [System.BitConverter]::ToSingle($tamponComplet, $o + 20)
                    $cx = [System.BitConverter]::ToSingle($tamponComplet, $o + 24); $cy = [System.BitConverter]::ToSingle($tamponComplet, $o + 28); $cz = [System.BitConverter]::ToSingle($tamponComplet, $o + 32)
                    # 27/10/2026 : CORRECTIF PERF - cette branche (lecture rapide en un seul
                    # bloc) construisait encore les triangles en PSCustomObject, oubliee lors du
                    # premier correctif Hashtable (qui n'avait touche que la branche
                    # d'echantillonnage juste en dessous). Or c'est justement CETTE branche qui
                    # est utilisee a chaque fusion locale (budget = tous les triangles, donc
                    # $pas=1 systematiquement) - donc l'ancien cout PSCustomObject s'appliquait
                    # toujours sur le chemin le plus frequent/le plus gros. Memes mesures/
                    # justification que l'autre correctif Hashtable (~3x plus rapide a construire).
                    $triangles.Add(@{
                        A = [double[]]@($ax, $ay, $az); B = [double[]]@($bx, $by, $bz); C = [double[]]@($cx, $cy, $cz)
                    })
                }
            } else {
                $tampon = New-Object byte[] 50
                for ($i = 0; $i -lt $nbTrianglesAnnonce; $i += $pas) {
                    $posOctet = 84 + ([int64]$i * 50)
                    # Securite : si le nombre de triangles annonce dans l'en-tete est errone
                    # (fichier tronque/mal forme), on ne depasse jamais la fin reelle du fichier.
                    if ($posOctet + 50 -gt $tailleFichier) { break }
                    [void]$flux.Seek($posOctet, [System.IO.SeekOrigin]::Begin)
                    $luOk = $flux.Read($tampon, 0, 50)
                    if ($luOk -lt 50) { break }
                    # Les 12 premiers octets du triangle (indices 0-11) sont la normale,
                    # ignoree/recalculee ailleurs - les sommets commencent a l'octet 12.
                    $ax = [System.BitConverter]::ToSingle($tampon, 12); $ay = [System.BitConverter]::ToSingle($tampon, 16); $az = [System.BitConverter]::ToSingle($tampon, 20)
                    $bx = [System.BitConverter]::ToSingle($tampon, 24); $by = [System.BitConverter]::ToSingle($tampon, 28); $bz = [System.BitConverter]::ToSingle($tampon, 32)
                    $cx = [System.BitConverter]::ToSingle($tampon, 36); $cy = [System.BitConverter]::ToSingle($tampon, 40); $cz = [System.BitConverter]::ToSingle($tampon, 44)
                    $triangles.Add(@{
                        A = [double[]]@($ax, $ay, $az); B = [double[]]@($bx, $by, $bz); C = [double[]]@($cx, $cy, $cz)
                    })
                }
            }
        } else {
            # ASCII confirme (signature "solid" en tete). Pas de saut possible facilement
            # (format texte en lignes de longueur variable) - on lit via un flux (pas tout le
            # fichier en memoire d'un coup) et on s'arrete des qu'on a assez de triangles.
            # Garde-fou supplementaire : meme en ASCII genuin, on ne scanne jamais plus de
            # 300 Mo du fichier (au-dela, on s'arrete avec ce qu'on a deja trouve) pour ne
            # jamais pouvoir reproduire le blocage "fichier entier scanne sans fin".
            [void]$flux.Seek(0, [System.IO.SeekOrigin]::Begin)
            $lecteur = New-Object System.IO.StreamReader($flux)
            $sommetsCourants = [System.Collections.Generic.List[double[]]]::new()
            $limiteOctets = 300MB
            while (-not $lecteur.EndOfStream -and $triangles.Count -lt $maxTriangles -and $flux.Position -lt $limiteOctets) {
                $ligne = $lecteur.ReadLine()
                $l = $ligne.Trim()
                if ($l.StartsWith("vertex")) {
                    $parties = $l -split '\s+'
                    if ($parties.Count -ge 4) {
                        $sommetsCourants.Add([double[]]@([double]$parties[1], [double]$parties[2], [double]$parties[3]))
                    }
                } elseif ($l.StartsWith("endfacet")) {
                    if ($sommetsCourants.Count -eq 3) {
                        $triangles.Add(@{ A = $sommetsCourants[0]; B = $sommetsCourants[1]; C = $sommetsCourants[2] })
                    }
                    $sommetsCourants.Clear()
                }
            }
            # ASCII : pas de total reel connu sans scanner tout le fichier (cout qu'on
            # evite ici) - on part du principe qu'on a tout lu (pas de gonflage applique
            # pour ce cas, rare dans ce pipeline).
            $nombreOriginal = $triangles.Count
        }
    } finally {
        $flux.Dispose()
    }
    return [PSCustomObject]@{ Triangles = $triangles; NombreOriginal = $nombreOriginal }
}

# ============================================================================
# LECTURE 3MF -> liste de triangles (avec resolution des transforms imbriques)
# ============================================================================
function Appliquer-Transform3MF([double[]]$point, [double[]]$m) {
    # m = 12 valeurs : M00 M01 M02 M10 M11 M12 M20 M21 M22 M30 M31 M32
    $x = $point[0]; $y = $point[1]; $z = $point[2]
    $nx = $x * $m[0] + $y * $m[3] + $z * $m[6] + $m[9]
    $ny = $x * $m[1] + $y * $m[4] + $z * $m[7] + $m[10]
    $nz = $x * $m[2] + $y * $m[5] + $z * $m[8] + $m[11]
    return [double[]]@($nx, $ny, $nz)
}

function Appliquer-Transform3MF-Triangles($triangles, [double[]]$m) {
    # 27/10/2026 : CORRECTIF PERF MAJEUR #2 - mesure reelle (chronos detailles ajoutes dans
    # Resoudre-Objet3MF) : sur une piece Bambu Studio typique (fichier .model EXTERNE +
    # transform de composant, confirme par diagnostic), appliquer la transformation via
    # Appliquer-Transform3MF (1 appel de fonction PAR SOMMET, donc 3 par triangle, chacun
    # recreant un [double[]]) prenait 10.2 s sur 200000 triangles synthetiques - coherent avec
    # les 10-80 s observees en reel sur 206104/565366 triangles. L'appel de fonction lui-meme
    # (meme minuscule) a un cout fixe non negligeable en PowerShell interprete, multiplie ici
    # par 3 x le nombre de triangles. Fix valide en bac a sable : UNE SEULE fonction qui prend
    # TOUTE la liste de triangles et fait la boucle EN INTERNE (zero appel de fonction par
    # sommet/triangle), modifie les tableaux A/B/C EN PLACE (zero reallocation), et sort les 12
    # valeurs de la matrice en variables locales une seule fois (zero indexation de tableau
    # repetee dans la boucle chaude) -> 0.6 s au lieu de 10.2 s sur le meme test (~17x).
    # 27/10/2026 : BUGFIX sur le correctif ci-dessus - plusieurs triangles d'un meme maillage
    # 3MF peuvent PARTAGER le meme tableau [double[]] pour un sommet commun (le sommet est lu
    # une fois depuis <vertices> puis reference par plusieurs triangles, sans etre clone). Muter
    # ce tableau EN PLACE sans precaution l'aurait transforme une fois par triangle qui le
    # reference (double/triple transformation sur les sommets partages - confirme par
    # test_3mf_correctness.ps1 qui a detecte des coordonnees doublees). On garde donc une liste
    # des tableaux DEJA transformes (comparaison par reference, pas par contenu) pour n'appliquer
    # la transformation qu'UNE SEULE FOIS par sommet physique, meme s'il est partage.
    # (HashSet[object] par defaut suffit : un tableau ([double[]]) n'implemente pas Equals/
    # GetHashCode par contenu, donc la comparaison par defaut est deja une comparaison par
    # reference - compatible PowerShell 5.1 Desktop et 7.x, pas besoin de ReferenceEqualityComparer.)
    $dejaTransformes = [System.Collections.Generic.HashSet[object]]::new()
    $m0=$m[0]; $m1=$m[1]; $m2=$m[2]; $m3=$m[3]; $m4=$m[4]; $m5=$m[5]
    $m6=$m[6]; $m7=$m[7]; $m8=$m[8]; $m9=$m[9]; $m10=$m[10]; $m11=$m[11]
    foreach ($tri in $triangles) {
        foreach ($cle in 'A', 'B', 'C') {
            $p = $tri[$cle]
            if (-not $dejaTransformes.Add($p)) { continue }
            $x = $p[0]; $y = $p[1]; $z = $p[2]
            $p[0] = $x * $m0 + $y * $m3 + $z * $m6 + $m9
            $p[1] = $x * $m1 + $y * $m4 + $z * $m7 + $m10
            $p[2] = $x * $m2 + $y * $m5 + $z * $m8 + $m11
        }
    }
}

function Lire-Document3MF([string]$cheminFichier) {
    # 26/10/2026 : REECRITURE - l'ancienne version chargeait tout le fichier en DOM ([xml] +
    # XmlNamespaceManager + SelectNodes XPath), ce qui est tres lent sur un maillage a plus
    # d'un million de triangles (vu en usage reel : devient le vrai goulot d'etranglement une
    # fois --orient desactive - plusieurs minutes juste pour cette lecture). On lit maintenant
    # le fichier en FLUX avec XmlReader (jamais tout le document en memoire, pas de recherche
    # XPath repetee) : un seul passage sequentiel qui construit directement les objets
    # (sommets/triangles/composants) et les items du <build>.
    $objets = @{}
    $items = [System.Collections.Generic.List[object]]::new()
    $objetCourant = $null
    $nsProduction = "http://schemas.microsoft.com/3dmanufacturing/production/2015/06"
    $reader = [System.Xml.XmlReader]::Create($cheminFichier)
    try {
        while ($reader.Read()) {
            if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            switch ($reader.LocalName) {
                'object' {
                    $objetCourant = [PSCustomObject]@{
                        Sommets   = [System.Collections.Generic.List[double[]]]::new()
                        Triangles = [System.Collections.Generic.List[int[]]]::new()
                        Composants = [System.Collections.Generic.List[object]]::new()
                    }
                    $objets[$reader.GetAttribute('id')] = $objetCourant
                }
                'vertex' {
                    if ($objetCourant) {
                        $objetCourant.Sommets.Add([double[]]@(
                            [double]$reader.GetAttribute('x'),
                            [double]$reader.GetAttribute('y'),
                            [double]$reader.GetAttribute('z')))
                    }
                }
                'triangle' {
                    if ($objetCourant) {
                        $objetCourant.Triangles.Add([int[]]@(
                            [int]$reader.GetAttribute('v1'),
                            [int]$reader.GetAttribute('v2'),
                            [int]$reader.GetAttribute('v3')))
                    }
                }
                'component' {
                    if ($objetCourant) {
                        # Un <component> peut pointer vers un objet DANS LE MEME document
                        # (attribut "objectid" seul) OU vers un objet dans un AUTRE fichier
                        # .model du package (attribut p:path, extension "Production" de
                        # Bambu Studio) - voir Charger-Document3MFExterne.
                        $objetCourant.Composants.Add([PSCustomObject]@{
                            IdObjet       = $reader.GetAttribute('objectid')
                            CheminExterne = $reader.GetAttribute('path', $nsProduction)
                            Transform     = $reader.GetAttribute('transform')
                        })
                    }
                }
                'item' {
                    $items.Add([PSCustomObject]@{
                        IdObjet   = $reader.GetAttribute('objectid')
                        Transform = $reader.GetAttribute('transform')
                    })
                }
            }
        }
    } finally {
        $reader.Close()
    }
    return [PSCustomObject]@{ Objets = $objets; Items = $items }
}

function Charger-Document3MFExterne([string]$cheminRelatif, [string]$dossierBase, [hashtable]$cacheDocuments, [hashtable]$chronos = $null) {
    # 26/10/2026 : Bambu Studio (extension "Production" du format 3MF) n'ecrit PAS forcement
    # tout le maillage dans 3D/3dmodel.model : chaque piece fusionnee se retrouve dans son
    # PROPRE fichier (3D/Objects/object_1.model, object_2.model, ...), et le fichier principal
    # se contente d'une reference <component p:path="/3D/Objects/object_1.model"
    # objectid="1" .../>. Avant un premier correctif, on ne lisait QUE le fichier principal ->
    # ces references ne resolvaient jamais vers un objet connu -> 0 triangle.
    if ($cacheDocuments.ContainsKey($cheminRelatif)) { return $cacheDocuments[$cheminRelatif] }
    $cheminFichier = Join-Path $dossierBase ($cheminRelatif.TrimStart('/') -replace '/', '\')
    $objets = @{}
    if (Test-Path $cheminFichier) {
        # 27/10/2026 : chrono ajoute ici - diagnostic montre que cette piece passe par un
        # fichier .model EXTERNE (extension "Production" de Bambu Studio), et que la quasi
        # totalite des 21-84 s mesurees sur "resolution objets+composants+transforms" se
        # trouve forcement soit ici (lecture/parsing du fichier externe), soit dans la boucle
        # de construction des triangles Hashtable juste apres, soit dans l'application de la
        # transformation du composant - on les distingue maintenant separement.
        $chronoLocal = [System.Diagnostics.Stopwatch]::StartNew()
        $objets = (Lire-Document3MF $cheminFichier).Objets
        $chronoLocal.Stop()
        if ($chronos) { $chronos.TempsChargementExterne += $chronoLocal.Elapsed.TotalSeconds }
    }
    $cacheDocuments[$cheminRelatif] = $objets
    return $objets
}

function Resoudre-Objet3MF([string]$idObjet, [hashtable]$objetsParId, [string]$dossierBase, [hashtable]$cacheDocuments, [hashtable]$chronos = $null) {
    if (-not $chronos) { $chronos = @{ TempsChargementExterne = 0.0; TempsConstructionTriangles = 0.0; TempsTransform = 0.0 } }
    $obj = $objetsParId[$idObjet]
    $resultat = [System.Collections.Generic.List[object]]::new()
    if (-not $obj) { return $resultat }

    if ($obj.Triangles.Count -gt 0) {
        # 27/10/2026 : CORRECTIF PERF MAJEUR - mesure reelle (chronos ajoutes dans
        # Traiter-UneFigurineCore) : relire le .3mf produit par Bambu Studio prenait 28.5 s sur
        # 206104 triangles et 130.3 s sur 565366 triangles - soit un temps qui grimpe PLUS VITE
        # que le nombre de triangles (x2.74 triangles -> x4.57 temps), ce qui expliquait a lui
        # seul la quasi-totalite du temps "manquant" rapporte par Tomyn sur les gros fichiers.
        # Cette boucle (UNE instanciation d'objet PAR TRIANGLE) en etait la cause : chaque
        # [PSCustomObject] passe par l'adaptateur de type PowerShell/PSObject, nettement plus
        # couteux qu'une simple Hashtable - et plus il y a de triangles vivants simultanement
        # en memoire, plus le garbage collector travaille (d'ou la courbe pire que lineaire).
        # Fix valide en bac a sable (100000 iterations) : Hashtable ~3x plus rapide a construire
        # que PSCustomObject, et l'acces/l'ECRITURE par point (.A, .A = ...) fonctionne a
        # l'identique sur une Hashtable - aucun autre endroit du code n'a besoin de changer.
        $chronoConstruction = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($t in $obj.Triangles) {
            $resultat.Add(@{ A = $obj.Sommets[$t[0]]; B = $obj.Sommets[$t[1]]; C = $obj.Sommets[$t[2]] })
        }
        $chronoConstruction.Stop()
        $chronos.TempsConstructionTriangles += $chronoConstruction.Elapsed.TotalSeconds
        return $resultat
    }

    foreach ($comp in $obj.Composants) {
        if ($comp.CheminExterne) {
            $objetsCible = Charger-Document3MFExterne -cheminRelatif $comp.CheminExterne -dossierBase $dossierBase -cacheDocuments $cacheDocuments -chronos $chronos
        } else {
            $objetsCible = $objetsParId
        }
        $sousTriangles = Resoudre-Objet3MF -idObjet $comp.IdObjet -objetsParId $objetsCible -dossierBase $dossierBase -cacheDocuments $cacheDocuments -chronos $chronos
        if ($comp.Transform) {
            $chronoTransform = [System.Diagnostics.Stopwatch]::StartNew()
            $m = ($comp.Transform -split '\s+') | ForEach-Object { [double]$_ }
            Appliquer-Transform3MF-Triangles $sousTriangles $m
            $chronoTransform.Stop()
            $chronos.TempsTransform += $chronoTransform.Elapsed.TotalSeconds
        }
        foreach ($tri in $sousTriangles) { $resultat.Add($tri) }
    }
    return $resultat
}

function Lire-Maillage3MF([string]$chemin3mf) {
    # 27/10/2026 : chronos DETAILLES ajoutes ici - le correctif Hashtable (voir
    # Resoudre-Objet3MF) a bien reduit le temps de cette fonction (confirme par Tomyn : 130.3 s
    # -> 83.2 s sur 565366 triangles) mais c'est encore beaucoup trop lent, donc il reste un
    # autre cout dominant a l'interieur. On decoupe maintenant cette fonction en ses 3 phases
    # (decompression du .3mf, lecture du XML principal, resolution des objets/composants +
    # transforms) pour savoir laquelle domine reellement, plutot que de deviner encore.
    $diagnostics = [System.Collections.Generic.List[string]]::new()
    $dossierTemp = Join-Path ([System.IO.Path]::GetTempPath()) ("fusion3mf_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $dossierTemp -Force | Out-Null
    try {
        $chronoZip = [System.Diagnostics.Stopwatch]::StartNew()
        [System.IO.Compression.ZipFile]::ExtractToDirectory($chemin3mf, $dossierTemp)
        $chronoZip.Stop()
        $cheminModele = Join-Path $dossierTemp "3D\3dmodel.model"
        if (-not (Test-Path $cheminModele)) { throw "3D\3dmodel.model introuvable dans le .3mf." }

        $chronoXmlPrincipal = [System.Diagnostics.Stopwatch]::StartNew()
        $document = Lire-Document3MF $cheminModele
        $chronoXmlPrincipal.Stop()
        $objetsParId = $document.Objets
        $cacheDocuments = @{}

        # Diagnostic de structure : un .3mf Bambu Studio peut eclater chaque piece dans son
        # PROPRE fichier .model externe (extension "Production") - si c'est le cas ici, la
        # resolution ci-dessous va re-declencher un Lire-Document3MF COMPLET par piece externe,
        # ce qui ne serait PAS visible dans un test synthetique a un seul objet.
        $nbComposantsExternes = 0
        foreach ($obj in $objetsParId.Values) {
            foreach ($comp in $obj.Composants) { if ($comp.CheminExterne) { $nbComposantsExternes++ } }
        }
        $diagnostics.Add("structure : $($objetsParId.Count) objet(s) dans le fichier principal, $($document.Items.Count) item(s) dans <build>, $nbComposantsExternes composant(s) pointant vers un fichier .model EXTERNE")

        $chronoResolution = [System.Diagnostics.Stopwatch]::StartNew()
        # 27/10/2026 : sous-chronos pour la resolution elle-meme (voir Resoudre-Objet3MF /
        # Charger-Document3MFExterne) - diagnostic a confirme que la piece passe par un fichier
        # .model EXTERNE (extension Production de Bambu Studio), donc la question suivante est
        # LAQUELLE de ses 3 sous-etapes (lire le fichier externe, construire les triangles,
        # appliquer la transformation du composant) domine les 21-84 s observes.
        $chronosDetailles = @{ TempsChargementExterne = 0.0; TempsConstructionTriangles = 0.0; TempsTransform = 0.0 }
        $resultatGlobal = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $document.Items) {
            $triangles = Resoudre-Objet3MF -idObjet $item.IdObjet -objetsParId $objetsParId -dossierBase $dossierTemp -cacheDocuments $cacheDocuments -chronos $chronosDetailles
            if ($item.Transform) {
                # 27/10/2026 : meme correctif que dans Resoudre-Objet3MF (CORRECTIF PERF MAJEUR #2)
                # - on applique la transformation de l'<item> avec la version "batch" au lieu de
                # 3 appels de fonction par triangle (voir Appliquer-Transform3MF-Triangles).
                $m = ($item.Transform -split '\s+') | ForEach-Object { [double]$_ }
                Appliquer-Transform3MF-Triangles $triangles $m
            }
            foreach ($tri in $triangles) { $resultatGlobal.Add($tri) }
        }
        $chronoResolution.Stop()
        $diagnostics.Add("decompression .3mf : $([Math]::Round($chronoZip.Elapsed.TotalSeconds, 1)) s / lecture XML principal : $([Math]::Round($chronoXmlPrincipal.Elapsed.TotalSeconds, 1)) s / resolution objets+composants+transforms : $([Math]::Round($chronoResolution.Elapsed.TotalSeconds, 1)) s")
        $diagnostics.Add("detail resolution -> lecture fichier(s) .model externe(s) : $([Math]::Round($chronosDetailles.TempsChargementExterne, 1)) s / construction triangles (Hashtable) : $([Math]::Round($chronosDetailles.TempsConstructionTriangles, 1)) s / application transform composant : $([Math]::Round($chronosDetailles.TempsTransform, 1)) s")
        return [PSCustomObject]@{ Triangles = $resultatGlobal; Diagnostics = $diagnostics }
    } finally {
        Remove-Item -Path $dossierTemp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================================
# ECRITURE STL BINAIRE
# ============================================================================
function Ecrire-STLBinaire($triangles, [string]$cheminSortie) {
    # 27/10/2026 : CORRECTIF PERF - l'ancienne version faisait UN appel BinaryWriter.Write PAR
    # VALEUR (13 appels/triangle : 3 floats normale + 9 floats sommets + 1 uint16), soit ~19
    # MILLIONS d'appels pour une figurine a 1.46M triangles. Chaque appel .Write() depuis du
    # code PowerShell interprete a un cout fixe non negligeable (dispatch de methode .NET) -
    # multiplie par des millions d'appels, ca explique probablement une bonne part du
    # ralentissement tres disproportionne que Tomyn a observe sur les gros maillages (206000
    # triangles en quelques secondes, 1.46M triangles toujours pas fini apres 15 minutes - un
    # ecart bien plus grand que le simple facteur x7 de triangles). Fix : on remplit UN SEUL
    # gros buffer en memoire (les conversions BitConverter restent necessaires, mais on elimine
    # les millions d'appels au FLUX lui-meme, bien plus couteux), puis UNE SEULE ecriture disque.
    $nb = $triangles.Count
    $tampon = New-Object byte[] (84 + ($nb * 50))
    [Array]::Copy([System.BitConverter]::GetBytes([uint32]$nb), 0, $tampon, 80, 4)
    $pos = 84
    # 27/10/2026 : CORRECTIF PERF #3 - benchmarke en bac a sable (bench_ecriture.ps1) : le
    # "foreach ($pt in @($tri.A, $tri.B, $tri.C))" ci-dessous creait un tableau TEMPORAIRE a 3
    # elements PAR TRIANGLE (1.46M fois sur la grosse figurine de Tomyn) juste pour pouvoir
    # boucler sur A/B/C - pur gaspillage, aucun besoin de boucler puisqu'on a deja 3 noms fixes.
    # Deroule explicitement (A puis B puis C, sans tableau intermediaire) : memes octets en
    # sortie (verifie octet-par-octet dans le benchmark), ~2x plus rapide mesure sur 500000
    # triangles synthetiques (1.65s -> 0.82s) - cette fonction est appelee DEUX FOIS par
    # figurine (fusion locale + ecriture .stl final), donc le gain s'applique aux deux.
    foreach ($tri in $triangles) {
        $a = $tri.A; $b = $tri.B; $c = $tri.C
        $ux = $b[0] - $a[0]; $uy = $b[1] - $a[1]; $uz = $b[2] - $a[2]
        $vx = $c[0] - $a[0]; $vy = $c[1] - $a[1]; $vz = $c[2] - $a[2]
        $nx = $uy * $vz - $uz * $vy; $ny = $uz * $vx - $ux * $vz; $nz = $ux * $vy - $uy * $vx
        $longueur = [Math]::Sqrt($nx * $nx + $ny * $ny + $nz * $nz)
        if ($longueur -gt 0) { $nx /= $longueur; $ny /= $longueur; $nz /= $longueur }
        [Array]::Copy([System.BitConverter]::GetBytes([float]$nx), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$ny), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$nz), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$a[0]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$a[1]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$a[2]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$b[0]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$b[1]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$b[2]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$c[0]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$c[1]), 0, $tampon, $pos, 4); $pos += 4
        [Array]::Copy([System.BitConverter]::GetBytes([float]$c[2]), 0, $tampon, $pos, 4); $pos += 4
        $pos += 2  # "attribute byte count" - deja a 0 par defaut dans un New-Object byte[]
    }
    [System.IO.File]::WriteAllBytes($cheminSortie, $tampon)
}

# ============================================================================
# APERCU ISOMETRIQUE STATIQUE (rendu logiciel simple, pas de librairie 3D)
# ============================================================================
function Rendre-ApercuIsometrique($triangles, [int]$largeur = 320, [int]$hauteur = 240, [int]$tailleBloc = 2) {
    # 26/10/2026 : REECRITURE COMPLETE. L'ancienne version dessinait chaque triangle avec
    # Graphics.FillPolygon (+ creation de plusieurs objets PointF/SolidBrush PAR triangle),
    # et compensait l'echantillonnage en "gonflant" chaque triangle garde autour de son
    # propre centre ecran. Double probleme :
    #  - LENT : FillPolygon + creation d'objets .NET en boucle PowerShell interpretee ne
    #    tient pas la charge au-dela de quelques dizaines de milliers de triangles (teste :
    #    x3 triangles = x3+ temps, linaire avec un facteur constant enorme).
    #  - PAS FIDELE : gonfler un triangle isole, sans tenir compte de son orientation reelle,
    #    produit un blob qui ne ressemble pas a la surface d'origine - sur des pieces
    #    detaillees (ex: ~950 000 triangles par piece observes dans ce projet), meme avec
    #    plus de triangles, le resultat restait de la "bouillie".
    # Nouvelle approche : on "asperge" directement les SOMMETS des triangles sur un buffer
    # de pixels BRUT (tableau d'octets, pas d'objets .NET ni d'appel GDI+ par sommet - juste
    # de l'arithmetique et de l'indexation de tableau), avec un simple test de profondeur
    # (z-buffer) pour l'occlusion. Chaque sommet coute une fraction du prix d'un FillPolygon,
    # donc on peut se permettre d'en afficher BEAUCOUP plus (toute la densite reelle du
    # maillage plutot qu'un triangle gonfle tous les N) - avec assez de sommets, la surface
    # parait pleine SANS avoir besoin d'un facteur de gonflage artificiel.
    $bmp = New-Object System.Drawing.Bitmap($largeur, $hauteur)

    if (-not $triangles -or $triangles.Count -eq 0) {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear($THEME.FondChamp)
        $g.DrawString("(pas d'apercu)", (New-Object System.Drawing.Font("Segoe UI", 9)), (New-Object System.Drawing.SolidBrush($THEME.TexteAttenue)), 10, 10)
        $g.Dispose()
        return $bmp
    }

    $cos30 = [Math]::Cos([Math]::PI / 6); $sin30 = [Math]::Sin([Math]::PI / 6)

    # Passe 1 (legere) : juste les bornes min/max pour calculer l'echelle d'affichage -
    # necessaire avant de poser le moindre pixel.
    $minX = [double]::MaxValue; $maxX = [double]::MinValue
    $minY = [double]::MaxValue; $maxY = [double]::MinValue
    foreach ($tri in $triangles) {
        foreach ($p in @($tri.A, $tri.B, $tri.C)) {
            $sx = ($p[0] - $p[1]) * $cos30
            $sy = ($p[0] + $p[1]) * $sin30 - $p[2]
            if ($sx -lt $minX) { $minX = $sx }; if ($sx -gt $maxX) { $maxX = $sx }
            if ($sy -lt $minY) { $minY = $sy }; if ($sy -gt $maxY) { $maxY = $sy }
        }
    }
    $largeurModele = $maxX - $minX; $hauteurModele = $maxY - $minY
    if ($largeurModele -le 0) { $largeurModele = 1 }; if ($hauteurModele -le 0) { $hauteurModele = 1 }
    $marge = 20
    $echelle = [Math]::Min(($largeur - 2 * $marge) / $largeurModele, ($hauteur - 2 * $marge) / $hauteurModele)
    $centreX = ($minX + $maxX) / 2; $centreY = ($minY + $maxY) / 2

    # Buffer pixel brut BGRA (ordre memoire de Format32bppArgb) + buffer de profondeur
    # (double par pixel) pour le test d'occlusion. Tout est initialise a la couleur de fond.
    $nbPixels = $largeur * $hauteur
    $couleurFond = $THEME.FondChamp
    $bufPixels = New-Object byte[] ($nbPixels * 4)
    for ($i = 0; $i -lt $nbPixels; $i++) {
        $o = $i * 4
        $bufPixels[$o] = $couleurFond.B; $bufPixels[$o + 1] = $couleurFond.G
        $bufPixels[$o + 2] = $couleurFond.R; $bufPixels[$o + 3] = 255
    }
    $bufProfondeur = New-Object double[] $nbPixels
    for ($i = 0; $i -lt $nbPixels; $i++) { $bufProfondeur[$i] = [double]::MinValue }

    $directionLumiere = [double[]]@(0.45, -0.55, 0.70)
    $longueurLumiere = [Math]::Sqrt(($directionLumiere | ForEach-Object { $_ * $_ } | Measure-Object -Sum).Sum)
    for ($k = 0; $k -lt $directionLumiere.Length; $k++) { $directionLumiere[$k] /= $longueurLumiere }
    $couleurBase = @(90, 150, 220)

    foreach ($tri in $triangles) {
        # Ombrage calcule une fois par triangle (normale), reutilise pour ses 3 sommets.
        $ux = $tri.B[0] - $tri.A[0]; $uy = $tri.B[1] - $tri.A[1]; $uz = $tri.B[2] - $tri.A[2]
        $vx = $tri.C[0] - $tri.A[0]; $vy = $tri.C[1] - $tri.A[1]; $vz = $tri.C[2] - $tri.A[2]
        $nx = $uy * $vz - $uz * $vy; $ny = $uz * $vx - $ux * $vz; $nz = $ux * $vy - $uy * $vx
        $nl = [Math]::Sqrt($nx * $nx + $ny * $ny + $nz * $nz)
        $ombrage = 0.6
        if ($nl -gt 0) {
            $nx /= $nl; $ny /= $nl; $nz /= $nl
            $eclat = ($nx * $directionLumiere[0] + $ny * $directionLumiere[1] + $nz * $directionLumiere[2])
            $ombrage = [Math]::Max(0.25, [Math]::Min(1.0, 0.45 + $eclat * 0.65))
        }
        $cB = [byte]([Math]::Min(255, $couleurBase[2] * $ombrage))
        $cG = [byte]([Math]::Min(255, $couleurBase[1] * $ombrage))
        $cR = [byte]([Math]::Min(255, $couleurBase[0] * $ombrage))

        # 26/10/2026 : "foreach ($p in @($tri.A,$tri.B,$tri.C))" + double "foreach (dx/dy in
        # 0,1)" allouaient un petit tableau .NET a CHAQUE sommet (plusieurs fois) - mesure
        # (sandbox) : deroulement manuel (pas de tableaux temporaires) ~35% plus rapide sur
        # 300000 triangles. Sur une boucle de cette taille en PowerShell interprete, ca compte.
        for ($iv = 0; $iv -lt 3; $iv++) {
            if ($iv -eq 0) { $px = $tri.A[0]; $py = $tri.A[1]; $pz = $tri.A[2] }
            elseif ($iv -eq 1) { $px = $tri.B[0]; $py = $tri.B[1]; $pz = $tri.B[2] }
            else { $px = $tri.C[0]; $py = $tri.C[1]; $pz = $tri.C[2] }
            $sx = ($px - $py) * $cos30
            $sy = ($px + $py) * $sin30 - $pz
            $exCentre = [int](($sx - $centreX) * $echelle + $largeur / 2)
            $eyCentre = [int](($sy - $centreY) * $echelle + $hauteur / 2)
            $profondeur = $px + $py + $pz
            # On pose un petit bloc NxN (pas juste 1 pixel) pour combler les trous entre
            # sommets voisins sur les zones moins denses, sans le cout d'un vrai FillPolygon.
            # 27/10/2026 : taille de bloc parametrable ($tailleBloc, 2 par defaut = identique
            # au comportement d'origine) - demande de Tomyn : l'apercu de piece individuelle
            # (budget de triangles plus bas pour rester rapide, voir Demarrer-ApercuAsync) avait
            # des "trous" bien plus visibles en mode clair qu'en sombre (le fond blanc ressort
            # entre les points epars). Des blocs plus gros compensent la densite de points plus
            # faible SANS avoir besoin de plus de triangles (donc sans perdre en vitesse).
            for ($dx = 0; $dx -lt $tailleBloc; $dx++) {
                $ex = $exCentre + $dx
                if ($ex -lt 0 -or $ex -ge $largeur) { continue }
                for ($dy = 0; $dy -lt $tailleBloc; $dy++) {
                    $ey = $eyCentre + $dy
                    if ($ey -lt 0 -or $ey -ge $hauteur) { continue }
                    $idx = $ey * $largeur + $ex
                    if ($profondeur -gt $bufProfondeur[$idx]) { $bufProfondeur[$idx] = $profondeur; $o = $idx * 4; $bufPixels[$o] = $cB; $bufPixels[$o + 1] = $cG; $bufPixels[$o + 2] = $cR; $bufPixels[$o + 3] = 255 }
                }
            }
        }
    }

    $rect = New-Object System.Drawing.Rectangle(0, 0, $largeur, $hauteur)
    $donnees = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $octetsParLigne = $largeur * 4
        if ($donnees.Stride -eq $octetsParLigne) {
            [System.Runtime.InteropServices.Marshal]::Copy($bufPixels, 0, $donnees.Scan0, $bufPixels.Length)
        } else {
            # Stride different de largeur*4 (arrondi/alignement) - copie ligne par ligne pour
            # rester correct quelle que soit la valeur exacte du stride.
            for ($ligne = 0; $ligne -lt $hauteur; $ligne++) {
                $dest = [System.IntPtr]::Add($donnees.Scan0, $ligne * $donnees.Stride)
                [System.Runtime.InteropServices.Marshal]::Copy($bufPixels, $ligne * $octetsParLigne, $dest, $octetsParLigne)
            }
        }
    } finally {
        $bmp.UnlockBits($donnees)
    }
    return $bmp
}

# ============================================================================
# APERCU VIA RENDU MATERIEL (WPF/Media3D) - rapide ET complet (pas d'echantillonnage)
# 26/10/2026 : contrairement au rendu logiciel ci-dessus (Rendre-ApercuIsometrique, qui
# doit projeter/calculer l'ombrage/trier CHAQUE triangle en PowerShell pur - lent sur un
# gros maillage), on delegue ici le rendu a la carte graphique via WPF (Media3D), comme le
# font les slicers pour leurs miniatures : le tri de profondeur et l'eclairage sont geres
# nativement (z-buffer), donc on peut afficher TOUS les triangles sans echantillonnage.
# Cote PowerShell, il ne reste qu'une lecture directe des sommets du/des fichier(s) STL
# (pas de calculs de projection/ombrage/tri a faire nous-memes = plus rapide aussi la).
# Si WPF n'est pas disponible ou echoue pour une raison quelconque, l'appelant doit
# prevoir un repli sur Rendre-ApercuIsometrique (voir Demarrer-ApercuAsync).
# ============================================================================
function Rendre-ApercuWPF {
    param(
        [string[]]$CheminsFichiers,
        [int]$Largeur = 320,
        [int]$Hauteur = 240,
        [hashtable]$THEME,
        # 26/10/2026 : meme avec le rendu GPU (pas de tri/ombrage/projection a calculer nous-
        # memes), chaque triangle lu paye quand meme le cout d'une iteration de boucle
        # PowerShell (langage interprete) pour construire les Point3D. Sur un maillage de
        # plusieurs millions de triangles (miniature tres detaillee), ca restait beaucoup
        # trop long (et a pu, selon le test, declencher une pression memoire/erreur mal
        # diagnostiquee avec une liste non bornee). On plafonne donc ici aussi - mais avec
        # un budget bien plus genereux qu'avant (pas de calcul couteux par triangle a payer
        # en plus), donc le resultat reste visuellement quasi complet dans l'immense
        # majorite des cas.
        [int]$MaxTrianglesParFichier = 150000
    )

    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

    $positions = [System.Collections.Generic.List[System.Windows.Media.Media3D.Point3D]]::new()

    foreach ($chemin in $CheminsFichiers) {
        $octets = [System.IO.File]::ReadAllBytes($chemin)
        $estBinaire = $false
        if ($octets.Length -ge 84) {
            $nbTri = [System.BitConverter]::ToUInt32($octets, 80)
            if ((84 + ($nbTri * 50)) -eq $octets.Length) { $estBinaire = $true }
        }
        if ($estBinaire) {
            # Saut direct par calcul d'offset (pas d'accumulation lineaire) : les triangles
            # sautes ne sont jamais decodes, c'est ce qui garde ca rapide meme plafonne haut.
            $pasLecture = [Math]::Max(1, [Math]::Floor($nbTri / [Math]::Max(1, $MaxTrianglesParFichier)))
            for ($i = 0; $i -lt $nbTri; $i += $pasLecture) {
                $pos = 84 + ($i * 50) + 12
                $ax = [System.BitConverter]::ToSingle($octets, $pos);      $ay = [System.BitConverter]::ToSingle($octets, $pos + 4);  $az = [System.BitConverter]::ToSingle($octets, $pos + 8)
                $bx = [System.BitConverter]::ToSingle($octets, $pos + 12); $by = [System.BitConverter]::ToSingle($octets, $pos + 16); $bz = [System.BitConverter]::ToSingle($octets, $pos + 20)
                $cx = [System.BitConverter]::ToSingle($octets, $pos + 24); $cy = [System.BitConverter]::ToSingle($octets, $pos + 28); $cz = [System.BitConverter]::ToSingle($octets, $pos + 32)
                $positions.Add([System.Windows.Media.Media3D.Point3D]::new($ax, $ay, $az))
                $positions.Add([System.Windows.Media.Media3D.Point3D]::new($bx, $by, $bz))
                $positions.Add([System.Windows.Media.Media3D.Point3D]::new($cx, $cy, $cz))
            }
        } else {
            $lignes = [System.IO.File]::ReadAllLines($chemin)
            $compteTri = 0
            foreach ($ligne in $lignes) {
                if ($compteTri -ge $MaxTrianglesParFichier) { break }
                $l = $ligne.Trim()
                if ($l.StartsWith("vertex")) {
                    $parties = $l -split '\s+'
                    if ($parties.Count -ge 4) {
                        $positions.Add([System.Windows.Media.Media3D.Point3D]::new([double]$parties[1], [double]$parties[2], [double]$parties[3]))
                        if ($positions.Count % 3 -eq 0) { $compteTri++ }
                    }
                }
            }
        }
    }

    if ($positions.Count -eq 0) {
        $bmp = New-Object System.Drawing.Bitmap($Largeur, $Hauteur)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear($THEME.FondChamp)
        $g.DrawString("(pas d'apercu)", (New-Object System.Drawing.Font("Segoe UI", 9)), (New-Object System.Drawing.SolidBrush($THEME.TexteAttenue)), 10, 10)
        $g.Dispose()
        return $bmp
    }

    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $mesh.Positions = New-Object System.Windows.Media.Media3D.Point3DCollection -ArgumentList (, $positions)
    $indices = [System.Linq.Enumerable]::Range(0, $positions.Count)
    $mesh.TriangleIndices = New-Object System.Windows.Media.Media3D.Int32Collection -ArgumentList (, $indices)

    $bounds = $mesh.Bounds
    $centre = New-Object System.Windows.Media.Media3D.Point3D(($bounds.X + $bounds.SizeX / 2), ($bounds.Y + $bounds.SizeY / 2), ($bounds.Z + $bounds.SizeZ / 2))
    $rayon = [Math]::Max(0.001, ([Math]::Sqrt($bounds.SizeX * $bounds.SizeX + $bounds.SizeY * $bounds.SizeY + $bounds.SizeZ * $bounds.SizeZ) / 2))

    $couleurBase = [System.Windows.Media.Color]::FromRgb($THEME.AccentPrincipal.R, $THEME.AccentPrincipal.G, $THEME.AccentPrincipal.B)
    $materiau = New-Object System.Windows.Media.Media3D.MaterialGroup
    $diffus = New-Object System.Windows.Media.Media3D.DiffuseMaterial(New-Object System.Windows.Media.SolidColorBrush($couleurBase))
    [void]$materiau.Children.Add($diffus)
    $specul = New-Object System.Windows.Media.Media3D.SpecularMaterial(
        (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)), 30
    )
    [void]$materiau.Children.Add($specul)

    $modele = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $materiau)
    $modele.BackMaterial = $diffus

    $groupe = New-Object System.Windows.Media.Media3D.Model3DGroup
    [void]$groupe.Children.Add($modele)
    [void]$groupe.Children.Add((New-Object System.Windows.Media.Media3D.AmbientLight([System.Windows.Media.Color]::FromRgb(100, 100, 100))))
    [void]$groupe.Children.Add((New-Object System.Windows.Media.Media3D.DirectionalLight(
        [System.Windows.Media.Colors]::White, (New-Object System.Windows.Media.Media3D.Vector3D(-0.5, -0.6, -0.7))
    )))

    $visuel = New-Object System.Windows.Media.Media3D.ModelVisual3D
    $visuel.Content = $groupe

    # Camera orthographique en vue "isometrique" (biais haut/cote), comme l'ancien rendu.
    $distanceCamera = $rayon * 3
    $posCamera = New-Object System.Windows.Media.Media3D.Point3D(
        ($centre.X + $distanceCamera * 0.7), ($centre.Y - $distanceCamera * 0.7), ($centre.Z + $distanceCamera * 0.6)
    )
    $camera = New-Object System.Windows.Media.Media3D.OrthographicCamera
    $camera.Position = $posCamera
    $direction = New-Object System.Windows.Media.Media3D.Vector3D(($centre.X - $posCamera.X), ($centre.Y - $posCamera.Y), ($centre.Z - $posCamera.Z))
    $direction.Normalize()
    $camera.LookDirection = $direction
    $camera.UpDirection = New-Object System.Windows.Media.Media3D.Vector3D(0, 0, 1)
    $camera.Width = $rayon * 2.3

    $viewport = New-Object System.Windows.Controls.Viewport3D
    $viewport.Camera = $camera
    [void]$viewport.Children.Add($visuel)
    $viewport.Width = $Largeur
    $viewport.Height = $Hauteur
    $viewport.Background = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromRgb($THEME.FondChamp.R, $THEME.FondChamp.G, $THEME.FondChamp.B)
    )

    $tailleRect = New-Object System.Windows.Size($Largeur, $Hauteur)
    $viewport.Measure($tailleRect)
    $viewport.Arrange((New-Object System.Windows.Rect(0, 0, $Largeur, $Hauteur)))
    $viewport.UpdateLayout()

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Largeur, $Hauteur, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($viewport)

    # Conversion RenderTargetBitmap (WPF) -> System.Drawing.Bitmap (ce que le PictureBox
    # WinForms attend), via un aller-retour PNG en memoire - simple et fiable.
    $encodeur = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    [void]$encodeur.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $flux = New-Object System.IO.MemoryStream
    $encodeur.Save($flux)
    $flux.Position = 0
    $bmpFinal = New-Object System.Drawing.Bitmap($flux)
    $flux.Dispose()
    return $bmpFinal
}

# ============================================================================
# FENETRE PRINCIPALE
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Fusion de figurines multi-pieces -> STL (via Bambu Studio)"
$form.Size = New-Object System.Drawing.Size(1180, 900)
$form.StartPosition = "CenterScreen"
$form.BackColor = $THEME.FondApp
$form.ForeColor = $THEME.Texte
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Creer-Label([string]$texte, [int]$x, [int]$y) {
    # 26/10/2026 : CORRECTIF affichage - le Label WinForms a AutoSize=true par defaut. Une
    # etiquette longue ("Dossier racine (sous-dossiers = figurines) :") s'agrandissait donc
    # au-dela de la largeur prevue et passait PAR-DESSUS le champ de texte voisin (ajoute
    # apres elle, donc dessine au-dessus), ce qui donnait un affichage charcute/illisible.
    # Fix : AutoSize explicite, et etiquette placee sur SA PROPRE LIGNE, au-dessus du champ,
    # jamais a cote - plus aucun risque de chevauchement quelle que soit la longueur du texte.
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Text = $texte
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.ForeColor = $THEME.TexteAttenue
    $form.Controls.Add($lbl)
    return $lbl
}

# 27/10/2026 : repeint tous les controles DEJA CREES avec les valeurs actuelles de $THEME -
# utilisee quand Tomyn bascule sombre/clair en cours de route (le premier affichage n'en a pas
# besoin, les controles lisent deja $THEME au moment de leur creation). Marche par TYPE de
# controle plutot que par une liste de variables a maintenir a la main - sauf pour les boutons
# a couleur d'accent FIXE (Scanner, Traiter, Arreter, Relancer les echecs...) qui portent
# .Tag = "AccentFixe" pose a leur creation : ceux-la gardent leur couleur semantique quel que
# soit le theme, seuls les boutons "neutres" basculent FondPanneau/Texte.
function Repeindre-ControlesTheme($collection) {
    foreach ($ctrl in $collection) {
        switch ($ctrl.GetType().Name) {
            "Label" { $ctrl.ForeColor = $THEME.TexteAttenue }
            "TextBox" { $ctrl.BackColor = $THEME.FondChamp; $ctrl.ForeColor = $THEME.Texte }
            "ListView" { $ctrl.BackColor = $THEME.FondChamp; $ctrl.ForeColor = $THEME.Texte }
            "CheckedListBox" { $ctrl.BackColor = $THEME.FondChamp; $ctrl.ForeColor = $THEME.Texte }
            "PictureBox" { $ctrl.BackColor = $THEME.FondChamp }
            "CheckBox" { $ctrl.ForeColor = $THEME.Texte; $ctrl.BackColor = [System.Drawing.Color]::Transparent }
            "Button" {
                if ($ctrl.Tag -ne "AccentFixe") {
                    $ctrl.BackColor = $THEME.FondPanneau
                    $ctrl.ForeColor = $THEME.Texte
                    # 27/10/2026 : une fine bordure en plus du fond deja rehausse - l'ecart de
                    # couleur seul restait limite (surtout remarque en mode clair sur "Tout
                    # reinclure"/"Nettoyer dossiers temp"), la bordure donne un contour net
                    # quel que soit le theme.
                    $ctrl.FlatAppearance.BorderSize = 1
                    $ctrl.FlatAppearance.BorderColor = $THEME.Bordure
                }
            }
        }
        if ($ctrl.Controls -and $ctrl.Controls.Count -gt 0) { Repeindre-ControlesTheme $ctrl.Controls }
    }
}

function Basculer-Theme {
    $script:Config.Theme = if ($script:Config.Theme -eq "Claire") { "Sombre" } else { "Claire" }
    $nouvellePalette = if ($script:Config.Theme -eq "Claire") { $script:PaletteClaire } else { $script:PaletteSombre }
    foreach ($cle in $nouvellePalette.Keys) { $THEME[$cle] = $nouvellePalette[$cle] }
    $form.BackColor = $THEME.FondApp
    $form.ForeColor = $THEME.Texte
    Repeindre-ControlesTheme $form.Controls
    $boutonTheme.Text = if ($script:Config.Theme -eq "Claire") { "Mode sombre" } else { "Mode clair" }
    Sauver-Config
    # 27/10/2026 : les apercus deja affiches ont ete DESSINES avec l'ancienne couleur de fond -
    # on les regenere pour la figurine/piece actuellement selectionnee, sinon ils gardent un
    # fond qui ne correspond plus au nouveau theme jusqu'au prochain clic de Tomyn.
    if ($liste.SelectedItems.Count -gt 0) {
        $grpSelectionne = $liste.SelectedItems[0].Tag
        if ($grpSelectionne) {
            Rafraichir-ApercusPour $grpSelectionne
            Remplir-ListePieces $grpSelectionne
        }
    }
}

$boutonTheme = New-Object System.Windows.Forms.Button
$boutonTheme.Text = if ($script:Config.Theme -eq "Claire") { "Mode sombre" } else { "Mode clair" }
$boutonTheme.Location = New-Object System.Drawing.Point(1020, 6)
$boutonTheme.Size = New-Object System.Drawing.Size(140, 24)
$boutonTheme.BackColor = $THEME.FondPanneau
$boutonTheme.ForeColor = $THEME.Texte
$boutonTheme.FlatStyle = "Flat"
$boutonTheme.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($boutonTheme)
Arrondir-Bouton $boutonTheme 6
$boutonTheme.Add_Click({ Basculer-Theme })

# 26/10/2026 : les deux champs ci-dessous etaient avant construits par une fonction
# Creer-ChampDossier commune, mais leur affichage restait casse (ordre/position
# imprevisibles) meme apres le correctif du Label. Ecrits maintenant en DIRECT, exactement
# comme le champ "Bambu Studio" juste en dessous (qui lui s'est toujours affiche
# correctement) - plus aucune fonction/closure partagee entre les deux champs.

[void](Creer-Label "Dossier racine (sous-dossiers = figurines) :" 12 12)
$champRacine = New-Object System.Windows.Forms.TextBox
$champRacine.Location = New-Object System.Drawing.Point(12, 32)
$champRacine.Size = New-Object System.Drawing.Size(900, 24)
$champRacine.BackColor = $THEME.FondChamp
$champRacine.ForeColor = $THEME.Texte
# 27/10/2026 : coins arrondis (Arrondir-Bouton, voir apres l'ajout au formulaire) sur les
# champs, listes et apercus, pour une allure moins "appli Windows par defaut". La bordure
# FixedSingle est GARDEE (contrairement a un premier essai sans bordure) : sans elle, le champ
# ne se distinguait du fond que par sa couleur, et assombrir le fond pour compenser rapprochait
# trop le mode clair du mode sombre (remarque de Tomyn) - la bordure delimite le champ sans
# avoir besoin de jouer sur les couleurs de fond.
$champRacine.BorderStyle = "FixedSingle"
$champRacine.Text = $script:Config.DossierRacine
$champRacine.AllowDrop = $true
$champRacine.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
})
$champRacine.Add_DragDrop({
    param($s, $e)
    $chemins = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if (-not $chemins -or $chemins.Count -eq 0) { return }
    $premier = $chemins[0]
    if (Test-Path -LiteralPath $premier -PathType Container) { $champRacine.Text = $premier }
    elseif (Test-Path -LiteralPath $premier -PathType Leaf) { $champRacine.Text = Split-Path $premier -Parent }
    # 27/10/2026 : demande de Tomyn - lance le scan automatiquement apres un glisser-deposer,
    # pas besoin de cliquer "Scanner le dossier racine" en plus.
    $boutonScanner.PerformClick()
})
$form.Controls.Add($champRacine)
Arrondir-Bouton $champRacine 6
$boutonRacineParcourir = New-Object System.Windows.Forms.Button
$boutonRacineParcourir.Text = "Parcourir..."
$boutonRacineParcourir.Location = New-Object System.Drawing.Point(920, 31)
$boutonRacineParcourir.Size = New-Object System.Drawing.Size(110, 26)
$boutonRacineParcourir.BackColor = $THEME.FondPanneau
$boutonRacineParcourir.ForeColor = $THEME.Texte
$boutonRacineParcourir.FlatStyle = "Flat"
$boutonRacineParcourir.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($champRacine.Text -and (Test-Path $champRacine.Text)) { $dlg.SelectedPath = $champRacine.Text }
    if ($dlg.ShowDialog() -eq "OK") { $champRacine.Text = $dlg.SelectedPath }
})
$form.Controls.Add($boutonRacineParcourir)
Arrondir-Bouton $boutonRacineParcourir

[void](Creer-Label "Dossier de sortie pour les .stl fusionnes :" 12 68)
$champSortie = New-Object System.Windows.Forms.TextBox
$champSortie.Location = New-Object System.Drawing.Point(12, 88)
$champSortie.Size = New-Object System.Drawing.Size(900, 24)
$champSortie.BackColor = $THEME.FondChamp
$champSortie.ForeColor = $THEME.Texte
$champSortie.BorderStyle = "FixedSingle"
$champSortie.Text = $script:Config.DossierSortie
$champSortie.AllowDrop = $true
$champSortie.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
})
$champSortie.Add_DragDrop({
    param($s, $e)
    $chemins = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if (-not $chemins -or $chemins.Count -eq 0) { return }
    $premier = $chemins[0]
    if (Test-Path -LiteralPath $premier -PathType Container) { $champSortie.Text = $premier }
    elseif (Test-Path -LiteralPath $premier -PathType Leaf) { $champSortie.Text = Split-Path $premier -Parent }
})
$form.Controls.Add($champSortie)
Arrondir-Bouton $champSortie 6
$boutonSortieParcourir = New-Object System.Windows.Forms.Button
$boutonSortieParcourir.Text = "Parcourir..."
$boutonSortieParcourir.Location = New-Object System.Drawing.Point(920, 87)
$boutonSortieParcourir.Size = New-Object System.Drawing.Size(110, 26)
$boutonSortieParcourir.BackColor = $THEME.FondPanneau
$boutonSortieParcourir.ForeColor = $THEME.Texte
$boutonSortieParcourir.FlatStyle = "Flat"
$boutonSortieParcourir.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($champSortie.Text -and (Test-Path $champSortie.Text)) { $dlg.SelectedPath = $champSortie.Text }
    if ($dlg.ShowDialog() -eq "OK") { $champSortie.Text = $dlg.SelectedPath }
})
$form.Controls.Add($boutonSortieParcourir)
Arrondir-Bouton $boutonSortieParcourir

# 26/10/2026 : ouvre directement le dossier de sortie dans l'Explorateur (demande de
# Tomyn) - evite de devoir naviguer a la main jusqu'au dossier ou les .stl/.3mf viennent
# d'etre ecrits.
$boutonSortieOuvrir = New-Object System.Windows.Forms.Button
$boutonSortieOuvrir.Text = "Ouvrir"
$boutonSortieOuvrir.Location = New-Object System.Drawing.Point(1038, 87)
$boutonSortieOuvrir.Size = New-Object System.Drawing.Size(90, 26)
$boutonSortieOuvrir.BackColor = $THEME.FondPanneau
$boutonSortieOuvrir.ForeColor = $THEME.Texte
$boutonSortieOuvrir.FlatStyle = "Flat"
$boutonSortieOuvrir.Add_Click({
    if (-not $champSortie.Text -or -not (Test-Path $champSortie.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Dossier de sortie introuvable (pas encore cree ? lance au moins un traitement, ou verifie le chemin).", "Fusion figurines")
        return
    }
    Start-Process -FilePath "explorer.exe" -ArgumentList "`"$($champSortie.Text)`""
})
$form.Controls.Add($boutonSortieOuvrir)
Arrondir-Bouton $boutonSortieOuvrir

[void](Creer-Label "Bambu Studio (.exe) :" 12 124)
$champBambuExe = New-Object System.Windows.Forms.TextBox
$champBambuExe.Location = New-Object System.Drawing.Point(12, 144)
$champBambuExe.Size = New-Object System.Drawing.Size(900, 24)
$champBambuExe.BackColor = $THEME.FondChamp
$champBambuExe.ForeColor = $THEME.Texte
$champBambuExe.BorderStyle = "FixedSingle"
$champBambuExe.Text = $script:Config.CheminBambuExe
$champBambuExe.AllowDrop = $true
$champBambuExe.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
})
$champBambuExe.Add_DragDrop({
    param($s, $e)
    $chemins = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if (-not $chemins -or $chemins.Count -eq 0) { return }
    $premier = $chemins[0]
    if ((Test-Path -LiteralPath $premier -PathType Leaf) -and $premier.ToLower().EndsWith(".exe")) {
        $champBambuExe.Text = $premier
    }
})
$form.Controls.Add($champBambuExe)
Arrondir-Bouton $champBambuExe 6
$boutonBambuExe = New-Object System.Windows.Forms.Button
$boutonBambuExe.Text = "Parcourir..."
$boutonBambuExe.Location = New-Object System.Drawing.Point(920, 143)
$boutonBambuExe.Size = New-Object System.Drawing.Size(110, 26)
$boutonBambuExe.BackColor = $THEME.FondPanneau
$boutonBambuExe.ForeColor = $THEME.Texte
$boutonBambuExe.FlatStyle = "Flat"
$boutonBambuExe.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Executable (*.exe)|*.exe"
    if ($dlg.ShowDialog() -eq "OK") { $champBambuExe.Text = $dlg.FileName }
})
$form.Controls.Add($boutonBambuExe)
Arrondir-Bouton $boutonBambuExe

$boutonScanner = New-Object System.Windows.Forms.Button
$boutonScanner.Text = "Scanner le dossier racine"
$boutonScanner.Location = New-Object System.Drawing.Point(12, 184)
$boutonScanner.Size = New-Object System.Drawing.Size(220, 32)
$boutonScanner.BackColor = $THEME.AccentPrincipal
$boutonScanner.ForeColor = $THEME.AccentPrincipalTxt
$boutonScanner.FlatStyle = "Flat"
# 27/10/2026 : couleur d'accent semantique (pas une couleur "neutre" de theme) - garde cette
# couleur quel que soit le mode sombre/clair, voir Repeindre-ControlesTheme.
$boutonScanner.Tag = "AccentFixe"
$form.Controls.Add($boutonScanner)
Arrondir-Bouton $boutonScanner

$boutonTraiter = New-Object System.Windows.Forms.Button
$boutonTraiter.Text = "Traiter la selection"
$boutonTraiter.Location = New-Object System.Drawing.Point(244, 184)
$boutonTraiter.Size = New-Object System.Drawing.Size(180, 32)
$boutonTraiter.BackColor = [System.Drawing.Color]::FromArgb(70, 160, 110)
$boutonTraiter.ForeColor = [System.Drawing.Color]::White
$boutonTraiter.FlatStyle = "Flat"
$boutonTraiter.Tag = "AccentFixe"
$form.Controls.Add($boutonTraiter)
Arrondir-Bouton $boutonTraiter

$boutonTraiterTout = New-Object System.Windows.Forms.Button
$boutonTraiterTout.Text = "Traiter tout"
$boutonTraiterTout.Location = New-Object System.Drawing.Point(436, 184)
$boutonTraiterTout.Size = New-Object System.Drawing.Size(140, 32)
$boutonTraiterTout.BackColor = [System.Drawing.Color]::FromArgb(70, 160, 110)
$boutonTraiterTout.ForeColor = [System.Drawing.Color]::White
$boutonTraiterTout.FlatStyle = "Flat"
$boutonTraiterTout.Tag = "AccentFixe"
$form.Controls.Add($boutonTraiterTout)
Arrondir-Bouton $boutonTraiterTout

$boutonArreter = New-Object System.Windows.Forms.Button
$boutonArreter.Text = "Arreter"
$boutonArreter.Location = New-Object System.Drawing.Point(588, 184)
$boutonArreter.Size = New-Object System.Drawing.Size(110, 32)
$boutonArreter.BackColor = [System.Drawing.Color]::FromArgb(190, 70, 70)
$boutonArreter.ForeColor = [System.Drawing.Color]::White
$boutonArreter.FlatStyle = "Flat"
$boutonArreter.Tag = "AccentFixe"
$boutonArreter.Enabled = $false
$form.Controls.Add($boutonArreter)
Arrondir-Bouton $boutonArreter

# 26/10/2026 : --orient (recherche de la meilleure orientation par Bambu Studio) peut
# prendre 10 a 25+ minutes PAR figurine sur un maillage a ~1M triangles (il teste plein
# d'orientations candidates et recalcule les surplombs/supports a chaque fois sur tout le
# maillage) - gel reel de l'app meme passe en arriere-plan, juste "sans bloquer la fenetre".
# Decochee par defaut (demande de Tomyn) : fusion + arrange + export restent automatiques et
# rapides (quelques secondes), l'orientation se fait alors a la main dans Bambu Studio comme
# avant si besoin. Cochable au cas par cas quand on est pret a attendre.
$caseOrientation = New-Object System.Windows.Forms.CheckBox
$caseOrientation.Text = "Orientation auto (lent, 10-25+ min/figurine sur gros maillage)"
$caseOrientation.Location = New-Object System.Drawing.Point(710, 189)
$caseOrientation.Size = New-Object System.Drawing.Size(420, 24)
$caseOrientation.ForeColor = $THEME.Texte
$caseOrientation.Checked = $false
$form.Controls.Add($caseOrientation)

# 27/10/2026 : demande de Tomyn - "reprise intelligente" (voir le filtrage dans
# Lancer-Traitement) : coche par defaut, decochable pour forcer un retraitement complet meme
# si la sortie semble deja a jour (ex: pour re-tester apres un correctif du script lui-meme).
$caseSauterDejaTraitees = New-Object System.Windows.Forms.CheckBox
$caseSauterDejaTraitees.Text = "Sauter les figurines deja traitees (sortie plus recente que les pieces)"
$caseSauterDejaTraitees.Location = New-Object System.Drawing.Point(710, 217)
$caseSauterDejaTraitees.Size = New-Object System.Drawing.Size(420, 24)
$caseSauterDejaTraitees.ForeColor = $THEME.Texte
$caseSauterDejaTraitees.Checked = [bool]$script:Config.SauterDejaTraitees
$caseSauterDejaTraitees.Add_CheckedChanged({
    $script:Config.SauterDejaTraitees = $caseSauterDejaTraitees.Checked
    Sauver-Config
})
$form.Controls.Add($caseSauterDejaTraitees)

# 27/10/2026 : sur un gros lot, inutile de tout relancer si seulement 1 ou 2 figurines ont
# echoue - ce bouton ne relance que celles dont le statut indique un echec/erreur.
$boutonRelancerEchecs = New-Object System.Windows.Forms.Button
$boutonRelancerEchecs.Text = "Relancer les echecs"
$boutonRelancerEchecs.Location = New-Object System.Drawing.Point(12, 220)
$boutonRelancerEchecs.Size = New-Object System.Drawing.Size(220, 32)
$boutonRelancerEchecs.BackColor = [System.Drawing.Color]::FromArgb(235, 165, 60)
$boutonRelancerEchecs.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 34)
$boutonRelancerEchecs.FlatStyle = "Flat"
$boutonRelancerEchecs.Tag = "AccentFixe"
$form.Controls.Add($boutonRelancerEchecs)
Arrondir-Bouton $boutonRelancerEchecs

# 27/10/2026 : chaque echec garde son dossier temporaire pour diagnostic (voir
# Traiter-UneFigurineCore) - ca s'accumule dans %TEMP% avec le temps. Nettoyage manuel plutot
# qu'automatique pour ne jamais supprimer un dossier que Tomyn voulait encore inspecter.
$boutonNettoyerTemp = New-Object System.Windows.Forms.Button
$boutonNettoyerTemp.Text = "Nettoyer dossiers temp"
$boutonNettoyerTemp.Location = New-Object System.Drawing.Point(244, 220)
$boutonNettoyerTemp.Size = New-Object System.Drawing.Size(220, 32)
$boutonNettoyerTemp.BackColor = $THEME.FondPanneau
$boutonNettoyerTemp.ForeColor = $THEME.Texte
$boutonNettoyerTemp.FlatStyle = "Flat"
$form.Controls.Add($boutonNettoyerTemp)
Arrondir-Bouton $boutonNettoyerTemp

# 27/10/2026 : demande de Tomyn - seuil de "piece pas vraiment en contact" (voir
# Test-PiecesEnContact) reglable depuis l'interface au lieu d'etre fige a 0.3mm en dur.
# Certaines figurines ont des pieces volontairement tres rapprochees (aimants, charnieres,
# encoches) qui declenchaient l'avertissement a tort avec une valeur unique pour tout le monde.
[void](Creer-Label "Seuil contact (mm) :" 486 224)
$champSeuilContact = New-Object System.Windows.Forms.NumericUpDown
$champSeuilContact.Location = New-Object System.Drawing.Point(486, 244)
$champSeuilContact.Size = New-Object System.Drawing.Size(70, 24)
$champSeuilContact.DecimalPlaces = 2
$champSeuilContact.Increment = 0.1
$champSeuilContact.Minimum = 0.0
$champSeuilContact.Maximum = 10.0
$champSeuilContact.BackColor = $THEME.FondChamp
$champSeuilContact.ForeColor = $THEME.Texte
$champSeuilContact.Value = [Math]::Max(0.0, [Math]::Min(10.0, [double]$script:Config.SeuilContactMm))
$champSeuilContact.Add_ValueChanged({
    $script:Config.SeuilContactMm = [double]$champSeuilContact.Value
    Sauver-Config
})
$form.Controls.Add($champSeuilContact)
$boutonNettoyerTemp.Add_Click({
    $dossiersTemp = Get-ChildItem -Path ([System.IO.Path]::GetTempPath()) -Directory -Filter "fusionfig_*" -ErrorAction SilentlyContinue
    if (-not $dossiersTemp -or $dossiersTemp.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun dossier temporaire de fusion a nettoyer.", "Fusion figurines")
        return
    }
    $tailleTotale = ($dossiersTemp | Get-ChildItem -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    $texteTaille = if ($tailleTotale) { "~$([Math]::Round($tailleTotale / 1MB, 1)) Mo" } else { "taille inconnue" }
    $reponse = [System.Windows.Forms.MessageBox]::Show(
        "$($dossiersTemp.Count) dossier(s) temporaire(s) trouve(s) ($texteTaille au total) - la plupart viennent de traitements en echec conserves pour diagnostic.`n`nLes supprimer ?",
        "Fusion figurines", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($reponse -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $nbSupprimes = 0; $nbEchecs = 0
    foreach ($d in $dossiersTemp) {
        try { Remove-Item -Path $d.FullName -Recurse -Force -ErrorAction Stop; $nbSupprimes++ }
        catch { $nbEchecs++ }
    }
    Ajouter-Journal "Nettoyage dossiers temp : $nbSupprimes supprime(s)$(if ($nbEchecs -gt 0) { ", $nbEchecs en erreur (probablement encore en cours d'utilisation)" })."
    [System.Windows.Forms.MessageBox]::Show("$nbSupprimes dossier(s) supprime(s).$(if ($nbEchecs -gt 0) { " $nbEchecs n'ont pas pu etre supprimes." })", "Fusion figurines")
})

$liste = New-Object System.Windows.Forms.ListView
$liste.Location = New-Object System.Drawing.Point(12, 270)
# 27/10/2026 : retrecie (520 -> 350) pour liberer de la place en dessous pour le panneau
# "pieces de la figurine" - voir plus bas (demande de Tomyn : le clic droit + menu + fenetre
# modale pour exclure une piece etait trop lourd, remplace par un panneau direct dans la
# fenetre principale).
$liste.Size = New-Object System.Drawing.Size(560, 350)
$liste.View = "Details"
$liste.FullRowSelect = $true
$liste.GridLines = $false
$liste.BorderStyle = "FixedSingle"
$liste.BackColor = $THEME.FondChamp
$liste.ForeColor = $THEME.Texte

# 27/10/2026 : BUG TROUVE - $menuListeContextuel etait utilise plus bas (assigne a
# $liste.ContextMenuStrip, puis des items ajoutes dessus) sans avoir jamais ete cree avec
# New-Object. En PowerShell, assigner une proprietes a une variable $null ne plante pas
# (ContextMenuStrip restait silencieusement a $null, donc le clic droit n'affichait AUCUN
# menu), mais l'ajout d'item ($menuListeContextuel.Items.Add(...)) levait une erreur "cannot
# call a method on a null-valued expression" qui passait inapercue (non bloquante au niveau
# du script, juste affichee en console si une console est visible). Resultat : clic droit sur
# une figurine ne faisait rigoureusement rien, impossible d'exclure des pieces ou d'ouvrir le
# resultat par ce biais (seul le double-clic pour ouvrir le resultat fonctionnait encore, lui
# cable directement sur $liste et non sur ce menu).
$menuListeContextuel = New-Object System.Windows.Forms.ContextMenuStrip
$liste.Columns.Add("Figurine", 190) | Out-Null
$liste.Columns.Add("Pieces", 55) | Out-Null
$liste.Columns.Add("Triangles", 85) | Out-Null
$liste.Columns.Add("Statut", 210) | Out-Null
$form.Controls.Add($liste)
Arrondir-Bouton $liste 6

# 27/10/2026 : demande de Tomyn - meme le bouton "Exclure la selection" etait encore une
# etape de trop : cocher une piece doit l'exclure IMMEDIATEMENT, sans clic supplementaire.
# Remplace par l'evenement ItemCheck de la CheckedListBox (voir plus bas, juste apres
# Remplir-ListePieces) qui applique l'exclusion des que l'etat d'une case change.
[void](Creer-Label "Pieces de la figurine selectionnee (coche = EXCLUE immediatement) :" 12 628)
$listePieces = New-Object System.Windows.Forms.CheckedListBox
$listePieces.Location = New-Object System.Drawing.Point(12, 650)
# 27/10/2026 : retrecie (440 -> 220) pour liberer la place de l'apercu de la piece
# selectionnee juste a droite (demande de Tomyn - savoir visuellement laquelle il regarde
# avant de decider de l'exclure). Hauteur agrandie (130 -> 160) en meme temps que l'apercu
# ci-dessous, demande trop petit pour etre lisible a la premiere taille.
$listePieces.Size = New-Object System.Drawing.Size(220, 160)
$listePieces.BackColor = $THEME.FondChamp
$listePieces.ForeColor = $THEME.Texte
$listePieces.BorderStyle = "FixedSingle"
# 29/10/2026 : demande de Tomyn - certains noms de pieces sont trop longs pour la largeur fixe
# de cette liste (220px) et se coupaient sans aucun moyen de voir la fin, rendant impossible de
# savoir laquelle est laquelle avant de l'exclure. HorizontalScrollbar calcule automatiquement
# (CheckedListBox n'est pas en dessin personnalise ici) la largeur necessaire au nom le plus
# long et affiche un ascenseur horizontal des que ca depasse, sans toucher a la taille du
# panneau lui-meme.
$listePieces.HorizontalScrollbar = $true
# 27/10/2026 : CheckOnClick = $true cochait/decochait des qu'on cliquait N'IMPORTE OU sur la
# ligne (meme juste pour la selectionner/la regarder) - Tomyn veut pouvoir selectionner une
# piece sans que ca change son exclusion, et ne (dé)cocher QUE quand il clique precisement sur
# la case. $false redonne ce comportement standard : clic sur le texte = simple selection,
# clic sur la case (ou barre espace sur l'item selectionne) = coche/decoche.
$listePieces.CheckOnClick = $false
$form.Controls.Add($listePieces)
Arrondir-Bouton $listePieces 6

# 27/10/2026 : demande de Tomyn - meme apercu que celui d'entree/sortie, mais pour la SEULE
# piece actuellement selectionnee dans $listePieces (pas toute la figurine), pour voir de
# quelle piece il s'agit avant de decider de l'exclure ou non.
[void](Creer-Label "Piece selectionnee :" 244 628)
$apercuPiece = New-Object System.Windows.Forms.PictureBox
$apercuPiece.Location = New-Object System.Drawing.Point(244, 650)
# 27/10/2026 : agrandi (160x130 -> 216x160) - demande de Tomyn, trop petit pour etre
# discernable confortablement. Le budget de triangles (voir l'appel a Demarrer-ApercuAsync
# plus bas) reste bas pour garder le rendu quasi instantane malgre la taille plus grande.
$apercuPiece.Size = New-Object System.Drawing.Size(216, 160)
$apercuPiece.BackColor = $THEME.FondChamp
$apercuPiece.BorderStyle = "FixedSingle"
$apercuPiece.SizeMode = "Zoom"
$form.Controls.Add($apercuPiece)
Arrondir-Bouton $apercuPiece 8

$boutonToutInclure = New-Object System.Windows.Forms.Button
$boutonToutInclure.Text = "Tout reinclure"
$boutonToutInclure.Location = New-Object System.Drawing.Point(472, 650)
$boutonToutInclure.Size = New-Object System.Drawing.Size(100, 76)
$boutonToutInclure.BackColor = $THEME.FondPanneau
$boutonToutInclure.ForeColor = $THEME.Texte
$boutonToutInclure.FlatStyle = "Flat"
$form.Controls.Add($boutonToutInclure)
Arrondir-Bouton $boutonToutInclure

# 29/10/2026 : demande de Tomyn - symetrique de "Tout reinclure" : utile quand il ne veut
# garder qu'une ou deux pieces sur un gros lot (plus rapide de tout exclure puis recocher les
# 1-2 bonnes que de decocher une a une toutes les mauvaises). Garde toujours la DERNIERE piece
# incluse (meme regle que le garde-fou de $listePieces.Add_ItemCheck plus bas : "il faut garder
# au moins une piece incluse") plutot que de se retrouver avec zero piece a fusionner.
# 29/10/2026 : CORRECTIF - place a cote de "Tout reinclure" (x=578), ca mordait sur l'apercu de
# sortie qui commence juste apres a x=584 (vu par Tomyn sur une photo de son ecran : le bouton
# recouvrait carrement la case "Apercu sortie"). Empile desormais SOUS "Tout reinclure" dans la
# meme colonne (x=472) au lieu de cote, pas de risque de deborder sur l'apercu.
$boutonToutExclure = New-Object System.Windows.Forms.Button
$boutonToutExclure.Text = "Tout exclure"
$boutonToutExclure.Location = New-Object System.Drawing.Point(472, 734)
$boutonToutExclure.Size = New-Object System.Drawing.Size(100, 76)
$boutonToutExclure.BackColor = $THEME.FondPanneau
$boutonToutExclure.ForeColor = $THEME.Texte
$boutonToutExclure.FlatStyle = "Flat"
$form.Controls.Add($boutonToutExclure)
Arrondir-Bouton $boutonToutExclure

# 27/10/2026 : demande de Tomyn - ouvrir directement le resultat (.3mf) d'une figurine deja
# traitee pour verifier l'assemblage dans Bambu Studio, sans fouiller dans l'explorateur.
# Double-clic sur la ligne OU clic droit -> "Ouvrir le resultat".
function Ouvrir-ResultatSelectionne {
    if ($liste.SelectedItems.Count -eq 0) { return }
    $grp = $liste.SelectedItems[0].Tag
    if (-not $grp.CheminSortieStl) {
        [System.Windows.Forms.MessageBox]::Show("Cette figurine n'a pas (encore) ete traitee avec succes.", "Fusion figurines")
        return
    }
    $chemin3mf = [System.IO.Path]::ChangeExtension($grp.CheminSortieStl, ".3mf")
    try {
        if (Test-Path $chemin3mf) { Start-Process -FilePath $chemin3mf }
        else { Start-Process -FilePath $grp.CheminSortieStl }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Impossible d'ouvrir le resultat : $($_.Exception.Message)", "Fusion figurines")
    }
}
$liste.ContextMenuStrip = $menuListeContextuel
$liste.Add_DoubleClick({ Ouvrir-ResultatSelectionne })
# 27/10/2026 : l'item de menu "Ouvrir le resultat" lui-meme n'avait jamais ete ajoute (seul le
# double-clic etait cable) - corrige en meme temps que le bug $menuListeContextuel ci-dessus.
[void]$menuListeContextuel.Items.Add("Ouvrir le resultat (.3mf)").Add_Click({ Ouvrir-ResultatSelectionne })
# 27/10/2026 : clic droit sur une ligne ne la selectionnait pas automatiquement (comportement
# par defaut d'un ListView WinForms - seul un clic GAUCHE selectionne) : le menu s'ouvrait donc
# sur la selection precedente (ou aucune selection), ce qui pouvait sembler "ne rien faire" si
# aucune ligne n'avait encore ete cliquee une premiere fois. On selectionne maintenant la ligne
# sous le curseur des le clic droit, avant que le menu ne s'affiche.
$liste.Add_MouseUp({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $item = $liste.GetItemAt($e.X, $e.Y)
        if ($item) {
            $liste.SelectedItems.Clear()
            $item.Selected = $true
            $item.Focused = $true
        }
    }
})

# 27/10/2026 : demande de Tomyn - l'ancienne version (clic droit -> menu -> fenetre modale a
# part) etait jugee trop lourde. Remplacee par le panneau $listePieces directement dans la
# fenetre principale (voir sa creation plus haut) : cette fonction le remplit avec les pieces
# de la figurine passee en parametre, coche = piece deja exclue.
function Remplir-ListePieces($grp) {
    $script:RemplissageListePieces = $true
    $listePieces.Items.Clear()
    if ($grp) {
        $exclues = if ($grp.PSObject.Properties['PiecesExclues']) { @($grp.PiecesExclues) } else { @() }
        foreach ($piece in $grp.Pieces) {
            $index = $listePieces.Items.Add($piece.Name)
            $listePieces.SetItemChecked($index, ($exclues -contains $piece.Name))
        }
    }
    $script:RemplissageListePieces = $false
    # 27/10/2026 : la liste a change de figurine (ou est videe) - l'apercu piece-par-piece
    # affiche encore l'ancienne selection, qui n'a plus de sens ici.
    Demarrer-ApercuAsync "piece" $apercuPiece @()
}

# 27/10/2026 : met a jour uniquement la colonne "Pieces" de la ligne concernee dans $liste,
# SANS passer par Rafraichir-ListeGroupes (qui fait $liste.Items.Clear() + re-ajoute tout -
# ca perdait la selection a CHAQUE case cochee, donc $listePieces se vidait immediatement
# apres chaque clic - exactement le genre de "faut cliquer 2 fois" que Tomyn ne voulait plus).
function Calculer-NombreTrianglesInclus($grp) {
    # 27/10/2026 : recalcule le nombre de triangles en ne comptant que les pieces NON exclues -
    # c'est ce nombre-la (pas le total de toutes les pieces) qui doit s'afficher dans le
    # tableau, puisque c'est ce qui sera reellement envoye a la fusion. Se base sur
    # TrianglesParPiece (compte par piece deja fait au scan, voir $boutonScanner.Add_Click) -
    # instantane, aucune relecture de fichier necessaire ici.
    $exclues = if ($grp.PSObject.Properties['PiecesExclues']) { @($grp.PiecesExclues) } else { @() }
    if (-not $grp.PSObject.Properties['TrianglesParPiece']) { return $grp.NombreTriangles }
    $total = 0
    foreach ($p in $grp.Pieces) {
        if ($exclues -contains $p.Name) { continue }
        if ($grp.TrianglesParPiece.ContainsKey($p.Name)) { $total += $grp.TrianglesParPiece[$p.Name] }
    }
    return $total
}

function Mettre-A-Jour-LigneGroupe($grp) {
    # 27/10/2026 : le nombre de triangles affiche doit suivre les exclusions (voir
    # Calculer-NombreTrianglesInclus) - recalcule et stocke dans NombreTriangles AVANT de
    # rafraichir la ligne, pour rester coherent avec Rafraichir-ListeGroupes et avec ce que lira
    # Traiter-UneFigurineCore au lancement.
    $grp.NombreTriangles = Calculer-NombreTrianglesInclus $grp
    foreach ($item in $liste.Items) {
        if ($item.Tag -eq $grp) {
            # 29/10/2026 : ajoute en meme temps que le renommage avant fusion (voir
            # Lancer-Traitement) - cette fonction ne touchait avant que les colonnes 2-4, jamais
            # le nom (colonne principale), donc un $grp.Nom change en cours de route ne se
            # voyait jamais dans le tableau.
            $item.Text = $grp.Nom
            $nbExclues = if ($grp.PSObject.Properties['PiecesExclues']) { $grp.PiecesExclues.Count } else { 0 }
            $item.SubItems[1].Text = if ($nbExclues -gt 0) { "$($grp.Pieces.Count - $nbExclues)/$($grp.Pieces.Count)" } else { [string]$grp.Pieces.Count }
            $texteTriangles = "{0:N0}" -f $grp.NombreTriangles
            if ($grp.NombreTriangles -ge $script:SeuilGrosMaillage) { $texteTriangles = "$texteTriangles !" }
            $item.SubItems[2].Text = $texteTriangles
            # 28/10/2026 : la colonne "Statut" n'etait jamais rafraichie par cette fonction -
            # necessaire pour que le statut "Deja a jour (sautee)" (reprise intelligente) soit
            # visible immediatement dans le tableau, sans attendre une autre action qui
            # redeclencherait un rafraichissement complet.
            $item.SubItems[3].Text = $grp.Statut
            break
        }
    }
}

# 27/10/2026 : demande de Tomyn - cocher une case doit exclure IMMEDIATEMENT, sans bouton a
# cliquer en plus. ItemCheck se declenche AVANT que l'etat de la case soit reellement applique
# (e.NewValue donne la nouvelle valeur, GetItemChecked($e.Index) donnerait encore l'ANCIENNE) -
# on calcule donc la liste finale en substituant $e.NewValue pour l'index concerne.
$script:RemplissageListePieces = $false
$listePieces.Add_ItemCheck({
    param($s, $e)
    if ($script:RemplissageListePieces) { return }
    if ($liste.SelectedItems.Count -eq 0) { return }
    $grp = $liste.SelectedItems[0].Tag
    if (-not $grp) { return }
    $nouvellesExclues = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $listePieces.Items.Count; $i++) {
        $coche = if ($i -eq $e.Index) { $e.NewValue -eq [System.Windows.Forms.CheckState]::Checked } else { $listePieces.GetItemChecked($i) }
        if ($coche) { $nouvellesExclues.Add($listePieces.Items[$i]) }
    }
    if ($nouvellesExclues.Count -ge $grp.Pieces.Count) {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
        [System.Windows.Forms.MessageBox]::Show("Il faut garder au moins une piece incluse.", "Fusion figurines")
        return
    }
    $grp.PiecesExclues = @($nouvellesExclues)
    # 29/10/2026 : CORRECTIF signale par Tomyn - l'apercu de sortie (voir Rafraichir-ApercusPour)
    # affiche le VRAI resultat des que $grp.CheminSortieStl pointe vers un fichier existant, sans
    # jamais vérifier s'il correspond encore a la selection ACTUELLE de pieces. Dans un dossier
    # avec plusieurs figurines (meme groupe reutilise avec des exclusions differentes a chaque
    # fois, voir le renommage plus bas dans Lancer-Traitement), changer la selection apres un
    # premier traitement gardait l'ancien resultat affiche au lieu de repasser sur l'estimation -
    # Tomyn voyait la fusion precedente au lieu d'un apercu de celle qu'il est en train de
    # preparer. On invalide donc ce resultat memorise des que la selection change : il ne
    # redeviendra valide qu'apres un nouveau traitement reussi (voir $res.CheminSortieStl plus
    # bas). Le statut repasse aussi a "En attente" pour la meme raison (plus de "OK" trompeur
    # affiche pour une selection qui n'a pas encore ete traitee).
    $grp.CheminSortieStl = $null
    $grp.Statut = "En attente"
    if ($nouvellesExclues.Count -gt 0) {
        Ajouter-Journal "[$($grp.Nom)] $($nouvellesExclues.Count) piece(s) exclue(s) de la fusion : $($nouvellesExclues -join ', ')"
    } else {
        Ajouter-Journal "[$($grp.Nom)] Toutes les pieces sont incluses."
    }
    Mettre-A-Jour-LigneGroupe $grp
    Rafraichir-ApercusPour $grp
})

# 27/10/2026 : demande de Tomyn - voir un apercu de la piece sur laquelle il clique (pas toute
# la figurine), pour savoir ce qu'il regarde avant de decider de l'exclure. Separe de
# Add_ItemCheck ci-dessus : ceci se declenche en cliquant sur le TEXTE de la ligne (simple
# selection, CheckOnClick = $false), l'autre en cochant la case.
$listePieces.Add_SelectedIndexChanged({
    if ($listePieces.SelectedIndex -lt 0 -or $liste.SelectedItems.Count -eq 0) {
        Demarrer-ApercuAsync "piece" $apercuPiece @()
        return
    }
    $grp = $liste.SelectedItems[0].Tag
    if (-not $grp) { Demarrer-ApercuAsync "piece" $apercuPiece @(); return }
    $nomPiece = $listePieces.SelectedItem
    $piece = $grp.Pieces | Where-Object { $_.Name -eq $nomPiece } | Select-Object -First 1
    if (-not $piece) { Demarrer-ApercuAsync "piece" $apercuPiece @(); return }
    # 27/10/2026 : demande de Tomyn - cet apercu-ci sert juste a identifier la piece, pas a
    # evaluer le resultat final comme entree/sortie : budget de triangles et resolution bien
    # plus bas (une seule piece, pas besoin de 80000 triangles/320x240 pour la reconnaitre)
    # pour que ca s'affiche quasi instantanement en cliquant.
    # 27/10/2026 : "rien ne change" rapporte avec la rotation Z precedente - confirme que ces
    # pieces (spirales/colonnes) sont a symetrie de revolution verticale, donc une rotation
    # autour de l'axe vertical ne POUVAIT rien changer a la silhouette projetee. Remplacee par
    # une inclinaison autour de l'axe X (on penche la piece vers la camera, 35 degres) qui,
    # elle, change toujours la silhouette (fait apparaitre le dessus/dessous).
    # 27/10/2026 : demande de Tomyn - en mode clair, les "trous" entre points epars (budget de
    # triangles bas = points moins denses) ressortaient bien plus que sur fond sombre. Bloc de
    # points plus gros (3x3 au lieu de 2x2) pour cet apercu-ci uniquement, pour compenser SANS
    # augmenter le budget de triangles (donc sans ralentir le clic).
    Demarrer-ApercuAsync "piece" $apercuPiece @($piece.FullName) $false 15000 216 160 35 5
})

$boutonToutInclure.Add_Click({
    if ($liste.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Selectionne d'abord une figurine dans la liste.", "Fusion figurines")
        return
    }
    $grp = $liste.SelectedItems[0].Tag
    $grp.PiecesExclues = @()
    # 29/10/2026 : voir le commentaire dans Add_ItemCheck plus haut - meme invalidation
    # necessaire ici, ce bouton change aussi la selection de pieces.
    $grp.CheminSortieStl = $null
    $grp.Statut = "En attente"
    Ajouter-Journal "[$($grp.Nom)] Toutes les pieces sont incluses."
    Remplir-ListePieces $grp
    Mettre-A-Jour-LigneGroupe $grp
    Rafraichir-ApercusPour $grp
})

$boutonToutExclure.Add_Click({
    if ($liste.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Selectionne d'abord une figurine dans la liste.", "Fusion figurines")
        return
    }
    $grp = $liste.SelectedItems[0].Tag
    if ($grp.Pieces.Count -le 1) {
        [System.Windows.Forms.MessageBox]::Show("Il faut garder au moins une piece incluse.", "Fusion figurines")
        return
    }
    # Garde la DERNIERE piece incluse (voir commentaire a la creation du bouton) - toutes les
    # autres passent en exclues.
    $grp.PiecesExclues = @($grp.Pieces | Select-Object -First ($grp.Pieces.Count - 1) | ForEach-Object { $_.Name })
    # 29/10/2026 : voir le commentaire dans Add_ItemCheck plus haut - meme invalidation
    # necessaire ici, ce bouton change aussi la selection de pieces.
    $grp.CheminSortieStl = $null
    $grp.Statut = "En attente"
    Ajouter-Journal "[$($grp.Nom)] Toutes les pieces exclues sauf '$($grp.Pieces[-1].Name)' (clic sur 'Tout exclure')."
    Remplir-ListePieces $grp
    Mettre-A-Jour-LigneGroupe $grp
    Rafraichir-ApercusPour $grp
})

[void](Creer-Label "Apercu entree (pieces telles que detectees)" 584 270)
$apercuEntree = New-Object System.Windows.Forms.PictureBox
$apercuEntree.Location = New-Object System.Drawing.Point(584, 292)
$apercuEntree.Size = New-Object System.Drawing.Size(320, 240)
$apercuEntree.BackColor = $THEME.FondChamp
$apercuEntree.BorderStyle = "FixedSingle"
$apercuEntree.SizeMode = "Zoom"
$form.Controls.Add($apercuEntree)
Arrondir-Bouton $apercuEntree 8

$labelApercuSortie = Creer-Label "Apercu sortie (estimation - pas encore traite)" 584 540
$apercuSortie = New-Object System.Windows.Forms.PictureBox
$apercuSortie.Location = New-Object System.Drawing.Point(584, 562)
$apercuSortie.Size = New-Object System.Drawing.Size(320, 240)
$apercuSortie.BackColor = $THEME.FondChamp
$apercuSortie.BorderStyle = "FixedSingle"
$apercuSortie.SizeMode = "Zoom"
$form.Controls.Add($apercuSortie)
Arrondir-Bouton $apercuSortie 8

[void](Creer-Label "Journal :" 920 270)
$journal = New-Object System.Windows.Forms.TextBox
$journal.Location = New-Object System.Drawing.Point(920, 292)
$journal.Size = New-Object System.Drawing.Size(240, 510)
$journal.Multiline = $true
$journal.ScrollBars = "Vertical"
$journal.ReadOnly = $true
$journal.BackColor = $THEME.FondChamp
$journal.ForeColor = $THEME.Texte
$journal.BorderStyle = "FixedSingle"
$journal.Font = New-Object System.Drawing.Font("Consolas", 8)
$form.Controls.Add($journal)
Arrondir-Bouton $journal 6

# 27/10/2026 : demande de Tomyn - exporter le contenu complet du journal dans un .txt,
# pratique pour me le transmettre en entier apres un traitement plutot que du copier-coller.
$boutonExporterJournal = New-Object System.Windows.Forms.Button
$boutonExporterJournal.Text = "Exporter journal (.txt)"
$boutonExporterJournal.Location = New-Object System.Drawing.Point(920, 810)
$boutonExporterJournal.Size = New-Object System.Drawing.Size(240, 30)
$boutonExporterJournal.BackColor = $THEME.FondPanneau
$boutonExporterJournal.ForeColor = $THEME.Texte
$boutonExporterJournal.FlatStyle = "Flat"
$form.Controls.Add($boutonExporterJournal)
Arrondir-Bouton $boutonExporterJournal
$boutonExporterJournal.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Fichier texte (*.txt)|*.txt"
    $dlg.FileName = "journal-fusion-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
    if ($champSortie.Text -and (Test-Path $champSortie.Text)) { $dlg.InitialDirectory = $champSortie.Text }
    if ($dlg.ShowDialog() -eq "OK") {
        try {
            [System.IO.File]::WriteAllText($dlg.FileName, $journal.Text, [System.Text.Encoding]::UTF8)
            [System.Windows.Forms.MessageBox]::Show("Journal exporte : $($dlg.FileName)", "Fusion figurines")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Echec de l'export du journal : $($_.Exception.Message)", "Fusion figurines")
        }
    }
})

function Ajouter-Journal([string]$texte) {
    $horodatage = Get-Date -Format "HH:mm:ss"
    $ligne = "[$horodatage] $texte"
    $journal.AppendText("$ligne`r`n")
    # 29/10/2026 : persiste aussi sur disque (voir $script:CheminJournalFichier plus haut) -
    # Add-Content ouvre/ferme le fichier a chaque appel plutot que de garder un handle ouvert,
    # ce qui coute un peu plus cher par ligne mais garantit que rien n'est perdu si l'appli
    # crashe juste apres (pas de buffer en memoire a vidanger qui ne le sera jamais).
    if ($script:CheminJournalFichier) {
        try { Add-Content -Path $script:CheminJournalFichier -Value $ligne -Encoding UTF8 -ErrorAction Stop }
        catch { $script:CheminJournalFichier = $null }
    }
}

# ============================================================================
# SCAN DES GROUPES DE PIECES
# ============================================================================
$script:Groupes = @()

function Rafraichir-ListeGroupes {
    # 29/10/2026 : CORRECTIF signale par Tomyn - $liste.Items.Clear() ci-dessous vide aussi la
    # SELECTION, ce qui cassait silencieusement le rafraichissement automatique de l'apercu
    # apres un traitement : le "finally" du minuteur de traitement appelle CETTE fonction, PUIS
    # verifie "$liste.SelectedItems[0].Tag -eq $grp" pour savoir s'il doit rebasculer l'apercu
    # sur le vrai resultat - sauf que plus rien n'est selectionne a ce moment-la (vide par la
    # ligne juste en dessous), donc cette condition etait TOUJOURS fausse et l'apercu restait
    # bloque sur l'ancien contenu ("estimation") meme apres un traitement reussi. On memorise
    # les groupes selectionnes avant de vider la liste, et on les re-selectionne une fois
    # repeuplee.
    $groupesSelectionnes = @($liste.SelectedItems | ForEach-Object { $_.Tag })
    $liste.Items.Clear()
    foreach ($grp in $script:Groupes) {
        $ligne = New-Object System.Windows.Forms.ListViewItem($grp.Nom)
        $nbExclues = if ($grp.PSObject.Properties['PiecesExclues']) { $grp.PiecesExclues.Count } else { 0 }
        $texteP = if ($nbExclues -gt 0) { "$($grp.Pieces.Count - $nbExclues)/$($grp.Pieces.Count)" } else { [string]$grp.Pieces.Count }
        [void]$ligne.SubItems.Add($texteP)
        $texteTriangles = if ($grp.PSObject.Properties['NombreTriangles']) { "{0:N0}" -f $grp.NombreTriangles } else { "?" }
        if ($grp.PSObject.Properties['NombreTriangles'] -and $grp.NombreTriangles -ge $script:SeuilGrosMaillage) { $texteTriangles = "$texteTriangles !" }
        [void]$ligne.SubItems.Add($texteTriangles)
        [void]$ligne.SubItems.Add($grp.Statut)
        $ligne.Tag = $grp
        $liste.Items.Add($ligne) | Out-Null
        if ($groupesSelectionnes -contains $grp) { $ligne.Selected = $true }
    }
}

# 26/10/2026 : au-dela de ce nombre de triangles (toutes pieces d'une figurine
# cumulees), on previent Tomyn par un toast : l'orientation automatique (--orient 1)
# devient tres lente sur un maillage de cette taille (voir journal du 26/10 : 10-25+
# minutes sur 1 030 520 triangles) - mieux vaut laisser la case decochee sur ces
# figurines-la et orienter a la main dans Bambu Studio si besoin.
$script:SeuilGrosMaillage = 300000

function Compter-TrianglesSTL([string]$chemin) {
    # Lecture rapide du nombre de triangles d'un STL BINAIRE : 80 octets d'en-tete puis
    # un uint32 (nombre de triangles) - PAS besoin de lire tout le maillage pour ca, donc
    # instantane meme sur un gros fichier (contrairement a Lire-Maillage3MF qui lui doit
    # tout parcourir). Retourne -1 si format non reconnu (ex: STL ASCII, rare en pratique
    # ici) pour ne pas fausser le total plutot que de compter 0 a tort.
    try {
        $flux = [System.IO.File]::OpenRead($chemin)
        try {
            if ($flux.Length -lt 84) { return -1 }
            $entete = New-Object byte[] 5
            [void]$flux.Read($entete, 0, 5)
            if ([System.Text.Encoding]::ASCII.GetString($entete) -eq "solid") { return -1 }
            [void]$flux.Seek(80, [System.IO.SeekOrigin]::Begin)
            $bufCompte = New-Object byte[] 4
            [void]$flux.Read($bufCompte, 0, 4)
            return [int][System.BitConverter]::ToUInt32($bufCompte, 0)
        } finally { $flux.Close() }
    } catch { return -1 }
}

# ============================================================================
# SUPPORT .OBJ (demande de Tomyn - certains createurs ne fournissent les pieces qu'en .obj,
# format texte "Wavefront", pas en .stl). Contrairement au .stl binaire, .obj n'a pas d'en-tete
# donnant le nombre de triangles ni de taille de record fixe permettant de "sauter" directement
# a un triangle donne (Seek) - on doit forcement lire le fichier ligne par ligne. Les sommets
# ("v x y z") et les faces ("f i1 i2 i3 ..." - indices 1-based vers les sommets, potentiellement
# suivis de "/texture/normale" a ignorer, et potentiellement plus de 3 indices pour un polygone
# non triangulaire - auquel cas on triangule en eventail depuis le premier sommet) sont lus une
# seule fois ; impossible d'echantillonner AVANT lecture comme pour le .stl binaire.
# ============================================================================
function Compter-TrianglesOBJ([string]$chemin) {
    # Equivalent de Compter-TrianglesSTL pour un .obj : pas d'en-tete a lire, donc on compte les
    # faces ligne par ligne (StreamReader, jamais tout le fichier en memoire d'un coup) en
    # tenant compte de la triangulation en eventail (une face a N sommets devient N-2 triangles).
    try {
        if (-not (Test-Path $chemin)) { return -1 }
        $total = 0
        $lecteur = New-Object System.IO.StreamReader($chemin)
        try {
            while (-not $lecteur.EndOfStream) {
                $ligne = $lecteur.ReadLine()
                if ($ligne.Length -ge 2 -and $ligne[0] -eq 'f' -and [char]::IsWhiteSpace($ligne[1])) {
                    $nbSommets = ($ligne.Trim() -split '\s+').Count - 1
                    if ($nbSommets -ge 3) { $total += ($nbSommets - 2) }
                }
            }
        } finally { $lecteur.Close() }
        return $total
    } catch { return -1 }
}

function Lire-TrianglesOBJEchantillon([string]$cheminFichier, [int]$maxTriangles) {
    # Pas de raccourci "Seek direct sur le Neme triangle" possible en texte (voir commentaire
    # plus haut) : on lit TOUS les sommets (necessaires, une face peut referencer n'importe quel
    # sommet deja vu, pas seulement les precedents) puis on triangule chaque face au fur et a
    # mesure. L'echantillonnage (budget $maxTriangles) ne reduit donc pas la LECTURE comme pour
    # le .stl binaire, seulement le nombre de triangles GARDES en memoire pour le rendu/la
    # fusion - un sur N au lieu de tous, meme principe que la branche ASCII du lecteur .stl.
    $sommets = [System.Collections.Generic.List[double[]]]::new()
    $triangles = [System.Collections.Generic.List[object]]::new()
    $nombreOriginal = 0
    try {
        if (-not (Test-Path $cheminFichier)) { return [PSCustomObject]@{ Triangles = $triangles; NombreOriginal = 0 } }
        # Premiere passe : compte total pour calculer le pas d'echantillonnage (coherent avec
        # Lire-TrianglesSTLEchantillon, qui connait le total via l'en-tete AVANT de lire).
        $nombreOriginal = Compter-TrianglesOBJ $cheminFichier
        if ($nombreOriginal -lt 0) { $nombreOriginal = 0 }
        $pas = [Math]::Max(1, [Math]::Floor($nombreOriginal / [Math]::Max(1, $maxTriangles)))
        $indexTriangle = -1
        $lecteur = New-Object System.IO.StreamReader($cheminFichier)
        try {
            while (-not $lecteur.EndOfStream) {
                $ligne = $lecteur.ReadLine()
                if ($ligne.Length -lt 2) { continue }
                if ($ligne[0] -eq 'v' -and [char]::IsWhiteSpace($ligne[1])) {
                    $parties = $ligne.Trim() -split '\s+'
                    if ($parties.Count -ge 4) {
                        $sommets.Add([double[]]@([double]$parties[1], [double]$parties[2], [double]$parties[3]))
                    }
                } elseif ($ligne[0] -eq 'f' -and [char]::IsWhiteSpace($ligne[1])) {
                    $jetons = ($ligne.Trim() -split '\s+') | Select-Object -Skip 1
                    # "i", "i/vt", "i/vt/vn" ou "i//vn" - seul l'indice de SOMMET (avant le
                    # premier "/" s'il y en a un) nous interesse. Indices 1-based (standard OBJ),
                    # et negatifs autorises par le format (compte depuis la fin de la liste des
                    # sommets DEJA lus) - les deux cas sont geres.
                    $indicesSommets = foreach ($j in $jetons) {
                        $brut = ($j -split '/')[0]
                        if ($brut) {
                            $i = [int]$brut
                            if ($i -gt 0) { $i - 1 } else { $sommets.Count + $i }
                        }
                    }
                    $indicesSommets = @($indicesSommets)
                    if ($indicesSommets.Count -ge 3) {
                        # Triangulation en eventail : sommet 0 commun a tous les triangles de la
                        # face, puis (1,2), (2,3), ... - correct pour un polygone convexe plan
                        # (quasi systematique sur ce genre de pieces imprimables), pas garanti
                        # pour un polygone concave, mais sans alternative moins couteuse sans
                        # vraie triangulation geometrique.
                        for ($k = 1; $k -lt $indicesSommets.Count - 1; $k++) {
                            $indexTriangle++
                            if (($indexTriangle % $pas) -ne 0) { continue }
                            if ($triangles.Count -ge $maxTriangles) { continue }
                            $ia = $indicesSommets[0]; $ib = $indicesSommets[$k]; $ic = $indicesSommets[$k + 1]
                            if ($ia -ge 0 -and $ia -lt $sommets.Count -and $ib -ge 0 -and $ib -lt $sommets.Count -and $ic -ge 0 -and $ic -lt $sommets.Count) {
                                $triangles.Add(@{ A = $sommets[$ia]; B = $sommets[$ib]; C = $sommets[$ic] })
                            }
                        }
                    }
                }
            }
        } finally { $lecteur.Close() }
    } catch { }
    return [PSCustomObject]@{ Triangles = $triangles; NombreOriginal = $nombreOriginal }
}

# 29/10/2026 : demande de Tomyn - detecte automatiquement .stl vs .obj a partir de l'extension
# et appelle la bonne paire de fonctions, pour que le reste du code (scan, apercus, fusion
# locale) n'ait jamais besoin de savoir quel format chaque piece utilise.
function Compter-TrianglesPiece([string]$chemin) {
    if ($chemin -match '\.obj$') { return Compter-TrianglesOBJ $chemin }
    return Compter-TrianglesSTL $chemin
}

function Lire-TrianglesPieceEchantillon([string]$chemin, [int]$maxTriangles) {
    if ($chemin -match '\.obj$') { return Lire-TrianglesOBJEchantillon $chemin $maxTriangles }
    return Lire-TrianglesSTLEchantillon $chemin $maxTriangles
}

# ============================================================================
# TOAST (notification non-bloquante, coin bas-droit de la fenetre, auto-fermeture)
# ============================================================================
function Afficher-Toast([string]$texte, [int]$dureeMs = 8000) {
    $toast = New-Object System.Windows.Forms.Form
    $toast.FormBorderStyle = "None"
    $toast.ShowInTaskbar = $false
    $toast.TopMost = $true
    $toast.StartPosition = "Manual"
    $toast.BackColor = [System.Drawing.Color]::FromArgb(70, 54, 20)
    $toast.Size = New-Object System.Drawing.Size(420, 120)

    $barreAccent = New-Object System.Windows.Forms.Panel
    $barreAccent.Location = New-Object System.Drawing.Point(0, 0)
    $barreAccent.Size = New-Object System.Drawing.Size(6, $toast.Height)
    $barreAccent.BackColor = [System.Drawing.Color]::FromArgb(235, 165, 60)
    $toast.Controls.Add($barreAccent)

    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text = "Attention"
    $lblTitre.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
    $lblTitre.ForeColor = [System.Drawing.Color]::FromArgb(235, 165, 60)
    $lblTitre.BackColor = [System.Drawing.Color]::Transparent
    $lblTitre.AutoSize = $false
    $lblTitre.Location = New-Object System.Drawing.Point(22, 10)
    $lblTitre.Size = New-Object System.Drawing.Size(380, 20)
    $toast.Controls.Add($lblTitre)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $texte
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(255, 225, 190)
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $lbl.AutoSize = $false
    $lbl.Location = New-Object System.Drawing.Point(22, 32)
    $lbl.Size = New-Object System.Drawing.Size(384, 78)
    $toast.Controls.Add($lbl)

    $toast.Location = New-Object System.Drawing.Point(
        ($form.Location.X + $form.Width - $toast.Width - 24),
        ($form.Location.Y + $form.Height - $toast.Height - 24))

    Arrondir-Bouton $toast 14

    $minuteurFermeture = New-Object System.Windows.Forms.Timer
    $minuteurFermeture.Interval = $dureeMs
    $gestionnaireFermeture = {
        $minuteurFermeture.Stop(); $minuteurFermeture.Dispose()
        try { $toast.Close() } catch {}
    }.GetNewClosure()
    $minuteurFermeture.Add_Tick($gestionnaireFermeture)

    # fermeture anticipee si on clique sur le toast
    $fermerAuClic = { $gestionnaireFermeture.Invoke() }.GetNewClosure()
    $toast.Add_Click($fermerAuClic)
    $lbl.Add_Click($fermerAuClic)
    $lblTitre.Add_Click($fermerAuClic)

    $toast.Show()
    $toast.BringToFront()
    $minuteurFermeture.Start()
}

$boutonScanner.Add_Click({
    $script:Config.DossierRacine = $champRacine.Text
    $script:Config.DossierSortie = $champSortie.Text
    $script:Config.CheminBambuExe = $champBambuExe.Text
    Sauver-Config

    if (-not (Test-Path $champRacine.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Dossier racine invalide.", "Fusion figurines")
        return
    }
    # 29/10/2026 : demande de Tomyn - un scan avec des .obj a mis pres de 10 minutes alors que
    # le comptage de triangles seul (mesure isolee) prend moins d'1s meme sur un .obj de
    # plusieurs Mo : la lenteur ne vient donc PAS forcement du parsing lui-meme, mais le journal
    # ne donnait jusqu'ici aucune info pendant le scan (silence total entre le clic et "Scan
    # termine"). On chronometre maintenant le scan en entier ET chaque piece individuellement,
    # pour que le PROCHAIN blocage pointe directement le fichier/l'etape responsable au lieu de
    # deviner a l'aveugle.
    $chronoScan = [System.Diagnostics.Stopwatch]::StartNew()
    $script:Groupes = @()

    # 26/10/2026 : CAS 1 - le dossier racine CONTIENT LUI-MEME directement les pieces (ex:
    # Tomyn pointe directement sur le dossier d'UNE figurine, pas sur son parent). Dans ce
    # cas on traite le dossier racine lui-meme comme UN groupe.
    # 29/10/2026 : demande de Tomyn - certaines figurines ne fournissent leurs pieces qu'en
    # .obj (pas de .stl du tout) - voir Compter-TrianglesPiece/Lire-TrianglesPieceEchantillon
    # plus haut, qui detectent le format a la volee pour le reste du pipeline. Le scan doit
    # donc chercher les DEUX extensions, pas juste .stl.
    $piecesRacine = Get-ChildItem -Path $champRacine.Text -File | Where-Object { $_.Extension -in '.stl', '.obj' }
    if ($piecesRacine.Count -ge 2) {
        $script:Groupes += [PSCustomObject]@{
            Nom = Split-Path $champRacine.Text -Leaf; Dossier = $champRacine.Text; Pieces = $piecesRacine; Statut = "En attente"; CheminSortieStl = $null; DureeSecondes = $null; PiecesExclues = @()
        }
    }

    # CAS 2 - le dossier racine contient PLUSIEURS figurines, chacune dans son propre
    # sous-dossier (usage normal pour un lot).
    # 29/10/2026 : demande de Tomyn - un sous-dossier avec 0 ou 1 .stl est ecarte ici SANS
    # aucune trace dans le journal, impossible jusqu'ici de distinguer "normal, figurine a
    # piece unique, rien a fusionner" de "louche, fichier manquant/mal copie". On garde la
    # liste pour un recap en fin de scan plutot que d'avertir dossier par dossier (ca
    # spammerait le journal dans un lot ou les pieces uniques sont frequentes et attendues).
    $sousDossiersIgnores = [System.Collections.Generic.List[string]]::new()
    $sousDossiers = Get-ChildItem -Path $champRacine.Text -Directory
    foreach ($sd in $sousDossiers) {
        $pieces = Get-ChildItem -Path $sd.FullName -File | Where-Object { $_.Extension -in '.stl', '.obj' }
        if ($pieces.Count -ge 2) {
            $script:Groupes += [PSCustomObject]@{
                Nom = $sd.Name; Dossier = $sd.FullName; Pieces = $pieces; Statut = "En attente"; CheminSortieStl = $null; DureeSecondes = $null; PiecesExclues = @()
            }
        } else {
            $sousDossiersIgnores.Add("$($sd.Name) ($($pieces.Count) .stl/.obj)")
        }
    }

    # 26/10/2026 : comptage rapide des triangles par figurine (juste les en-tetes STL,
    # instantane) pour avertir sur les gros maillages AVANT de lancer un traitement avec
    # orientation auto qui serait tres long.
    $figurinesVolumineuses = [System.Collections.Generic.List[string]]::new()
    # 27/10/2026 : "Poseidons Brutes - Swords" a deja casse la ligne de commande a cause d'un
    # " - " dans le nom de dossier (voir plus bas Traiter-UneFigurineCore, correctif des
    # guillemets). Les guillemets sont desormais poses nous-memes, mais un GUILLEMET DROIT
    # DANS le nom lui-meme casserait a nouveau le guillemetage - et quelques autres caracteres
    # speciaux a Windows (< > | ^ %) peuvent aussi poser probleme selon le contexte. On
    # avertit au scan plutot que de laisser decouvrir ca a l'echec du traitement.
    $caracteresProblematiques = '["<>|\^%]'
    foreach ($grp in $script:Groupes) {
        $total = 0
        # 27/10/2026 : avant, un .stl vide/corrompu ne se decouvrait qu'au moment de son
        # traitement reel - potentiellement apres une longue attente dans la file. On le
        # signale maintenant des le scan : -1 = format non reconnu (STL ASCII, fichier trop
        # petit, erreur de lecture), 0 = STL binaire valide mais sans triangle - dans les deux
        # cas la piece ne contribuera rien a la fusion.
        $piecesSuspectes = [System.Collections.Generic.List[string]]::new()
        # 27/10/2026 : demande de Tomyn - le nombre de triangles affiche dans le tableau doit
        # refleter ce qui sera REELLEMENT fusionne (pieces incluses seulement), pas le total de
        # toutes les pieces qu'on exclue ou pas. On garde donc le compte PAR PIECE (deja lu ici,
        # quasi gratuit) pour pouvoir recalculer instantanement a chaque case cochee/decochee
        # (voir Calculer-NombreTrianglesInclus) sans avoir a rouvrir les fichiers.
        $trianglesParPiece = @{}
        foreach ($p in $grp.Pieces) {
            # 29/10/2026 : CAUSE TROUVEE du blocage signale par Tomyn - pas un bug, un fichier
            # .obj de 51 millions de points (export ZBrush brut, "NOT_DECIMATION") qui a mis
            # 113s a etre juste COMPTE, pendant lesquels rien ne s'affiche (voir plus bas, le
            # chrono ne logue qu'APRES coup). Un .obj n'a pas d'en-tete a lire comme un .stl : sa
            # taille sur disque EST deja l'info qu'on peut utiliser pour avertir AVANT de se
            # lancer dans la lecture, au lieu de laisser Tomyn croire que tout est plante.
            if ($p.Extension -eq '.obj' -and $p.Length -ge 80MB) {
                Ajouter-Journal "[$($grp.Nom)] '$($p.Name)' fait $([Math]::Round($p.Length / 1MB)) Mo (.obj) - le comptage va prendre du temps (pas un blocage, verifie si c'est bien un fichier decime/utile a garder)."
            }
            $chronoPiece = [System.Diagnostics.Stopwatch]::StartNew()
            $n = Compter-TrianglesPiece $p.FullName
            $chronoPiece.Stop()
            # 29/10/2026 : seuil de 1s - voir commentaire sur $chronoScan plus haut. On ne logue
            # QUE les pieces anormalement lentes (pas chacune) pour ne pas noyer le journal sur
            # un gros lot ou tout va bien.
            if ($chronoPiece.Elapsed.TotalSeconds -ge 1) {
                Ajouter-Journal "[$($grp.Nom)] Comptage de '$($p.Name)' anormalement lent : $([Math]::Round($chronoPiece.Elapsed.TotalSeconds, 1))s."
            }
            if ($n -ge 0) { $total += $n }
            if ($n -le 0) { $piecesSuspectes.Add($p.Name) }
            $trianglesParPiece[$p.Name] = [Math]::Max(0, $n)
        }
        if ($piecesSuspectes.Count -gt 0) {
            Ajouter-Journal "[$($grp.Nom)] ATTENTION : piece(s) vide(s) ou illisible(s) - $($piecesSuspectes -join ', ') - verifie le(s) fichier(s) ou exclus-le(s) via clic droit > 'Choisir les pieces a inclure' avant de lancer."
        }
        $grp | Add-Member -NotePropertyName TrianglesParPiece -NotePropertyValue $trianglesParPiece -Force
        $grp | Add-Member -NotePropertyName NombreTriangles -NotePropertyValue $total -Force
        if ($total -ge $script:SeuilGrosMaillage) {
            $figurinesVolumineuses.Add("$($grp.Nom) ($([int]($total/1000))k triangles)")
            Ajouter-Journal "[$($grp.Nom)] ATTENTION gros maillage : $total triangles - laisser 'Orientation auto' decochee, sinon ca peut prendre 10-25+ minutes."
        }
        if ($grp.Dossier -match $caracteresProblematiques) {
            Ajouter-Journal "[$($grp.Nom)] ATTENTION : le chemin du dossier contient un caractere qui peut casser la ligne de commande Bambu Studio ($($grp.Dossier)) - si le traitement echoue bizarrement, essaie de renommer le dossier sans guillemets/chevrons/barres."
        }
    }

    Rafraichir-ListeGroupes
    Remplir-ListePieces $null
    $chronoScan.Stop()
    Ajouter-Journal "Scan termine en $([Math]::Round($chronoScan.Elapsed.TotalSeconds, 1))s : $($script:Groupes.Count) figurine(s) multi-pieces detectee(s) (sous-dossiers avec 2 .stl ou plus)."
    if ($sousDossiersIgnores.Count -gt 0) {
        Ajouter-Journal "$($sousDossiersIgnores.Count) sous-dossier(s) ignore(s) (0 ou 1 .stl, rien a fusionner) : $($sousDossiersIgnores -join ', ')"
    }

    if ($figurinesVolumineuses.Count -gt 0) {
        Afficher-Toast "Maillage volumineux detecte :`r`n$($figurinesVolumineuses -join "`r`n")`r`n`r`nLaisse 'Orientation auto' decochee pour ces figurines (sinon ca peut prendre 10-25+ minutes chacune)."
    }
})

# ============================================================================
# APERCU ASYNCHRONE
# 26/10/2026 : CORRECTIF "fenetre ne repond plus" - avant, l'apercu (lecture STL +
# rendu isometrique) se faisait directement dans le gestionnaire SelectedIndexChanged,
# donc SUR LE THREAD DE L'INTERFACE : tant que le calcul tournait (gros maillage =
# plusieurs secondes, voire plus), Windows considerait la fenetre comme plantee
# ("ne repond pas") puisque plus aucun message (clic, redessin...) n'etait traite.
# Fix : le calcul se fait dans un Runspace separe (arriere-plan), et un Timer
# WinForms (qui, lui, tourne sur le thread de l'UI via la boucle de messages, donc
# sans bloquer) verifie regulierement si c'est termine pour poser l'image resultante.
# ============================================================================
$script:ApercuJobs = @{}
$script:ApercuRunspacesPartagees = @{}

function Obtenir-RunspacePartagee([string]$cle) {
    # 26/10/2026 : CORRECTIF lenteur systematique (meme sur un petit fichier) - creer un
    # Runspace "par defaut" via [runspacefactory]::CreateRunspace() (sans argument)
    # charge/scanne TOUS les modules PowerShell installes sur la machine avant de pouvoir
    # executer quoi que ce soit (souvent plusieurs secondes, parfois beaucoup plus avec un
    # antivirus qui inspecte chaque module). Comme ca se refaisait a CHAQUE clic dans la
    # liste, chaque apercu payait ce cout de demarrage, meme pour un fichier minuscule.
    # Fix : un runspace "minimal" (CreateDefault2 - pas de chargement de modules), cree
    # UNE SEULE FOIS et reutilise pour tous les apercus suivants.
    if (-not $script:ApercuRunspacesPartagees.ContainsKey($cle)) {
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
        $rs = [runspacefactory]::CreateRunspace($iss)
        # 26/10/2026 : STA -> MTA. Le STA ne servait qu'au rendu WPF/Media3D (retire du
        # pipeline, voir Rendre-ApercuWPF). Ce qui tourne ici desormais (lecture STL +
        # System.Drawing/GDI+ pur) n'a aucun besoin d'un appartement COM STA. Or creer un
        # Runspace STA depuis le thread UI (qui est LUI-MEME STA, comme tout WinForms) est
        # un scenario classique de deadlock croise entre appartements STA si les deux
        # threads doivent se synchroniser (meme sans COM explicite, le simple fait d'ouvrir
        # un 2e appartement STA peut bloquer en attendant une pompe de messages qui ne vient
        # jamais) - ca correspond exactement au symptome observe (appli entierement figee,
        # meme la capture d'ecran bloque). MTA supprime ce risque : les threads MTA n'ont
        # pas cette contrainte de synchronisation par appartement.
        $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
        $rs.Open()
        $script:ApercuRunspacesPartagees[$cle] = $rs
    }
    return $script:ApercuRunspacesPartagees[$cle]
}

function Fermer-RunspacesPartagees {
    foreach ($rs in $script:ApercuRunspacesPartagees.Values) {
        try { $rs.Close(); $rs.Dispose() } catch {}
    }
    $script:ApercuRunspacesPartagees.Clear()
}

function Creer-ImagePlaceholder([string]$texte, [int]$largeur = 320, [int]$hauteur = 240) {
    $bmp = New-Object System.Drawing.Bitmap($largeur, $hauteur)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear($THEME.FondChamp)
    $g.DrawString($texte, (New-Object System.Drawing.Font("Segoe UI", 9)), (New-Object System.Drawing.SolidBrush($THEME.TexteAttenue)), 10, 10)
    $g.Dispose()
    return $bmp
}

function Arreter-ApercuJob([string]$cle) {
    $job = $script:ApercuJobs[$cle]
    if (-not $job) { return }
    try { $job.Minuteur.Stop() } catch {}
    try { $job.Minuteur.Dispose() } catch {}
    try { $job.PS.Stop() } catch {}
    try { $job.PS.Dispose() } catch {}
    # 29/10/2026 : CORRECTIF - signale par Tomyn (cascade d'erreurs "fichier introuvable" puis
    # "expression Null" puis "depassement de capacite", a des lignes DIFFERENTES de la MEME
    # fonction, pour UN SEUL fichier qui existait pourtant bel et bien, juste apres avoir coche/
    # decoche plusieurs cases d'exclusion rapidement). Cause : $job.PS.Stop() ci-dessus NE
    # GARANTIT PAS que le script encore en cours s'arrete reellement avant de rendre la main -
    # une E/S bloquante (OpenRead) n'est pas interrompue par Stop(). Avant ce correctif, le
    # Runspace PARTAGE etait reutilise IMMEDIATEMENT pour le PROCHAIN apercu (commentaire
    # original : "pour eviter de repayer son cout de demarrage") alors que l'ancien script
    # pouvait encore tourner dessus en arriere-plan - deux scripts executes EN MEME TEMPS sur
    # le MEME Runspace corrompent son etat interne, ce qui correspond exactement aux erreurs en
    # cascade observees. On abandonne desormais ce Runspace (sans le FERMER de force - Close()
    # sur un Runspace occupe par une E/S bloquante risquerait a son tour de geler l'interface,
    # meme probleme que Stop() plus haut) : le prochain apercu via Obtenir-RunspacePartagee en
    # recreera un NOUVEAU et PROPRE. L'ancien finira (ou se fera tuer) de son cote et sera
    # recupere par le garbage collector, sans jamais etre reutilise en parallele d'un autre.
    $script:ApercuRunspacesPartagees.Remove($cle)
    $script:ApercuJobs.Remove($cle)
}

function Demarrer-ApercuAsync([string]$cle, [System.Windows.Forms.PictureBox]$picturebox, [string[]]$cheminsFichiers, [bool]$EcarterPieces = $false, [int]$BudgetTotal = 80000, [int]$Largeur = 320, [int]$Hauteur = 240, [double]$InclinaisonXDegres = 0, [int]$TailleBloc = 2) {
    # Annule le calcul precedent pour cette image (ex: l'utilisateur a change de
    # selection dans la liste avant que le rendu precedent ait fini).
    Arreter-ApercuJob $cle

    $ancienneImage = $picturebox.Image
    if (-not $cheminsFichiers -or $cheminsFichiers.Count -eq 0) {
        $picturebox.Image = $null
        if ($ancienneImage) { $ancienneImage.Dispose() }
        return
    }
    $picturebox.Image = Creer-ImagePlaceholder "Chargement de l'apercu..."
    if ($ancienneImage) { $ancienneImage.Dispose() }

    # 26/10/2026 : le rendu WPF/GPU (Rendre-ApercuWPF) a ete RETIRE du pipeline - echec
    # reproductible (meme avec un maillage plafonne/raisonnable) qui echappait au try/catch
    # PowerShell (probablement un probleme de thread/boucle de messages WPF specifique a ce
    # contexte d'arriere-plan sans fenetre). La fonction reste dans le fichier pour reference
    # future mais n'est plus appelee. On revient au rendu logiciel echantillonne + gonflage
    # (Lire-TrianglesSTLEchantillon + Rendre-ApercuIsometrique) comme chemin UNIQUE : plus
    # simple, et on sait qu'il produit effectivement une image.
    #
    # On recopie les VRAIES definitions actuelles de ces deux fonctions (celles deja
    # chargees dans CETTE session) pour que le calcul en arriere-plan fasse exactement la
    # meme chose que le code principal - pas de copie figee qui risquerait de diverger.
    $defLireSTL = (Get-Item function:Lire-TrianglesSTLEchantillon).Definition
    $defRendre = (Get-Item function:Rendre-ApercuIsometrique).Definition
    # 29/10/2026 : support .obj (voir Compter-TrianglesPiece/Lire-TrianglesPieceEchantillon
    # plus haut) - ce runspace a besoin de sa propre copie de ces fonctions aussi, meme
    # principe de transmission que pour les deux ci-dessus.
    $defCompterOBJ = (Get-Item function:Compter-TrianglesOBJ).Definition
    $defLireOBJ = (Get-Item function:Lire-TrianglesOBJEchantillon).Definition
    $defCompterPiece = (Get-Item function:Compter-TrianglesPiece).Definition
    $defLirePiece = (Get-Item function:Lire-TrianglesPieceEchantillon).Definition
    # 26/10/2026 : le test a 120000 (x3) etait encore plus lent ET toujours de la "bouillie"
    # - ca confirme que FillPolygon+gonflage etait le vrai probleme (vitesse ET fidelite), pas
    # le budget. Nouveau rendu par sommets (buffer brut, pas d'objets .NET ni d'appel GDI+
    # par sommet - voir Rendre-ApercuIsometrique) : mesure isolee (hors GUI, juste la logique
    # de rendu) sur 300000 triangles synthetiques = 19-30s rien que pour cette passe - encore
    # trop pour "quelques secondes". On part sur une valeur plus mesuree (80000, ~5-8s
    # estimes par extrapolation) plutot que de re-tomber dans le piege "ca va forcement aller
    # plus vite" sans mesure. A ajuster avec les stats + la duree reelle que tu rapporteras.
    # 27/10/2026 : $BudgetTotal est maintenant parametrable - l'apercu d'une SEULE piece
    # (panneau d'exclusion) n'a besoin que d'identifier visuellement la piece, pas d'une
    # fidelite "qualite fusion" : Tomyn le veut quasi instantane, donc appele avec un budget
    # bien plus bas (voir le gestionnaire de selection de $listePieces).
    $budgetParFichier = [Math]::Max(1000, [Math]::Floor($BudgetTotal / [Math]::Max(1, $cheminsFichiers.Count)))

    # IMPORTANT : $THEME est une variable GLOBALE du script principal - elle n'existe PAS
    # dans ce runspace separe (session independante). Rendre-ApercuIsometrique l'utilise
    # (couleurs de fond/texte), donc on doit la lui transmettre explicitement, sinon
    # $THEME.FondChamp vaut $null et ca plante.
    $scriptTexte = 'param($CheminsFichiers, $BudgetParFichier, $THEME, $EcarterPieces, $Largeur, $Hauteur, $InclinaisonXDegres, $TailleBloc)' + "`r`n" +
        'function Lire-TrianglesSTLEchantillon {' + "`r`n" + $defLireSTL + "`r`n" + '}' + "`r`n" +
        'function Compter-TrianglesOBJ {' + "`r`n" + $defCompterOBJ + "`r`n" + '}' + "`r`n" +
        'function Lire-TrianglesOBJEchantillon {' + "`r`n" + $defLireOBJ + "`r`n" + '}' + "`r`n" +
        'function Compter-TrianglesPiece {' + "`r`n" + $defCompterPiece + "`r`n" + '}' + "`r`n" +
        'function Lire-TrianglesPieceEchantillon {' + "`r`n" + $defLirePiece + "`r`n" + '}' + "`r`n" +
        'function Rendre-ApercuIsometrique {' + "`r`n" + $defRendre + "`r`n" + '}' + "`r`n" +
        @'
$tousTriangles = [System.Collections.Generic.List[object]]::new()
$sommeOriginal = 0
$sommeEchantillon = 0
# 27/10/2026 : demande de Tomyn - l'apercu d'une piece ISOLEE (pas en groupe) peut tomber
# visuellement "a plat" avec la camera isometrique fixe. PREMIERE tentative : tourner autour de
# l'axe Z (vertical) - SANS EFFET VISIBLE rapporte par Tomyn, ce qui confirme l'hypothese : une
# piece a symetrie de revolution autour de l'axe vertical (frequent sur ce genre de figurines -
# spirales, colonnes, tiges) garde exactement la MEME silhouette quel que soit l'angle de
# rotation autour de CET axe-la, donc ca ne pouvait rien changer. Corrige en tournant plutot
# autour de l'axe X (on "penche" la piece vers la camera) : ca, ca change TOUJOURS la
# silhouette projetee, symetrie de revolution verticale ou pas, puisque ca fait apparaitre le
# dessus/dessous de la piece. N'affecte que l'apercu de piece individuelle (rotation = 0 pour
# entree/sortie), ne touche pas au vrai fichier.
$angleRotX = $InclinaisonXDegres * [Math]::PI / 180.0
$cosRotX = [Math]::Cos($angleRotX); $sinRotX = [Math]::Sin($angleRotX)
# 26/10/2026 : demande de Tomyn - l'apercu d'ENTREE doit montrer les pieces ECARTEES (comme
# si elles n'etaient pas encore assemblees), pas a leurs positions d'origine qui se touchent/
# s'emboitent deja (ca ressemblait a la figurine finale, trompeur pour un "avant"). L'apercu
# de SORTIE, lui, continue de montrer le vrai resultat fusionne tel quel ($EcarterPieces
# vaut $false pour cet appel - voir Demarrer-ApercuAsync). On pose juste chaque piece a cote
# de la precedente le long de X, avec un petit espace, en utilisant SA PROPRE boite englobante
# - purement visuel, ca ne touche pas du tout au vrai traitement (Bambu Studio continue de
# recevoir les fichiers originaux, intacts).
$decalageXSuivant = 0.0
foreach ($chemin in $CheminsFichiers) {
    $resultatLecture = Lire-TrianglesPieceEchantillon $chemin $BudgetParFichier
    if ($angleRotX -ne 0) {
        foreach ($t in $resultatLecture.Triangles) {
            foreach ($pt in @($t.A, $t.B, $t.C)) {
                $y0 = $pt[1]; $z0 = $pt[2]
                $pt[1] = $y0 * $cosRotX - $z0 * $sinRotX
                $pt[2] = $y0 * $sinRotX + $z0 * $cosRotX
            }
        }
    }
    if ($EcarterPieces -and $resultatLecture.Triangles.Count -gt 0) {
        $minX = [double]::MaxValue; $maxX = [double]::MinValue
        foreach ($t in $resultatLecture.Triangles) {
            foreach ($pt in @($t.A, $t.B, $t.C)) {
                if ($pt[0] -lt $minX) { $minX = $pt[0] }
                if ($pt[0] -gt $maxX) { $maxX = $pt[0] }
            }
        }
        $largeurPiece = $maxX - $minX
        $decalage = $decalageXSuivant - $minX
        foreach ($t in $resultatLecture.Triangles) {
            foreach ($pt in @($t.A, $t.B, $t.C)) { $pt[0] += $decalage }
        }
        $decalageXSuivant += $largeurPiece + [Math]::Max($largeurPiece * 0.25, 1.0)
    }
    foreach ($t in $resultatLecture.Triangles) { $tousTriangles.Add($t) }
    $sommeOriginal += $resultatLecture.NombreOriginal
    $sommeEchantillon += $resultatLecture.Triangles.Count
}
# 26/10/2026 : plus de "facteur de gonflage" - le nouveau rendu asperge les sommets
# directement (voir Rendre-ApercuIsometrique), la densite de points fait le travail de
# combler les trous, pas une deformation geometrique artificielle apres coup.
$bmpRendu = Rendre-ApercuIsometrique $tousTriangles $Largeur $Hauteur $TailleBloc
# On renvoie aussi les stats (pas juste l'image) pour pouvoir les loguer cote UI et ajuster
# le budget de triangles avec des chiffres reels au lieu de deviner depuis une capture d'ecran.
[PSCustomObject]@{ Image = $bmpRendu; NombreOriginal = $sommeOriginal; NombreEchantillon = $sommeEchantillon }
'@

    $rs = Obtenir-RunspacePartagee $cle
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($scriptTexte).AddArgument(@($cheminsFichiers)).AddArgument($budgetParFichier).AddArgument($THEME).AddArgument($EcarterPieces).AddArgument($Largeur).AddArgument($Hauteur).AddArgument($InclinaisonXDegres).AddArgument($TailleBloc)
    $resultatAsync = $ps.BeginInvoke()

    # 27/10/2026 : capture LOCALE (pas $script:) de la Hashtable partagee, prise AVANT la
    # creation de la closure plus bas - une reassignation/mutation faite via "$script:..."
    # DEPUIS L'INTERIEUR d'un .GetNewClosure() ne se propage pas de maniere fiable vers la
    # vraie variable de script (bug reel trouve et corrige le 27/10/2026 sur le traitement
    # principal - voir le commentaire a la declaration de $script:EtatTraitement). On applique
    # ici le meme correctif par precaution, pour que le nettoyage de $script:ApercuJobs a la
    # fin de l'apercu (plus bas, "finally") soit bien visible de l'exterieur de la closure.
    $apercuJobsPartage = $script:ApercuJobs

    $minuteur = New-Object System.Windows.Forms.Timer
    $minuteur.Interval = 200
    $gestionnaireTick = {
        if (-not $resultatAsync.IsCompleted) { return }
        $minuteur.Stop()
        try {
            $sortie = $ps.EndInvoke($resultatAsync)
            if ($ps.HadErrors) {
                $msgs = $ps.Streams.Error | ForEach-Object {
                    $texte = $_.ToString()
                    if ($_.Exception -and $_.Exception.Message) { $texte = "$texte [$($_.Exception.GetType().FullName)] $($_.Exception.Message)" }
                    if ($_.ScriptStackTrace) { $texte = "$texte -- $($_.ScriptStackTrace -replace '[\r\n]+', ' > ')" }
                    $texte
                }
                if (-not $msgs -or $msgs.Count -eq 0) { $msgs = @("(HadErrors=true mais Streams.Error vide - cause inconnue)") }
                Ajouter-Journal "Erreur apercu ($cle) : $($msgs -join ' | ')"
                $img = $picturebox.Image
                $picturebox.Image = Creer-ImagePlaceholder "(erreur apercu - voir journal)"
                if ($img) { $img.Dispose() }
            } elseif ($sortie -and $sortie.Count -gt 0 -and $sortie[0].Image -is [System.Drawing.Bitmap]) {
                $res = $sortie[0]
                $img = $picturebox.Image
                $picturebox.Image = $res.Image
                if ($img) { $img.Dispose() }
                $pourcent = if ($res.NombreOriginal -gt 0) { [Math]::Round(100.0 * $res.NombreEchantillon / $res.NombreOriginal, 1) } else { 100 }
                Ajouter-Journal "Apercu ($cle) : $($res.NombreEchantillon) / $($res.NombreOriginal) triangles gardes ($pourcent %)"
            } else {
                $img = $picturebox.Image
                $picturebox.Image = Creer-ImagePlaceholder "(pas d'apercu)"
                if ($img) { $img.Dispose() }
            }
        } catch {
            # IMPORTANT : $_.Exception.Message peut etre VIDE pour certaines exceptions
            # levees par EndInvoke (ex: PipelineStoppedException generique) alors que le
            # vrai detail est dans $ps.Streams.Error - on affiche donc TOUJOURS les deux.
            $detailFlux = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join ' | '
            $texteErreur = "[$($_.Exception.GetType().FullName)] $($_.Exception.Message)"
            if ($detailFlux) { $texteErreur = "$texteErreur -- flux erreurs : $detailFlux" }
            Ajouter-Journal "Erreur apercu ($cle) : $texteErreur"
            $img = $picturebox.Image
            $picturebox.Image = Creer-ImagePlaceholder "(erreur apercu - voir journal)"
            if ($img) { $img.Dispose() }
        } finally {
            # Le Runspace ($rs) N'est PAS ferme ici : il est partage/reutilise (voir
            # Obtenir-RunspacePartagee), seul le pipeline PowerShell ($ps) est jetable.
            try { $ps.Dispose() } catch {}
            $minuteur.Dispose()
            $apercuJobsPartage.Remove($cle)
        }
    }.GetNewClosure()
    $minuteur.Add_Tick($gestionnaireTick)

    $script:ApercuJobs[$cle] = [PSCustomObject]@{ Minuteur = $minuteur; PS = $ps }
    $minuteur.Start()
}

function Rafraichir-ApercusPour([PSCustomObject]$grp) {
    # 27/10/2026 : si des pieces ont ete exclues ("Choisir les pieces a inclure..."), les
    # apercus doivent refleter UNIQUEMENT les pieces qui seront reellement envoyees au
    # traitement - sinon l'apercu mentirait sur ce qui va etre fusionne.
    $exclues = if ($grp.PSObject.Properties['PiecesExclues']) { @($grp.PiecesExclues) } else { @() }
    $piecesIncluses = @($grp.Pieces | Where-Object { $exclues -notcontains $_.Name })
    # 27/10/2026 : Tomyn demande d'appliquer aussi aux apercus entree/sortie
    # l'agrandissement des points (3x3 au lieu de 2x2) fait pour l'apercu de
    # piece individuelle - ca ne ralentit pas (la boucle generique est meme
    # plus rapide que l'ancien code deroule a la main), donc autant en profiter
    # partout pour combler les "trous" en mode clair.
    Demarrer-ApercuAsync "entree" $apercuEntree ($piecesIncluses | ForEach-Object { $_.FullName }) $true 20000 320 240 0 5

    # 26/10/2026 : demande de Tomyn - pas la peine d'attendre la fin du traitement Bambu
    # Studio (qui peut prendre plusieurs minutes) pour avoir une idee du resultat assemble :
    # tant qu'il n'y a pas encore de vrai .stl de sortie, on affiche une ESTIMATION en
    # rendant les pieces d'origine SANS l'ecartement (donc a leurs positions telles
    # qu'importees, qui s'emboitent deja normalement - voir EcarterPieces plus haut). Des que
    # le vrai resultat traite existe, on bascule dessus (plus fiable que l'estimation : vraie
    # orientation si la case est cochee, vrai export Bambu Studio).
    if ($grp.CheminSortieStl -and (Test-Path $grp.CheminSortieStl)) {
        $labelApercuSortie.Text = "Apercu sortie (resultat reel, fusionne + traite)"
        Demarrer-ApercuAsync "sortie" $apercuSortie @($grp.CheminSortieStl) $false 20000 320 240 0 5
    } else {
        $labelApercuSortie.Text = "Apercu sortie (estimation - pas encore traite)"
        Demarrer-ApercuAsync "sortie" $apercuSortie ($piecesIncluses | ForEach-Object { $_.FullName }) $false 20000 320 240 0 5
    }
}

$liste.Add_SelectedIndexChanged({
    if ($liste.SelectedItems.Count -eq 0) { Remplir-ListePieces $null; return }
    $grp = $liste.SelectedItems[0].Tag
    if (-not $grp) { Remplir-ListePieces $null; return }
    Rafraichir-ApercusPour $grp
    Remplir-ListePieces $grp
})

# ============================================================================
# DETECTION DE CONTACT REEL ENTRE DEUX PIECES
# ============================================================================
# 27/10/2026 : demande de Tomyn - certaines figurines ont des pieces visuellement
# proches/alignees (l'apercu "entree" les montre cote a cote comme prevu par le
# concepteur) mais dont les surfaces ne se TOUCHENT pas reellement une fois
# fusionnees - le concepteur n'avait pas pense a un assemblage imprime en un seul
# bloc. Le controle existant plus bas (boitesParPiece/$seuilEcart) ne detecte que
# les pieces GROSSIEREMENT ecartees (> 15% de la diagonale globale) - il ne voit
# rien quand l'ecart ne fait que quelques dixiemes/quelques mm, exactement le cas
# qui interesse Tomyn ici. Cette fonction repond a une question plus fine : existe-
# t-il au moins UN point de la piece A a moins de $seuilMm d'un point de la piece B ?
# Approche : grille spatiale (hash) de taille $seuilMm sur la piece la plus petite,
# puis on teste chaque point de l'autre piece contre les 27 cellules voisines (3x3x3)
# et on sort DES qu'un contact est trouve - cas normal (pieces bien assemblees), le
# premier point teste tombe generalement deja dans la zone de contact -> quasi
# instantane. Seules les pieces REELLEMENT disjointes (le cas qu'on cherche a
# detecter) paient le cout du parcours complet.
function Test-PiecesEnContact([object[]]$pointsA, [object[]]$pointsB, [double]$seuilMm) {
    if ($pointsA.Count -eq 0 -or $pointsB.Count -eq 0) { return $false }
    if ($pointsA.Count -gt $pointsB.Count) { $tmp = $pointsA; $pointsA = $pointsB; $pointsB = $tmp }
    $taille = [Math]::Max($seuilMm, 0.01)
    $grille = @{}
    foreach ($p in $pointsA) {
        $cle = "$([Math]::Floor($p[0]/$taille))|$([Math]::Floor($p[1]/$taille))|$([Math]::Floor($p[2]/$taille))"
        if (-not $grille.ContainsKey($cle)) { $grille[$cle] = [System.Collections.Generic.List[object]]::new() }
        $grille[$cle].Add($p)
    }
    $seuilCarre = $seuilMm * $seuilMm
    foreach ($p in $pointsB) {
        $cxB = [Math]::Floor($p[0]/$taille); $cyB = [Math]::Floor($p[1]/$taille); $czB = [Math]::Floor($p[2]/$taille)
        for ($dx = -1; $dx -le 1; $dx++) {
            for ($dy = -1; $dy -le 1; $dy++) {
                for ($dz = -1; $dz -le 1; $dz++) {
                    $cle = "$($cxB+$dx)|$($cyB+$dy)|$($czB+$dz)"
                    if ($grille.ContainsKey($cle)) {
                        foreach ($q in $grille[$cle]) {
                            $ex = $p[0]-$q[0]; $ey = $p[1]-$q[1]; $ez = $p[2]-$q[2]
                            if (($ex*$ex + $ey*$ey + $ez*$ez) -le $seuilCarre) { return $true }
                        }
                    }
                }
            }
        }
    }
    return $false
}

# ============================================================================
# TRAITEMENT D'UNE FIGURINE
# ============================================================================
function Traiter-UneFigurineCore([string]$nomFigurine, [string[]]$cheminsPieces, [string]$dossierSortie, [string]$cheminBambuExe, [bool]$orientationAuto, [hashtable]$etatAnnulation = $null, [datetime]$heureEnvoi = (Get-Date), [double]$seuilContactMm = 0.3) {
    # 27/10/2026 : Tomyn a releve ~30s d'ecart entre "Demarrage" (cote UI, juste apres
    # BeginInvoke) et la premiere ligne reellement ecrite par cette fonction, alors que la
    # fusion locale + Bambu Studio ne prennent que quelques secondes une fois mesures separement
    # (voir $chronoFusion/$chronoBambu plus bas). Ca pointe vers un delai de DISPATCH vers le
    # runspace (parsing du gros bloc de texte contenant plusieurs fonctions, JIT, etc.), pas vers
    # notre propre code. Premiere ligne journalisee plus bas (une fois $journalLignes cree) :
    # combien de temps s'est ecoule AVANT que ce code ne commence vraiment a s'executer.
    $delaiDispatch = (Get-Date) - $heureEnvoi
    # 27/10/2026 : CORRECTIF - "Arreter" ne faisait RIEN tant qu'une figurine etait deja en
    # cours de traitement (Tomyn : "meme le bouton arreter... il va meme pas s'arreter") -
    # $script:DoitArreter n'etait verifie qu'AVANT de lancer la figurine SUIVANTE dans la file,
    # jamais pendant que celle en cours tournait. $etatAnnulation (Hashtable MUTEE, jamais
    # reassignee - meme technique que $etatPartage ailleurs dans le fichier) est le MEME objet
    # cote thread UI et cote ce runspace d'arriere-plan (un objet .NET passe par reference via
    # AddArgument reste le meme objet d'un runspace a l'autre, pas une copie) : le thread UI y
    # pose DoitArreter=$true au clic sur "Arreter", et on le relit ici pour tuer reellement le
    # processus bambu-studio.exe en cours (voir plus bas, pres de Start-Process).
    # 26/10/2026 : version "pure" (sans toucher aux controles WinForms) de l'ancienne
    # Traiter-UneFigurine, pour pouvoir tourner dans un runspace d'arriere-plan - voir
    # Demarrer-ProchainTraitement. Ne fait AUCUN appel a Ajouter-Journal/$grp directement
    # (interdit depuis un autre thread) : accumule les lignes de journal dans une liste et
    # renvoie tout (statut, lignes de journal, chemin de sortie) au thread UI qui les
    # appliquera lui-meme a la fin.
    $journalLignes = [System.Collections.Generic.List[string]]::new()
    $journalLignes.Add("[$nomFigurine] (delai avant demarrage reel du traitement en arriere-plan : $([Math]::Round($delaiDispatch.TotalSeconds, 1)) s)")
    # 27/10/2026 : $delaiDispatch mesure ~0.1s (dispatch runspace quasi instantane) alors que
    # ~29s restent inexpliques entre "Demarrage" et la fin reelle - la SEULE chose qui se passe
    # entre le calcul de $delaiDispatch et le debut de $chronoFusion plus bas, c'est la creation
    # de ce dossier temporaire. Chronometre a part pour confirmer/ecarter (suspect : dossier
    # TEMP charge de dossiers fusionfig_* accumules pendant toute cette session de tests, ou
    # antivirus qui scanne chaque creation de dossier).
    $chronoTempDir = [System.Diagnostics.Stopwatch]::StartNew()
    $dossierTemp = Join-Path ([System.IO.Path]::GetTempPath()) ("fusionfig_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $dossierTemp -Force | Out-Null
    $chronoTempDir.Stop()
    $journalLignes.Add("[$nomFigurine] (duree creation dossier temporaire : $([Math]::Round($chronoTempDir.Elapsed.TotalSeconds, 1)) s)")
    $chemin3mf = Join-Path $dossierTemp "$nomFigurine.3mf"
    # 26/10/2026 : DEUX bugs empilés, trouvés en executant la commande a la main dans un
    # terminal (merci Tomyn) :
    #  1) --orient n'est PAS un interrupteur : comme --arrange, il attend une valeur
    #     ("0-disable, 1-enable, others-auto" - voir --help). On le passait seul -> "setup
    #     params error" -> sortie immediate (code -2). Corrige (voir plus bas : "1").
    #  2) MEME apres ce premier correctif, Start-Process -ArgumentList avec un TABLEAU de
    #     chaines NE MET PAS automatiquement de guillemets autour des elements contenant des
    #     espaces (contrairement a l'intuition) - chaque chemin de fichier etait donc envoye
    #     "brut". Windows decoupe alors la ligne de commande sur chaque espace, et un dossier
    #     nomme "...Poseidons Brutes - Swords..." se retrouve fragmente en plusieurs
    #     arguments separes - dont un "-" isole (venant de " - ") que Bambu Studio prend pour
    #     une option invalide ("Invalid option --"). Confirme en tapant la commande a la main
    #     AVEC guillemets (qui marchait) vs en la laissant passer par Start-Process SANS
    #     guillemets explicites (qui echouait a l'identique du bug reel). Fix : on entoure
    #     chaque chemin de guillemets nous-memes avant de les donner a -ArgumentList.
    # 26/10/2026 : --orient "1" force la recherche de la meilleure orientation par Bambu
    # Studio, qui peut prendre 10-25+ minutes sur un gros maillage (il teste plusieurs
    # orientations candidates et recalcule les surplombs/supports a chaque fois). Case
    # "Orientation auto" decochee par defaut dans l'UI -> on passe "0" (desactive) : la
    # fusion/export restent automatiques et rapides, l'orientation se fait a la main dans
    # Bambu Studio si besoin, comme avant ce correctif.
    # 27/10/2026 : CORRECTIF DE FOND - "--arrange" s'est revele INCOHERENT d'un fichier a
    # l'autre : "--arrange 0" a donne des pieces eclatees sur "Xpose Deadpool Statue", corrige
    # en passant a "--arrange 1"... qui a ensuite donne des pieces NON assemblees sur
    # "pawn-spiral-passthrough-fidget-toy" (capture d'ecran Tomyn : les 2 pieces cote a cote
    # sur le plateau, pas imbriquees), alors que l'apercu de sortie (base sur les positions
    # BRUTES des pieces, sans aucune repositionnement) les montrait bien assemblees. Conclusion :
    # aucune valeur fixe de --arrange n'est fiable pour TOUS les fichiers - Bambu Studio decide
    # de reorganiser (ou pas) les pieces selon une logique interne qu'on ne controle pas de
    # maniere previsible depuis la ligne de commande, des qu'on lui donne PLUSIEURS fichiers en
    # entree.
    # SOLUTION ROBUSTE : on ne laisse plus JAMAIS Bambu Studio decider du positionnement relatif
    # des pieces. On fusionne nous-memes les pieces en un SEUL maillage (on lit chaque .stl tel
    # quel, EXACTEMENT comme pour l'apercu "Aperçu entree/sortie" deja valide par Tomyn comme
    # fidele a l'assemblage reel - on ne fait QUE les recoller ensemble, aucune translation/
    # rotation appliquee) et on n'envoie a Bambu Studio qu'un SEUL fichier .stl deja assemble.
    # Avec un seul objet en entree, il n'y a plus de "plusieurs pieces a arranger" du tout :
    # le probleme est elimine a la racine plutot que rafistole flag par flag. Bambu ne sert plus
    # qu'a (a) chercher l'orientation si la case est cochee, et (b) exporter le .3mf.
    $journalLignes.Add("[$nomFigurine] Fusion de $($cheminsPieces.Count) piece(s) (positions d'origine conservees)...")
    # 27/10/2026 : demande de Tomyn - un fidget toy de ~200000 triangles (petit) a quand meme
    # pris plus d'une minute, ce qui ecarte l'explication "juste un gros maillage" et suggere un
    # cout FIXE (independant de la taille) quelque part - candidat le plus probable : le
    # demarrage de Bambu Studio CLI lui-meme (chargement du moteur de slicing/prereglages), pas
    # notre fusion locale. Plutot que deviner, on chronometre separement les deux etapes et on
    # journalise les deux durees - ca tranchera au prochain test, meme sur un petit fichier.
    $chronoFusion = [System.Diagnostics.Stopwatch]::StartNew()
    $triangles = [System.Collections.Generic.List[object]]::new()
    # 27/10/2026 : au passage, on releve la boite englobante de CHAQUE piece (quasi gratuit,
    # on parcourt deja tous les sommets pour la fusion) pour pouvoir detecter, juste apres,
    # une piece anormalement ecartee des autres - plutot que de ne le decouvrir qu'en ouvrant
    # le resultat dans Bambu Studio plusieurs minutes plus tard.
    $boitesParPiece = [System.Collections.Generic.List[object]]::new()
    foreach ($cheminPiece in $cheminsPieces) {
        $nbTri = Compter-TrianglesPiece $cheminPiece
        if ($nbTri -le 0) { $nbTri = 1 }
        $lu = Lire-TrianglesPieceEchantillon $cheminPiece ([Math]::Max($nbTri, 1))
        $minX = [double]::MaxValue; $maxX = [double]::MinValue
        $minY = [double]::MaxValue; $maxY = [double]::MinValue
        $minZ = [double]::MaxValue; $maxZ = [double]::MinValue
        # 27/10/2026 : CORRECTIF PERF #3 - meme gaspillage que dans Ecrire-STLBinaire (tableau
        # temporaire @($t.A,$t.B,$t.C) PAR TRIANGLE juste pour boucler) sur cette boucle qui
        # tourne sur TOUS les triangles de la piece (pas un echantillon) - un simple "$a,$b,$c"
        # sans @() cree quand meme un tableau (meme mecanisme), donc deroule explicitement les
        # 3 points (A puis B puis C) sans aucune construction de tableau intermediaire.
        foreach ($t in $lu.Triangles) {
            $pt = $t.A
            if ($pt[0] -lt $minX) { $minX = $pt[0] }; if ($pt[0] -gt $maxX) { $maxX = $pt[0] }
            if ($pt[1] -lt $minY) { $minY = $pt[1] }; if ($pt[1] -gt $maxY) { $maxY = $pt[1] }
            if ($pt[2] -lt $minZ) { $minZ = $pt[2] }; if ($pt[2] -gt $maxZ) { $maxZ = $pt[2] }
            $pt = $t.B
            if ($pt[0] -lt $minX) { $minX = $pt[0] }; if ($pt[0] -gt $maxX) { $maxX = $pt[0] }
            if ($pt[1] -lt $minY) { $minY = $pt[1] }; if ($pt[1] -gt $maxY) { $maxY = $pt[1] }
            if ($pt[2] -lt $minZ) { $minZ = $pt[2] }; if ($pt[2] -gt $maxZ) { $maxZ = $pt[2] }
            $pt = $t.C
            if ($pt[0] -lt $minX) { $minX = $pt[0] }; if ($pt[0] -gt $maxX) { $maxX = $pt[0] }
            if ($pt[1] -lt $minY) { $minY = $pt[1] }; if ($pt[1] -gt $maxY) { $maxY = $pt[1] }
            if ($pt[2] -lt $minZ) { $minZ = $pt[2] }; if ($pt[2] -gt $maxZ) { $maxZ = $pt[2] }
        }
        if ($lu.Triangles.Count -gt 0) {
            # 27/10/2026 : echantillon reduit de sommets (pas toute la piece - inutile et
            # couteux pour une simple detection de contact) reutilise par Test-PiecesEnContact
            # plus bas. On tape dans $lu.Triangles deja charge en memoire, aucune relecture du
            # fichier - juste un sous-echantillonnage par pas.
            $pointsContact = [System.Collections.Generic.List[object]]::new()
            $budgetContact = 20000
            $pas = [Math]::Max(1, [Math]::Floor($lu.Triangles.Count / $budgetContact))
            for ($k = 0; $k -lt $lu.Triangles.Count; $k += $pas) { $pointsContact.Add($lu.Triangles[$k].A) }
            $boitesParPiece.Add([PSCustomObject]@{
                Nom = (Split-Path $cheminPiece -Leaf)
                MinX = $minX; MaxX = $maxX; MinY = $minY; MaxY = $maxY; MinZ = $minZ; MaxZ = $maxZ
                PointsContact = $pointsContact
            })
        }
        $triangles.AddRange($lu.Triangles)
    }
    if ($boitesParPiece.Count -ge 2) {
        # taille globale approximative (diagonale de la boite qui englobe toutes les pieces),
        # utilisee comme reference pour juger si un ecart entre deux pieces est "anormal" -
        # un seuil absolu en mm n'aurait pas de sens entre une petite et une grosse figurine.
        $diagX = ($boitesParPiece | Measure-Object MaxX -Maximum).Maximum - ($boitesParPiece | Measure-Object MinX -Minimum).Minimum
        $diagY = ($boitesParPiece | Measure-Object MaxY -Maximum).Maximum - ($boitesParPiece | Measure-Object MinY -Minimum).Minimum
        $diagZ = ($boitesParPiece | Measure-Object MaxZ -Maximum).Maximum - ($boitesParPiece | Measure-Object MinZ -Minimum).Minimum
        $diagGlobale = [Math]::Sqrt($diagX * $diagX + $diagY * $diagY + $diagZ * $diagZ)
        $seuilEcart = [Math]::Max(1.0, $diagGlobale * 0.15)
        for ($i = 0; $i -lt $boitesParPiece.Count; $i++) {
            $ecartMin = [double]::MaxValue
            for ($j = 0; $j -lt $boitesParPiece.Count; $j++) {
                if ($i -eq $j) { continue }
                $a = $boitesParPiece[$i]; $b = $boitesParPiece[$j]
                $dx = [Math]::Max(0.0, [Math]::Max($a.MinX - $b.MaxX, $b.MinX - $a.MaxX))
                $dy = [Math]::Max(0.0, [Math]::Max($a.MinY - $b.MaxY, $b.MinY - $a.MaxY))
                $dz = [Math]::Max(0.0, [Math]::Max($a.MinZ - $b.MaxZ, $b.MinZ - $a.MaxZ))
                $dist = [Math]::Sqrt($dx * $dx + $dy * $dy + $dz * $dz)
                if ($dist -lt $ecartMin) { $ecartMin = $dist }
            }
            if ($ecartMin -gt $seuilEcart) {
                $journalLignes.Add("[$nomFigurine] ATTENTION : la piece '$($boitesParPiece[$i].Nom)' semble ecartee des autres (~$([Math]::Round($ecartMin, 1)) mm, seuil $([Math]::Round($seuilEcart, 1)) mm) - verifie l'assemblage dans le resultat, le fichier source est peut-etre incomplet/mal positionne.")
            }
        }
        # 27/10/2026 : demande de Tomyn - certaines pieces passent le controle ci-dessus
        # (pas "grossierement" ecartees) mais ne se TOUCHENT pas vraiment une fois fusionnees
        # (le concepteur n'avait pas pense l'assemblage comme un seul bloc imprimable). Seuil
        # de contact fixe a 0.3mm (en dessous de la largeur de buse habituelle ~0.4mm - un
        # ecart plus petit que ca a de bonnes chances d'etre comble/fusionne a l'impression,
        # un ecart plus grand ne le sera pas). AVERTISSEMENT seulement, pas un blocage : c'est
        # une heuristique basee sur un echantillon de points, pas une garantie absolue - une
        # piece designee expres pour etre separee (encoche d'aimant, etc.) declenchera aussi
        # ce message a tort.
        # 27/10/2026 : desormais reglable depuis l'interface (voir $champSeuilContact) au lieu
        # d'etre fige ici - recupere via le parametre $seuilContactMm transmis au runspace
        # (cette fonction tourne dans une session isolee, sans acces a $script:Config).
        for ($i = 0; $i -lt $boitesParPiece.Count; $i++) {
            $enContactAvecUneAutre = $false
            for ($j = 0; $j -lt $boitesParPiece.Count; $j++) {
                if ($i -eq $j) { continue }
                if (Test-PiecesEnContact $boitesParPiece[$i].PointsContact $boitesParPiece[$j].PointsContact $seuilContactMm) {
                    $enContactAvecUneAutre = $true
                    break
                }
            }
            if (-not $enContactAvecUneAutre) {
                $journalLignes.Add("[$nomFigurine] ATTENTION : la piece '$($boitesParPiece[$i].Nom)' ne semble REELLEMENT en contact avec aucune autre piece (aucun point a moins de $seuilContactMm mm) - meme si elle parait assemblee dans l'apercu, le resultat fusionne risque d'etre imprime comme des objets physiquement separes/detaches plutot qu'un seul bloc.")
            }
        }
    }
    if ($triangles.Count -eq 0) {
        $journalLignes.Add("[$nomFigurine] ECHEC : aucun triangle lu dans les pieces d'origine.")
        return [PSCustomObject]@{ Statut = "Echec - pieces vides"; CheminSortieStl = $null; Journal = $journalLignes }
    }
    $cheminStlFusionne = Join-Path $dossierTemp "$nomFigurine-fusion.stl"
    Ecrire-STLBinaire $triangles $cheminStlFusionne
    $chronoFusion.Stop()
    # 27/10/2026 : garde en memoire le nombre de triangles AVANT l'appel a Bambu Studio -
    # $triangles sera ecrase plus bas par la relecture du .3mf produit, et c'est la
    # comparaison entre les deux qui permet de reperer une fusion qui aurait silencieusement
    # perdu ou duplique de la matiere (triangles degeneres supprimes par Bambu, pièce ignoree,
    # etc.) meme quand Bambu Studio se termine sans erreur.
    $nbTrianglesOrigine = $triangles.Count
    $journalLignes.Add("[$nomFigurine] Fusion locale OK : $nbTrianglesOrigine triangles -> un seul fichier envoye a Bambu Studio. (duree fusion locale : $([Math]::Round($chronoFusion.Elapsed.TotalSeconds, 1)) s)")

    $valeurOrient = if ($orientationAuto) { "1" } else { "0" }
    $arguments = @("--orient", $valeurOrient, "--arrange", "0", "--export-3mf", "`"$chemin3mf`"", "`"$cheminStlFusionne`"")

    $journalLignes.Add("[$nomFigurine] Lancement Bambu Studio (--orient, 1 fichier deja fusionne)...")
    $journalLignes.Add("[$nomFigurine] Commande : `"$cheminBambuExe`" $($arguments -join ' ')")
    $stdoutLog = Join-Path $dossierTemp "stdout.log"
    $stderrLog = Join-Path $dossierTemp "stderr.log"
    $chronoBambu = [System.Diagnostics.Stopwatch]::StartNew()
    $futArrete = $false
    try {
        # 27/10/2026 : "-Wait" bloquait TOUT le code de cette fonction jusqu'a la fin du
        # processus, rendant impossible de l'interrompre depuis l'exterieur (voir correctif
        # "Arreter" plus haut) - on enleve "-Wait", on recupere le PID immediatement via
        # -PassThru, on le publie dans $etatAnnulation pour que le thread UI puisse aussi le
        # tuer directement si besoin, puis on attend nous-memes en sondant periodiquement si
        # l'utilisateur a demande l'arret.
        $proc = Start-Process -FilePath $cheminBambuExe -ArgumentList $arguments -PassThru -NoNewWindow `
            -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog
        if ($etatAnnulation) { $etatAnnulation.BambuProcessId = $proc.Id }
        while (-not $proc.HasExited) {
            if ($etatAnnulation -and $etatAnnulation.DoitArreter) {
                $futArrete = $true
                try { $proc.Kill($true) } catch {}
                break
            }
            Start-Sleep -Milliseconds 300
        }
        $proc.WaitForExit()
        if ($etatAnnulation) { $etatAnnulation.BambuProcessId = $null }
        $chronoBambu.Stop()
        if ($futArrete) {
            $journalLignes.Add("[$nomFigurine] Traitement arrete par l'utilisateur (Bambu Studio interrompu apres $([Math]::Round($chronoBambu.Elapsed.TotalSeconds, 1)) s).")
            return [PSCustomObject]@{ Statut = "Arrete par l'utilisateur"; CheminSortieStl = $null; Journal = $journalLignes }
        }
        $journalLignes.Add("[$nomFigurine] Bambu Studio termine (duree : $([Math]::Round($chronoBambu.Elapsed.TotalSeconds, 1)) s)")
    } catch {
        $chronoBambu.Stop()
        if ($etatAnnulation) { $etatAnnulation.BambuProcessId = $null }
        $journalLignes.Add("[$nomFigurine] ERREUR au lancement de Bambu Studio : $($_.Exception.Message)")
        return [PSCustomObject]@{ Statut = "Erreur - lancement Bambu Studio"; CheminSortieStl = $null; Journal = $journalLignes }
    }

    if (-not (Test-Path $chemin3mf)) {
        $journalLignes.Add("[$nomFigurine] ECHEC : aucun .3mf produit (code sortie $($proc.ExitCode)).")
        $erreurs = Get-Content $stderrLog -ErrorAction SilentlyContinue
        $sorties = Get-Content $stdoutLog -ErrorAction SilentlyContinue
        if ($erreurs) { $journalLignes.Add("[$nomFigurine] stderr : $($erreurs -join ' | ')") }
        if ($sorties) { $journalLignes.Add("[$nomFigurine] stdout : $($sorties -join ' | ')") }
        if (-not $erreurs -and -not $sorties) {
            $journalLignes.Add("[$nomFigurine] (stdout et stderr vides -> verifie que Bambu Studio n'etait pas DEJA OUVERT; si oui, ferme-le completement - y compris l'icone dans la barre des taches - puis reessaie.)")
        }
        $journalLignes.Add("[$nomFigurine] Dossier temporaire conserve pour diagnostic : $dossierTemp")
        return [PSCustomObject]@{ Statut = "Echec - pas de .3mf"; CheminSortieStl = $null; Journal = $journalLignes }
    }

    try {
        # 27/10/2026 : chrono ajoute ici - c'etait la DERNIERE grosse portion du traitement
        # qui n'etait pas encore mesuree (delai avant demarrage, dossier temporaire, fusion
        # locale et Bambu Studio lui-meme etaient deja chronometres et ne representaient a eux
        # quatre qu'une petite fraction du temps reel total observe sur les gros fichiers -
        # l'essentiel du temps manquant doit forcement etre ici : relire le .3mf produit par
        # Bambu Studio (XmlReader en flux + reconstruction d'un objet PowerShell par triangle)
        # et reecrire le .stl final.
        $chronoLecture3mf = [System.Diagnostics.Stopwatch]::StartNew()
        $resultatLecture3mf = Lire-Maillage3MF $chemin3mf
        $chronoLecture3mf.Stop()
        $triangles = $resultatLecture3mf.Triangles
        foreach ($ligneDiag in $resultatLecture3mf.Diagnostics) { $journalLignes.Add("[$nomFigurine] (3mf) $ligneDiag") }
        if ($triangles.Count -eq 0) {
            $journalLignes.Add("[$nomFigurine] ECHEC : le .3mf produit ne contient aucun triangle.")
            $journalLignes.Add("[$nomFigurine] Dossier temporaire conserve pour diagnostic : $dossierTemp")
            return [PSCustomObject]@{ Statut = "Echec - .3mf vide"; CheminSortieStl = $null; Journal = $journalLignes }
        }
        if (-not (Test-Path $dossierSortie)) { New-Item -ItemType Directory -Path $dossierSortie -Force | Out-Null }

        # 26/10/2026 : Tomyn veut garder le .3mf ET le .stl cote a cote en sortie (pas l'un
        # OU l'autre) - le .3mf exporte par Bambu Studio (qui contient deja la fusion +
        # l'arrangement, et l'orientation si la case correspondante est cochee) est donc
        # copie tel quel dans le dossier de sortie, en plus de la conversion en .stl.
        $chronoCopie3mf = [System.Diagnostics.Stopwatch]::StartNew()
        $chemin3mfSortie = Join-Path $dossierSortie "$nomFigurine.3mf"
        Copy-Item -Path $chemin3mf -Destination $chemin3mfSortie -Force
        $chronoCopie3mf.Stop()

        $chronoEcritureFinale = [System.Diagnostics.Stopwatch]::StartNew()
        $cheminStlSortie = Join-Path $dossierSortie "$nomFigurine.stl"
        Ecrire-STLBinaire $triangles $cheminStlSortie
        $chronoEcritureFinale.Stop()
        $journalLignes.Add("[$nomFigurine] (duree lecture .3mf resultat : $([Math]::Round($chronoLecture3mf.Elapsed.TotalSeconds, 1)) s / copie .3mf : $([Math]::Round($chronoCopie3mf.Elapsed.TotalSeconds, 1)) s / ecriture .stl final : $([Math]::Round($chronoEcritureFinale.Elapsed.TotalSeconds, 1)) s)")
        $journalLignes.Add("[$nomFigurine] OK : $($triangles.Count) triangles -> $cheminStlSortie (+ $chemin3mfSortie)")

        # 27/10/2026 : comparaison entree/sortie - un ecart important est suspect (Bambu a
        # peut-etre perdu ou ignore de la matiere) meme si le traitement se termine "OK".
        # Une petite difference est normale (triangles degeneres nettoyes, re-triangulation).
        if ($nbTrianglesOrigine -gt 0) {
            $ecartPourcent = [Math]::Abs($triangles.Count - $nbTrianglesOrigine) / $nbTrianglesOrigine * 100.0
            if ($ecartPourcent -gt 2.0) {
                $journalLignes.Add("[$nomFigurine] ATTENTION : $nbTrianglesOrigine triangles envoyes a Bambu Studio mais $($triangles.Count) dans le resultat (ecart $([Math]::Round($ecartPourcent, 1))%) - verifie que rien n'a ete perdu dans le resultat final.")
            }
        }

        Remove-Item -Path $dossierTemp -Recurse -Force -ErrorAction SilentlyContinue
        return [PSCustomObject]@{ Statut = "OK"; CheminSortieStl = $cheminStlSortie; Journal = $journalLignes }
    } catch {
        $journalLignes.Add("[$nomFigurine] ERREUR lecture/ecriture : $($_.Exception.Message)")
        $journalLignes.Add("[$nomFigurine] Dossier temporaire conserve pour diagnostic : $dossierTemp")
        return [PSCustomObject]@{ Statut = "Erreur - conversion 3mf->stl"; CheminSortieStl = $null; Journal = $journalLignes }
    }
}

$script:DoitArreter = $false
# 27/10/2026 : CORRECTIF - $script:DoitArreter (ci-dessus) n'etait verifie QU'AVANT de lancer
# la figurine SUIVANTE de la file (voir Demarrer-ProchainTraitement) : il ne faisait RIEN tant
# qu'une figurine etait DEJA en cours de traitement - exactement le cas rapporte par Tomyn
# ("meme le bouton arreter... il va meme pas s'arreter") avec une seule grosse figurine en
# cours. $script:EtatAnnulationPartage est un objet DIFFERENT, transmis AU runspace d'arriere-
# plan (voir AddArgument plus bas) et relu PENDANT l'attente de Bambu Studio - Traiter-
# UneFigurineCore peut ainsi reellement tuer le processus en cours. Reinitialise (mute, jamais
# reassigne pendant qu'un traitement tourne) a chaque nouvelle figurine.
$script:EtatAnnulationPartage = @{ DoitArreter = $false; BambuProcessId = $null }
$script:FileTraitement = [System.Collections.Generic.Queue[object]]::new()
# 27/10/2026 : CAUSE TROUVEE (et reproduite isolement hors de l'appli) du blocage "un
# traitement est deja en cours" qui persistait meme apres un traitement bien termine : une
# simple REASSIGNATION "$script:TraitementEnCours = ..." faite DEPUIS L'INTERIEUR d'un
# scriptblock construit avec .GetNewClosure() (c'est le cas du minuteur de traitement, voir
# Demarrer-ProchainTraitement) NE SE PROPAGE PAS vers la vraie variable de script - la
# closure obtient sa PROPRE copie isolee de la portee "script:", deconnectee de l'originale
# des sa creation. Donc le "$script:TraitementEnCours = $null" fait a la fin de chaque
# traitement (dans le minuteur) semblait fonctionner cote journal, mais ne remettait JAMAIS
# a jour la vraie variable lue par le garde-fou de Lancer-Traitement - d'ou le blocage
# systematique au traitement suivant. Solution : ne plus REASSIGNER une variable $script:
# depuis l'interieur d'une closure, mais MUTER un objet partage (ici une Hashtable, qui est
# un type reference) capture dans une variable LOCALE juste avant la creation de la closure -
# la closure recoit alors sa propre COPIE DE LA REFERENCE, qui pointe vers le MEME objet en
# memoire, donc modifier son contenu (pas la variable elle-meme) est bien visible partout.
$script:EtatTraitement = @{ EnCours = $null }

# 29/10/2026 : demande de Tomyn - le renommage avant fusion (voir plus bas dans
# Lancer-Traitement) proposait au depart le MEME nom deja pris comme valeur par defaut dans la
# boite de dialogue, ce qui l'obligeait a taper un nom lui-meme a chaque fois. Cette fonction
# propose maintenant automatiquement le prochain nom libre ("Nom (2)", "Nom (3)", etc., en
# testant .stl ET .3mf a chaque fois) - Tomyn n'a plus qu'a valider, tout en gardant la
# possibilite de taper autre chose dans la boite si ce nom automatique ne lui convient pas.
function Obtenir-ProchainNomLibre([string]$nomBase, [string]$dossierSortie) {
    $compteur = 2
    $nomCandidat = $nomBase
    while ((Test-Path (Join-Path $dossierSortie "$nomCandidat.stl")) -or (Test-Path (Join-Path $dossierSortie "$nomCandidat.3mf"))) {
        $nomCandidat = "$nomBase ($compteur)"
        $compteur++
    }
    return $nomCandidat
}

function Lancer-Traitement([array]$cibles) {
    if ($cibles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Rien a traiter - scanne d'abord un dossier.", "Fusion figurines")
        return
    }
    # 28/10/2026 : figurines sautees par la "reprise intelligente" (voir plus bas) - gardees a
    # part puis fusionnees dans $script:FigurinesCeLot juste apres sa (re)initialisation, pour
    # qu'elles apparaissent elles aussi dans le rapport CSV de fin de lot.
    $script:FigurinesSauteesCeLot = [System.Collections.Generic.List[object]]::new()
    # 26/10/2026 : garde-fou - empeche de lancer un DEUXIEME traitement par-dessus un premier
    # deja en cours dans CETTE MEME fenetre (ce qui ferait tourner 2 processus Bambu Studio +
    # 2 minuteurs en parallele, et donnerait un journal illisible avec des temps qui semblent
    # "reculer" puisque ce sont en fait deux chronometres differents entrelaces). Les boutons
    # sont normalement deja desactives pendant un traitement, donc ceci est une securite
    # supplementaire plutot que le cas attendu.
    if ($script:EtatTraitement.EnCours -or ($script:FileTraitement -and $script:FileTraitement.Count -gt 0)) {
        # 27/10/2026 : vrai bug identifie et corrige (voir le commentaire a la declaration de
        # $script:EtatTraitement plus haut) - ce garde-fou ne devrait donc plus jamais se
        # declencher apres un traitement reellement termine. On garde quand meme un "forcer
        # quand meme" en secours (Oui/Non) au lieu d'un simple blocage, par prudence : si ca
        # se reproduit malgre tout, Tomyn n'a plus besoin de fermer/rouvrir l'appli pour
        # continuer a travailler, et le diagnostic affiche permet de me le signaler avec les
        # valeurs exactes plutot qu'une simple description.
        $nomBloquant = if ($script:EtatTraitement.EnCours) { $script:EtatTraitement.EnCours.Nom } else { "(aucun)" }
        $tailleFile = if ($script:FileTraitement) { $script:FileTraitement.Count } else { 0 }
        $reponse = [System.Windows.Forms.MessageBox]::Show(
            "Un traitement est deja en cours - attends qu'il se termine (ou clique Arreter) avant d'en lancer un autre.`n`n[Diagnostic] EnCours = $nomBloquant / FileTraitement.Count = $tailleFile`n`nSi tu es sur qu'aucun traitement n'est reellement en cours (Bambu Studio ferme, barre de progression disparue), clique Oui pour forcer quand meme.",
            "Fusion figurines", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($reponse -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Ajouter-Journal "Etat de traitement force a zero manuellement (etait : EnCours=$nomBloquant, FileTraitement.Count=$tailleFile)."
        $script:EtatTraitement.EnCours = $null
        $script:FileTraitement = [System.Collections.Generic.Queue[object]]::new()
    }
    $cheminBambuExe = $champBambuExe.Text
    if (-not (Test-Path $cheminBambuExe)) {
        [System.Windows.Forms.MessageBox]::Show("Chemin de bambu-studio.exe introuvable : $cheminBambuExe", "Fusion figurines")
        return
    }

    # 27/10/2026 : tentative de verif "--help" abandonnee - teste en reel, --help de Bambu
    # Studio ne liste PAS --orient/--arrange/--export-3mf alors que ces options fonctionnent
    # tres bien (ce sont des options internes a la CLI, pas documentees dans l'aide generale).
    # Le controle etait donc un faux positif systematique - bloquant a chaque lancement pour
    # rien. Supprime plutot que laisse en place a moitie casse.
    $dossierSortie = $champSortie.Text
    $orientationAuto = $caseOrientation.Checked

    # 27/10/2026 : demande de Tomyn - "reprise intelligente" : si on relance un lot qui a deja
    # ete traite en partie (ex: figurines ajoutees au dossier source apres un premier passage),
    # on sautait jusqu'ici TOUT refaire depuis zero, pieces deja fusionnees comprises. On sautera
    # desormais une figurine si sa sortie (.stl ET .3mf) existe deja dans $dossierSortie et est
    # PLUS RECENTE que toutes ses pieces source incluses - signe qu'elle n'a pas change depuis
    # son dernier traitement reussi. Volontairement CONSERVATEUR : si des pieces ont ete exclues/
    # reincluses (PiecesExclues non vide), on ne se fie pas aux dates (exclure une piece ne change
    # pas sa date de modification) et on retraite systematiquement. Case decochable si Tomyn veut
    # forcer un retraitement complet malgre tout.
    if ($caseSauterDejaTraitees.Checked) {
        $ciblesFiltrees = [System.Collections.Generic.List[object]]::new()
        $nbSautees = 0
        foreach ($grp in $cibles) {
            $dejaTraitee = $false
            $aDesExclusions = $grp.PSObject.Properties['PiecesExclues'] -and $grp.PiecesExclues.Count -gt 0
            if (-not $aDesExclusions) {
                $cheminStlAttendu = Join-Path $dossierSortie "$($grp.Nom).stl"
                $chemin3mfAttendu = Join-Path $dossierSortie "$($grp.Nom).3mf"
                if ((Test-Path $cheminStlAttendu) -and (Test-Path $chemin3mfAttendu)) {
                    $dateSortie = [System.IO.File]::GetLastWriteTimeUtc($cheminStlAttendu)
                    # 27/10/2026 : CORRECTIF - $grp.Pieces vient du scan initial (boutonScanner),
                    # ses objets FileInfo gardent la date de modification figee a CE moment-la.
                    # Si Tomyn modifie une piece APRES le scan mais AVANT de lancer le traitement,
                    # .LastWriteTimeUtc en cache reste l'ANCIENNE date et la figurine serait sautee
                    # a tort (confirme par test_reprise_intelligente.ps1 en bac a sable). On relit
                    # donc la date directement sur le disque ici plutot que de faire confiance au
                    # cache de l'objet FileInfo.
                    $dateSourcePlusRecente = ($grp.Pieces | ForEach-Object { [System.IO.File]::GetLastWriteTimeUtc($_.FullName) } | Measure-Object -Maximum).Maximum
                    if ($dateSortie -ge $dateSourcePlusRecente) { $dejaTraitee = $true }
                }
            }
            if ($dejaTraitee) {
                $nbSautees++
                # 28/10/2026 : CORRECTIF - une figurine sautee ici disparaissait purement et
                # simplement : pas de statut mis a jour dans le tableau (la ligne gardait son
                # ancien statut, potentiellement perime ou juste "En attente" pour toujours),
                # et aucune trace dans le rapport CSV de fin de lot (seul un compte TOTAL
                # apparaissait dans le journal, sans dire LESQUELLES). On met maintenant a jour
                # son statut tout de suite (visible immediatement, pas besoin d'attendre la fin
                # du lot) et on la garde a part pour l'ajouter au rapport CSV plus bas.
                $grp.Statut = "Deja a jour (sautee)"
                $grp.CheminSortieStl = $cheminStlAttendu
                $grp.DureeSecondes = $null
                Mettre-A-Jour-LigneGroupe $grp
                $script:FigurinesSauteesCeLot.Add($grp)
            } else {
                $ciblesFiltrees.Add($grp)
            }
        }
        if ($nbSautees -gt 0) {
            Ajouter-Journal "Reprise intelligente : $nbSautees figurine(s) deja traitee(s) (sortie plus recente que les pieces source) sautee(s) - decoche 'Sauter les figurines deja traitees' pour forcer un retraitement complet."
        }
        $cibles = @($ciblesFiltrees)
        if ($cibles.Count -eq 0) {
            Ajouter-Journal "Rien a traiter - toutes les figurines selectionnees sont deja a jour."
            [System.Windows.Forms.MessageBox]::Show("Toutes les figurines selectionnees sont deja a jour (sortie plus recente que les pieces source).`n`nDecoche 'Sauter les figurines deja traitees' pour forcer un retraitement complet.", "Fusion figurines")
            return
        }
    }

    # 29/10/2026 : demande de Tomyn - cas "plusieurs figurines (ex: une figurine + sa base,
    # une autre figurine + une autre base, etc.) melangees dans le MEME dossier source". Le
    # scan ne cree qu'UN groupe par dossier (voir boutonScanner.Add_Click) : pour fusionner la
    # 2e figurine du dossier, Tomyn reutilise ce MEME groupe en changeant juste les pieces
    # exclues (celles de la 1ere figurine passent exclues, celles de la 2e sont reincluses).
    # Probleme : la sortie s'appelle toujours "<NomDossier>.stl" des deux cotes - la 2e fusion
    # ECRASERAIT silencieusement le resultat de la 1ere. On demande donc, quand une fusion AVEC
    # exclusions va produire un nom de sortie qui existe deja sur le disque, si Tomyn veut
    # renommer cette fusion-ci avant de continuer. Oui = demande un nom et continue sous ce nom.
    # Non = cette figurine est ignoree pour ce lot (le temps qu'il decide d'un nom lui-meme).
    # Limite a "a des exclusions" pour ne jamais interrompre le cas normal (une figurine = un
    # dossier, reprocessee/ecrasee volontairement sans qu'on lui demande rien).
    $ciblesApresRenommage = [System.Collections.Generic.List[object]]::new()
    foreach ($grp in $cibles) {
        $aDesExclusions = $grp.PSObject.Properties['PiecesExclues'] -and $grp.PiecesExclues.Count -gt 0
        $cheminStlExistant = Join-Path $dossierSortie "$($grp.Nom).stl"
        $chemin3mfExistant = Join-Path $dossierSortie "$($grp.Nom).3mf"
        if ($aDesExclusions -and ((Test-Path $cheminStlExistant) -or (Test-Path $chemin3mfExistant))) {
            $nomSuggere = Obtenir-ProchainNomLibre $grp.Nom $dossierSortie
            $reponse = [System.Windows.Forms.MessageBox]::Show(
                "Un resultat '$($grp.Nom).stl' existe deja dans le dossier de sortie.`n`nSi c'est une AUTRE figurine du meme dossier source (plusieurs figurines/bases melangees dans un seul dossier, avec des pieces exclues differentes a chaque fois), cette fusion-ci peut etre renommee automatiquement en '$nomSuggere' pour ne pas ecraser le resultat precedent (tu pourras changer ce nom propose si tu preferes autre chose).`n`nRenommer et continuer ?`n(Non = cette figurine est ignoree pour ce lot)",
                "Fusion figurines - resultat deja existant", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($reponse -eq [System.Windows.Forms.DialogResult]::Yes) {
                $nouveauNom = [Microsoft.VisualBasic.Interaction]::InputBox("Nom propose automatiquement (modifie-le si tu veux autre chose) :", "Renommer avant fusion", $nomSuggere)
                $nouveauNom = $nouveauNom.Trim()
                if ($nouveauNom) {
                    $ancienNom = $grp.Nom
                    $grp.Nom = $nouveauNom
                    Mettre-A-Jour-LigneGroupe $grp
                    Ajouter-Journal "[$ancienNom] Renommee en '$nouveauNom' avant traitement (evite d'ecraser le resultat existant)."
                    $ciblesApresRenommage.Add($grp)
                } else {
                    Ajouter-Journal "[$($grp.Nom)] Ignoree pour ce lot (renommage annule ou nom vide)."
                    $grp.Statut = "Ignoree (renommage annule)"
                    Mettre-A-Jour-LigneGroupe $grp
                }
            } else {
                Ajouter-Journal "[$($grp.Nom)] Ignoree pour ce lot (resultat existant, pas de renommage demande)."
                $grp.Statut = "Ignoree (resultat existant)"
                Mettre-A-Jour-LigneGroupe $grp
            }
        } else {
            $ciblesApresRenommage.Add($grp)
        }
    }
    $cibles = @($ciblesApresRenommage)
    if ($cibles.Count -eq 0) {
        Ajouter-Journal "Rien a traiter apres renommage/ignore."
        return
    }

    # 27/10/2026 : demande de Tomyn - jusqu'ici, Bambu Studio deja ouvert en arriere-plan ne
    # se decouvrait qu'APRES l'echec (stdout/stderr vides, voir le message de diagnostic dans
    # Traiter-UneFigurineCore) - on le detecte maintenant AVANT de lancer quoi que ce soit.
    $nomProcessus = [System.IO.Path]::GetFileNameWithoutExtension($cheminBambuExe)
    $processusOuverts = Get-Process -Name $nomProcessus -ErrorAction SilentlyContinue
    if ($processusOuverts) {
        $reponse = [System.Windows.Forms.MessageBox]::Show(
            "Bambu Studio semble deja ouvert (processus '$nomProcessus' detecte, $($processusOuverts.Count) instance(s)).`n`nFerme-le COMPLETEMENT (y compris l'icone dans la barre des taches Windows) avant de lancer, sinon le traitement va tres probablement echouer silencieusement (stdout/stderr vides).`n`nContinuer quand meme ?",
            "Fusion figurines", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($reponse -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    # 27/10/2026 : verification grossiere de l'espace disque - une fusion ecrit temporairement
    # un .stl fusionne + un .3mf (dossier temp Windows) PUIS copie le .3mf et le .stl finaux
    # dans le dossier de sortie : jusqu'a ~2.5x la taille des pieces d'origine en pointe, sur
    # potentiellement DEUX lecteurs differents (celui de %TEMP% et celui du dossier de sortie).
    # Pas une garantie exacte (on ne connait pas la taille reelle du .3mf avant de l'avoir
    # produit), juste un garde-fou pour eviter de decouvrir "disque plein" apres 20 minutes
    # d'attente sur un gros lot.
    try {
        $tailleOrigine = ($cibles | ForEach-Object { ($_.Pieces | Measure-Object -Property Length -Sum).Sum }) | Measure-Object -Sum | Select-Object -ExpandProperty Sum
        if ($tailleOrigine -gt 0) {
            $tailleEstimee = [double]$tailleOrigine * 2.5
            $lecteursAVerifier = @()
            try { $lecteursAVerifier += [System.IO.Path]::GetPathRoot($dossierSortie) } catch {}
            try { $lecteursAVerifier += [System.IO.Path]::GetPathRoot([System.IO.Path]::GetTempPath()) } catch {}
            foreach ($racine in ($lecteursAVerifier | Select-Object -Unique)) {
                if (-not $racine) { continue }
                try {
                    $lecteur = New-Object System.IO.DriveInfo($racine)
                    if ($lecteur.IsReady -and $lecteur.AvailableFreeSpace -lt $tailleEstimee) {
                        $reponse = [System.Windows.Forms.MessageBox]::Show(
                            "Espace disque disponible sur $racine : $([Math]::Round($lecteur.AvailableFreeSpace / 1GB, 1)) Go, estimation necessaire pour ce lot : ~$([Math]::Round($tailleEstimee / 1GB, 1)) Go (approximatif).`n`nContinuer quand meme ?",
                            "Fusion figurines", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        if ($reponse -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                    }
                } catch {}
            }
        }
    } catch {}
    $script:DoitArreter = $false
    $boutonArreter.Enabled = $true
    $boutonTraiter.Enabled = $false
    $boutonTraiterTout.Enabled = $false
    $boutonScanner.Enabled = $false

    # 26/10/2026 : CORRECTIF GEL DE 10+ MINUTES - l'ancienne version appelait
    # Traiter-UneFigurine directement sur le thread de l'interface, avec -Wait sur
    # Start-Process : Bambu Studio (recherche d'orientation sur ~1M triangles) peut
    # legitimement prendre plusieurs minutes PAR figurine, et pendant tout ce temps la
    # fenetre WinForms ne recoit plus aucun message (pas de redessin, "ne repond pas" dans
    # le Gestionnaire des taches) - meme symptome et meme cause que le gel de l'apercu
    # corrige plus haut (voir Obtenir-RunspacePartagee / Demarrer-ApercuAsync). Fix
    # identique : chaque figurine est traitee dans le runspace partage "fusion", un
    # Timer WinForms attend la fin sans bloquer, et on enchaine sur la figurine suivante
    # de la file une fois le resultat applique - l'UI reste reactive (boutons, log,
    # Arreter) pendant tout le traitement, qu'une figurine prenne 10 secondes ou 20 minutes.
    $script:FileTraitement = [System.Collections.Generic.Queue[object]]::new()
    foreach ($grp in $cibles) { $script:FileTraitement.Enqueue($grp) }
    $script:TotalFigurinesTraitement = $cibles.Count
    $script:IndexFigurineTraitement = 0
    # 27/10/2026 : liste (dans l'ordre de traitement) utilisee pour ecrire le rapport CSV de
    # fin de lot - voir la fin de Demarrer-ProchainTraitement.
    $script:FigurinesCeLot = [System.Collections.Generic.List[object]]::new()
    foreach ($g in $script:FigurinesSauteesCeLot) { $script:FigurinesCeLot.Add($g) }
    Ajouter-Journal "Orientation automatique : $(if ($orientationAuto) { 'activee (peut etre lent sur gros maillages)' } else { 'desactivee - a faire a la main dans Bambu Studio si besoin' })"
    Afficher-OverlayProgression "Preparation..."
    Demarrer-ProchainTraitement $dossierSortie $cheminBambuExe $orientationAuto
}

function Demarrer-ProchainTraitement([string]$dossierSortie, [string]$cheminBambuExe, [bool]$orientationAuto) {
    if ($script:DoitArreter -or $script:FileTraitement.Count -eq 0) {
        $boutonArreter.Enabled = $false
        $boutonTraiter.Enabled = $true
        $boutonTraiterTout.Enabled = $true
        $boutonScanner.Enabled = $true
        if ($script:DoitArreter) {
            Ajouter-Journal "Arret demande."
            Mettre-A-Jour-OverlayProgression ([int](100 * $script:IndexFigurineTraitement / [Math]::Max(1, $script:TotalFigurinesTraitement))) "Arrete."
        } else {
            Mettre-A-Jour-OverlayProgression 100 "Traitement termine"
        }
        Ajouter-Journal "=== Traitement termine ==="
        # 27/10/2026 : rapport CSV de fin de lot (demande de Tomyn) - recapitule chaque
        # figurine traitee dans CE lot (nom, nb pieces, nb triangles, statut, duree, chemin de
        # sortie), pour voir d'un coup d'oeil ce qui a reussi/echoue sur un gros lot sans
        # devoir rescroller tout le journal. Ecrit uniquement si au moins une figurine a ete
        # traitee (pas de rapport vide si on a clique "Arreter" avant meme le debut).
        if ($script:FigurinesCeLot -and $script:FigurinesCeLot.Count -gt 0) {
            try {
                if (-not (Test-Path $dossierSortie)) { New-Item -ItemType Directory -Path $dossierSortie -Force | Out-Null }
                $cheminRapport = Join-Path $dossierSortie "rapport-fusion-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
                $script:FigurinesCeLot | Select-Object `
                    @{Name = "Figurine"; Expression = { $_.Nom } },
                    @{Name = "Pieces"; Expression = { $_.Pieces.Count } },
                    @{Name = "Triangles"; Expression = { $_.NombreTriangles } },
                    @{Name = "Statut"; Expression = { $_.Statut } },
                    @{Name = "DureeSecondes"; Expression = { if ($null -ne $_.DureeSecondes) { [Math]::Round($_.DureeSecondes, 1) } else { "" } } },
                    @{Name = "FichierSortie"; Expression = { $_.CheminSortieStl } } |
                    Export-Csv -Path $cheminRapport -Delimiter ";" -NoTypeInformation -Encoding UTF8
                Ajouter-Journal "Rapport du lot ecrit : $cheminRapport"
            } catch {
                Ajouter-Journal "ATTENTION : echec de l'ecriture du rapport CSV : $($_.Exception.Message)"
            }
        }
        # 27/10/2026 : notification Windows native - seulement si la fenetre n'a pas le focus
        # (minimisee ou autre fenetre active), pour ne pas doubler l'info quand Tomyn regarde
        # deja l'overlay/le journal en direct.
        if ($script:FigurinesCeLot -and $script:FigurinesCeLot.Count -gt 0 -and -not $form.ContainsFocus) {
            try {
                $nbOk = @($script:FigurinesCeLot | Where-Object { $_.Statut -eq "OK" }).Count
                $texteNotif = "$nbOk / $($script:FigurinesCeLot.Count) figurine(s) traitee(s) avec succes."
                $script:IconeNotification.BalloonTipTitle = "Fusion figurines - traitement termine"
                $script:IconeNotification.BalloonTipText = $texteNotif
                $script:IconeNotification.ShowBalloonTip(8000)
            } catch {}
        }
        Masquer-OverlayProgressionApresDelai 3
        return
    }

    $grp = $script:FileTraitement.Dequeue()
    $script:FigurinesCeLot.Add($grp)
    $script:EtatTraitement.EnCours = $grp
    # 27/10/2026 : capture LOCALE (pas $script:) de l'objet partage, prise juste avant la
    # creation de la closure du minuteur plus bas - voir le commentaire a la declaration de
    # $script:EtatTraitement : c'est cette variable locale, et non "$script:EtatTraitement"
    # directement, qu'il faut utiliser DANS la closure pour que la mutation soit bien visible
    # depuis l'exterieur (Lancer-Traitement).
    $etatPartage = $script:EtatTraitement
    # 27/10/2026 : reinitialise (MUTE, pas reassigne) pour cette nouvelle figurine - voir le
    # correctif "Arreter" plus haut (declaration de $script:EtatAnnulationPartage).
    $script:EtatAnnulationPartage.DoitArreter = $false
    $script:EtatAnnulationPartage.BambuProcessId = $null
    $etatAnnulationPartage = $script:EtatAnnulationPartage
    $script:IndexFigurineTraitement++
    $pctAvant = [int](100 * ($script:IndexFigurineTraitement - 1) / [Math]::Max(1, $script:TotalFigurinesTraitement))
    Mettre-A-Jour-OverlayProgression $pctAvant "Figurine $($script:IndexFigurineTraitement) / $($script:TotalFigurinesTraitement) : $($grp.Nom)"

    # Le runspace n'a pas acces aux fonctions du script principal (session independante) -
    # on lui transmet le CORPS de chaque fonction necessaire, exactement comme pour l'apercu.
    # 26/10/2026 : Lire-Maillage3MF appelle a son tour Resoudre-Objet3MF, qui appelle lui-meme
    # Appliquer-Transform3MF et Charger-Document3MFExterne (reference "component" vers un
    # fichier .model externe, extension Production de Bambu Studio) - TOUTE la chaine doit
    # etre transmise au runspace, sinon "terme non reconnu" a l'execution (vu en test : bug
    # corrige ici, seuls Lire-Maillage3MF/Ecrire-STLBinaire avaient ete transmis au premier essai).
    $defCore = (Get-Item function:Traiter-UneFigurineCore).Definition
    $defLireMaillage = (Get-Item function:Lire-Maillage3MF).Definition
    $defEcrireStl = (Get-Item function:Ecrire-STLBinaire).Definition
    $defAppliquerTransform = (Get-Item function:Appliquer-Transform3MF).Definition
    # 27/10/2026 : CORRECTIF - Appliquer-Transform3MF-Triangles (nouvelle fonction du correctif
    # perf #2) etait appelee par Resoudre-Objet3MF/Lire-Maillage3MF mais oubliee ici -> "terme
    # non reconnu" a l'execution reelle (vu par Tomyn), exactement le meme type de bug que
    # Test-PiecesEnContact plus haut. Meme chaine de transmission que les autres fonctions 3MF.
    $defAppliquerTransformTriangles = (Get-Item function:Appliquer-Transform3MF-Triangles).Definition
    $defLireDocument3MF = (Get-Item function:Lire-Document3MF).Definition
    $defChargerDocExterne = (Get-Item function:Charger-Document3MFExterne).Definition
    $defResoudreObjet = (Get-Item function:Resoudre-Objet3MF).Definition
    # 27/10/2026 : Traiter-UneFigurineCore lit maintenant lui-meme les .stl d'origine (fusion
    # locale avant d'appeler Bambu Studio - voir plus haut) : Compter-TrianglesSTL et
    # Lire-TrianglesSTLEchantillon doivent donc aussi etre transmis au runspace.
    $defCompterTriangles = (Get-Item function:Compter-TrianglesSTL).Definition
    $defLireTrianglesEchantillon = (Get-Item function:Lire-TrianglesSTLEchantillon).Definition
    # 29/10/2026 : support .obj (voir Compter-TrianglesPiece/Lire-TrianglesPieceEchantillon) -
    # Traiter-UneFigurineCore appelle maintenant ces deux-la au lieu des fonctions STL directes
    # (voir plus bas, boucle de lecture des pieces pour la fusion locale), donc elles doivent
    # etre transmises a CE runspace aussi, meme principe que les deux lignes au-dessus.
    $defCompterOBJ = (Get-Item function:Compter-TrianglesOBJ).Definition
    $defLireOBJ = (Get-Item function:Lire-TrianglesOBJEchantillon).Definition
    $defCompterPiece = (Get-Item function:Compter-TrianglesPiece).Definition
    $defLirePiece = (Get-Item function:Lire-TrianglesPieceEchantillon).Definition
    # 27/10/2026 : CORRECTIF CRASH - Test-PiecesEnContact (detection "pieces proches mais pas
    # vraiment en contact") etait appelee par Traiter-UneFigurineCore mais jamais transmise a
    # CE runspace (session isolee, voir le commentaire juste au-dessus) : "Test-PiecesEnContact
    # n'est pas reconnu..." a chaque traitement, decouvert par Tomyn sur le fidget toy.
    $defTestPiecesEnContact = (Get-Item function:Test-PiecesEnContact).Definition
    $scriptTexte = 'param($NomFigurine, $CheminsPieces, $DossierSortie, $CheminBambuExe, $OrientationAuto, $EtatAnnulation, $HeureEnvoi, $SeuilContactMm)' + "`r`n" +
        'function Appliquer-Transform3MF {' + "`r`n" + $defAppliquerTransform + "`r`n" + '}' + "`r`n" +
        'function Appliquer-Transform3MF-Triangles {' + "`r`n" + $defAppliquerTransformTriangles + "`r`n" + '}' + "`r`n" +
        'function Lire-Document3MF {' + "`r`n" + $defLireDocument3MF + "`r`n" + '}' + "`r`n" +
        'function Charger-Document3MFExterne {' + "`r`n" + $defChargerDocExterne + "`r`n" + '}' + "`r`n" +
        'function Resoudre-Objet3MF {' + "`r`n" + $defResoudreObjet + "`r`n" + '}' + "`r`n" +
        'function Lire-Maillage3MF {' + "`r`n" + $defLireMaillage + "`r`n" + '}' + "`r`n" +
        'function Ecrire-STLBinaire {' + "`r`n" + $defEcrireStl + "`r`n" + '}' + "`r`n" +
        'function Compter-TrianglesSTL {' + "`r`n" + $defCompterTriangles + "`r`n" + '}' + "`r`n" +
        'function Lire-TrianglesSTLEchantillon {' + "`r`n" + $defLireTrianglesEchantillon + "`r`n" + '}' + "`r`n" +
        'function Compter-TrianglesOBJ {' + "`r`n" + $defCompterOBJ + "`r`n" + '}' + "`r`n" +
        'function Lire-TrianglesOBJEchantillon {' + "`r`n" + $defLireOBJ + "`r`n" + '}' + "`r`n" +
        'function Compter-TrianglesPiece {' + "`r`n" + $defCompterPiece + "`r`n" + '}' + "`r`n" +
        'function Lire-TrianglesPieceEchantillon {' + "`r`n" + $defLirePiece + "`r`n" + '}' + "`r`n" +
        'function Test-PiecesEnContact {' + "`r`n" + $defTestPiecesEnContact + "`r`n" + '}' + "`r`n" +
        'function Traiter-UneFigurineCore {' + "`r`n" + $defCore + "`r`n" + '}' + "`r`n" +
        'Traiter-UneFigurineCore $NomFigurine $CheminsPieces $DossierSortie $CheminBambuExe $OrientationAuto $EtatAnnulation $HeureEnvoi $SeuilContactMm'

    $rs = Obtenir-RunspacePartagee "fusion"
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    # 27/10/2026 : ne transmet au traitement que les pieces NON exclues (voir "Choisir les
    # pieces a inclure..." dans le menu contextuel de la liste).
    $exclues = if ($grp.PSObject.Properties['PiecesExclues']) { @($grp.PiecesExclues) } else { @() }
    $cheminsPieces = @($grp.Pieces | Where-Object { $exclues -notcontains $_.Name } | ForEach-Object { $_.FullName })
    # 27/10/2026 : Tomyn a releve ~30s d'ecart entre "Demarrage" et la premiere ligne reellement
    # ecrite par Traiter-UneFigurineCore, alors que la fusion locale + Bambu Studio ne prennent
    # que 7s mesures. Pour savoir si ce delai vient du DISPATCH vers le runspace (parsing du
    # script texte, JIT, etc.) ou d'autre chose, on transmet l'heure d'envoi et la fonction
    # logue elle-meme, en tout premier, le temps ecoule avant qu'elle ne commence vraiment.
    $heureEnvoiRunspace = Get-Date
    # 27/10/2026 : seuil de contact lu depuis la config (regle en direct par $champSeuilContact,
    # voir sa creation plus haut) - transmis en argument puisque ce runspace n'a pas acces a
    # $script:Config (session isolee, voir le commentaire sur $defCore plus haut).
    $seuilContactMmActuel = [double]$script:Config.SeuilContactMm
    [void]$ps.AddScript($scriptTexte).AddArgument($grp.Nom).AddArgument($cheminsPieces).AddArgument($dossierSortie).AddArgument($cheminBambuExe).AddArgument($orientationAuto).AddArgument($etatAnnulationPartage).AddArgument($heureEnvoiRunspace).AddArgument($seuilContactMmActuel)
    $resultatAsync = $ps.BeginInvoke()
    $debutFigurine = Get-Date
    # 27/10/2026 : variables LOCALES (pas $script:) - elles ne servent qu'au throttling du
    # journal PENDANT cette figurine, jamais lues depuis l'exterieur de la closure, donc pas
    # besoin (et pas fiable, voir plus haut) de passer par $script:.
    $dernierPalierJournal = 0
    $maxSecondesAffichees = 0

    Ajouter-Journal "[$($grp.Nom)] Demarrage (en arriere-plan, l'interface reste utilisable)..."

    $minuteur = New-Object System.Windows.Forms.Timer
    $minuteur.Interval = 300
    $gestionnaireTick = {
        # 27/10/2026 : ce bloc tourne protege par un try/catch/finally global (voir
        # $doitNettoyer plus bas) pour qu'une erreur imprevue dans l'animation ne puisse
        # jamais empecher le nettoyage final de s'executer. ET surtout : toute variable
        # partagee avec l'EXTERIEUR de cette closure (ex: l'etat "traitement en cours" lu par
        # Lancer-Traitement) doit passer par $etatPartage (capture locale d'un objet
        # Hashtable MUTE, jamais reassigne) et NON par "$script:..." - voir le commentaire a
        # la declaration de $script:EtatTraitement plus haut dans le fichier pour le pourquoi
        # (bug reel trouve et corrige le 27/10/2026 : une reassignation $script: depuis une
        # closure .GetNewClosure() ne se propage pas vers la vraie variable de script).
        $doitNettoyer = $false
        try {
            if (-not $resultatAsync.IsCompleted) {
                # tant qu'on attend (lancement Bambu Studio + fusion/conversion, qui peut
                # prendre plusieurs minutes sur un gros maillage - pas de vraie mesure de
                # progression possible depuis ici), on anime la barre en va-et-vient et on
                # affiche le temps ecoule, plutot que de la laisser figee a 0% (Tomyn : "ca
                # donne l'impression que c'est bloque"). On ajoute aussi une ligne de journal
                # toutes les 30s pour que le journal lui-meme montre des signes de vie.
                $ecouleFigurine = (Get-Date) - $debutFigurine
                $largeurPiste = $pisteOverlay.Width
                $largeurBloc = 90
                $periodeMs = 1600.0
                $fracCycle = ($ecouleFigurine.TotalMilliseconds % $periodeMs) / $periodeMs
                $frac = if ($fracCycle -lt 0.5) { $fracCycle * 2 } else { (1 - $fracCycle) * 2 }
                $remplissageOverlay.Width = $largeurBloc
                $remplissageOverlay.Left = [int]($frac * [Math]::Max(0, $largeurPiste - $largeurBloc))
                # garde-fou "jamais en arriere" - Tomyn a vu le temps affiche redescendre (ex:
                # 2 min 45 s puis 2 min 16 s) alors que (Get-Date) - $debutFigurine ne peut
                # mathematiquement pas diminuer pour un seul minuteur/point de depart. Cause
                # exacte non reproduite ici, mais on se protege quand meme : on n'affiche/
                # journalise jamais un temps plus petit que le plus grand deja vu pour cette
                # figurine.
                $secondesEcoulees = [int]$ecouleFigurine.TotalSeconds
                if ($secondesEcoulees -lt $maxSecondesAffichees) { $secondesEcoulees = $maxSecondesAffichees }
                else { $maxSecondesAffichees = $secondesEcoulees }
                $labelOverlayPct.Text = "en cours depuis $(Formater-DureeMinSec $secondesEcoulees)"

                # 27/10/2026 : CORRECTIF - [int] en PowerShell ARRONDIT (round-to-even), il ne
                # TRONQUE PAS comme en C/C# : [int](16/30) = 1, pas 0 ! Ca declenchait ce palier
                # en realite a 16s/46s/76s... (toujours ~16s, quelle que soit la taille du
                # fichier - ce que Tomyn a repere), au lieu de 30s/60s/90s comme prevu a
                # l'origine. [Math]::Floor() tronque vraiment et corrige le declenchement.
                # 27/10/2026 (suite) : Tomyn trouve que 30s de silence dans le journal avant la
                # premiere ligne donne justement l'impression inverse - "rien ne se passe"
                # pendant une demi-minute (la barre anime mais ne rassure pas autant qu'une
                # ligne de texte). Palier resserre a 5s : signe de vie frequent sans spammer sur
                # les tres longs traitements (un job de 90s donne ~18 lignes, raisonnable).
                $intervalleJournalSecondes = 1
                $paliers30s = [Math]::Floor($secondesEcoulees / $intervalleJournalSecondes)
                if ($paliers30s -gt $dernierPalierJournal) {
                    $dernierPalierJournal = $paliers30s
                    Ajouter-Journal "[$($grp.Nom)] ... toujours en cours ($(Formater-DureeMinSec $secondesEcoulees) ecoulees)"
                }
                return
            }
            $doitNettoyer = $true
            $minuteur.Stop()
            $remplissageOverlay.Left = 0
            # 27/10/2026 : duree reelle de cette figurine, pour le rapport CSV de fin de lot.
            $grp.DureeSecondes = ((Get-Date) - $debutFigurine).TotalSeconds
            $sortie = $ps.EndInvoke($resultatAsync)
            if ($ps.HadErrors) {
                $msgs = $ps.Streams.Error | ForEach-Object {
                    $texte = $_.ToString()
                    if ($_.Exception -and $_.Exception.Message) { $texte = "$texte [$($_.Exception.GetType().FullName)] $($_.Exception.Message)" }
                    if ($_.ScriptStackTrace) { $texte = "$texte -- $($_.ScriptStackTrace -replace '[\r\n]+', ' > ')" }
                    $texte
                }
                if (-not $msgs -or $msgs.Count -eq 0) { $msgs = @("(HadErrors=true mais Streams.Error vide - cause inconnue)") }
                Ajouter-Journal "[$($grp.Nom)] Erreur traitement : $($msgs -join ' | ')"
                $grp.Statut = "Erreur - voir journal"
            } elseif ($sortie -and $sortie.Count -gt 0) {
                $res = $sortie[0]
                foreach ($ligne in $res.Journal) { Ajouter-Journal $ligne }
                $grp.Statut = $res.Statut
                if ($res.CheminSortieStl) { $grp.CheminSortieStl = $res.CheminSortieStl }
                # 27/10/2026 : demande de Tomyn - l'avertissement "piece pas reellement en
                # contact" doit sauter aux yeux meme pour quelqu'un qui ne regarde jamais le
                # journal (pense a publier l'appli sur un forum - un autre utilisateur ne lira
                # pas le journal). Contrairement a la notification de fin de lot plus bas, on
                # l'affiche MEME si la fenetre a le focus : Tomyn dit explicitement qu'il ne
                # suit pas le journal en continu meme quand l'appli est ouverte.
                $avertissementsContact = @($res.Journal | Where-Object { $_ -match "ne semble REELLEMENT en contact" })
                if ($avertissementsContact.Count -gt 0) {
                    try {
                        $script:IconeNotification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
                        $script:IconeNotification.BalloonTipTitle = "$($grp.Nom) : pieces peut-etre non assemblees"
                        $texteContact = if ($avertissementsContact.Count -eq 1) { "1 piece ne semble pas reellement en contact avec le reste - verifie le journal avant d'imprimer." } else { "$($avertissementsContact.Count) pieces ne semblent pas reellement en contact avec le reste - verifie le journal avant d'imprimer." }
                        $script:IconeNotification.BalloonTipText = $texteContact
                        $script:IconeNotification.ShowBalloonTip(10000)
                    } catch {}
                }
            } else {
                Ajouter-Journal "[$($grp.Nom)] Aucun resultat renvoye par le traitement."
                $grp.Statut = "Erreur - pas de resultat"
            }
        } catch {
            $doitNettoyer = $true
            try { $minuteur.Stop() } catch {}
            try { $grp.DureeSecondes = ((Get-Date) - $debutFigurine).TotalSeconds } catch {}
            $detailFlux = ""
            try { $detailFlux = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join ' | ' } catch {}
            $texteErreur = "[$($_.Exception.GetType().FullName)] $($_.Exception.Message)"
            if ($detailFlux) { $texteErreur = "$texteErreur -- flux erreurs : $detailFlux" }
            Ajouter-Journal "[$($grp.Nom)] Erreur traitement (tick) : $texteErreur"
            $grp.Statut = "Erreur - voir journal"
        } finally {
            if ($doitNettoyer) {
                # 27/10/2026 : chrono ajoute ici - c'est la DERNIERE zone non mesuree : tout ce
                # qui tourne sur le thread UI entre la fin du traitement en arriere-plan
                # (EndInvoke deja fait plus haut) et le vrai retour au message-loop. Hypothese a
                # verifier avec de vrais chiffres (pas de deviner) : Arreter-ApercuJob appelle
                # $job.PS.Stop() de facon SYNCHRONE/BLOQUANTE sur ce meme thread UI - si un
                # apercu precedent tourne encore, Stop() peut attendre que ce calcul se termine
                # de lui-meme (aucun point de controle d'annulation dans la boucle de rendu).
                $chronoNettoyageUI = [System.Diagnostics.Stopwatch]::StartNew()
                try { $ps.Dispose() } catch {}
                try { $minuteur.Dispose() } catch {}
                try { Rafraichir-ListeGroupes } catch {}
                # si la figurine qui vient de finir est celle actuellement affichee dans les
                # apercus, on bascule automatiquement l'apercu de sortie sur le vrai resultat
                # (sans ca, il faudrait recliquer sur la ligne pour voir le changement).
                try {
                    if ($liste.SelectedItems.Count -gt 0 -and $liste.SelectedItems[0].Tag -eq $grp) {
                        Rafraichir-ApercusPour $grp
                    }
                } catch {}
                $chronoNettoyageUI.Stop()
                Ajouter-Journal "[$($grp.Nom)] (duree nettoyage interface apres traitement - liste + apercus : $([Math]::Round($chronoNettoyageUI.Elapsed.TotalSeconds, 1)) s)"
                # $etatPartage (PAS $script:EtatTraitement) - voir commentaire plus haut.
                $etatPartage.EnCours = $null
                Demarrer-ProchainTraitement $dossierSortie $cheminBambuExe $orientationAuto
            }
        }
    }.GetNewClosure()
    $minuteur.Add_Tick($gestionnaireTick)
    $minuteur.Start()
}

# ============================================================================
# OVERLAY DE PROGRESSION (meme esprit visuel que Coloriser-STL-GUI/Melange-Couleurs-GUI :
# fenetre sans bordure, flottant au-dessus de la fenetre principale, barre "maison" en
# Panel plutot que la ProgressBar grise de Windows) - demande de Tomyn plutot qu'un
# simple message dans le journal.
# ============================================================================
$overlayProgression = New-Object System.Windows.Forms.Form
$overlayProgression.FormBorderStyle = "None"
$overlayProgression.ShowInTaskbar = $false
$overlayProgression.TopMost = $true
$overlayProgression.StartPosition = "Manual"
$overlayProgression.BackColor = [System.Drawing.Color]::FromArgb(26, 26, 30)
$overlayProgression.Size = New-Object System.Drawing.Size(440, 160)

$labelOverlayTitre = New-Object System.Windows.Forms.Label
$labelOverlayTitre.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$labelOverlayTitre.ForeColor = $THEME.Texte
$labelOverlayTitre.BackColor = [System.Drawing.Color]::Transparent
$labelOverlayTitre.AutoEllipsis = $true
$labelOverlayTitre.TextAlign = "MiddleCenter"
$labelOverlayTitre.Location = New-Object System.Drawing.Point(20, 20)
$labelOverlayTitre.Size = New-Object System.Drawing.Size(400, 44)
$overlayProgression.Controls.Add($labelOverlayTitre)

$pisteOverlay = New-Object System.Windows.Forms.Panel
$pisteOverlay.Location = New-Object System.Drawing.Point(30, 75)
$pisteOverlay.Size = New-Object System.Drawing.Size(380, 10)
$pisteOverlay.BackColor = $THEME.FondChamp
$overlayProgression.Controls.Add($pisteOverlay)

$remplissageOverlay = New-Object System.Windows.Forms.Panel
$remplissageOverlay.Location = New-Object System.Drawing.Point(0, 0)
$remplissageOverlay.Size = New-Object System.Drawing.Size(0, 10)
$remplissageOverlay.BackColor = $THEME.AccentPrincipal
$pisteOverlay.Controls.Add($remplissageOverlay)

$labelOverlayPct = New-Object System.Windows.Forms.Label
$labelOverlayPct.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$labelOverlayPct.ForeColor = $THEME.TexteAttenue
$labelOverlayPct.BackColor = [System.Drawing.Color]::Transparent
$labelOverlayPct.TextAlign = "MiddleCenter"
$labelOverlayPct.Location = New-Object System.Drawing.Point(20, 95)
$labelOverlayPct.Size = New-Object System.Drawing.Size(400, 22)
$overlayProgression.Controls.Add($labelOverlayPct)

$btnArreterOverlay = New-Object System.Windows.Forms.Button
$btnArreterOverlay.Text = "Arreter"
$btnArreterOverlay.Location = New-Object System.Drawing.Point(170, 122)
$btnArreterOverlay.Size = New-Object System.Drawing.Size(100, 28)
$btnArreterOverlay.FlatStyle = "Flat"
$btnArreterOverlay.FlatAppearance.BorderColor = $THEME.Bordure
$btnArreterOverlay.ForeColor = [System.Drawing.Color]::White
$btnArreterOverlay.BackColor = [System.Drawing.Color]::FromArgb(190, 70, 70)
$btnArreterOverlay.Cursor = "Hand"
$overlayProgression.Controls.Add($btnArreterOverlay)
$btnArreterOverlay.Add_Click({
    # 27/10/2026 : en plus d'empecher la figurine SUIVANTE de demarrer ($script:DoitArreter),
    # on signale aussi l'arret a la figurine EN COURS via l'etat partage avec le runspace
    # (voir Traiter-UneFigurineCore) - sans ca, "Arreter" ne faisait rien tant qu'un
    # traitement etait deja en vol (rapporte par Tomyn).
    $script:DoitArreter = $true
    $script:EtatAnnulationPartage.DoitArreter = $true
    $btnArreterOverlay.Enabled = $false; $btnArreterOverlay.Text = "Arret..."
})

function Afficher-OverlayProgression([string]$titre) {
    $labelOverlayTitre.Text = $titre
    $remplissageOverlay.Width = 0
    $labelOverlayPct.Text = "0%"
    $btnArreterOverlay.Enabled = $true
    $btnArreterOverlay.Text = "Arreter"
    $overlayProgression.Location = New-Object System.Drawing.Point(
        ($form.Location.X + [int](($form.Width - $overlayProgression.Width) / 2)),
        ($form.Location.Y + [int](($form.Height - $overlayProgression.Height) / 2)))
    Arrondir-Bouton $overlayProgression 14
    Arrondir-Bouton $btnArreterOverlay 6
    $overlayProgression.Show()
    $overlayProgression.BringToFront()
}

function Mettre-A-Jour-OverlayProgression([int]$pct, [string]$titre = $null) {
    if ($titre) { $labelOverlayTitre.Text = $titre }
    $pctBorne = [Math]::Max(0, [Math]::Min(100, $pct))
    $remplissageOverlay.Left = 0
    $remplissageOverlay.Width = [int]($pisteOverlay.Width * $pctBorne / 100)
    $labelOverlayPct.Text = "$pctBorne%"
}

function Masquer-OverlayProgression {
    $overlayProgression.Hide()
}

function Masquer-OverlayProgressionApresDelai([int]$delaiSecondes = 3) {
    $minuteurFin = New-Object System.Windows.Forms.Timer
    $minuteurFin.Interval = $delaiSecondes * 1000
    $minuteurFin.Add_Tick({
        $minuteurFin.Stop(); $minuteurFin.Dispose()
        Masquer-OverlayProgression
    }.GetNewClosure())
    $minuteurFin.Start()
}

$boutonTraiter.Add_Click({
    $cibles = @($liste.SelectedItems | ForEach-Object { $_.Tag })
    if ($cibles.Count -eq 0) { $cibles = $script:Groupes }
    Lancer-Traitement $cibles
})
$boutonTraiterTout.Add_Click({ Lancer-Traitement $script:Groupes })
$boutonArreter.Add_Click({ $script:DoitArreter = $true; $script:EtatAnnulationPartage.DoitArreter = $true })
$boutonRelancerEchecs.Add_Click({
    $echecs = @($script:Groupes | Where-Object { $_.Statut -and ($_.Statut -like "Erreur*" -or $_.Statut -like "Echec*") })
    if ($echecs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucune figurine en echec/erreur a relancer.", "Fusion figurines")
        return
    }
    Lancer-Traitement $echecs
})

# 27/10/2026 : demande de Tomyn - une vraie notification Windows (pas juste le toast interne
# a l'appli, invisible si la fenetre est minimisee/en arriere-plan) quand un traitement se
# termine, vu que certains runs (orientation auto sur gros maillage) peuvent prendre 10-25+
# minutes. NotifyIcon + ShowBalloonTip est traduit automatiquement en notification native du
# Centre de notifications sur Windows 10/11.
$script:IconeNotification = New-Object System.Windows.Forms.NotifyIcon
$script:IconeNotification.Icon = [System.Drawing.SystemIcons]::Application
$script:IconeNotification.Text = "Fusion figurines"
$script:IconeNotification.Visible = $true

$form.Add_FormClosing({
    foreach ($cle in @($script:ApercuJobs.Keys)) { Arreter-ApercuJob $cle }
    Fermer-RunspacesPartagees
    $script:IconeNotification.Visible = $false
    $script:IconeNotification.Dispose()
})

# 27/10/2026 : applique UNE FOIS, apres que tous les controles existent, les finitions que
# Repeindre-ControlesTheme ajoute (bordure fine sur les boutons neutres) mais que les sites de
# creation individuels ne posent pas eux-memes - evite de dupliquer cette logique partout.
Repeindre-ControlesTheme $form.Controls

$form.Add_Shown({
    $form.Activate()
    if ($script:CheminJournalFichier) {
        Ajouter-Journal "Journal egalement enregistre sur disque : $script:CheminJournalFichier"
    } else {
        Ajouter-Journal "ATTENTION : impossible de creer le fichier journal sur disque (dossier '$script:DossierLogs' inaccessible) - journal visible ici uniquement, perdu si l'appli crashe."
    }
})
[System.Windows.Forms.Application]::Run($form)
