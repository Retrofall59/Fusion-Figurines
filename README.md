# Fusion-Figurines-GUI

Interface graphique PowerShell pour fusionner en masse des figurines 3D multi-pièces (plusieurs fichiers `.stl` par figurine) en un seul fichier imprimable, via le moteur de Bambu Studio en ligne de commande.

## Pourquoi

Beaucoup de figurines trouvées en ligne (Thingiverse, MyMiniFactory, etc.) sont livrées découpées en plusieurs `.stl` (socle, bras, armes, accessoires...) qui s'assemblent à l'impression. Les recoller un par un dans un slicer, figurine après figurine, devient vite insupportable quand on en a des dizaines. Cet outil scanne un dossier entier, détecte automatiquement les figurines multi-pièces, et les fusionne en lot.

## Fonctionnalités

- **Scan automatique** d'une arborescence de dossiers : chaque sous-dossier contenant 2 `.stl` ou plus est traité comme une figurine.
- **Fusion locale rapide** des pièces (positions d'origine conservées) avant envoi à Bambu Studio.
- **Détection de pièces mal assemblées** : alerte si une pièce semble trop écartée des autres, ou ne touche vraiment aucune autre pièce (risque d'impression en morceaux détachés).
- **Aperçus visuels** (entrée/sortie) générés sans dépendance 3D externe.
- **Exclusion de pièces à la volée** (clic sur une case) sans relancer tout le scan, avec mise à jour immédiate du nombre de triangles réellement concerné.
- **Avertissement sur les gros maillages** avant de lancer une orientation automatique potentiellement très longue.
- **Traitement en arrière-plan** : l'interface reste utilisable pendant qu'un lot tourne.
- **Rapport CSV** en fin de lot.

## Prérequis

- Windows, PowerShell 5.1 ou supérieur.
- [Bambu Studio](https://bambulab.com/en/download/studio) installé (le chemin vers `bambu-studio.exe` se configure dans l'interface).

## Utilisation

Double-clique `lanceur-fusion-figurines.vbs` (évite la fenêtre console et les soucis de politique d'exécution PowerShell), ou lance directement :

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Fusion-Figurines-GUI.ps1
```

Renseigne le dossier racine contenant tes figurines, le dossier de sortie, et le chemin vers Bambu Studio, puis lance le scan.

## Licence

MIT - fais-en ce que tu veux, aucune garantie.
