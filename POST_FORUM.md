## [Partage] Fusion-Figurines-GUI : parce que recoller 40 pièces de STL à la main, c'est un plan de carrière que j'ai décliné

Bonsoir la populace de l'impression 3D,

Soyons honnêtes trente secondes, ce qui n'arrivera plus jamais après ce post : si vous téléchargez des figurines multi-pièces sur Internet comme tout le monde, vous passez votre vie à faire "Importer → sélectionner les 6 fichiers → Oui je veux les fusionner → Orienter → Trancher" à la main, pièce par pièce, dossier par dossier, en vous demandant pourquoi vous avez choisi ce loisir plutôt que la pêche. Moi aussi. Jusqu'à ce que je décide que mon temps libre méritait mieux qu'un clic répétitif, et que j'aille emmerder une IA pendant plusieurs heures pour qu'elle le fasse à ma place. Résultat : **Fusion-Figurines-GUI**, un outil PowerShell qui scanne un dossier entier de figurines et les fusionne en lot via Bambu Studio, pendant que vous, vous faites autre chose de votre vie (regarder l'écran en attendant, par exemple).

### Ça fait quoi, concrètement

- Vous lui donnez un dossier plein de sous-dossiers, chacun contenant les morceaux d'une figurine (socle, bras, épée, tête qui a roulé sous le meuble, etc.).
- Il détecte tout seul lesquels sont des figurines multi-pièces (2 `.stl` ou plus dans le dossier - ouais, le seuil d'intelligence requis n'est pas vertigineux, mais c'est largement suffisant).
- Il les recolle localement, les envoie à Bambu Studio en ligne de commande pour l'orientation/export, et vous ressort un `.stl` unique prêt à trancher. En lot. Sans vous demander votre avis à chaque figurine, ce qui est précisément le but.
- Il vous prévient si une pièce a l'air de flotter dans le vide à 2mm de sa copine au lieu d'être vraiment en contact (parce que "ça a l'air collé dans l'aperçu" et "c'est collé en vrai" sont deux affirmations qui n'ont strictement rien à voir).
- Il vous engueule gentiment si le maillage est énorme et que vous avez laissé "orientation automatique" cochée, parce que Bambu Studio sur un gros fichier avec l'auto-orient, c'est un peu le mode "va chercher un café, reviens dans 20 minutes, reviens encore dans 20 minutes".
- Vous pouvez décocher une pièce foireuse d'un clic sans tout re-scanner, et le compteur de triangles affiché suit enfin la réalité au lieu de vous afficher fièrement un nombre qui ne correspond à rien (oui, il y a eu un bug là-dessus, je l'assume, on va y revenir).

### Le chapitre "performance" que personne n'a demandé mais que je vous raconte quand même

Petite figurine, 200 000 triangles : traité en quelques secondes, tout va bien, le monde est beau.
Grosse figurine, 1,46 million de triangles : l'outil part en retraite anticipée. Dix, vingt, vingt-cinq minutes. Peut-être plus. J'ai arrêté de chronométrer par pudeur.

Ce qui m'a vraiment agacé, ce n'est pas que ce soit plus long sur un plus gros fichier - ça, c'est la vie, la physique, tout ça. C'est que ce soit **bien plus lent que ça n'aurait dû l'être**, style le temps n'augmentait pas avec le nombre de triangles, il explosait avec. J'ai donc fini par traquer ça sérieusement : chronomètres plantés partout dans le script, comparaison ligne par ligne, et découverte qu'une bonne partie du problème venait de PowerShell qui, en interne, préfère visiblement construire des objets avec emballage cadeau et petit mot de remerciement (`PSCustomObject`) plutôt qu'un simple tableau de courses (`Hashtable`) - multiplié par plusieurs millions de triangles, ça chiffre vite. Ajoutez à ça des appels de fonction à l'unité pour transformer chaque sommet un par un (3 appels par triangle, fois 1,46 million, merci bonsoir), et une fonction d'écriture du `.stl` final qui créait un petit tableau jetable à chaque triangle juste pour le plaisir de boucler dessus.

Résultat après correctifs successifs (benchmarkés en bac à sable avant d'être appliqués en vrai, parce que "je pense que ça devrait aller plus vite" n'est pas une méthode de travail, c'est une prière) : la même figurine de 1,46 million de triangles, qui ne finissait pas en 25 minutes au début, tourne maintenant en **39 secondes**. Pas d'extrapolation foireuse, pas de "ça devrait être mieux en théorie" : mesuré, journal à l'appui, sur la machine réelle.

Et au passage, le compteur "ça tourne depuis X secondes" dans le journal affichait religieusement sa première ligne à 16 secondes pile, quelle que soit la taille du fichier. Spoiler : ce n'était pas un mystérieux délai de démarrage, c'était PowerShell qui arrondit `[int]` au lieu de tronquer comme à peu près tout le reste de l'informatique depuis 1970. `[int](16/30)` vaut 1. Demandez-lui pas pourquoi, il ne le sait pas lui-même.

### Prérequis

- Windows (évidemment, on parle de PowerShell et de `.exe`, pas de rêver).
- Bambu Studio installé, avec son exécutable accessible (le chemin se configure dans l'interface, pas besoin de bricoler le PATH comme un sauvage).
- Des figurines multi-pièces qui vous ont fait perdre la volonté de vivre à force de les fusionner à la main.

### Où c'est

Code sur GitHub, lien plus bas. C'est gratuit, c'est du PowerShell donc moche à regarder mais ça marche, et c'est fourni "en l'état" - si ça met le feu à votre imprimante, à votre PC ou à votre couple, je ne suis responsable de rien, lisez la licence MIT, elle est très claire sur ce point précis.

Bon tranchage.
