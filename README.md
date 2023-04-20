# XlsxKanapeche
Centralise facilement les données éparpillées sur plusieurs fichiers Excels

Fonctionne pour un nombre variable de fichiers sources, et un nombre variables d'onglets sources de même format dans chaque fichier

Concatène les entrées de tableaux sur les onglets dont le nom répond à une condition, dans divers fichiers dits "Sources" vers ce présent fichier dit "Cible"
Validé le 19/04/2023


# Package opérationnel :

Trame de remontage.xlsx

    > contextuelle à un chantier de centralisation d'exigences, validée à cet usage le 19/04/2023 sur la base des fichiers test "Source 1.xlsx" et "Source 2.xlsx" (dépersonnalisés des infos projet)
    
    > trame sur la plage de cellules A1:B23
    
    > Modules validés : Module 2 [ImportDataFromMultipleFiles()]
    
    > Modules immatures : Module 1 en chantier (routine de remise à zéro du fichier pour confort de test, pas fonctionnel)

# Package de test :
Source 1.xlsx : 2 onglets ciblés par la condition sur le nom "Résultat exigences"
Source 2.xlsx : 1 onglet ciblé par la condition sur le nom "Résultat exigences"

# Prérequis :
- utiliser une trame harmonieuse entre les fichiers sources et le fichier cible
- déposer les fichiers sources dans un même dossier au même niveau


# Flexibilité :
- il n'est pas nécessaire que ce fichier soit dans le même dossier que les sources
- d'autres fichiers peuvent coexister dans le fichier sources


# Paramètres du code :
- le chemin du dossier contenant les fichiers sources
> variable strPath = "Y:\L3-Data\Job STB NUWARDrevD\2-Remontages\Remontage 20230216\Sources\"]

- les dimensions de la trame du tableau
> variable premierRangLibre doit prendre la valeur 3 s'il y a deux lignes de trame
> largeur de trame exprimée en dur, ici 23 colonnes

- la condition sur le nom des onglets dont le contenu est collecté
> argument du test utilisé pour déterminer la valeur de la variable testWs juste à l'entrée de la boucle qui parcours

# WARNING : 
le module 2 ouvre et ferme l'ensemble des fichiers sources -> VEILLER A ENREGISTRER VOTRE TRAVAIL SUR L'ENSEMBLE DES FICHIERS SOURCES AVANT DE L'UTILISER
