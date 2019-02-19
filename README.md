# quadrafix
================
Ligne de commande pour automatiser l'import d'écriture dans Quadra Compta.
Cet outil vient compléter la fonction Import ASCII de quadra qui ne permet d'importer des écritures que via la GUI.
Dans une configuration où l'on souhaite automatiser l'import d'écritures de bout en bout ce type d'outil est nécessaire.
Le programme suit les étapes suivantes :
* ouverture du dossier qcompta.mdb pour récupérer les informations de base du dossier (date d'exercice, plan comptable, sortie monoposte, etc.)
* mise à jour de la table Ecritures
    * vérification de la période (cloturée ou non)
    * création du compte si absent
    * affectation analytique en fonction du code fournit ou des affectations par défaut
    * copie de la pièce comptable dans le dossier \Images, si présente
* mise à jour table Centralisateurs
* mise à jour soldes dans la table Comptes

__Attention__ : l'outil n'utilise pas le mécanisme de verrouillage QLocks (pipe). Ce sera à l'utilisateur de vérifier qu'il n'y aura pas de collision avec unsaisie en cours.

exemples :
`python.exe quadrafix.py -i mono.ipl -b dc -d 752 -f .\samples\ecr-pj.csv`
`quadrafix.exe -i quadra.ipl -d DOSSIER -b DC -f ecr.csv`

Pour les utilisateurs de Quadra Compta (Cegid), cet outil permet d'insérer dans la table Ecritures des écritures issues d'un fichier CSV.

usage: quadrafix.py [-h] [-f FICHIER] [-d DOSSIER] [-b BASE] [-i IPL] [-v]
                    [--version]

optional arguments:
  -h, --help            show this help message and exit
  -f FICHIER, --fichier FICHIER
                        chemin du fichier CSV contant les écritures
  -d DOSSIER, --dossier DOSSIER
                        code dossier quadra
  -b BASE, --base BASE  base dossier quadra (DC, DS201812, ...)
  -i IPL, --ipl IPL     chemin fichier IPL
  -v, --verbose         mode debug
  --version             show program's version number and exit

  Le fichier CSV doit prendre le format suivant :
> journal; date; compte; libellé, débit; crédit; pièce; image ; centre

Détail du fichier CSV:
* journal = code journal (3 car.)
* date = jj/mm/aaaa
* compte = numéro compte (8 car. max)
* libellé = libellé (30 car. max)
* débit = montant débit
* crédit = montant crédit
* piece = numéro de piece
* image = fichier de la pièce comptable
* centre = code analytique
