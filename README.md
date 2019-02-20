# quadrafix

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

Les écritures sont fournies sous la forme d'un fichier au format CSV

__Attention__ : l'outil n'utilise pas le mécanisme de verrouillage QLocks (pipe). Ce sera à l'utilisateur de vérifier qu'il n'y aura pas de collision avec une saisie en cours.

Le fichier quadrafix.exe est téléchargeable dans le dossier /dist

__Utilisation__

`quadrafix.exe -i quadra.ipl -d DOSSIER -b DC -f ecr.csv`

usage: quadrafix.exe [-h] [-f FICHIER] [-d DOSSIER] [-b BASE] [-i IPL] [-v]
                     [--version]

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

__Exemple__
> AC;01/01/2019;0ZZZZZZZ;Achat PC;;999,99;"012345";infac.pdf
> AC;01/01/2019;60110000;Achat PC;999,99;;"012345";infac.pdf;004

__Prérequis__ : vous devrez installer le Microsoft Access Database Engine pour que Quadrafix puisse accéder à un qcompta.mdb.
https://www.microsoft.com/fr-fr/download/details.aspx?id=13255

