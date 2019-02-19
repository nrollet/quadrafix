# quadrafix
ligne de commande pour automatiser l'import d'écriture dans Quadra Compta


exemples :
    python.exe quadrafix.py -i mono.ipl -b dc -d 752 -f .\samples\ecr-pj.csv
    quadrafix.exe -i quadra.ipl -d DOSSIER -b DC -f ecr.csv

Pour les utilisateurs de Quadra Compta (Cegid), cet outil permet d'insérer dans la table Ecritures des écritures issues d'un fichier CSV.
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