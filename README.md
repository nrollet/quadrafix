# quadrafix
ligne de commande pour automatiser l'import d'écriture dans Quadra Compta

exemple :
    quadrafix.exe -i quadra.ipl -d DOSSIER -b DC -f ecr.csv

Pour les utilisateurs de Quadra Compta (Cegid), cet outil permet d'insérer dans la table Ecritures des écritures issues d'un fichier CSV.
Le fichier CSV doit prendre le format suivant :
code journal; date; compte; libellé, débit; crédit; image ; pièce

