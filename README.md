# SIFACpy
Routine pour automatiser le traitement des extractions SIFAC grâce a une base de données des libellés et des noms de personnes d'un laboratoire.

## Installation

Assurez-vous d'avoir une installation de python3 fonctionnelle, avec les librairies suivantes : numpy, pandas, glob, re, xlrd

Copier coller le dossier SIFACpy dans le repertoire contenant les extractions SIFAC

Dans le dossier SIFACpy, modifier le fichier parametres_sifac.xls pour avoir les bons noms de fichiers, correspondant aux extractions a traiter. Attention, si les fichiers sont dans un repertoire un niveau en dessous du repertoire SIFACpy, il faut indiquer ../nom_de_fichier.xls comme nom de fichier.

## Lancement 

Apres chaque nouvelle extraction, lancer le traitement en double-cliquant sur sifac.py. Un fichier extraction_proc.xls sera créé, avec les associations automatiquement insérées dans une premiere colonne.



