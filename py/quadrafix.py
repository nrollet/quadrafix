import argparse
import logging
import os
from preparecsv import Prepare_Csv
from quadratools import QueryCompta, quadra_env
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def progressbar(count, total):
    """
    Pour l'affichage d'une barre de progression
    pendant l'insert des écritures
    """
    bar = "#" * 40
    level = int((count * 40) / total) + 1
    tail = "\r"
    if count == total:
        tail = "\n"
    print(
        "[{}] {}/{}".format(
            bar[0:level].ljust(40), str(count).zfill(len(str(total))), str(total)
        ),
        end=tail,
    )

########### CLI PARSER

parser = argparse.ArgumentParser(description="Outil en ligne de commande pour l'import d'écritures dans Quadra Compta")
parser.add_argument("-f", "--fichier", required=False, help="chemin du fichier CSV contant les écritures")
parser.add_argument("-d", "--dossier", required=False, help="code dossier quadra")
parser.add_argument(
    "-b", "--base", required=False, help="base dossier quadra (DC, DS201812, ...)"
)
parser.add_argument("-i", "--ipl", required=False, help="chemin fichier IPL")
parser.add_argument("-v", "--verbose", action="count", help="mode debug")
parser.add_argument("--version", action="version", version="%(prog)s 1.0")

args = vars(parser.parse_args())

if args["verbose"]:
    log_level = logging.DEBUG
else:
    log_level = logging.INFO
logging.basicConfig(
    level=log_level,
    format="%(asctime)s %(module)s:%(funcName)s\t%(levelname)s -- %(message)s",
    datefmt="%m-%d %H:%M",
    filename="quadrafix.log",
    filemode="w"
)

Tk().withdraw()

ipl = ""
if args["ipl"]:
    ipl = args["ipl"]
else:
    ipl = askopenfilename(title="Indiquez le fichier IPL", filetypes=[("IPL", "*.ipl")])
if not ipl:
    logging.warning("Annulation")
    exit(1)

fichier = ""
if args["fichier"]:
    fichier = args["fichier"]
else:
    while not fichier:
        fichier = askopenfilename(
            title="Indiquez le fichier CSV",
            filetypes=[("TXT", "*.txt"), ("CSV", "*.csv")],
        )

dossier = ""
if args["dossier"]:
    dossier = args["dossier"]
else:
    while not dossier:
        dossier = input("Code dossier : ")
dossier = dossier.zfill(6)

base = ""
if args["base"]:
    base = args["base"]
else:
    while not base:
        base = input("Code de la base comptable(DC, DA, DS) : ")

########### PREP FICHIER D'ECRITURES

csv = Prepare_Csv(fichier)

if csv.check_format():
    list = csv.read()

########## PREP ENVIR. QUADRA

qenv = quadra_env(ipl)

if qenv["SERVEUR"]:
    logging.info("mode serveur")
else:
    logging.info("mode monoposte")

chemin_cpta = qenv["RACDATACPTA"]

chemin_piece = os.path.dirname(os.path.realpath(fichier))

chemin_mdb = "{}{}/{}/qcompta.mdb".format(chemin_cpta, base, dossier)

########## ACTIONS

Q = QueryCompta()
Q.load_params(chemin_mdb)

if Q.param_doss["datesortie"]:
    logging.error("Le dossier est sorti")
    Q.close()
    exit()

i = 1

logging.info("Mise à jour de la table Ecritures")
for journal, date, compte, libelle, debit, credit, piece, image, centre in list:
    logging.debug("insert n°{}".format(i))
    Q.insert_ecrit(
        journal=journal,
        date=date,
        compte=compte,
        libelle=libelle,
        debit=debit,
        credit=credit,
        piece=piece,
        image=image,
        centre=centre,
        folio="000",
        image_root=chemin_piece,
    )
    progressbar(i, len(list))
    i += 1

########## FINITIONS

Q.maj_centralisateurs()
Q.maj_solde_comptes()
Q.close()
