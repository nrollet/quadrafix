import csv
import logging
from datetime import datetime


class fecparser(object):
    """
    Outil pour convertir le fichier FEC en dictionary
    Va compléter s'il manque des colonnes
    Va convertir certaines string en date ou float
    """
    def __init__(self, filepath):

        self.csv = open(filepath, "r")
        self.pc = {}
        self.ecr = []

    def read(self):

        headline = self.csv.readline()

        # test du délimiteur
        if len(headline.split("\t")) > 0:
            delimiter = "\t"
            delim_is = "tab"
        elif len(headline.split("|")) > 0:
            delimiter = "|"
            delim_is = "pipe"

        logging.info("delimiter : {}".format(delim_is))

        self.csv.seek(0)

        for row in csv.DictReader(self.csv, delimiter=delimiter):

            full_fec_row = {
                "JournalCode": "",
                "JournalLib": "",
                "EcritureNum": "",
                "EcritureDate": "",
                "CompteNum": "",
                "CompteLib": "",
                "CompAuxNum": "",
                "CompAuxLib": "",
                "PieceRef": "",
                "PieceDate": "",
                "EcritureLib": "",
                "Debit": "",
                "Credit": "",
                "EcritureLet": "",
                "DateLet": "",
                "ValidDate": "",
                "Montantdevise": "",
                "Idevise": "",
            }

            for key in row.keys():
                if key in full_fec_row.keys() and row[key]:
                    if key in ["Debit", "Credit"]:
                        row[key] = float(row[key].replace(",", "."))
                    if key in ["EcritureDate", "PieceDate", "DateLet", "ValidDate"]:
                        row[key] = datetime.strptime(row[key], "%Y%m%d")
                    full_fec_row.update({key: row[key]})

            self.ecr.append(full_fec_row)

        return self.ecr

    def plan_compt(self):

        for ecr in self.ecr:
            self.pc.setdefault(ecr["CompteNum"], ["", 0.0, 0.0])
            self.pc[ecr["CompteNum"]][0] = ecr["CompteLib"]
            self.pc[ecr["CompteNum"]][1] += float(ecr["Debit"])
            self.pc[ecr["CompteNum"]][2] += float(ecr["Credit"])

        return self.pc

    def resultat(self):

        debit, credit = 0.0, 0.0
        for compte, val in self.pc.items():
            if compte.startswith("6") or compte.startswith("7"):
                debit += val[1]
                credit += val[2]

        return debit - credit

    def close(self):
        self.csv.close()


if __name__ == "__main__":

    import pprint

    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(
        level=logging.DEBUG, format="%(funcName)s\t\t%(levelname)s - %(message)s"
    )

    fec = "./samples/FEC.txt"

    o = fecparser(fec)
    # pp.pprint(o.read())
    o.read()
    pp.pprint(o.plan_compt())
    print(o.resultat())
    o.close()
