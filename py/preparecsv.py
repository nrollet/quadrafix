import csv
import logging
import re
from datetime import datetime
import pprint

pp = pprint.PrettyPrinter(indent=4)


class Prepare_Csv(object):
    def __init__(self, filepath):

        self.filepath = filepath

        with open(filepath, "r", encoding="cp1252") as f:
            self.lines = f.readlines()

    def check_format(self):

        COL_JRN = 0
        COL_DAT = 1
        COL_CPT = 2
        COL_LIB = 3
        COL_DEB = 4
        COL_CRE = 5
        COL_PCE = 6
        COL_IMG = 7
        COL_ANA = 8

        if len(self.lines) > 0:
            logging.info("{}, {} lignes".format(self.filepath, len(self.lines)))
        else:
            logging.error("{} est vide".format(self.filepath))
            return False

        sep_type = ""
        if len(self.lines[0].split(";")) > 1:
            sep_type = ";"
            sep_intit = sep_type
        if len(self.lines[0].split("\t")) > 1:
            sep_type = "\t"
            sep_intit = "tab"

        if sep_type:
            logging.info("séparateur : {}".format(sep_intit))
        else:
            logging.error("séparateur introuvable")
            return False

        self.data = []
        i = 1
        for line in self.lines:
            line = line.strip()
            splited = line.split(sep_type)

            if len(splited):

                splited.extend([""] * (9 - len(splited)))
                print(splited)

                journal = ""
                date = ""
                compte = ""
                debit = 0.0
                credit = 0.0
                image = ""
                piece = ""
                libelle = splited[COL_LIB][:30]
                piece = splited[COL_PCE][:10]
                centre = ""

                # nettoyage du libelle
                for ch in ["'", ";", "%"]:
                    if ch in libelle:
                        libelle = libelle.replace(ch, " ")

                if re.match(r"\w{1,3}", splited[COL_JRN]):
                    journal = splited[COL_JRN]
                else:
                    logging.error("journal incompatible : {}".format(splited[COL_JRN]))
                    return False

                if re.match(r"[0-3][0-9]/(0[1-9]|1[0-2])/20[0-9]{2}", splited[COL_DAT]):
                    date = datetime.strptime(splited[COL_DAT], "%d/%m/%Y")
                else:
                    logging.error("date incompatible : {}".format(splited[COL_DAT]))
                    return False

                if re.match(r"\w{1,8}", splited[COL_CPT]):
                    compte = splited[COL_CPT].upper()
                else:
                    logging.error("compte incompatible : {}".format(splited[COL_CPT]))
                    return False

                if re.match(r"(\d+(,|\.)?\d+)", splited[COL_DEB]):
                    debit = float(splited[COL_DEB].replace(",", "."))

                if re.match(r"(\d+(,|\.)?\d+)", splited[COL_CRE]):
                    credit = float(splited[COL_CRE].replace(",", "."))

                if re.match(
                    r"([^:^\?^\"]*)\.(pdf|tif|doc|docx|xls|xlsx)",
                    splited[COL_IMG],
                    re.IGNORECASE,
                ):
                    image = splited[COL_IMG]

                if re.match(r"\w{1,5}", splited[COL_ANA]):
                    centre = splited[COL_ANA]

                # if splited[COL_LET]:
                #     lettre = splited[COL_LET]

            self.data.append(
                [journal, date, compte, libelle, debit, credit, piece, image, centre]
            )
        i += 1
        return True

    def read(self):
        return self.data


if __name__ == "__main__":

    import pprint
    import os

    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(
        level=logging.DEBUG, format="%(funcName)s\t\t%(levelname)s - %(message)s"
    )

    csv_path = "ecr.csv"
    o = Prepare_Csv(csv_path)
    o.check_format()
    # pp.pprint(o.read())
    # pp.pprint(o.check_balance())

