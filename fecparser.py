import csv
import logging
from datetime import datetime


class fecparser(object):
    def __init__(self, filepath):

        self.csv = open(filepath, "r")
        self.data = {}
        self.pc = []
        self.ecr = []

    def read(self):

        headline = self.csv.readline()

        if len(headline.split("\t")) > 0:
            delimiter = "\t"
            delim_is = "tab"
        elif len(headline.split("|")) > 0:
            delimiter = "|"
            delim_is = "pipe"

        logging.info("delimiter : {}".format(delim_is))

        for row in csv.DictReader(self.csv, delimiter=delimiter):
            print("-" * 20)
            print(row)

            ecr_dic = {}

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
                if key in full_fec_row:
                    ecr_dic.update(key, row[key])


            # for key in row.keys():
            #     print(key, row[key])

            # JournalCode = row["JournalCode"]
            # if "JournalLib" in headers
            # JournalLib = ""
            # EcritureDate = ""
            # CompteNum = ""
            # CompteLib = ""
            # CompAuxNum = ""
            # CompAuxLib = ""
            # PieceRef = ""
            # PieceDate = ""
            # EcritureLib = ""
            # Debit = ""
            # Credit = ""
            # EcritureLet = ""

            # self.pc.append((row["CompteNum"], row["CompteLib"]))
        # print (headers)

        # self.f.seek(0)
        # self.data = csv.DictReader(self.f, delimiter=delimiter)
        # count = len(list(self.data))
        # self.f.seek(0)

        # if not count:
        #     logging.warning("0 line to process")
        #     return False
        # else:
        #     logging.info("nb lines : {}".format(count))
        #     # headers = self.data.
        #     # print(headers)
        #     for row in self.data:
        #         print(row)
        # self.pc.append((row["CompteNum"],row["CompteLib"]))
        # self.ecr.append([
        #     row["JournalCode"],
        #     datetime.strptime(row["EcritureDate"], "%Y%m%d"),
        #     row["CompteNum"],
        #     row["PieceRef"],
        #     row["EcritureLib"],

        # ]

        # )

        # self.f.seek(0)
        # for row in self.data:
        #     print(row["EcritureLib"])

    def plan_compt(self):
        return self.pc

    def close(self):
        self.csv.close()


if __name__ == "__main__":

    import pprint

    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(
        level=logging.DEBUG, format="%(funcName)s\t\t%(levelname)s - %(message)s"
    )

    fec = "./samples/FECs.txt"

    o = fecparser(fec)
    o.read()
    print(o.plan_compt())
    o.close()
