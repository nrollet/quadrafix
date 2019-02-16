import csv
import logging
from datetime import datetime


class fecparser(object):
    def __init__(self):

        self.data = []

    def read(self, filepath):

        with open(filepath, "r") as f:
            heads = f.readline()
            print(heads)

            if len(heads.split("\t")) > 0:
                delimiter = "\t"
                delim_is = "tab"
            elif len(heads.split("|")) > 0:
                delimiter = "|"
                delim_is = "pipe"

            logging.info("delimiter : {}".format(delim_is))

            # data = csv.reader(f, delimiter=delimiter)
            f.seek(0)
            data = csv.DictReader(f, delimiter=delimiter)
            for row in data:
                print(row)


if __name__ == "__main__":

    import pprint

    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(
        level=logging.DEBUG, format="%(funcName)s\t\t%(levelname)s - %(message)s"
    )

    fec = "./samples/FECs.txt"

    o = fecparser()
    o.read(fec)
