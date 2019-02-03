import pyodbc
import logging
import sys
import random
import string
import os
import shutil
from datetime import datetime

def progressbar(count, total):
    """
    Pour l'affichage d'une barre de progression
    pendant l'insert des écritures
    """
    bar = "#"*40
    level = int((count*40)/total)+1
    tail = "\r"
    if count == total:
        tail = "\n"

    print("[{}] {}/{}".format(bar[0:level].ljust(40),
                            str(count).zfill(len(str(total))),
                            str(total)), end=tail)

def gen_image_name(reserved):
    """
    Création aléatoire de nom pour les images pieces
    Avec vérif si le fichier existe déjà
    """
    name = "".join(random.choice(string.ascii_uppercase) for _ in range(10))
    if name in reserved:
        gen_image_name(reserved)
    else:
        return name



def quadra_env(ipl):
    """
    Collecte des chemins dossiers à partir du fichier IPL
    """
    with open(ipl) as f:
        txt = f.readlines()
    
    dic = {}

    for line in txt:
        line = line.replace("\n", "")
        line = line.replace("\\", "/")
        if "=" in line : 
            cle, valeur = line.split("=")[0:2]
            dic.update({cle : valeur})

    return dic



class QueryCompta(object):

    def __init__(self):

        self.param_doss = {
            "raisonsoc": "Société1",
            "exedebut": datetime(year=1890, month=1, day=1),
            "exefin": datetime(year=1890, month=12, day=31),
            "datevalid": "",
            "dateclot": "",
            "collectfrn": "40100000",
            "collectcli": "41100000",
            "prefxfrn": "0",
            "prefxcli": "9",
            "plan": {},
            "affect": {},
            "images": []
        }

    def load_params(self, chem_base):
        self.chem_base = chem_base.lower()

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + \
            self.chem_base
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(self.chem_base))
            self.cursor = self.conx.cursor()

        except pyodbc.Error:
            print("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
        except:
            print("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))

        # Paramètres table Dossier1
        sql = """
            SELECT
            RaisonSociale, DebutExercice, FinExercice,
            PeriodeValidee, PeriodeCloturee
            FROM Dossier1
        """
        self.cursor.execute(sql)
        for rs, exd, exf, pv, pc in self.cursor.fetchall():
            self.param_doss["raisonsoc"] = rs
            self.param_doss["exedebut"] = exd
            self.param_doss["exefin"] = exf
            self.param_doss["datevalid"] = pv
            self.param_doss["dateclot"] = pc

        # Paramètres table Dossier2
        sql = """
            SELECT
            CollectifFrnDefaut, CollectifClientDefaut,
            CodifClasse0Seule, CodifClasse9Seule
            FROM Dossier2
        """
        self.cursor.execute(sql)
        for colfrn, colcli, cl0, _ in self.cursor.fetchall():
            self.param_doss["collectfrn"] = colfrn
            self.param_doss["collectcli"] = colcli
            if cl0 == "C":
                self.param_doss["prefxfrn"] = "9"
                self.param_doss["prefxcli"] = "0"

        # Plan comptable
        dic_pc = {}
        sql = """
            SELECT
            Numero, Type, Intitule, NbEcritures, ProchaineLettre
            FROM Comptes
        """
        self.cursor.execute(sql)
        for num, _, intit, nbecr, lettr in self.cursor.fetchall():
            dic_pc.setdefault(num, [])
            dic_pc[num] = {"intitule": intit,
                        "nbrecr": nbecr,
                        "lettrage": lettr
                        }
        self.param_doss["plan"] = dic_pc

        # Affectations Ana
        dic_affect = {}
        sql = """SELECT NumCompte, CodeCentre FROM AffectationAna"""
        self.cursor.execute(sql)
        data = self.cursor.fetchall()
        if data:
            for compte, centre in data:
                dic_affect.update({compte : centre})
        self.param_doss["affect"] = dic_affect

        # Liste des images existantes
        img_list = []
        sql = """SELECT DISTINCT RefImage FROM Ecritures"""
        self.cursor.execute(sql)
        for image in self.cursor.fetchall():
            img_list.append(image[0].split(".")[0])
        self.param_doss["images"] = img_list

        logging.info("Nom : {}".format(self.param_doss["raisonsoc"]))
        logging.info("Exercice : {} {}".format(
            self.param_doss["exedebut"].strftime("%d/%m/%Y"),
            self.param_doss["exefin"].strftime("%d/%m/%Y")))
        logging.info("Coll. fourn : {}".format(self.param_doss["collectfrn"]))
        logging.info("Préfixe fourn : {}".format(self.param_doss["prefxfrn"]))
        logging.info("Préfixe client : {}".format(self.param_doss["prefxcli"]))
        logging.info("Comptes : {}".format(len(self.param_doss["plan"])))        

        return self.param_doss


    def verif_journal(self, journal):
        """
        Vérifie si compte présent dans le PC
        Return boolean
        """
        check = True
        sSQL = (
            "SELECT Code FROM Journaux "
            "WHERE code='{}'".format(journal)
        )
        data = self.cursor.execute(sSQL).fetchall()
        if not data:
            check = False
        return check

    def journal(self, code_journal):
        """
        Récupère le solde de tous les comptes d'un journal donné
        renvoi dictionnaire dont la clé est le n° de pièce
        """

        sql = """
            SELECT
            E.NumeroPiece, E.NumeroCompte, C.Intitule, E.PeriodeEcriture,
            SUM(E.MontantTenuDebit-E.MontantTenuCredit) AS Solde
            FROM Ecritures E
            LEFT JOIN Comptes C ON E.NumeroCompte=C.Numero
            WHERE CodeJournal='{}'
            AND TypeLigne='E'
            GROUP BY NumeroPiece, NumeroCompte, Intitule, PeriodeEcriture
        """.format(code_journal)

        data = self.cursor.execute(sql).fetchall()
        dic = {}
        for piece, compte, intitule, periode, solde in data:
            dic.setdefault(compte, {})
            dic[compte].update({"intitule": intitule})
            dic[compte].setdefault("piece", [])
            dic[compte]["piece"].append([piece, periode, solde])
        return dic

    def get_solde_compte(self, compte):

        sql = """
            SELECT
            SUM(MontantTenuDebit) AS Debit,
            SUM(MontantTenuCredit) AS Credit
            FROM Ecritures
            WHERE NumeroCompte='{}'
            AND TypeLigne='E'
        """.format(compte)
        self.cursor.execute(sql)
        return self.cursor.fetchall()[0]

    def close(self):
        logging.info('fermeture de la base')
        self.conx.commit()
        self.conx.close()

    def insert_compte(self, compte):
        """
        Ajout d'un nouveau compte dans la table Comptes
        """

        # Détection de la nature du compte
        if compte.startswith(self.param_doss["prefxfrn"]):
            type_cpt = "F"
            Collectif = self.param_doss["collectfrn"]
        elif compte.startswith(self.param_doss["prefxcli"]):
            type_cpt = "C"
            Collectif = self.param_doss["collectcli"]
        elif compte[0] in ["1", "2", "3", "4", "5", "6", "7"]:
            type_cpt = "G"
            Collectif = "NULL"
        else:
            logging.error("Compte {} hors PC".format(compte))
            return False

        Collectif = "'{}'".format(Collectif)

        TypeCollectif = False
        Debit = 0.0
        Credit = 0.0
        DebitHorsEx = 0.0
        CreditHorsEx = 0.0
        Debit_1 = 0.0
        Credit_1 = 0.0
        Collectif_1 = Collectif
        Debit_2 = 0.0
        Credit_2 = 0.0
        Collectif_2 = Collectif
        NbEcritures = False
        DetailCloture = 1
        ALettrerAuto = 1
        CentraliseGdLivre = False
        SuiviQuantite = False
        CumulPiedJournal = False
        TvaEncaissement = False
        DateSysCreation = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ProchaineLettre = "AA"
        NoProchainLettrage = False
        CodeTva = False
        EditM2 = False
        Franchise = False
        IntraCom = False
        NiveauDroit = 0
        GererIntCptCour = False
        MargeTheorique = 100
        TvaDOM = False
        Periodicite = False
        SuiviDevises = False
        CompteInactif = False
        BonAPayer = False
        QuantiteNbEntier = 5
        QuantiteNbDec = 2
        PrixMoyenNbEntier = 5
        PrixMoyenNbDec = 2
        PersonneMorale = False
        SuiviQuantite2 = False
        QuantiteNbEntier2 = 5
        QuantiteNbDec2 = 2
        PrixMoyenNbEntier2 = 5
        PrixMoyenNbDec2 = 2
        RepartitionAna = 0
        RepartitionAuto = False
        ActiverLotTrace = False
        TvaAutresOpeImpos = False
        ProchaineLettreTiers = "AAA"
        PrestaTel = False
        TypeIntraCom = 0
        Prestataire = 0
        CptParticulier = False

        logging.info("Ajout d'un nouveau compte : {} ({})".format(compte, type_cpt))

        sql = f"""
            INSERT INTO Comptes
            (Numero, Type, TypeCollectif,
            Debit, Credit, DebitHorsEx, CreditHorsEx,
            Collectif,
            Debit_1, Credit_1, Collectif_1,
            Debit_2, Credit_2, Collectif_2,
            NbEcritures, DetailCloture,
            ALettrerAuto, CentraliseGdLivre,
            SuiviQuantite, CumulPiedJournal,
            TvaEncaissement, DateSysCreation,
            ProchaineLettre, NoProchainLettrage,
            CodeTva,EditM2, Franchise,
            IntraCom, NiveauDroit, GererIntCptCour,
            MargeTheorique, TvaDOM, Periodicite,
            SuiviDevises, CompteInactif, BonAPayer,
            QuantiteNbEntier, QuantiteNbDec,
            PrixMoyenNbEntier, PrixMoyenNbDec,
            PersonneMorale, SuiviQuantite2,
            QuantiteNbEntier2, QuantiteNbDec2,
            PrixMoyenNbEntier2, PrixMoyenNbDec2,
            RepartitionAna, RepartitionAuto,
            ActiverLotTrace, TvaAutresOpeImpos,
            ProchaineLettreTiers, PrestaTel,
            TypeIntraCom,Prestataire,CptParticulier)
            VALUES
            ('{compte}', '{type_cpt}', {TypeCollectif},
            {Debit}, {Credit}, {DebitHorsEx}, {CreditHorsEx},
            {Collectif}, 
            {Debit_1}, {Credit_1}, {Collectif_1},
            {Debit_2}, {Credit_2}, {Collectif_2},
            {NbEcritures}, {DetailCloture},
            {ALettrerAuto}, {CentraliseGdLivre},
            {SuiviQuantite}, {CumulPiedJournal},
            {TvaEncaissement}, #{DateSysCreation}#,
            '{ProchaineLettre}', {NoProchainLettrage},
            {CodeTva}, {EditM2}, {Franchise},
            {IntraCom}, {NiveauDroit}, {GererIntCptCour},
            {MargeTheorique}, {TvaDOM}, {Periodicite},
            {SuiviDevises}, {CompteInactif}, {BonAPayer},
            {QuantiteNbEntier}, {QuantiteNbDec},
            {PrixMoyenNbEntier}, {PrixMoyenNbDec},
            {PersonneMorale}, {SuiviQuantite2},
            {QuantiteNbEntier2}, {QuantiteNbDec2},
            {PrixMoyenNbEntier2}, {PrixMoyenNbDec2},
            {RepartitionAna}, {RepartitionAuto},
            {ActiverLotTrace}, {TvaAutresOpeImpos},
            '{ProchaineLettreTiers}', {PrestaTel},
            {TypeIntraCom}, {Prestataire}, {CptParticulier})
        """
        try:
            self.cursor.execute(sql)
            logging.info("Mise à jour table Comptes OK")
            self.param_doss["plan"].update({compte : ""})
            return True
        except pyodbc.Error:
            logging.debug(sql)
            logging.error("erreur insert Comptes {}".format(sys.exc_info()[1]))
            return False

    def get_last_uniq(self):
        """
        Retourne le dernier identifiant unique
        de la liste des écritures
        """
        sSQL = "SELECT MAX(NumUniq) FROM Ecritures"
        self.cursor.execute(sSQL)
        numero = self.cursor.fetchall()[0][0]
        if numero == None: numero = 0
        return numero

    def get_last_lignefolio(self, journal, periode):
        """
        Retourne le dernier numero de ligne folio
        utilisé dans un journal
        """
        sSQL = ("SELECT MAX(LigneFolio) FROM Ecritures "
                "WHERE CodeJournal='{}' "
                "AND PeriodeEcriture=#{}#").format(journal, periode)
        self.cursor.execute(sSQL)
        self.cursor.execute(sSQL)
        last_lfolio = list(self.cursor)[0][0]  
        if last_lfolio is None: last_lfolio = 0
        return last_lfolio            

    def get_affectation_ana(self, compte):
        """
        Pour un compte donné, retourne le code analytique
        par défaut
        """

        sSQL = ("SELECT CodeCentre "
                "FROM AffectationAna "
                "WHERE Numcompte='{}'"
                .format(compte))
        
        self.cursor.execute(sSQL)
        self.cursor.execute(sSQL)
        data = list(self.cursor)

        if data :
            affect = data[0][0]
        else:
            affect = ""

        return affect    

    def insert_ecrit(self, compte, journal, folio, date,
                     libelle, debit, credit, piece, image, centre, image_root):
        """
        Insere une nouvelle ligne dans la table ecritures de Quadra.
        Si le compte possède une affectation analytique, une deuxème
        ligne est insérée avec les données analytiques
        """
        # Contrôle piece-jointe
        source_image_path = os.path.join(image_root, image)
        if image and not os.path.isfile(source_image_path):
            logging.error("fichier absent {}".format(os.path.join(image_root, image)))
            image = ""

        # Folio = "000"
        MontantSaisiDebit = 0
        MontantSaisiCredit = 0
        Quantite = 0
        NumLettrage = 0
        RapproBancaireOk = False
        NoLotEcritures = 0
        PieceInterne = 0
        CodeOperateur = "ABAQ"
        Etat = 0
        NumLigne = 0
        TypeLigne = "E"
        Actif = False
        PrctRepartition = 0
        ClientOuFrn = 0
        MontantAna = 0
        MilliemesAna = 0
        CodeTva = -1
        BonsAPayer = False
        MtDevForce = False
        EnLitige = False
        Quantite2 = 0
        NumEcrEco = 0
        NoLotFactor = 0
        Validee = False
        NoLotIs = 0
        NumMandat = 0
        DateSysSaisie = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        CompteContrepartie = ""   

        # Elements analytique invar.
        A_NumLigne = 1
        A_TypeLigne = "A"
        A_Nature = '*'
        A_PrctRepartition = 100
        A_TypeSaisie = 'P'  

        periode = datetime(date.year,
                           date.month,1)
        jour = date.day
        libelle = libelle[:30]
        piece = piece[0:10]

        uid = self.get_last_uniq() + 1
        lfolio = self.get_last_lignefolio(journal, periode) + 10

        # Vérif affect analytique
        if not centre:
            if compte in self.param_doss["affect"].keys():
                centre = self.param_doss["affect"][compte]

        # Vérif existence du compte
        if not compte in self.param_doss["plan"]:
            # Si insert ok on met à jour les param_doss
            if self.insert_compte(compte):
                self.param_doss["plan"].update({compte : {"initule": "", "nbrecr": 0, "lettrage": ""}})

        sql = f"""
            INSERT INTO Ecritures
            (NumUniq, NumeroCompte, 
            CodeJournal, Folio, 
            LigneFolio, PeriodeEcriture, 
            JourEcriture, Libelle, 
            MontantTenuDebit, MontantTenuCredit, 
            MontantSaisiDebit, MontantSaisiCredit, 
            CompteContrepartie, Quantite, 
            NumeroPiece, NumLettrage, 
            RapproBancaireOk, NoLotEcritures, 
            PieceInterne, CodeOperateur, 
            DateSysSaisie, Etat, 
            NumLigne, TypeLigne, 
            Actif, PrctRepartition, 
            ClientOuFrn, RefImage, 
            MontantAna, MilliemesAna, 
            CentreSimple, CodeTva, 
            BonsAPayer, MtDevForce, 
            EnLitige, Quantite2, 
            NumEcrEco, NoLotFactor, 
            Validee, NoLotIs, 
            NumMandat) 
            VALUES 
            ({uid}, '{compte}', 
            '{journal}', {folio}, 
            {lfolio}, #{periode}#, 
            {jour}, '{libelle}', 
            {debit}, {credit}, 
            {MontantSaisiDebit}, {MontantSaisiCredit}, 
            '{CompteContrepartie}', {Quantite}, 
            '{piece}', {NumLettrage}, 
            {RapproBancaireOk}, {NoLotEcritures}, 
            {PieceInterne}, '{CodeOperateur}', 
            #{DateSysSaisie}#, {Etat}, 
            {NumLigne}, '{TypeLigne}', 
            {Actif}, {PrctRepartition}, 
            {ClientOuFrn}, '{image}', 
            {MontantAna}, {MilliemesAna}, 
            '{centre}', {CodeTva}, 
            {BonsAPayer}, {MtDevForce}, 
            {EnLitige}, {Quantite2}, 
            {NumEcrEco}, {NoLotFactor}, 
            {Validee}, {NoLotIs}, 
            {NumMandat})
            """
        try:
            self.cursor.execute(sql)
        except pyodbc.Error:
            logging.error("erreur insert Ecritures {}".format(sys.exc_info()[1]))
            logging.debug(sql)
            return False

        if centre:
            montant = abs(debit - credit)
            uid += 1

            sql = f"""
                INSERT INTO Ecritures
                (NumUniq, NumeroCompte,
                CodeJournal, Folio,
                LigneFolio,PeriodeEcriture,
                JourEcriture, MontantTenuDebit,
                MontantTenuCredit, MontantSaisiDebit,
                MontantSaisiCredit, Quantite,
                NumLettrage, RapproBancaireOk,
                NoLotEcritures, PieceInterne,
                NumLigne, TypeLigne,
                Actif, Centre,
                Nature, PrctRepartition,
                TypeSaisie, MontantAna,
                MilliemesAna, CodeTva,
                BonsAPayer, MtDevForce,
                EnLitige, Quantite2,
                NumEcrEco, NoLotFactor,
                Validee, NoLotIs,
                NumMandat)
                VALUES
                ({uid}, '{compte}',
                '{journal}', {folio},
                {lfolio}, #{periode}#,
                {jour}, {debit}, 
                {credit}, {MontantSaisiDebit},
                {MontantSaisiCredit}, {Quantite},
                {NumLettrage}, {RapproBancaireOk},
                {NoLotEcritures}, {PieceInterne},
                {A_NumLigne}, '{A_TypeLigne}',
                {Actif}, '{centre}',
                '{A_Nature}', {A_PrctRepartition},
                '{A_TypeSaisie}', {montant},
                {MilliemesAna}, {CodeTva},
                {BonsAPayer}, {MtDevForce},
                {EnLitige}, {Quantite2},
                {NumEcrEco}, {NoLotFactor},
                {Validee}, {NoLotIs}, {NumMandat})
                """           
            try:
                self.cursor.execute(sql)
            except pyodbc.Error:
                logging.error("erreur insert Ecritures {} (A)".format(sys.exc_info()[1]))
                logging.debug(sql)
                return False

            # TRAITEMENT IMAGE
            dest_image_dir = self.chem_base.replace("qcompta.mdb", "images")
            if os.path.isfile(source_image_path):
                if os.path.isdir(dest_image_dir): 
                    try:  
                        shutil.copy(source_image_path, dest_image_dir+"/"+image)  
                    except IOError as e:
                        print ("Echec copie {}, {}".format(source_image_path, e))

            return True

    def maj_centralisateurs(self):
        """
        Mise à jour de toute la table des centralisateurs
        SAUF le journal des a-nouveaux
        """
        sql = """
        SELECT
            NBL.CodeJournal, NBL.PeriodeEcriture, NBL.Folio,
            NBL.NbLigneFolio, PRL.ProchaineLigne,
            CLI.DebitClient, CLI.CreditClient,
            FRN.DebitFournisseur, FRN.CreditFournisseur,
            BIL.DebitClasse15, BIL.CreditClasse15,
            EXP.DebitClasse67, EXP.CreditClasse67
        FROM
            (((((SELECT CodeJournal, PeriodeEcriture, Folio, COUNT(*) AS NbLigneFolio
            FROM Ecritures
            WHERE TypeLigne='E'
            GROUP BY CodeJournal, PeriodeEcriture, Folio) NBL
            LEFT JOIN
            (SELECT CodeJournal, PeriodeEcriture, Folio, (MAX(LigneFolio) + 10) AS ProchaineLigne
            FROM Ecritures
            WHERE TypeLigne='E'
            GROUP BY CodeJournal, PeriodeEcriture, Folio) PRL
            ON NBL.CodeJournal=PRL.CodeJournal AND
                NBL.PeriodeEcriture=PRL.PeriodeEcriture AND
                NBL.Folio=PRL.Folio)
            LEFT JOIN
            (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitClient, SUM(MontantTenuCredit) AS CreditClient
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '9%'
            GROUP BY CodeJournal, PeriodeEcriture, Folio) CLI
            ON NBL.CodeJournal=CLI.CodeJournal AND
                NBL.PeriodeEcriture=CLI.PeriodeEcriture AND
                NBL.Folio=CLI.Folio)
            LEFT JOIN
            (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitFournisseur, SUM(MontantTenuCredit) AS CreditFournisseur
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '0%'
            GROUP BY CodeJournal, PeriodeEcriture, Folio) FRN
            ON NBL.CodeJournal=FRN.CodeJournal AND
                NBL.PeriodeEcriture=FRN.PeriodeEcriture AND
                NBL.Folio=FRN.Folio)
            LEFT JOIN
            (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitClasse15, SUM(MontantTenuCredit) AS CreditClasse15
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '[1-5]%'
            GROUP BY CodeJournal, PeriodeEcriture, Folio) BIL
            ON NBL.CodeJournal=BIL.CodeJournal AND
                NBL.PeriodeEcriture=BIL.PeriodeEcriture AND
                NBL.Folio=BIL.Folio)
            LEFT JOIN
            (SELECT CodeJournal, PeriodeEcriture, Folio, SUM(MontantTenuDebit) AS DebitClasse67, SUM(MontantTenuCredit) AS CreditClasse67
            FROM Ecritures
            WHERE TypeLigne='E'
                AND NumeroCompte LIKE '[6-7]%'
            GROUP BY CodeJournal, PeriodeEcriture, Folio) EXP
            ON NBL.CodeJournal=EXP.CodeJournal AND
                NBL.PeriodeEcriture=EXP.PeriodeEcriture AND
                NBL.Folio=EXP.Folio
        WHERE NBL.CodeJournal<>'AN'
        """
        data = self.cursor.execute(sql).fetchall()

        logging.info("Mise à jour de la table Centralisateurs")
        count = 1

        for (journal, periode, folio,
             nbligne, prligne, 
             debitcli, creditcli,
             debitfrn, creditfrn, 
             debit15, credit15, 
             debit67, credit67) in data:

            if folio == None: folio = 0
            if debitcli == None: debitcli = 0
            if creditcli == None: creditcli = 0
            if debitfrn == None: debitfrn = 0
            if creditfrn == None: creditfrn = 0
            if debit15 == None: debit15 = 0
            if credit15 == None: credit15 = 0
            if debit67 == None: debit67 = 0
            if credit67 == None: credit67 = 0
            
            sql = f"""
            SELECT * FROM Centralisateur 
            WHERE CodeJournal='{journal}' 
            AND Periode=#{periode}# 
            AND Folio={folio} 
            """
            self.cursor.execute(sql)
            if list(self.cursor):
                logging.debug("update central {} {}".format(journal, periode))
                sql = f"""
                    UPDATE Centralisateur 
                    SET NbLigneFolio={nbligne}, 
                    ProchaineLigne={prligne}, 
                    DebitClient={debitcli}, 
                    CreditClient={creditcli}, 
                    DebitFournisseur={debitfrn}, 
                    CreditFournisseur={creditfrn}, 
                    DebitClasse15={debit15}, 
                    CreditClasse15={credit15}, 
                    DebitClasse67={debit67}, 
                    CreditClasse67={credit67} 
                    WHERE 
                    CodeJournal='{journal}' 
                    AND Periode=#{periode}# 
                    AND Folio={folio}
                    """
            else:
                logging.debug("insert central {} {}".format(journal, periode))
                sql = f"""
                    INSERT INTO Centralisateur 
                    (CodeJournal, Periode, Folio, 
                    NbLigneFolio, ProchaineLigne, 
                    DebitClient, CreditClient, 
                    DebitFournisseur, CreditFournisseur, 
                    DebitClasse15, CreditClasse15, 
                    DebitClasse67, CreditClasse67) 
                    VALUES 
                    ('{journal}', #{periode}#, {folio}, 
                    {nbligne}, {prligne}, 
                    {debitcli}, {creditcli}, 
                    {debitfrn}, {creditfrn}, 
                    {debit15}, {credit15}, 
                    {debit67}, {credit67})
                    """
            try:
                self.cursor.execute(sql)
                progressbar(count, len(data))
                count += 1
            except pyodbc.Error:
                logging.error("erreur insert Centralisateurs {}".format(sys.exc_info()[1]))
                logging.debug(sql)
                return False  

        logging.info("Mise à jour table Centralisateurs OK")
        return True                        

    def maj_solde_comptes(self):

        sql = f"""
        SELECT NB.NumeroCompte,
            N.debit, N.credit,
            N1.debit, N1.credit,
            NB.NbEcritures
        FROM ((
            SELECT NumeroCompte, COUNT(*) AS NbEcritures
            FROM Ecritures
            WHERE TypeLigne='E'
            GROUP BY NumeroCompte) NB
            LEFT JOIN(
            SELECT NumeroCompte, SUM(MontantTenuDebit) AS debit, SUM(MontantTenuCredit) as credit
            FROM Ecritures
            WHERE TypeLigne='E'
            GROUP BY NumeroCompte) N
            ON NB.NumeroCompte=N.Numerocompte)
            LEFT JOIN (
            SELECT NumeroCompte, SUM(MontantTenuDebit) AS debit, SUM(MontantTenuCredit) as credit
            FROM Ecritures
            WHERE PeriodeEcriture>=#{self.param_doss["exefin"]}# 
            AND TypeLigne='E'
            GROUP BY NumeroCompte) N1
            ON NB.NumeroCompte=N1.NumeroCompte
            """
        data = self.cursor.execute(sql).fetchall()
        count = 1
        logging.info("Mise à jour de la table Comptes")

        for (Numero, Debit, Credit,
             DebitHorsEx, CreditHorsEx, NbEcritures) in data:

            if Debit==None: Debit = 0
            if Credit==None: Credit = 0
            if DebitHorsEx==None: DebitHorsEx = 0
            if CreditHorsEx==None: CreditHorsEx = 0
            if NbEcritures==None: NbEcritures = 0

            sql = f"""
            UPDATE Comptes 
            SET 
            Debit={Debit},
            Credit={Credit},
            DebitHorsEx={DebitHorsEx},
            CreditHorsEx={CreditHorsEx},
            NbEcritures={NbEcritures}
            WHERE Numero='{Numero}'
            """
            try:
                self.cursor.execute(sql)
                progressbar(count, len(data))
                count += 1
            except pyodbc.Error:
                logging.error("erreur update Comptes {}".format(sys.exc_info()[1]))
                logging.debug(sql)
                return False  

        logging.info("Mise à jour table Comptes OK")
        return True    

if __name__ == '__main__':

    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    logging.basicConfig(level=logging.DEBUG,
                        format='%(funcName)s\t\t%(levelname)s - %(message)s')

    # cpta = "//srvquadra/qappli/quadra/database/cpta/da209912/Form05/qcompta.mdb"
    # cpta = "//srvquadra/qappli/quadra/database/cpta/ds2099/000175/qcompta.mdb"
    cpta = "C:/quadra/database/cpta/DC/000752/qcompta.mdb"

    o = QueryCompta()
    data = o.load_params(cpta)
    pp.pprint(data["images"])

    # compte = "60110000"

    # print(o.get_affectation_ana(compte))
    # print(o.get_last_uniq())
    # o.insert_compte("42501300")
    # pp.pprint(o.maj_solde_comptes())
    # print(o.get_last_lignefolio("AC", periode))
    # periode = datetime(year=2018, month=12, day=1)
    # o.insert_ecrit("0YYYYYY7", "CA", "000", periode, "ABAQUE", 999.99, 0, "XXXXX", "", "")
    # o.insert_ecrit("60110000", "CA", "000", periode, "ABAQUE", 0, 999.99, "XXXXX", "", "")
    # o.maj_central_all()
    # o.maj_central("AC", periode, 0)

    o.close()
