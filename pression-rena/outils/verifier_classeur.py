#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Verifie un classeur « Pression Rena » RECALCULE, contre les donnees d'origine.

    soffice --headless --convert-to xlsx --outdir recalc Pression_Rena.xlsx
    python3 verifier_classeur.py <source.xlsx> recalc/Pression_Rena.xlsx

Le classeur produit par construire_classeur.py ne contient que des formules,
sans valeur en cache : il faut le faire recalculer par un tableur avant de le
verifier. Ce script relit ensuite les valeurs obtenues et les compare a ce que
Python calcule de son cote, a partir du fichier source.

Il controle :
  1. que les 3613 mesures sont toutes arrivees, a la bonne date et au bon rang ;
  2. que l'invariant deux-lignes-par-jour tient sur toute la feuille Datas ;
  3. que les extremes du jour de la feuille Calcul valent bien les extremes ;
  4. que les moyennes lissees valent la moyenne des N derniers jours mesures ;
  5. que les quatre series empilees redonnent les quatre courbes ;
  6. que le resume de la feuille Saisie correspond a la fenetre choisie ;
  7. qu'aucune cellule ne porte une valeur d'erreur autre que #N/A ;
  8. que les etiquettes de date des deux graphiques sont lisibles.
"""

import datetime
import sys

import openpyxl

LIGNE_DEB = 3
TOLERANCE = 5e-6


# #N/A est VOULU : c'est ainsi qu'un jour sans mesure, ou une moyenne lissee
# qui n'a pas encore assez d'histoire derriere elle, laisse un trou dans la
# courbe au lieu d'une valeur inventee. Toute AUTRE erreur est un defaut.
ERREURS_ADMISES = ("#N/A",)
ERREURS_EXCEL = ("#N/A", "#NAME?", "#NOM?", "#VALUE!", "#VALEUR!", "#REF!", "#DIV/0!",
                 "#NUM!", "#NOMBRE!", "#NULL!", "#NUL!", "#GETTING_DATA", "Err:")


def valeur(cellule):
    """Lit une cellule : chaine vide et #N/A valent « pas de valeur »."""
    v = cellule.value
    if v == "" or v in ERREURS_ADMISES:
        return None
    return v


def erreurs_de_la_feuille(ws, limite=12):
    """Toute valeur d'erreur autre que #N/A, avec son adresse."""
    trouvees = []
    for rangee in ws.iter_rows():
        for cellule in rangee:
            v = cellule.value
            if isinstance(v, str) and v not in ERREURS_ADMISES:
                if any(v.startswith(e) for e in ERREURS_EXCEL):
                    trouvees.append("%s!%s = %s" % (ws.title, cellule.coordinate, v))
                    if len(trouvees) >= limite:
                        return trouvees
    return trouvees


class Bilan(object):
    def __init__(self):
        self.erreurs = []
        self.controles = 0

    def verifier(self, condition, message):
        self.controles += 1
        if not condition:
            self.erreurs.append(message)

    def proche(self, obtenu, attendu, message):
        self.controles += 1
        if obtenu is None and attendu is None:
            return
        if obtenu is None or attendu is None:
            self.erreurs.append("%s : obtenu %r, attendu %r" % (message, obtenu, attendu))
            return
        if abs(float(obtenu) - float(attendu)) > TOLERANCE:
            self.erreurs.append("%s : obtenu %r, attendu %r" % (message, obtenu, attendu))


def jour(valeur):
    return valeur.date() if isinstance(valeur, datetime.datetime) else valeur


def lire_mesures_source(chemin):
    """Les mesures du fichier d'origine, dans l'ordre : (date, rang, sys, dia, pouls)."""
    ws = openpyxl.load_workbook(chemin, data_only=True)["Datas"]
    mesures = []
    ligne = LIGNE_DEB
    while ligne + 1 <= ws.max_row:
        date = ws.cell(ligne, 1).value
        if date is None:
            break
        for rang in (0, 1):
            sys_ = ws.cell(ligne + rang, 5).value
            if sys_ is not None:
                mesures.append(
                    (
                        date.date(),
                        rang,
                        sys_,
                        ws.cell(ligne + rang, 6).value,
                        ws.cell(ligne + rang, 7).value,
                    )
                )
        ligne += 2
    return mesures


def main(argv):
    if len(argv) != 3:
        raise SystemExit(__doc__)
    source, produit = argv[1], argv[2]

    attendues = lire_mesures_source(source)
    wb = openpyxl.load_workbook(produit, data_only=True)
    datas, calcul, saisie, reglages = wb["Datas"], wb["Calcul"], wb["Saisie"], wb["Reglages"]
    b = Bilan()

    # --- 1 et 2 : les mesures et l'invariant ---------------------------------
    premier = jour(datas.cell(LIGNE_DEB, 1).value)
    ligne, reprises, ecarts_gabarit = LIGNE_DEB, [], 0
    while True:
        date = datas.cell(ligne, 1).value
        if date is None:
            break
        attendu = premier + datetime.timedelta(days=(ligne - LIGNE_DEB) // 2)
        if jour(date) != attendu:
            ecarts_gabarit += 1
        sys_ = datas.cell(ligne, 5).value
        if sys_ is not None:
            reprises.append(
                (
                    jour(date),
                    (ligne - LIGNE_DEB) % 2,
                    sys_,
                    datas.cell(ligne, 6).value,
                    datas.cell(ligne, 7).value,
                )
            )
        ligne += 1
    derniere_ligne_datas = ligne - 1

    b.verifier(ecarts_gabarit == 0,
               "invariant rompu : %d lignes ne tombent pas sur leur jour" % ecarts_gabarit)
    b.verifier((derniere_ligne_datas - LIGNE_DEB + 1) % 2 == 0,
               "la feuille Datas ne porte pas un nombre pair de lignes")
    b.verifier(len(reprises) == len(attendues),
               "mesures reprises : %d, attendues : %d" % (len(reprises), len(attendues)))
    manquantes = [m for m in attendues if m not in set(reprises)]
    b.verifier(not manquantes,
               "%d mesures perdues, par exemple %s" % (len(manquantes), manquantes[:3]))

    # --- 3 : les extremes du jour --------------------------------------------
    par_jour = {}
    for date, _, sys_, dia, pouls in attendues:
        par_jour.setdefault(date, []).append((sys_, dia, pouls))

    jours_calcul, ligne = [], LIGNE_DEB
    while calcul.cell(ligne, 1).value is not None:
        jours_calcul.append(jour(calcul.cell(ligne, 1).value))
        ligne += 1
    derniere_ligne_calcul = ligne - 1

    b.verifier(
        len(jours_calcul) * 2 == derniere_ligne_datas - LIGNE_DEB + 1,
        "Calcul porte %d jours pour %d lignes de Datas"
        % (len(jours_calcul), derniere_ligne_datas - LIGNE_DEB + 1),
    )

    for i, date in enumerate(jours_calcul):
        lg = LIGNE_DEB + i
        mesures = par_jour.get(date, [])
        nb = calcul.cell(lg, 3).value
        b.proche(nb, len(mesures), "Calcul!C%d, nombre de mesures du %s" % (lg, date))
        for col, extrait in ((4, max), (5, min)):
            obtenu = valeur(calcul.cell(lg, col))
            attendu = extrait(m[0] for m in mesures) if mesures else None
            b.proche(obtenu, attendu, "Calcul!%s%d, systolique du %s"
                     % ("DE"[col - 4], lg, date))
        for col, extrait in ((6, max), (7, min)):
            obtenu = valeur(calcul.cell(lg, col))
            attendu = extrait(m[1] for m in mesures) if mesures else None
            b.proche(obtenu, attendu, "Calcul!%s%d, diastolique du %s"
                     % ("FG"[col - 6], lg, date))

    # --- 4 : les moyennes lissees --------------------------------------------
    nb_jours = int(saisie.cell(6, 2).value)
    brutes = {
        10: [valeur(calcul.cell(LIGNE_DEB + i, 4)) for i in range(len(jours_calcul))],
        11: [valeur(calcul.cell(LIGNE_DEB + i, 5)) for i in range(len(jours_calcul))],
        12: [valeur(calcul.cell(LIGNE_DEB + i, 6)) for i in range(len(jours_calcul))],
        13: [valeur(calcul.cell(LIGNE_DEB + i, 7)) for i in range(len(jours_calcul))],
    }
    lisses = {}
    for col, serie in brutes.items():
        colonne = []
        for i in range(len(serie)):
            lg = LIGNE_DEB + i
            obtenu = valeur(calcul.cell(lg, col))
            # La fenetre couvre les nb_jours lignes qui finissent a celle-ci.
            # Tant qu'elle ne tient pas entiere DANS LES DONNEES — et non dans
            # la feuille — la moyenne n'existe pas : une moyenne sur sept jours
            # calculee sur trois n'en serait pas une.
            if lg < LIGNE_DEB + nb_jours - 1:
                attendu = None
            else:
                fenetre = [v for v in serie[max(0, i - nb_jours + 1):i + 1] if v is not None]
                attendu = sum(fenetre) / len(fenetre) if fenetre else None
            b.proche(obtenu, attendu, "Calcul!%s%d, moyenne lissée" % ("JKLM"[col - 10], lg))
            colonne.append(obtenu)
        lisses[col] = colonne

    # --- 5 : les quatre series empilees --------------------------------------
    # Empilees de bas en haut, elles doivent redonner exactement les quatre
    # courbes : c'est ce qui fait tomber les bandes coloriees au bon endroit.
    for i in range(len(jours_calcul)):
        lg = LIGNE_DEB + i
        z = [valeur(calcul.cell(lg, c)) for c in (14, 15, 16, 17)]
        sys_haut, sys_bas = lisses[10][i], lisses[11][i]
        dia_haut, dia_bas = lisses[12][i], lisses[13][i]
        if None in (sys_haut, sys_bas, dia_haut, dia_bas):
            continue
        b.proche(z[0], dia_bas, "Calcul!N%d, bas de la bande bleue" % lg)
        b.proche(z[0] + z[1], dia_haut, "Calcul!O%d, haut de la bande bleue" % lg)
        b.proche(z[0] + z[1] + z[2], sys_bas, "Calcul!P%d, bas de la bande verte" % lg)
        b.proche(sum(z), sys_haut, "Calcul!Q%d, haut de la bande verte" % lg)

    # --- 6 : le resume de la feuille Saisie -----------------------------------
    deb, fin = jour(reglages.cell(10, 2).value), jour(reglages.cell(11, 2).value)
    fenetre = [m for m in attendues if deb <= m[0] <= fin]
    b.proche(valeur(saisie.cell(14, 2)), len(fenetre), "Saisie!B14, nombre de mesures")
    if fenetre:
        b.proche(valeur(saisie.cell(15, 2)),
                 sum(m[2] for m in fenetre) / len(fenetre), "Saisie!B15, systolique moyenne")
        b.proche(valeur(saisie.cell(16, 2)),
                 sum(m[3] for m in fenetre) / len(fenetre), "Saisie!B16, diastolique moyenne")
        b.proche(valeur(saisie.cell(15, 5)),
                 max(m[2] for m in fenetre), "Saisie!E15, systolique maximale")
        b.proche(valeur(saisie.cell(16, 5)),
                 max(m[3] for m in fenetre), "Saisie!E16, diastolique maximale")
    jours_mesures = len({m[0] for m in fenetre})
    b.proche(valeur(saisie.cell(14, 5)), jours_mesures, "Saisie!E14, jours mesurés")

    # les quinze dernieres mesures, de la plus recente a la plus ancienne
    recentes = sorted(reprises, key=lambda m: (m[0], m[1]), reverse=True)[:15]
    for i, attendu in enumerate(recentes):
        lg = 22 + i
        b.verifier(jour(saisie.cell(lg, 1).value) == attendu[0],
                   "Saisie!A%d, date de la %d-ième dernière mesure : %s au lieu de %s"
                   % (lg, i + 1, jour(saisie.cell(lg, 1).value), attendu[0]))
        b.proche(valeur(saisie.cell(lg, 4)), attendu[2], "Saisie!D%d, systole" % lg)

    # --- 8 : les etiquettes de date -------------------------------------------
    # Elles ne doivent contenir aucun code de format reste en clair : c'est ce
    # que produit TEXTE(...;"jjj jj.mm.aa") hors d'un Excel francais.
    jours_courts = ("lun", "mar", "mer", "jeu", "ven", "sam", "dim")
    for feuille, colonne, ligne_essai in ((datas, 4, LIGNE_DEB), (calcul, 2, LIGNE_DEB)):
        etiquette = str(feuille.cell(ligne_essai, colonne).value or "")
        b.verifier(
            etiquette[:3] in jours_courts and "jj" not in etiquette and "aa" not in etiquette,
            "étiquette de date illisible dans %s : %r" % (feuille.title, etiquette),
        )

    b.verifier(str(reglages.cell(18, 2).value).startswith("Structure correcte"),
               "contrôle de structure : %r" % reglages.cell(18, 2).value)

    # --- 7 : aucune valeur d'erreur inattendue dans tout le classeur ---------
    for ws in (saisie, datas, calcul, reglages):
        fautes = erreurs_de_la_feuille(ws)
        b.verifier(not fautes, "valeurs d'erreur dans %s : %s" % (ws.title, fautes))

    print("Contrôles       : %d" % b.controles)
    print("Mesures source  : %d" % len(attendues))
    print("Mesures reprises: %d" % len(reprises))
    print("Jours de Calcul : %d" % len(jours_calcul))
    print("Fenêtre effective: %s -> %s  (%d mesures, lissage %d jours)"
          % (deb, fin, len(fenetre), nb_jours))
    if b.erreurs:
        print("\nÉCHEC : %d anomalies" % len(b.erreurs))
        for message in b.erreurs[:25]:
            print("  - %s" % message)
        if len(b.erreurs) > 25:
            print("  ... et %d autres" % (len(b.erreurs) - 25))
        return 1
    print("\nOK : aucune anomalie.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
