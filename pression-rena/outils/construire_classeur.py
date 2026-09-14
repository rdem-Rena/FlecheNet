#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Construit le classeur « Pression Rena » a partir du fichier de releves existant.

    python3 construire_classeur.py <source.xlsx> <destination.xlsx>

CE QUE LE SCRIPT PRODUIT
------------------------
Un classeur de six feuilles :

    Saisie          les reglages, le resume, les dernieres mesures
    Datas           les releves, deux lignes par jour (matin, soir)
    Calcul          un jour par ligne : extremes du jour et moyennes lissees
    Reglages        les cellules derivees, lues par les plages nommees
    Graph_Mesures   le graphique des mesures brutes (feuille graphique)
    Graph_Lissage   le graphique des moyennes lissees (feuille graphique)

L'INVARIANT DU CLASSEUR
-----------------------
« Datas » porte EXACTEMENT deux lignes par jour calendaire, sans trou et sans
doublon, de la premiere a la derniere date du gabarit. Ligne 3 + 2k et ligne
4 + 2k sont les deux demi-journees du k-ieme jour.

C'est cet invariant qui permet aux plages nommees de se calculer par une
simple soustraction de dates, sans MATCH ni RECHERCHE : le decalage d'un jour
vaut deux lignes dans « Datas » et une ligne dans « Calcul ». Il rend aussi le
graphique honnete — une periode non mesuree apparait comme un trou, alors que
le fichier d'origine la faisait disparaitre en rapprochant les deux dates qui
l'encadrent.

Le fichier d'origine respectait deja cet invariant partout sauf entre le
31.12.2020 et le 01.03.2021, ou 59 jours manquaient : ils sont retablis, vides.
La feuille « Saisie » porte un controle qui verifie l'invariant a chaque
ouverture.
"""

import datetime
import sys

import openpyxl
from openpyxl.chart import AreaChart, BarChart, LineChart
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.legend import Legend, LegendEntry
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Reperes du classeur. Tout le reste s'en deduit.
# ---------------------------------------------------------------------------
LIGNE_ENTETE = 2                         # les titres de colonnes, dans les 3 feuilles
LIGNE_DEB = 3                            # la premiere ligne de donnees
FIN_GABARIT = datetime.date(2028, 1, 5)  # horizon pre-rempli, repris du fichier source

# Le titre de la feuille Datas est REPRIS DU FICHIER SOURCE, et non ecrit ici :
# ce depot est public, et le nom de l'auteur n'a pas a y figurer.
TITRE_DEFAUT = "Suivi de la pression artérielle"

# Colonnes de « Datas ». L'ordre des quatorze premieres est celui du fichier
# d'origine : les reperes visuels de l'auteur sont conserves. « Poids » est
# ajoute en quinzieme, il n'existait pas.
COLS_DATAS = [
    ("Date", 11, "dd.mm.yyyy"),
    ("Heure", 7, "hh:mm"),
    ("Activité", 16, None),
    ("Date-Compo", 20, None),
    ("Systole", 8, "0"),
    ("Diastole", 9, "0"),
    ("Pouls", 7, "0"),
    ("Basse B.", 8, "0"),
    ("Basse H.", 8, "0"),
    ("Haute B.", 8, "0"),
    ("Haute H.", 8, "0"),
    ("Demi", 7, None),
    ("Change", 8, "0"),
    ("Commentaire", 40, None),
    ("Poids", 8, "0.0"),
]
C_DATE, C_HEURE, C_ACT, C_COMPO, C_SYS, C_DIA, C_POULS = 1, 2, 3, 4, 5, 6, 7
C_BB, C_BH, C_HB, C_HH, C_DEMI, C_CHANGE, C_COM, C_POIDS = 8, 9, 10, 11, 12, 13, 14, 15

# Colonnes de « Calcul ». Un jour par ligne.
COLS_CALCUL = [
    ("Jour", 11, "dd.mm.yyyy"),
    ("Libellé", 15, None),
    ("Nb", 5, "0"),
    ("Sys haut", 9, "0"),
    ("Sys bas", 9, "0"),
    ("Dia haut", 9, "0"),
    ("Dia bas", 9, "0"),
    ("Pouls moy", 10, "0"),
    ("Poids", 8, "0.0"),
    ("Sys haut lissé", 13, "0.0"),
    ("Sys bas lissé", 13, "0.0"),
    ("Dia haut lissé", 13, "0.0"),
    ("Dia bas lissé", 13, "0.0"),
    ("Z dia bas", 10, "0.0"),
    ("Z dia", 10, "0.0"),
    ("Z écart", 10, "0.0"),
    ("Z sys", 10, "0.0"),
    ("Seuil dia bas", 12, "0"),
    ("Seuil dia haut", 12, "0"),
    ("Seuil sys bas", 12, "0"),
    ("Seuil sys haut", 12, "0"),
]
K_JOUR, K_LIB, K_NB = 1, 2, 3
K_SYSH, K_SYSB, K_DIAH, K_DIAB, K_POULS, K_POIDS = 4, 5, 6, 7, 8, 9
K_LSYSH, K_LSYSB, K_LDIAH, K_LDIAB = 10, 11, 12, 13
K_ZDIAB, K_ZDIA, K_ZECART, K_ZSYS = 14, 15, 16, 17
K_SDB, K_SDH, K_SSB, K_SSH = 18, 19, 20, 21

# --- Palette -----------------------------------------------------------------
# Les quatre premieres sont celles du graphique d'origine : le graphique des
# mesures doit rester reconnaissable au premier coup d'oeil.
COUL_SYSTOLE = "0070C0"      # bleu
COUL_DIASTOLE = "00B050"     # vert
COUL_SEUIL_BAS = "FFCC00"    # jaune, les deux lignes 80 et 90
COUL_SEUIL_HAUT = "FF0000"   # rouge, les deux lignes 120 et 130
COUL_CHANGE = "BFBFBF"       # gris : les barres de changement de situation

# Le graphique lisse suit la demande : bande verte pour la systolique, bande
# bleue pour la diastolique. C'est l'inverse des couleurs de courbes du premier
# graphique, ou la systolique est bleue — voir docs/GRAPHIQUES.md.
COUL_BANDE_SYS = "D5E8C8"    # vert pale
COUL_BANDE_DIA = "CFE2F3"    # bleu pale
COUL_TRAIT_SYS = "548235"    # vert fonce
COUL_TRAIT_DIA = "2E75B6"    # bleu fonce

# --- Habillage des feuilles ---------------------------------------------------
COUL_BANDEAU = "1B365D"
COUL_TITRE_TXT = "FFFFFF"
COUL_ENTETE = "E9EEF6"
COUL_ENTETE_TXT = "44566C"
COUL_SAISIE = "FCFDFF"
COUL_BORD = "D6DEE9"
COUL_SECTION = "7A8CA6"

FONT_TITRE = Font(name="Segoe UI", size=13, bold=True, color=COUL_TITRE_TXT)
FONT_SECTION = Font(name="Segoe UI", size=9, bold=True, color=COUL_SECTION)
FONT_ENTETE = Font(name="Segoe UI", size=9, bold=True, color=COUL_ENTETE_TXT)
FONT_NORMAL = Font(name="Segoe UI", size=10)
FONT_VALEUR = Font(name="Segoe UI", size=12, bold=True, color="202D3E")
FONT_NOTE = Font(name="Segoe UI", size=8, italic=True, color=COUL_SECTION)

BORD_CHAMP = Border(*(Side(style="thin", color=COUL_BORD),) * 4)


# ---------------------------------------------------------------------------
# L'ETIQUETTE DE DATE DES DEUX GRAPHIQUES
# ---------------------------------------------------------------------------
# Le fichier d'origine ecrivait TEXT(A3;"jjj jj.mm.aa"). Les codes de format
# passes a TEXT ne sont PAS traduits a l'ouverture : ils sont stockes tels
# qu'ils ont ete tapes. « jjj » et « aa » ne veulent donc rien dire ailleurs
# que dans un Excel francais, et le meme fichier ouvert dans un Excel anglais
# — ou dans LibreOffice — affiche « jjj jj.08.aa » en clair sur tout l'axe.
#
# CHOISIR + JOURSEM donne le jour abrege sans code de format, et TEXTE(n;"00")
# ne contient que des zeros et un point : ni l'un ni l'autre ne change d'une
# langue a l'autre. Le resultat est identique a celui du fichier d'origine —
# « lun 31.08.20 » — mais il l'est partout.
def etiquette_date(cellule):
    return (
        'CHOOSE(WEEKDAY({0},2),"lun","mar","mer","jeu","ven","sam","dim")'
        '&" "&TEXT(DAY({0}),"00")&"."&TEXT(MONTH({0}),"00")'
        '&"."&TEXT(YEAR({0})-2000,"00")'
    ).format(cellule)


# ===========================================================================
# LECTURE DE LA SOURCE
# ===========================================================================
def lire_source(chemin):
    """Releve les mesures du fichier d'origine, par paires de lignes.

    Rend un dictionnaire date -> [demi-journee 1, demi-journee 2], chaque
    demi-journee etant un dictionnaire des valeurs saisies. Les deux lignes
    d'une paire gardent leur ordre : c'est lui, et non l'heure, qui dit
    laquelle est la premiere — huit jours du fichier portent deux mesures du
    matin, et les rejouer par l'heure en perdrait une.
    """
    wb = openpyxl.load_workbook(chemin, data_only=True)
    ws = wb["Datas"]
    titre = ws.cell(1, 1).value or TITRE_DEFAUT
    jours = {}
    ligne = LIGNE_DEB
    while ligne + 1 <= ws.max_row:
        date1 = ws.cell(ligne, C_DATE).value
        if date1 is None:
            break
        date2 = ws.cell(ligne + 1, C_DATE).value
        if date2 is None or date2.date() != date1.date():
            raise SystemExit(
                "Ligne %d : la paire de lignes ne porte pas la meme date (%s / %s). "
                "Le gabarit deux-lignes-par-jour n'est pas respecte dans la source."
                % (ligne, date1, date2)
            )
        demi = []
        for lg in (ligne, ligne + 1):
            demi.append(
                {
                    "heure": ws.cell(lg, C_HEURE).value,
                    "activite": ws.cell(lg, C_ACT).value,
                    "sys": ws.cell(lg, C_SYS).value,
                    "dia": ws.cell(lg, C_DIA).value,
                    "pouls": ws.cell(lg, C_POULS).value,
                    "change": ws.cell(lg, C_CHANGE).value,
                    "commentaire": ws.cell(lg, C_COM).value,
                }
            )
        jours[date1.date()] = demi
        ligne += 2
    return jours, str(titre)


# ===========================================================================
# HABILLAGE
# ===========================================================================
def bandeau(ws, derniere_colonne, texte):
    """Pose le bandeau de titre sur la premiere ligne."""
    ws.cell(1, 1, texte).font = FONT_TITRE
    for col in range(1, derniere_colonne + 1):
        ws.cell(1, col).fill = PatternFill("solid", fgColor=COUL_BANDEAU)
    ws.row_dimensions[1].height = 26


def entetes(ws, colonnes):
    """Ecrit la ligne d'en-tete et regle la largeur des colonnes."""
    for i, (titre, largeur, _) in enumerate(colonnes, start=1):
        cel = ws.cell(LIGNE_ENTETE, i, titre)
        cel.font = FONT_ENTETE
        cel.fill = PatternFill("solid", fgColor=COUL_ENTETE)
        cel.alignment = Alignment(horizontal="center", vertical="center")
        cel.border = BORD_CHAMP
        ws.column_dimensions[get_column_letter(i)].width = largeur
    ws.row_dimensions[LIGNE_ENTETE].height = 20


def champ_saisie(ws, ligne, colonne, valeur, fmt=None):
    """Une cellule que l'utilisateur est invite a modifier."""
    cel = ws.cell(ligne, colonne, valeur)
    cel.font = FONT_VALEUR
    cel.fill = PatternFill("solid", fgColor=COUL_SAISIE)
    cel.border = BORD_CHAMP
    cel.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        cel.number_format = fmt
    return cel


def libelle(ws, ligne, colonne, texte):
    cel = ws.cell(ligne, colonne, texte)
    cel.font = FONT_NORMAL
    cel.alignment = Alignment(horizontal="right", vertical="center")
    return cel


def section(ws, ligne, texte):
    cel = ws.cell(ligne, 1, texte.upper())
    cel.font = FONT_SECTION
    return cel


# ===========================================================================
# FEUILLE « Datas »
# ===========================================================================
def ecrire_datas(ws, jours, premier, dernier, titre):
    """Le calendrier complet, deux lignes par jour, mesures reportees."""
    bandeau(ws, len(COLS_DATAS), titre)
    entetes(ws, COLS_DATAS)
    ws.freeze_panes = "A3"

    formats = {i: f for i, (_, _, f) in enumerate(COLS_DATAS, start=1) if f}
    ligne = LIGNE_DEB
    jour = premier
    while jour <= dernier:
        demi = jours.get(jour)
        for rang in (0, 1):
            src = demi[rang] if demi else None
            ws.cell(ligne, C_DATE, jour).number_format = formats[C_DATE]
            if src:
                if src["heure"] is not None:
                    ws.cell(ligne, C_HEURE, src["heure"]).number_format = formats[C_HEURE]
                for col, cle in (
                    (C_ACT, "activite"),
                    (C_SYS, "sys"),
                    (C_DIA, "dia"),
                    (C_POULS, "pouls"),
                    (C_CHANGE, "change"),
                    (C_COM, "commentaire"),
                ):
                    if src[cle] is not None:
                        cel = ws.cell(ligne, col, src[cle])
                        if col in formats:
                            cel.number_format = formats[col]

            # « Demi » nomme la demi-journee d'apres l'heure. Le fichier
            # d'origine rendait « Erreur » pour une ligne sans heure, ce qui
            # remplissait la colonne de fautes sur toute la partie a venir du
            # calendrier ; une ligne vide rend desormais une chaine vide.
            ws.cell(ligne, C_DEMI, '=IF(B{0}="","",IF(B{0}<0.5,"matin","Soir"))'.format(ligne))

            # L'etiquette de l'axe des abscisses du premier graphique.
            ws.cell(
                ligne,
                C_COMPO,
                '=IF(A{0}="","",{1}&" "&L{0}&" - "&C{0})'.format(
                    ligne, etiquette_date("A{0}".format(ligne))
                ),
            )

            # Les quatre lignes de reference du graphique. Elles renvoient aux
            # cellules de reglage : changer un seuil sur « Saisie » deplace la
            # ligne sur les deux graphiques, sans toucher a cette colonne.
            ws.cell(ligne, C_BB, "=Seuil_Dia_Bas")
            ws.cell(ligne, C_BH, "=Seuil_Dia_Haut")
            ws.cell(ligne, C_HB, "=Seuil_Sys_Bas")
            ws.cell(ligne, C_HH, "=Seuil_Sys_Haut")
            for col in (C_BB, C_BH, C_HB, C_HH):
                ws.cell(ligne, col).number_format = "0"

            ligne += 1
        jour += datetime.timedelta(days=1)
    return ligne - 1


# ===========================================================================
# FEUILLE « Calcul »
# ===========================================================================
def ecrire_calcul(ws, premier, dernier, dern_ligne_datas):
    """Un jour par ligne : les extremes du jour, puis les moyennes lissees.

    AUCUNE FONCTION POSTERIEURE A EXCEL 2007 N'EST EMPLOYEE ICI. MAXIFS, MINIFS
    et AGGREGATE feraient l'affaire et se liraient mieux, mais un .xlsx ne les
    accepte que prefixees « _xlfn. » : ecrites en clair, elles rendent #NOM?
    dans Excel comme ailleurs. Le classeur etant produit par ce script et non
    saisi dans Excel, la faute passerait inapercue jusqu'a l'ouverture.

    Leur remplacement est d'ailleurs PLUS RAPIDE : l'invariant deux-lignes-par-
    jour donne les deux lignes du jour par le calcul, la ou MAXIFS relisait les
    5368 lignes de la feuille pour chacun des 2684 jours.
    """
    bandeau(ws, len(COLS_CALCUL), "Agrégation par jour et moyennes lissées")
    entetes(ws, COLS_CALCUL)
    ws.freeze_panes = "C3"

    def tranche(col):
        """Les deux lignes de Datas qui portent le jour de la ligne courante.

        Ecrite avec LIGNE() plutot qu'avec les numeros de ligne en dur : la
        formule est alors la meme partout, et une recopie vers le bas reste
        juste. INDEX(...):INDEX(...) construit la plage sans passer par
        DECALER, qui est volatile et ferait tout recalculer a chaque frappe.
        """
        plage = "Datas!${0}${1}:${0}${2}".format(
            get_column_letter(col), LIGNE_DEB, dern_ligne_datas
        )
        return "INDEX({0},(ROW()-{1})*2+1):INDEX({0},(ROW()-{1})*2+2)".format(plage, LIGNE_DEB)

    t_sys, t_dia = tranche(C_SYS), tranche(C_DIA)
    t_pouls, t_poids = tranche(C_POULS), tranche(C_POIDS)

    formats = {i: f for i, (_, _, f) in enumerate(COLS_CALCUL, start=1) if f}
    ligne = LIGNE_DEB
    jour = premier
    while jour <= dernier:
        ws.cell(ligne, K_JOUR, jour).number_format = formats[K_JOUR]
        ws.cell(ligne, K_LIB, "=" + etiquette_date("A{0}".format(ligne)))

        # Le nombre de mesures du jour, calcule UNE fois : les quatre extremes
        # s'y referent au lieu de refaire chacun son propre comptage.
        ws.cell(ligne, K_NB, "=COUNT({0})".format(t_sys))

        # Un jour sans mesure rend une chaine VIDE, et non #N/A : ces quatre
        # colonnes ne sont pas tracees, elles alimentent les moyennes lissees,
        # et SOMME propagerait l'erreur sur toute la courbe alors qu'elle
        # ignore le texte.
        for col, fonction, source in (
            (K_SYSH, "MAX", t_sys),
            (K_SYSB, "MIN", t_sys),
            (K_DIAH, "MAX", t_dia),
            (K_DIAB, "MIN", t_dia),
        ):
            ws.cell(
                ligne,
                col,
                '=IF($C{0}=0,"",{1}({2}))'.format(ligne, fonction, source),
            )
        for col, source in ((K_POULS, t_pouls), (K_POIDS, t_poids)):
            ws.cell(
                ligne,
                col,
                '=IF(COUNT({0})=0,"",AVERAGE({0}))'.format(source),
            )

        # LA MOYENNE LISSEE, sur les Nb_Jours jours qui precedent, celui-ci
        # compris. SOMME et NB ignorent l'un comme l'autre les cellules de
        # texte : une fenetre qui contient des jours sans mesure rend donc la
        # moyenne des autres, et non une erreur.
        #
        # Les Nb_Jours - 1 premieres lignes de la feuille n'ont pas assez
        # d'histoire derriere elles : elles rendent #N/A, que le graphique
        # montre comme un trou. Une « moyenne sur sept jours » calculee sur
        # trois n'en serait pas une.
        for col, source in (
            (K_LSYSH, K_SYSH),
            (K_LSYSB, K_SYSB),
            (K_LDIAH, K_DIAH),
            (K_LDIAB, K_DIAB),
        ):
            fenetre = "OFFSET({0}{1},1-Nb_Jours,0,Nb_Jours,1)".format(
                get_column_letter(source), ligne
            )
            ws.cell(
                ligne,
                col,
                "=IFERROR(IF(ROW()<{0}+Nb_Jours-1,NA(),IF(COUNT({1})=0,NA(),"
                "SUM({1})/COUNT({1}))),NA())".format(LIGNE_DEB, fenetre),
            )

        # LES QUATRE SERIES DE L'AIRE EMPILEE. Excel ne sait pas colorer
        # l'espace entre deux courbes. On empile donc quatre aires — invisible,
        # bleue, invisible, verte — dont les hauteurs cumulees redonnent les
        # quatre courbes :
        #
        #   Z dia bas  la diastolique basse, sous la bande bleue : transparente
        #   Z dia      l'epaisseur de la bande bleue
        #   Z ecart    ce qui separe la bande bleue de la verte : transparent
        #   Z sys      l'epaisseur de la bande verte
        #
        # MAX(0; ...) protege le cas ou les deux bandes se toucheraient : une
        # hauteur negative decalerait tout ce qui est empile au-dessus.
        ws.cell(ligne, K_ZDIAB, "=IFERROR(M{0},NA())".format(ligne))
        ws.cell(ligne, K_ZDIA, "=IFERROR(MAX(0,L{0}-M{0}),NA())".format(ligne))
        ws.cell(ligne, K_ZECART, "=IFERROR(MAX(0,K{0}-L{0}),NA())".format(ligne))
        ws.cell(ligne, K_ZSYS, "=IFERROR(MAX(0,J{0}-K{0}),NA())".format(ligne))

        for col, nom in (
            (K_SDB, "Seuil_Dia_Bas"),
            (K_SDH, "Seuil_Dia_Haut"),
            (K_SSB, "Seuil_Sys_Bas"),
            (K_SSH, "Seuil_Sys_Haut"),
        ):
            ws.cell(ligne, col, "=" + nom)

        for col, fmt in formats.items():
            if col != K_JOUR:
                ws.cell(ligne, col).number_format = fmt

        ligne += 1
        jour += datetime.timedelta(days=1)
    return ligne - 1



# ===========================================================================
# FEUILLE « Saisie »
# ===========================================================================
# Les trois reglages que l'auteur manipule — debut, fin, lissage — sont ici, en
# haut, et non sur la feuille technique : ce sont eux qui commandent les deux
# graphiques, et ils doivent se trouver sans chercher.
LG_SECTION_REGLAGES = 3
LG_DEB, LG_FIN, LG_LISSAGE = 4, 5, 6
LG_SECTION_SEUILS = 9
LG_SEUIL_BAS, LG_SEUIL_HAUT = 10, 11
LG_SECTION_RESUME = 13
LG_RESUME = 14                 # cinq lignes : 14 a 18
LG_SECTION_DERNIERES = 20
LG_DERNIERES_ENT = 21
LG_DERNIERES = 22              # quinze lignes : 22 a 36
NB_DERNIERES = 15
LG_SECTION_CONTROLE = 38
LG_CONTROLE = 39

COLS_DERNIERES = [
    ("Date", "A", "dd.mm.yyyy"),
    ("Heure", "B", "hh:mm"),
    ("Demi", "L", None),
    ("Systole", "E", "0"),
    ("Diastole", "F", "0"),
    ("Pouls", "G", "0"),
    ("Poids", "O", "0.0"),
    ("Activité", "C", None),
    ("Commentaire", "N", None),
]


def ecrire_saisie(ws, dern_datas, premier, dernier):
    """La page d'accueil : les reglages, le resume, les dernieres mesures."""
    bandeau(ws, 6, "Pression Rena  —  suivi de la pression artérielle")
    for col, largeur in (("A", 30), ("B", 15), ("C", 3), ("D", 28), ("E", 15), ("F", 46)):
        ws.column_dimensions[col].width = largeur

    p_poids = "Datas!$O${0}:$O${1}".format(LIGNE_DEB, dern_datas)

    # --- les trois reglages ---------------------------------------------------
    section(ws, LG_SECTION_REGLAGES, "Réglages des deux graphiques")
    libelle(ws, LG_DEB, 1, "Date de début")
    champ_saisie(ws, LG_DEB, 2, max(premier, dernier - datetime.timedelta(days=365)), "dd.mm.yyyy")
    libelle(ws, LG_FIN, 1, "Date de fin")
    champ_saisie(ws, LG_FIN, 2, dernier, "dd.mm.yyyy")
    libelle(ws, LG_LISSAGE, 1, "Jours de la moyenne lissée")
    champ_saisie(ws, LG_LISSAGE, 2, 7, "0")

    libelle(ws, LG_DEB, 4, "Première mesure du classeur")
    libelle(ws, LG_FIN, 4, "Dernière mesure saisie")
    libelle(ws, LG_LISSAGE, 4, "Jours affichés")
    ws.cell(LG_DEB, 5, "=Premier_Jour").number_format = "dd.mm.yyyy"
    ws.cell(LG_FIN, 5, "=Derniere_Mesure").number_format = "dd.mm.yyyy"
    ws.cell(LG_LISSAGE, 5, "=Nb_J").number_format = "0"
    for lg in (LG_DEB, LG_FIN, LG_LISSAGE):
        ws.cell(lg, 5).font = FONT_VALEUR
        ws.cell(lg, 5).alignment = Alignment(horizontal="center")

    ws.cell(LG_DEB, 6, "Les deux graphiques ne montrent que cet intervalle.").font = FONT_NOTE
    ws.cell(LG_FIN, 6, "Une date hors du calendrier est ramenée à ses bornes.").font = FONT_NOTE
    ws.cell(
        LG_LISSAGE, 6, "1 = aucun lissage. 7 = une semaine. La courbe démarre après ce nombre de jours."
    ).font = FONT_NOTE

    val_date = DataValidation(
        type="date", operator="between", formula1="Premier_Jour", formula2="Dernier_Jour",
        allow_blank=False, showErrorMessage=True,
        errorTitle="Date hors calendrier",
        error="La date doit tomber entre le premier et le dernier jour de la feuille Datas.",
    )
    ws.add_data_validation(val_date)
    val_date.add(ws.cell(LG_DEB, 2))
    val_date.add(ws.cell(LG_FIN, 2))

    val_liss = DataValidation(
        type="whole", operator="between", formula1="1", formula2="90",
        allow_blank=False, showErrorMessage=True,
        errorTitle="Lissage hors limites",
        error="Le nombre de jours de la moyenne lissée doit tenir entre 1 et 90.",
    )
    ws.add_data_validation(val_liss)
    val_liss.add(ws.cell(LG_LISSAGE, 2))

    # --- les quatre seuils ----------------------------------------------------
    section(ws, LG_SECTION_SEUILS, "Seuils de référence (mmHg)")
    libelle(ws, LG_SEUIL_BAS, 1, "Diastolique — ligne basse")
    champ_saisie(ws, LG_SEUIL_BAS, 2, 80, "0")
    libelle(ws, LG_SEUIL_HAUT, 1, "Diastolique — ligne haute")
    champ_saisie(ws, LG_SEUIL_HAUT, 2, 90, "0")
    libelle(ws, LG_SEUIL_BAS, 4, "Systolique — ligne basse")
    champ_saisie(ws, LG_SEUIL_BAS, 5, 120, "0")
    libelle(ws, LG_SEUIL_HAUT, 4, "Systolique — ligne haute")
    champ_saisie(ws, LG_SEUIL_HAUT, 5, 130, "0")
    ws.cell(LG_SEUIL_BAS, 6, "Les quatre lignes de repère des deux graphiques.").font = FONT_NOTE
    ws.cell(
        LG_SEUIL_HAUT, 6, "Repères de lecture, et non un diagnostic : voir docs/GRAPHIQUES.md."
    ).font = FONT_NOTE

    # --- le resume de la periode choisie --------------------------------------
    # Tout passe par les plages nommees, donc par la fenetre EXACTE des deux
    # graphiques : le resume ne peut pas annoncer autre chose que ce qu'ils
    # montrent. MAX rend zero sur une plage vide, d'ou le NB() qui le precede ;
    # MOYENNE rend #DIV/0!, que SIERREUR suffit a couvrir.
    section(ws, LG_SECTION_RESUME, "Résumé de la période affichée")
    gauche = [
        ("Mesures", "=COUNT(Systole)", "0"),
        ("Systolique moyenne", '=IFERROR(AVERAGE(Systole),"")', "0.0"),
        ("Diastolique moyenne", '=IFERROR(AVERAGE(Diastole),"")', "0.0"),
        ("Pouls moyen", '=IFERROR(AVERAGE(Pouls),"")', "0.0"),
        (
            "Au-dessus d'un seuil haut",
            '=IFERROR(SUMPRODUCT((Systole>0)*((Systole>Seuil_Sys_Haut)'
            '+(Diastole>Seuil_Dia_Haut)>0))/COUNT(Systole),"")',
            "0.0%",
        ),
    ]
    droite = [
        ("Jours mesurés", '=COUNTIF(Jours_Nb,">0")', "0"),
        ("Systolique maximale", '=IF(COUNT(Systole)=0,"",MAX(Systole))', "0"),
        ("Diastolique maximale", '=IF(COUNT(Diastole)=0,"",MAX(Diastole))', "0"),
        ("Poids, dernier relevé", '=IFERROR(LOOKUP(2,1/({0}<>""),{0}),"")'.format(p_poids), "0.0"),
        ("Poids moyen", '=IFERROR(AVERAGE(Poids_Mesure),"")', "0.0"),
    ]
    for i, ((lib_g, f_g, fmt_g), (lib_d, f_d, fmt_d)) in enumerate(zip(gauche, droite)):
        lg = LG_RESUME + i
        libelle(ws, lg, 1, lib_g)
        cel = ws.cell(lg, 2, f_g)
        cel.number_format, cel.font = fmt_g, FONT_VALEUR
        cel.alignment = Alignment(horizontal="center")
        libelle(ws, lg, 4, lib_d)
        cel = ws.cell(lg, 5, f_d)
        cel.number_format, cel.font = fmt_d, FONT_VALEUR
        cel.alignment = Alignment(horizontal="center")

    # --- les quinze dernieres mesures -----------------------------------------
    section(ws, LG_SECTION_DERNIERES, "Les quinze dernières mesures")
    for i, (titre, _, _) in enumerate(COLS_DERNIERES, start=1):
        cel = ws.cell(LG_DERNIERES_ENT, i, titre)
        cel.font = FONT_ENTETE
        cel.fill = PatternFill("solid", fgColor=COUL_ENTETE)
        cel.alignment = Alignment(horizontal="center")
        cel.border = BORD_CHAMP
    for i in range(NB_DERNIERES):
        lg = LG_DERNIERES + i
        for j, (_, col_datas, fmt) in enumerate(COLS_DERNIERES, start=1):
            cel = ws.cell(
                lg,
                j,
                '=IF(Reglages!$B${0}=0,"",IFERROR(INDEX(Datas!${1}:${1},Reglages!$B${0}),""))'.format(
                    LG_REG_LIGNES + i, col_datas
                ),
            )
            cel.font = FONT_NORMAL
            if fmt:
                cel.number_format = fmt

    # --- le controle de structure ---------------------------------------------
    section(ws, LG_SECTION_CONTROLE, "Contrôle du classeur")
    cel = ws.cell(LG_CONTROLE, 1, "=Controle")
    cel.font = Font(name="Segoe UI", size=10, bold=True)
    ws.cell(
        LG_CONTROLE + 1,
        1,
        "« Datas » doit porter exactement deux lignes par jour. N'insérez et ne "
        "supprimez jamais de ligne : remplissez celle du jour.",
    ).font = FONT_NOTE

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"


# ===========================================================================
# FEUILLE « Reglages »
# ===========================================================================
LG_REG_BORNES = 4        # quatre lignes : 4 a 7
LG_REG_FENETRE = 10      # six lignes : 10 a 15
LG_REG_CONTROLE = 18
LG_REG_LIGNES = 21       # quinze lignes : 21 a 35


def ecrire_reglages(ws, dern_datas, dern_calcul):
    """Les cellules derivees. Personne ne les saisit : les plages nommees les lisent."""
    bandeau(ws, 3, "Cellules dérivées  —  ne rien modifier ici")
    for col, largeur in (("A", 34), ("B", 16), ("C", 60)):
        ws.column_dimensions[col].width = largeur

    p_date = "Datas!$A${0}:$A${1}".format(LIGNE_DEB, dern_datas)
    p_sys = "Datas!$E${0}:$E${1}".format(LIGNE_DEB, dern_datas)

    lignes = [
        (LG_REG_BORNES, "Premier jour du calendrier", "=Datas!$A${0}".format(LIGNE_DEB), "dd.mm.yyyy",
         "Lu sur la feuille Datas ; il n'est pas saisi deux fois."),
        (LG_REG_BORNES + 1, "Dernier jour du calendrier", "=Datas!$A${0}".format(dern_datas), "dd.mm.yyyy",
         "Fin du gabarit pré-rempli."),
        (LG_REG_BORNES + 2, "Dernière mesure saisie",
         '=IFERROR(LOOKUP(2,1/({0}<>""),{1}),Premier_Jour)'.format(p_sys, p_date), "dd.mm.yyyy",
         "La dernière ligne portant une systole."),
        (LG_REG_BORNES + 3, "Jours du calendrier", "=Dernier_Jour-Premier_Jour+1", "0",
         "Doit valoir la moitié des lignes de Datas."),

        (LG_REG_FENETRE, "Début effectif", "=MAX(Date_Deb,Premier_Jour)", "dd.mm.yyyy",
         "La date de début, ramenée dans le calendrier."),
        (LG_REG_FENETRE + 1, "Fin effective", "=MIN(MAX(Fin_Demandee,Deb_Eff),Dernier_Jour)", "dd.mm.yyyy",
         "La date de fin, ramenée dans le calendrier et jamais avant le début."),
        (LG_REG_FENETRE + 2, "Fin demandée", "=Date_Fin", "dd.mm.yyyy",
         "Recopie de la cellule de saisie : elle évite une référence circulaire."),
        (LG_REG_FENETRE + 3, "Décalage, en jours", "=MAX(0,Deb_Eff-Premier_Jour)", "0",
         "Nombre de lignes à sauter dans Calcul."),
        (LG_REG_FENETRE + 4, "Nombre de jours", "=MAX(1,Fin_Eff-Deb_Eff+1)", "0",
         "Hauteur des plages du graphique lissé."),
        (LG_REG_FENETRE + 5, "Décalage, en lignes", "=Dec_J*2", "0",
         "Deux lignes de Datas par jour : c'est tout le calcul."),
        (LG_REG_FENETRE + 6, "Nombre de lignes", "=Nb_J*2", "0",
         "Hauteur des plages du graphique des mesures."),
    ]
    section(ws, LG_REG_BORNES - 1, "Bornes du calendrier")
    section(ws, LG_REG_FENETRE - 1, "Fenêtre des graphiques")
    for lg, lib, formule, fmt, note in lignes:
        libelle(ws, lg, 1, lib)
        cel = ws.cell(lg, 2, formule)
        cel.number_format, cel.font = fmt, FONT_VALEUR
        cel.alignment = Alignment(horizontal="center")
        cel.border = BORD_CHAMP
        ws.cell(lg, 3, note).font = FONT_NOTE

    # --- le controle de structure ---------------------------------------------
    # Il verifie ce dont depend tout le classeur : deux lignes de Datas par
    # ligne de Calcul, et les memes dates aux deux bouts. Une ligne inseree ou
    # supprimee fait tomber le compte, et le message le dit avant que les
    # courbes ne se decalent en silence.
    section(ws, LG_REG_CONTROLE - 1, "Contrôle de structure")
    libelle(ws, LG_REG_CONTROLE, 1, "État")
    ws.cell(
        LG_REG_CONTROLE,
        2,
        '=IF(AND(COUNT({0})=2*COUNT(Calcul!$A${1}:$A${2}),Datas!$A${1}=Calcul!$A${1},'
        'Datas!$A${3}=Calcul!$A${2}),"Structure correcte : "&COUNT(Calcul!$A${1}:$A${2})&'
        '" jours, 2 lignes par jour.","ATTENTION : la feuille Datas ne porte plus exactement '
        'deux lignes par jour. Les graphiques sont décalés. Rétablissez les lignes supprimées '
        'ou insérées, ou relancez outils/construire_classeur.py.")'.format(
            p_date, LIGNE_DEB, dern_calcul, dern_datas
        ),
    ).font = FONT_NORMAL

    # --- les lignes des quinze dernieres mesures -------------------------------
    # RECHERCHE(2;1/(plage<>"");...) est le tour classique qui rend la DERNIERE
    # ligne non vide d'une plage : la division fabrique des 1 et des #DIV/0!, et
    # RECHERCHE, qui ignore les erreurs, s'arrete sur le dernier 1. Il vaut ici
    # mieux qu'AGGREGATE en forme matricielle, que toutes les versions
    # n'evaluent pas de la meme facon.
    #
    # Chaque ligne repart de la precedente en RACCOURCISSANT la plage : la
    # deuxieme cherche au-dessus de la premiere, et ainsi de suite. D'ou
    # DECALER, seul moyen de donner a une plage une hauteur calculee.
    section(ws, LG_REG_LIGNES - 1, "Lignes des quinze dernières mesures")
    plage_haute = 'OFFSET(Datas!$E${0},0,0,MAX(1,B{1}-{0}),1)'
    for i in range(NB_DERNIERES):
        lg = LG_REG_LIGNES + i
        ws.cell(lg, 1, i + 1).font = FONT_NORMAL
        if i == 0:
            formule = '=IFERROR(LOOKUP(2,1/({0}<>""),ROW({0})),0)'.format(p_sys)
        else:
            haute = plage_haute.format(LIGNE_DEB, lg - 1)
            formule = (
                '=IF(B{0}<={1},0,IFERROR(LOOKUP(2,1/({2}<>""),ROW({2})),0))'.format(
                    lg - 1, LIGNE_DEB, haute
                )
            )
        ws.cell(lg, 2, formule).font = FONT_NORMAL
    ws.cell(LG_REG_LIGNES, 3, "Numéro de ligne dans Datas, de la plus récente à la plus ancienne.").font = FONT_NOTE

    ws.sheet_view.showGridLines = False


# ===========================================================================
# PLAGES NOMMEES
# ===========================================================================
def poser_noms(wb):
    """Les plages nommees, toutes dynamiques.

    Les neuf premieres portent les noms du fichier d'origine : un graphique
    recree ailleurs sur les memes noms continue de fonctionner.

    Aucune n'utilise EQUIV. L'invariant deux-lignes-par-jour transforme la
    recherche d'une date en soustraction, et une soustraction ne peut pas
    rendre #N/A quand la date cherchee ne figure pas dans la colonne.
    """
    from openpyxl.workbook.defined_name import DefinedName

    noms = {
        "Date_Deb": "Saisie!$B${0}".format(LG_DEB),
        "Date_Fin": "Saisie!$B${0}".format(LG_FIN),
        "Nb_Jours": "Saisie!$B${0}".format(LG_LISSAGE),
        "Seuil_Dia_Bas": "Saisie!$B${0}".format(LG_SEUIL_BAS),
        "Seuil_Dia_Haut": "Saisie!$B${0}".format(LG_SEUIL_HAUT),
        "Seuil_Sys_Bas": "Saisie!$E${0}".format(LG_SEUIL_BAS),
        "Seuil_Sys_Haut": "Saisie!$E${0}".format(LG_SEUIL_HAUT),
        "Premier_Jour": "Reglages!$B${0}".format(LG_REG_BORNES),
        "Dernier_Jour": "Reglages!$B${0}".format(LG_REG_BORNES + 1),
        "Derniere_Mesure": "Reglages!$B${0}".format(LG_REG_BORNES + 2),
        "Deb_Eff": "Reglages!$B${0}".format(LG_REG_FENETRE),
        "Fin_Eff": "Reglages!$B${0}".format(LG_REG_FENETRE + 1),
        "Fin_Demandee": "Reglages!$B${0}".format(LG_REG_FENETRE + 2),
        "Dec_J": "Reglages!$B${0}".format(LG_REG_FENETRE + 3),
        "Nb_J": "Reglages!$B${0}".format(LG_REG_FENETRE + 4),
        "Dec_M": "Reglages!$B${0}".format(LG_REG_FENETRE + 5),
        "Nb_M": "Reglages!$B${0}".format(LG_REG_FENETRE + 6),
        "Controle": "Reglages!$B${0}".format(LG_REG_CONTROLE),
    }

    # Les mesures brutes : deux lignes par jour, donc Dec_M et Nb_M.
    for nom, col in (
        ("Plage_X", C_COMPO), ("Systole", C_SYS), ("Diastole", C_DIA), ("Pouls", C_POULS),
        ("Basse_Bas", C_BB), ("Basse_Hau", C_BH), ("Haute_Bas", C_HB), ("Haute_Hau", C_HH),
        ("Change", C_CHANGE), ("Poids_Mesure", C_POIDS),
    ):
        noms[nom] = "OFFSET(Datas!${0}${1},Dec_M,0,Nb_M,1)".format(
            get_column_letter(col), LIGNE_DEB
        )

    # Les moyennes lissees : une ligne par jour, donc Dec_J et Nb_J.
    for nom, col in (
        ("Jour_X", K_LIB), ("Jours_Nb", K_NB),
        ("Sys_Haut_L", K_LSYSH), ("Sys_Bas_L", K_LSYSB),
        ("Dia_Haut_L", K_LDIAH), ("Dia_Bas_L", K_LDIAB),
        ("Z_Dia_Bas", K_ZDIAB), ("Z_Dia", K_ZDIA), ("Z_Ecart", K_ZECART), ("Z_Sys", K_ZSYS),
        ("L_Basse_Bas", K_SDB), ("L_Basse_Hau", K_SDH),
        ("L_Haute_Bas", K_SSB), ("L_Haute_Hau", K_SSH),
    ):
        noms[nom] = "OFFSET(Calcul!${0}${1},Dec_J,0,Nb_J,1)".format(
            get_column_letter(col), LIGNE_DEB
        )

    for nom, ref in noms.items():
        wb.defined_names.add(DefinedName(nom, attr_text=ref))


# ===========================================================================
# LES DEUX GRAPHIQUES
# ===========================================================================
# Les series ne sont pas construites par Reference : elles designent des PLAGES
# NOMMEES, et la fabrique de series d'openpyxl n'accepte qu'une adresse de
# cellules. On assemble donc l'objet Series a la main — c'est la seule facon
# d'obtenir <c:f>[0]!Systole</c:f>, la reference qui rend le graphique
# dependant des trois reglages plutot que d'un nombre de lignes fige.
def serie(nom_plage, titre, entete=None):
    ser = Series()
    ser.val = NumDataSource(numRef=NumRef(f="[0]!" + nom_plage))
    ser.cat = AxDataSource(numRef=NumRef(f="[0]!" + titre))
    if entete:
        ser.tx = SeriesLabel(strRef=StrRef(f=entete))
    return ser


def trait(ser, couleur, epaisseur=12700, pointille=False, marqueur="none"):
    """Habille une serie de courbe."""
    ligne = LineProperties(solidFill=couleur, w=epaisseur)
    if pointille:
        ligne.prstDash = "sysDash"
    ser.graphicalProperties = GraphicalProperties(ln=ligne)
    ser.marker = Marker(symbol=marqueur, size=4)
    ser.smooth = False
    return ser


def aire(ser, couleur):
    """Habille une serie d'aire. couleur = None pour une aire invisible."""
    if couleur is None:
        ser.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    else:
        ser.graphicalProperties = GraphicalProperties(
            solidFill=couleur, ln=LineProperties(noFill=True)
        )
    return ser


def numeroter(chart):
    """Renumerote idx et order sur tout le graphique, groupes confondus.

    Deux series qui partagent un idx font qu'Excel n'en affiche qu'une, et
    signale un fichier illisible. La numerotation doit donc etre faite APRES
    la reunion des groupes, jamais groupe par groupe.
    """
    n = 0
    for groupe in [chart] + list(chart._charts[1:] if hasattr(chart, "_charts") else []):
        for ser in groupe.series:
            ser.idx = n
            ser.order = n
            n += 1
    return n


def cacher_legende(chart, indices):
    """Retire de la legende les series qui ne sont la que pour l'empilement."""
    chart.legend = Legend()
    chart.legend.position = "b"
    chart.legend.legendEntry = [LegendEntry(idx=i, delete=True) for i in indices]


def regler_axes(chart, titre_y="mmHg"):
    chart.y_axis.scaling.min = 60
    chart.y_axis.scaling.max = 170
    chart.y_axis.majorUnit = 10
    chart.y_axis.title = titre_y
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.dispBlanksAs = "gap"
    chart.height = 18
    chart.width = 32


def graphique_mesures():
    """Le graphique d'origine : chaque mesure, matin et soir, et ses reperes."""
    barres = BarChart()
    barres.type = "col"
    barres.grouping = "clustered"
    barres.gapWidth = 40
    ser = serie("Change", "Plage_X", "Datas!$M$2")
    ser.graphicalProperties = GraphicalProperties(
        solidFill=COUL_CHANGE, ln=LineProperties(noFill=True)
    )
    barres.series.append(ser)

    courbes = LineChart()
    courbes.grouping = "standard"
    for plage, entete, couleur, marqueur, epaisseur in (
        ("Systole", "$E$2", COUL_SYSTOLE, "x", 12700),
        ("Diastole", "$F$2", COUL_DIASTOLE, "x", 12700),
        ("Basse_Bas", "$H$2", COUL_SEUIL_BAS, "none", 9525),
        ("Basse_Hau", "$I$2", COUL_SEUIL_BAS, "none", 9525),
        ("Haute_Bas", "$J$2", COUL_SEUIL_HAUT, "none", 9525),
        ("Haute_Hau", "$K$2", COUL_SEUIL_HAUT, "none", 9525),
    ):
        courbes.series.append(
            trait(serie(plage, "Plage_X", "Datas!" + entete), couleur, epaisseur, marqueur=marqueur)
        )

    barres += courbes
    barres.title = "Mesures  —  chaque relevé, matin et soir"
    regler_axes(barres)
    numeroter(barres)
    cacher_legende(barres, [])
    return barres


def graphique_lissage():
    """Le graphique demande : deux bandes lissees, verte et bleue.

    L'AIRE EMPILEE D'ABORD, LES COURBES ENSUITE. L'ordre des groupes dans le
    fichier est l'ordre de dessin : les quatre aires posees en premier passent
    derriere, les huit courbes par-dessus. Inverser les deux ferait disparaitre
    les courbes sous les aplats.
    """
    bandes = AreaChart()
    bandes.grouping = "stacked"
    bandes.overlap = 100
    for plage, entete, couleur in (
        ("Z_Dia_Bas", None, None),
        ("Z_Dia", "Diastolique", COUL_BANDE_DIA),
        ("Z_Ecart", None, None),
        ("Z_Sys", "Systolique", COUL_BANDE_SYS),
    ):
        ser = serie(plage, "Jour_X")
        if entete:
            ser.tx = SeriesLabel(v="Plage " + entete.lower())
        bandes.series.append(aire(ser, couleur))

    courbes = LineChart()
    courbes.grouping = "standard"
    for plage, titre, couleur, epaisseur in (
        ("Sys_Haut_L", "Systolique haute lissée", COUL_TRAIT_SYS, 19050),
        ("Sys_Bas_L", "Systolique basse lissée", COUL_TRAIT_SYS, 19050),
        ("Dia_Haut_L", "Diastolique haute lissée", COUL_TRAIT_DIA, 19050),
        ("Dia_Bas_L", "Diastolique basse lissée", COUL_TRAIT_DIA, 19050),
        ("L_Basse_Bas", "Seuil dia. bas", COUL_SEUIL_BAS, 9525),
        ("L_Basse_Hau", "Seuil dia. haut", COUL_SEUIL_BAS, 9525),
        ("L_Haute_Bas", "Seuil sys. bas", COUL_SEUIL_HAUT, 9525),
        ("L_Haute_Hau", "Seuil sys. haut", COUL_SEUIL_HAUT, 9525),
    ):
        ser = serie(plage, "Jour_X")
        ser.tx = SeriesLabel(v=titre)
        courbes.series.append(trait(ser, couleur, epaisseur))

    bandes += courbes
    bandes.title = "Moyennes lissées  —  la bande va du plus bas au plus haut du jour"
    regler_axes(bandes)
    numeroter(bandes)
    # Les deux aires invisibles ne sont la que pour l'empilement : elles
    # n'apparaissent pas dans la legende.
    cacher_legende(bandes, [0, 2])
    return bandes


# ===========================================================================
# POINT D'ENTREE
# ===========================================================================
def main(argv):
    if len(argv) != 3:
        raise SystemExit(__doc__)
    source, destination = argv[1], argv[2]

    jours, titre = lire_source(source)
    premier = min(jours)
    dernier = max(max(jours), FIN_GABARIT)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_saisie = wb.create_sheet("Saisie")
    ws_datas = wb.create_sheet("Datas")
    ws_calcul = wb.create_sheet("Calcul")
    ws_reglages = wb.create_sheet("Reglages")

    dern_datas = ecrire_datas(ws_datas, jours, premier, dernier, titre)
    dern_calcul = ecrire_calcul(ws_calcul, premier, dernier, dern_datas)
    # La fenetre par defaut se cale sur la derniere MESURE, et non sur la fin du
    # gabarit : celui-ci court jusqu'en 2028 et ouvrirait les deux graphiques
    # sur une annee entierement vide.
    derniere_mesure = max(
        jour for jour, demi in jours.items() if any(d["sys"] is not None for d in demi)
    )
    ecrire_saisie(ws_saisie, dern_datas, premier, derniere_mesure)
    ecrire_reglages(ws_reglages, dern_datas, dern_calcul)
    poser_noms(wb)

    wb.create_chartsheet("Graph_Mesures").add_chart(graphique_mesures())
    wb.create_chartsheet("Graph_Lissage").add_chart(graphique_lissage())

    # Les formules sont ecrites sans valeur en cache : sans cet indicateur,
    # Excel afficherait des cellules vides jusqu'a la premiere modification.
    wb.calculation.fullCalcOnLoad = True
    wb.save(destination)

    mesures = sum(1 for demi in jours.values() for d in demi if d["sys"] is not None)
    print("Classeur ecrit      : %s" % destination)
    print("Jours du calendrier : %d  (%s -> %s)" % (
        (dernier - premier).days + 1, premier, dernier))
    print("Lignes de Datas     : %d  (derniere ligne %d)" % (dern_datas - 2, dern_datas))
    print("Lignes de Calcul    : %d  (derniere ligne %d)" % (dern_calcul - 2, dern_calcul))
    print("Mesures reprises    : %d" % mesures)


if __name__ == "__main__":
    main(sys.argv)
