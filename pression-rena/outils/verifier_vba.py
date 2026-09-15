#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Controle les modules VBA de « Pression Rena » sans ouvrir Excel.

    python3 verifier_vba.py ../src

VBA ne signale a la compilation ni un nom public employe deux fois, ni un
controle appele sur un formulaire qui ne le porte pas : la premiere faute
donne une erreur a l'import, la seconde une erreur 438 a l'execution, sur une
ligne d'apparence irreprochable. Ce script les cherche a la lecture.

Il verifie :
  1. qu'aucun nom public — procedure ou constante — n'est declare deux fois ;
  2. que tout controle appele par le formulaire est bien cree par le generateur ;
  3. que toute procedure appelee par le code genere existe, et est publique ;
  4. qu'aucune variable locale ne masque une procedure du meme module ;
  4bis. qu'aucun menu deroulant FERME ne recoit une affectation de .Text ;
  5. que la geometrie tient dans le formulaire ;
  6. que chaque fichier est bien en Windows-1252 / CRLF, l'encodage attendu par
     l'editeur VBA.
"""

import glob
import os
import re
import sys

ENCODAGE = "cp1252"

RE_PROC = re.compile(
    r"^(Public |Private )?(Sub|Function)\s+([A-Za-z_]\w*)", re.M)
RE_CONST = re.compile(
    r"^(Public|Private)\s+Const\s+([A-Za-z_]\w*)", re.M)
RE_TYPE = re.compile(r"^(Public|Private)\s+Type\s+([A-Za-z_]\w*)", re.M)
RE_AJ = re.compile(r'Aj\(\s*\w+\s*,\s*"[^"]+"\s*,\s*"([^"]+)"')
RE_CTRL_POINT = re.compile(r"\bf\.([A-Za-z_]\w*)")
RE_CTRL_NOM = re.compile(r'f\.Controls\(\s*"([^"]+)"')
RE_CTRL_CALC = re.compile(r'f\.Controls\(\s*"([A-Za-z_]+)_"')
RE_PROC_EVT = re.compile(r'Proc\s+"[^"]*"\s*,\s*"([A-Za-z_]\w*)')
RE_DIM = re.compile(r"^\s*Dim\s+([^\n]+)", re.M)
RE_LISTE = re.compile(r"^\s*Liste\s+\w+\s*,\s*(True|False)\s*$")

# Les membres du formulaire qui ne sont pas des controles.
MEMBRES_FORM = {"Controls", "Show", "Hide", "Repaint", "Caption", "Name", "Tag"}


# Les cinq octets que Windows-1252 ne definit pas. Leur presence signale un
# fichier qui n'est pas dans cet encodage, quoi qu'en dise son extension.
OCTETS_INDEFINIS = {0x81, 0x8D, 0x8F, 0x90, 0x9D}


def modules(dossier):
    """Les modules, lus en Windows-1252 : c'est ainsi qu'ils sont enregistres."""
    for chemin in sorted(glob.glob(os.path.join(dossier, "modPression_*.bas"))):
        with open(chemin, "rb") as f:
            octets = f.read()
        yield os.path.basename(chemin), (octets.decode(ENCODAGE, "replace"), octets)


def main(argv):
    dossier = argv[1] if len(argv) > 1 else "src"
    lus = dict(modules(dossier))
    if not lus:
        raise SystemExit("aucun module modPression_*.bas dans %s" % dossier)
    sources = {nom: texte for nom, (texte, _) in lus.items()}

    fautes = []
    controles = 0

    # --- 1. les noms publics, une seule fois chacun --------------------------
    publics = {}
    for nom, texte in sources.items():
        for portee, _, proc in RE_PROC.findall(texte):
            if portee.strip() != "Private":
                publics.setdefault(proc, []).append(nom)
        for portee, const in RE_CONST.findall(texte):
            if portee == "Public":
                publics.setdefault(const, []).append(nom)
        for portee, typ in RE_TYPE.findall(texte):
            if portee == "Public":
                publics.setdefault(typ, []).append(nom)
    for nom, ou in publics.items():
        controles += 1
        if len(ou) > 1:
            fautes.append("nom public « %s » déclaré dans %s" % (nom, " et ".join(ou)))

    # --- 2. les controles appeles existent-ils ? -----------------------------
    generateur = sources.get("modPression_Generateur.bas", "")
    crees = set(RE_AJ.findall(generateur))
    # les controles de la grille sont crees dans une boucle : on releve leur
    # prefixe plutot que chacun de leurs noms
    prefixes_grille = {n.split("_")[0] for n in crees if "_" in n}
    prefixes_grille |= {"lblPL", "lblP", "lblPEnt"}

    formulaire = sources.get("modPression_Formulaire.bas", "")
    appeles = set(RE_CTRL_POINT.findall(formulaire)) - MEMBRES_FORM
    appeles |= set(RE_CTRL_NOM.findall(formulaire))
    for appel in sorted(appeles):
        controles += 1
        if appel in crees:
            continue
        if appel.split("_")[0] in prefixes_grille:
            continue
        fautes.append("le formulaire appelle « %s », que le générateur ne crée pas" % appel)

    # --- 3. les procedures du code genere existent-elles ? -------------------
    for proc in sorted(set(RE_PROC_EVT.findall(generateur))):
        controles += 1
        if proc not in publics:
            fautes.append("le code généré appelle « %s », qui n'existe nulle part" % proc)
        elif "modPression_Formulaire.bas" not in publics[proc]:
            fautes.append("« %s » n'est pas dans modPression_Formulaire" % proc)

    # --- 4. une locale qui masque une procedure du meme module ---------------
    # Une variable nommee comme une procedure du module MASQUE cette procedure :
    # VBA ne dit rien a la compilation, transforme l'appel en acces tardif sur
    # l'objet, et MSForms repond a l'execution « propriete ou methode non geree
    # par cet objet ».
    for nom, texte in sources.items():
        procs = {p for _, _, p in RE_PROC.findall(texte)}
        for declaration in RE_DIM.findall(texte):
            for morceau in declaration.split(","):
                variable = morceau.strip().split()[0] if morceau.strip() else ""
                variable = variable.split("(")[0]
                controles += 1
                if variable and variable in procs:
                    fautes.append(
                        "%s : la variable « %s » masque la procédure du même nom"
                        % (nom, variable))

    # --- 4bis. un menu ferme ne recoit jamais .Text -------------------------
    # Un ComboBox en fmStyleDropDownList REFUSE l'affectation de .Text : MSForms
    # repond « erreur 380, valeur de propriete non valide » des que la chaine ne
    # figure pas dans la liste — et une chaine vide n'y figure jamais. Il faut
    # passer par ListIndex, qui accepte -1 pour « rien de choisi ».
    #
    # Rien ne le signale a la compilation : la faute n'apparait qu'a
    # l'ouverture du formulaire, sur une ligne d'apparence irreprochable.
    fermes = set()
    dernier = None
    for ligne in generateur.splitlines():
        trouve = RE_AJ.search(ligne)
        if trouve:
            dernier = trouve.group(1)
        m = RE_LISTE.match(ligne)
        if m and dernier:
            if m.group(1) == "True":
                fermes.add(dernier)
            dernier = None
    for nom in sorted(fermes):
        controles += 1
        for module, texte in sources.items():
            if re.search(r"\bf\.%s\.Text\s*=" % re.escape(nom), texte):
                fautes.append(
                    "%s affecte .Text à « %s », un menu fermé : erreur 380 à "
                    "l'exécution (passer par ListIndex)" % (module, nom))

    # --- 5. la geometrie ----------------------------------------------------
    fautes.extend(verifier_geometrie(sources))
    controles += 6

    # --- 6. l'encodage ------------------------------------------------------
    # L'editeur VBA lit les .bas en Windows-1252. Un fichier reenregistre en
    # UTF-8 s'importe sans rien dire et deforme tous les accents des libelles
    # et des messages : c'est la faute qui se voit le plus tard.
    for nom, (texte, octets) in lus.items():
        controles += 3
        indefinis = sorted({b for b in octets if b in OCTETS_INDEFINIS})
        if indefinis:
            fautes.append("%s : octets %s, indéfinis en %s"
                          % (nom, [hex(b) for b in indefinis], ENCODAGE))
        if b"\r\n" not in octets:
            fautes.append("%s : pas de fin de ligne CRLF" % nom)
        hauts = [b for b in octets if b >= 0x80]
        if hauts:
            try:
                octets.decode("utf-8")
                fautes.append("%s : le fichier est en UTF-8, pas en %s "
                              "(les accents seraient déformés à l'import)" % (nom, ENCODAGE))
            except UnicodeDecodeError:
                pass
        try:
            texte.encode(ENCODAGE)
        except UnicodeEncodeError as e:
            fautes.append("%s : « %s » ne s'écrit pas en %s (l'écrire en ChrW)"
                          % (nom, texte[e.start:e.end], ENCODAGE))

    print("Modules      : %d" % len(sources))
    print("Contrôles    : %d" % controles)
    if fautes:
        print("\nÉCHEC : %d anomalies" % len(fautes))
        for f in fautes:
            print("  - %s" % f)
        return 1
    print("\nOK : aucune anomalie.")
    return 0


def constantes(sources):
    """Les constantes numeriques publiques, evaluees."""
    valeurs = {}
    motif = re.compile(
        r"^Public\s+Const\s+([A-Za-z_]\w*)\s+As\s+\w+\s*=\s*([^'\n]+)", re.M)
    for texte in sources.values():
        for nom, expr in motif.findall(texte):
            expr = expr.strip()
            try:
                valeurs[nom] = eval(expr, {"__builtins__": {}}, dict(valeurs))
            except Exception:
                pass
    return valeurs


def verifier_geometrie(sources):
    """La geometrie du formulaire tient-elle dans ses bords ?"""
    c = constantes(sources)
    fautes = []

    def borne(nom, fin, limite, quoi):
        if fin > limite + 0.01:
            fautes.append("%s : %s finit à %.1f pt pour %.1f disponibles"
                          % (nom, quoi, fin, limite))

    # les trois cartes et la barre de boutons, dans la hauteur utile
    borne("zone 1", c["P_Z1_TOP"] + c["P_Z1_HAUT"], c["P_Z2_TOP"], "le bandeau")
    borne("zone 2", c["P_Z2_TOP"] + c["P_Z2_HAUT"], c["P_Z3_TOP"], "la fiche")
    borne("zone 3", c["P_Z3_TOP"] + c["P_Z3_HAUT"], c["P_BT_TOP"], "le tableau")
    borne("boutons", c["P_BT_TOP"] + c["P_BT_HAUT"], c["P_HAUTEUR"], "la barre de boutons")

    # la fiche : deux blocs de quatre lignes
    largeur_blocs = (c["P_GR_X"] * 2 + c["P_NB_BLOCS"] * c["P_GR_BLOC"]
                     + (c["P_NB_BLOCS"] - 1) * c["P_GR_GOUTTIERE"])
    borne("fiche", largeur_blocs, c["P_CARTE_LARG"], "les blocs")
    bas_fiche = (c["P_GR_Y"] + (c["P_NB_LIGNES"] - 1) * c["P_GR_LIGNE"]
                 + c["P_LBL_HAUT"] + c["P_CTL_HAUT"])
    borne("fiche", bas_fiche, c["P_Z2_HAUT"], "la dernière ligne")

    # le tableau : les colonnes, puis les lignes
    largeurs = re.search(r"PDernieresLargeurs\s*=\s*Array\(([^)]+)\)",
                         sources["modPression_Theme.bas"])
    total = sum(float(x) for x in largeurs.group(1).split(","))
    borne("tableau", total + 2 * c["P_PAD_X"], c["P_CARTE_LARG"] - 2, "les colonnes")
    bas_grille = (c["P_TITRE_HAUT"] + c["P_ENTETE_HAUT"]
                  + c["P_NB_DERNIERES"] * c["P_LIGNE_H"])
    borne("tableau", bas_grille, c["P_Z3_HAUT"], "les lignes")

    # les trois boutons de droite ne doivent pas mordre sur celui de gauche
    droite = (c["P_LARGEUR"] - c["P_MARGE"]
              - 3 * c["P_BT_LARG"] - 2 * c["P_BT_GOUTTIERE"])
    if droite < c["P_MARGE"]:
        fautes.append("boutons : les trois boutons de droite débordent à gauche")

    # le nombre de colonnes doit etre le meme dans les quatre tableaux
    tailles = {}
    for fonction in ("PDernieresColonnes", "PDernieresLibelles",
                     "PDernieresLargeurs", "PDernieresAlignements"):
        m = re.search(fonction + r"\s*=\s*Array\((.*?)\)\s*\n",
                      sources["modPression_Theme.bas"], re.S)
        tailles[fonction] = m.group(1).count(",") + 1 if m else -1
    if len(set(tailles.values())) != 1:
        fautes.append("tableau : les quatre tableaux de colonnes n'ont pas la même "
                      "longueur : %s" % tailles)

    return fautes


if __name__ == "__main__":
    sys.exit(main(sys.argv))
