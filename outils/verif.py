# -*- coding: utf-8 -*-
"""Controles structurels sur les modules VBA de src/.

VBA ne dit presque jamais ce qui ne va pas : « erreur de syntaxe » sans montrer
le nom fautif, « nom ambigu » sans dire ou, « propriete non geree par cet
objet » sans nommer le controle. Chaque controle ci-dessous est ne d'un defaut
reel rencontre sur ce classeur, et chacun a ete valide en reintroduisant ce
defaut.

    python3 outils/verif.py
"""
import re, os, sys, collections

SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src")
rd = lambda f: open(os.path.join(SRC, f), 'rb').read().decode('cp1252')
FICHIERS = sorted(f for f in os.listdir(SRC) if f.endswith(".bas"))
pb = []

# ---------------------------------------------------------------------------
# Lignes LOGIQUES : une suite terminee par « _ » n'en fait qu'une pour VBA.
# ---------------------------------------------------------------------------
def lignes_logiques(texte):
    out, cur, debut = [], "", 0
    for no, l in enumerate(texte.replace("\r\n", "\n").split("\n"), 1):
        if not cur:
            debut = no
        s = l.rstrip()
        if s.endswith(" _"):
            cur += s[:-1]
        else:
            out.append((debut, cur + s))
            cur = ""
    if cur:
        out.append((debut, cur))
    return out

def sans_commentaire(l):
    """retire le commentaire, en laissant les apostrophes des chaines"""
    out, dans = [], False
    for c in l:
        if c == '"':
            dans = not dans
        elif c == "'" and not dans:
            break
        out.append(c)
    return "".join(out)

SIGNAT = re.compile(r'^\s*(?:Public |Private |Friend )?(?:Static )?'
                    r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)', re.I)
FIN_PROC = re.compile(r'^\s*End\s+(?:Sub|Function|Property)\b', re.I)

# ---------------------------------------------------------------------------
# 1. EQUILIBRE DES PARENTHESES SUR UNE SIGNATURE
#
# Une signature coupee par « _ » se lit comme UNE ligne. Y glisser une
# instruction au milieu -- ce qui arrive vite en editant par script -- donne
# « Sub ou Fonction non definie », sans dire laquelle.
# ---------------------------------------------------------------------------
for f in FICHIERS:
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        if SIGNAT.match(c) and c.count("(") != c.count(")"):
            pb.append("%s:%d signature aux parentheses desequilibrees" % (f, no))

# ---------------------------------------------------------------------------
# 2. NOMS RESERVES PAR VBA
#
# Une variable ne peut porter le nom d'une fonction integree (cDate) ni d'un
# mot-cle du langage (local) : « erreur de syntaxe », sans plus.
#
# Ne figurent dans MOTS_CLES que les mots TOUJOURS reserves. Ceux qui ne le
# sont que dans une instruction precise -- Lib dans Declare, Base dans Option
# Base, Read et Write dans Open -- s'emploient tres bien comme variables, et ce
# classeur les emploie deja. La regle : un mot que le classeur utilise et qui
# compile n'a pas sa place ici.
# ---------------------------------------------------------------------------
FONCTIONS = set(x.lower() for x in """
Abs Array Asc AscB AscW Atn CBool CByte CCur CDate CDbl CDec CInt CLng CSng CStr CVar CVErr
Choose Chr ChrB ChrW Command Cos CreateObject CurDir Date DateAdd DateDiff DatePart DateSerial
DateValue Day DDB Dir DoEvents Environ EOF Error Exp FileAttr FileDateTime FileLen Filter Fix
Format FormatCurrency FormatDateTime FormatNumber FormatPercent FreeFile FV GetAllSettings
GetAttr GetObject GetSetting Hex Hour IIf IMEStatus Input InputB InputBox InStr InStrB InStrRev
Int IPmt IRR IsArray IsDate IsEmpty IsError IsMissing IsNull IsNumeric IsObject Join LBound
LCase Left LeftB Len LenB Loc LOF Log LTrim Mid MidB Minute MIRR Month MonthName MsgBox Now
NPer NPV Oct Partition Pmt PPmt PV QBColor Rate Replace RGB Right RightB Rnd Round RTrim Second
Seek Sgn Shell Sin SLN Space Spc Split Sqr Str StrComp StrConv String StrReverse Switch SYD Tab
Tan Time Timer TimeSerial TimeValue Trim TypeName UBound UCase Val VarType Weekday WeekdayName Year
""".split())

MOTS_CLES = set(x.lower() for x in """
And As Attribute Boolean ByRef ByVal Byte Call Case Close Const Currency Debug Decimal
Declare Dim Do Double Each Else ElseIf Empty End Enum Eqv Erase Event Exit Explicit False
For Friend Function Get Global GoSub GoTo If Imp Implements In Integer Is Let Like Line
Load Local Lock Long Loop LSet Me Mod New Next Not Nothing Null Object On Open Option
Optional Or ParamArray Preserve Print Private Property Public Put RaiseEvent ReDim Rem
Reset Resume Return RSet Select Set Single Static Stop Sub Then To True Type TypeOf Unload
Until Variant Wend While With WithEvents Xor
""".split())

DECL = re.compile(r'^\s*(?:Public |Private |Friend )?(?:Dim|Static|Const)\s+(.*)$', re.I)
UN_NOM = re.compile(r'^\s*(\w+)\s*(?:\([^)]*\))?\s*(?:As\b|=|$)', re.I)

def noms_de_la_liste(reste):
    """les noms d'une liste « a As X, b(1 To 3) As Y, c »"""
    out, prof, cur = set(), 0, ""
    for ch in reste + ",":
        if ch == "(":
            prof += 1
        elif ch == ")":
            prof -= 1
        if ch == "," and prof == 0:
            mn = UN_NOM.match(cur)
            if mn:
                out.add(mn.group(1))
            cur = ""
        else:
            cur += ch
    return out

def noms_declares(ligne):
    """une declaration porte souvent PLUSIEURS noms : Dim a As X, b As Y"""
    md = DECL.match(ligne)
    return noms_de_la_liste(md.group(1)) if md else set()

# Une variable de MODULE n'a pas de Dim : « Private mMarque As String ». Il faut
# donc ecarter a la main tout ce qu'un Public ou un Private peut annoncer
# d'autre -- une constante, un type, une procedure.
DECL_MOD = re.compile(r'^\s*(?:Public|Private|Global|Dim|Static)\s+'
                      r'(?:WithEvents\s+)?'
                      r'(?!Const\b|Type\b|Enum\b|Sub\b|Function\b'
                      r'|Property\b|Declare\b|Event\b)(.*)$', re.I)

def noms_module(ligne):
    md = DECL_MOD.match(ligne)
    return noms_de_la_liste(md.group(1)) if md else set()

def code_seul(bloc):
    """le corps d'une procedure sans ses commentaires : chercher un nom dans
    un bloc brut le trouve aussi dans la phrase qui l'explique, et le
    controle passe alors sur du code absent"""
    return "\n".join(sans_commentaire(l)
                     for l in bloc.replace("\r\n", "\n").split("\n"))

def sans_chaine(l):
    """retire le contenu des chaines : un nom cite dans un message n'est pas
    un nom employe"""
    out, dans = [], False
    for c in l:
        if c == '"':
            dans = not dans
            continue
        if not dans:
            out.append(c)
    return "".join(out)

def noms_parametres(signature):
    return set(re.findall(r'[(,]\s*(?:ByVal |ByRef |Optional |ParamArray )*(\w+)\s+As\b',
                          signature, re.I))

for f in FICHIERS:
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        noms = noms_declares(c)
        ms = SIGNAT.match(c)
        if ms:
            noms.add(ms.group(1))
            noms |= noms_parametres(c)
        for n in noms:
            if n.lower() in FONCTIONS:
                pb.append("%s:%d %s porte le nom d'une fonction VBA integree : "
                          "erreur de syntaxe a la compilation" % (f, no, n))
            elif n.lower() in MOTS_CLES:
                pb.append("%s:%d %s est un MOT-CLE du langage VBA : erreur de "
                          "syntaxe, sans dire lequel" % (f, no, n))

# ---------------------------------------------------------------------------
# 3. DEUX MODULES NE DOIVENT PAS EXPORTER LE MEME NOM
#
# VBA repond « Nom ambigu detecte » et le projet ne compile plus. Le cas arrive
# aussi a l'import : VBA NE REMPLACE JAMAIS un module, il en ajoute un second
# en collant un chiffre au nom.
# ---------------------------------------------------------------------------
PROC_MOD, PUB_OU = {}, {}
for f in FICHIERS:
    PROC_MOD[f] = set()
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            PROC_MOD[f].add(ms.group(1).lower())
            if not re.match(r'^\s*Private\b', c, re.I):
                PUB_OU.setdefault(ms.group(1), set()).add(f)

for nom in sorted(PUB_OU):
    if len(PUB_OU[nom]) > 1:
        pb.append("%s est publique dans %s : VBA repondra « Nom ambigu »"
                  % (nom, " et ".join(sorted(PUB_OU[nom]))))

# ---------------------------------------------------------------------------
# 4. UNE VARIABLE NE DOIT PAS PORTER LE NOM D'UNE PROCEDURE
#
# Une locale de meme nom MASQUE la procedure. VBA compile « Zone c, x » en appel
# tardif sur l'objet, et c'est a l'execution que MSForms repond « propriete ou
# methode non geree par cet objet » -- une erreur 438 tres loin de la cause.
# ---------------------------------------------------------------------------
PUB_BAS = {n.lower(): sorted(v)[0] for n, v in PUB_OU.items()}
for f in FICHIERS:
    proc = None
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            proc, noms = ms.group(1), noms_parametres(c)
        elif FIN_PROC.match(c):
            proc = None
            continue
        elif proc:
            noms = noms_declares(c)
        else:
            continue
        for n in noms:
            b = n.lower()
            if proc and b == proc.lower():
                continue                      # la valeur de retour d'une Function
            if b in PROC_MOD[f]:
                pb.append("%s:%d %s porte le nom d'une procedure du meme module : "
                          "dans %s elle la masque (erreur 438 a l'execution)"
                          % (f, no, n, proc))
            elif b in PUB_BAS and PUB_BAS[b] != f:
                pb.append("%s:%d %s porte le nom d'une procedure publique de %s : "
                          "dans %s elle la masque" % (f, no, n, PUB_BAS[b], proc))

# ---------------------------------------------------------------------------
# 5. On Error Resume Next DOIT SE REFERMER
#
# Il fait AVANCER le programme sur une instruction qui a echoue : utile sur une
# ligne ou deux, desastreux sur une procedure entiere -- le travail se poursuit
# sur un etat a moitie fait, et l'erreur remonte sur une ligne qui n'y est pour
# rien.
# ---------------------------------------------------------------------------
OER = re.compile(r'^\s*On\s+Error\s+Resume\s+Next\b', re.I)
OEG = re.compile(r'^\s*On\s+Error\s+GoTo\b', re.I)
for f in FICHIERS:
    proc, ouvert = None, 0
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            proc, ouvert = ms.group(1), 0
        elif proc and OER.match(c):
            ouvert = no
        elif proc and OEG.match(c):
            ouvert = 0
        elif FIN_PROC.match(c):
            if ouvert:
                pb.append("%s:%d On Error Resume Next ouvert dans %s n'est jamais "
                          "referme" % (f, ouvert, proc))
            proc, ouvert = None, 0

# ---------------------------------------------------------------------------
# 6. Err SE RELEVE AVANT TOUTE INSTRUCTION On Error
#
# TOUTE instruction On Error remet Err a zero. Un gestionnaire qui nettoie
# quelque chose avant de lire Err.Number affiche « 0 - » au lieu du diagnostic,
# et precisement quand on en a besoin.
# ---------------------------------------------------------------------------
ETIQ = re.compile(r'^\s*([A-Za-z]\w*):\s*$')
LIT_ERR = re.compile(r'\bErr\.(Number|Description|Source)\b', re.I)
ON_ERR = re.compile(r'^\s*On\s+Error\b', re.I)
for f in FICHIERS:
    texte = rd(f)
    vises = set(re.findall(r'On\s+Error\s+GoTo\s+([A-Za-z]\w+)', texte, re.I))
    lignes = lignes_logiques(texte)
    for idx, (no, l) in enumerate(lignes):
        et = ETIQ.match(sans_commentaire(l))
        if not et or et.group(1) not in vises:
            continue
        vu_on_error = False
        for no2, l2 in lignes[idx + 1:]:
            c = sans_commentaire(l2)
            if FIN_PROC.match(c):
                break
            if LIT_ERR.search(c):
                if vu_on_error:
                    pb.append("%s:%d le gestionnaire %s lit Err APRES une "
                              "instruction On Error, qui l'a remis a zero"
                              % (f, no2, et.group(1)))
                break
            if ON_ERR.match(c):
                vu_on_error = True

# ---------------------------------------------------------------------------
# 7. LES SIX INDICATEURS DE LA FICHE 2 SE CALCULENT TOUS
#
# Ils venaient d'une feuille de totaux tenue par des formules Excel, supprimee
# depuis. Une cle declaree au schema mais oubliee par le calcul afficherait un
# tiret a la place d'un montant, sans la moindre erreur.
# ---------------------------------------------------------------------------
schema = rd("modInterv_Schema.bas")
donnees = rd("modInterv_Donnees.bas")
form = rd("modInterv_Formulaire.bas")

cles = re.findall(r'Public Const (ITU_\w+) As String', schema)
if not cles:
    pb.append("aucune cle ITU_* au schema : les tuiles ne savent plus quoi calculer")
mt = re.search(r'Function Interv_TotauxTuiles\(.*?\nEnd Function', donnees, re.S)
if not mt:
    pb.append("Interv_TotauxTuiles a disparu")
else:
    for cle in cles:
        if cle not in mt.group(0):
            pb.append("%s est declaree au schema mais Interv_TotauxTuiles ne la "
                      "calcule pas : la tuile afficherait un tiret" % cle)
mf = re.search(r'Function TuilesStatistiques\(.*?\nEnd Function', form, re.S)
if mf:
    for cle in cles:
        if cle not in mf.group(0):
            pb.append("%s se calcule mais aucune tuile ne l'affiche" % cle)

# ---------------------------------------------------------------------------
# 8. PLUS RIEN NE DOIT VISER LA FEUILLE DE TOTAUX SUPPRIMEE
# ---------------------------------------------------------------------------
for f in FICHIERS:
    texte = rd(f)
    for no, l in enumerate(texte.replace("\r\n", "\n").split("\n"), 1):
        for mot in ("Tableau7", "NOM_TABLE_GRAPH", "NOM_FEUILLE_STATS"):
            if mot in l:
                pb.append("%s:%d vise encore %s : la feuille de totaux a ete "
                          "supprimee, tout se calcule depuis TblInterv"
                          % (f, no, mot))

# ---------------------------------------------------------------------------
# 9. ENCODAGE ET FINS DE LIGNE
#
# L'editeur VBA attend du Windows-1252 et des fins de ligne CRLF. Un fichier
# converti en UTF-8 deforme tous les accents a l'import ; des fins de ligne
# melangees font coller deux instructions.
# ---------------------------------------------------------------------------
for f in FICHIERS:
    brut = open(os.path.join(SRC, f), 'rb').read()
    try:
        brut.decode('cp1252')
    except UnicodeDecodeError as e:
        pb.append("%s n'est pas lisible en Windows-1252 (%s)" % (f, e))
        continue
    if b"\r\n" in brut and brut.replace(b"\r\n", b"").count(b"\n"):
        pb.append("%s melange les fins de ligne CRLF et LF" % f)
    elif b"\n" in brut and b"\r\n" not in brut:
        pb.append("%s est en fins de ligne LF : l'editeur VBA attend CRLF" % f)

# ---------------------------------------------------------------------------
# 10. L'APPLICATION NE VA PLUS CHERCHER LES DONNEES CHEZ ELLE
#
# Les donnees ont quitte le .xlsm pour FlecheNettoyageSA-AAAA.xlsx. Un
# ThisWorkbook.Worksheets oublie quelque part lirait donc le classeur de
# l'application -- qui ne contient plus aucun tableau. Pas d'erreur : un
# ListObject introuvable, une liste vide, un formulaire qui s'ouvre sur rien.
#
# Trois modules y ont droit, et eux seuls : les deux qui s'occupent de la
# feuille d'accueil, qui est bien dans l'application, et celui qui tient le
# classeur de donnees.
# ---------------------------------------------------------------------------
CHEZ_SOI = ("modAccueil_Generateur.bas", "modAccueil_Verrou.bas",
            "modDatas_Classeur.bas")
TW = re.compile(r'ThisWorkbook\s*\.\s*(?:Worksheets|Sheets|Names)\b', re.I)
for f in FICHIERS:
    if f in CHEZ_SOI:
        continue
    for no, l in lignes_logiques(rd(f)):
        if TW.search(sans_commentaire(l)):
            pb.append("%s:%d lit les feuilles ou les noms de ThisWorkbook : les "
                      "donnees sont dans le classeur de Datas_Classeur, celui-ci "
                      "n'en contient plus" % (f, no))

# ---------------------------------------------------------------------------
# 11. TOUTE ECRITURE DANS UN TABLEAU PASSE PAR LE GARDE-FOU
#
# Une annee passee s'ouvre en consultation, et l'annee en cours peut etre prise
# par un autre poste. Une fonction qui ecrirait sans demander Datas_PeutEcrire
# modifierait un classeur ouvert en lecture seule : Excel leve une 1004 au
# milieu de la saisie, ou -- pire -- l'ecriture passe dans une copie locale que
# OneDrive transformera en « fichier en conflit » que personne ne relira.
# ---------------------------------------------------------------------------
ECRIT = re.compile(r'(?:ListRows\s*\.\s*Add'
                   r'|ListRows\s*\([^)]*\)\s*\.\s*Delete'
                   r'|\.Range\s*\.\s*Value\s*=)', re.I)
for f in FICHIERS:
    proc, corps, debut = None, [], 0
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            proc, corps, debut = ms.group(1), [], no
        elif FIN_PROC.match(c):
            if proc and any(ECRIT.search(x) for x in corps) \
                    and not any("Datas_PeutEcrire" in x for x in corps):
                pb.append("%s:%d %s ecrit dans un tableau sans passer par "
                          "Datas_PeutEcrire : elle ecrirait dans des donnees "
                          "ouvertes en consultation seule" % (f, debut, proc))
            proc, corps = None, []
        elif proc:
            corps.append(c)

# ---------------------------------------------------------------------------
# 12. CHAQUE CELLULE NOMMEE EST LUE DU BON COTE
#
# Deux fichiers, donc deux lecteurs. Les confondre ne leve aucune erreur : le
# nom n'existe pas dans l'autre classeur, la fonction rend Empty, et l'objectif
# annuel vaut 0 ou le mot de passe devient vide -- ce dernier DEVERROUILLANT LE
# CLASSEUR SANS RIEN DEMANDER.
#
# Interv_CelluleNommee, qui lisait indifferemment l'un pour l'autre, n'existe
# plus : la citer est en soi le defaut.
# ---------------------------------------------------------------------------
COTE = {"CEL_MOT_DE_PASSE": "App_CelluleNommee",
        "CEL_DOSSIER_DATAS": "App_CelluleNommee",
        "CEL_TITRE": "App_CelluleNommee",
        "CEL_OBJECTIF": "Datas_CelluleNommee",
        "CEL_DATAS_VERSION": "Datas_CelluleNommee"}
LECTEURS = ("App_CelluleNommee", "Datas_CelluleNommee")
for f in FICHIERS:
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        if "Interv_CelluleNommee" in c:
            pb.append("%s:%d cite Interv_CelluleNommee : elle lisait les deux "
                      "classeurs sans distinguer, et n'existe plus -- choisir "
                      "App_CelluleNommee ou Datas_CelluleNommee" % (f, no))
            continue
        if re.match(r'\s*Public\s+Const\b', c):
            continue
        for cle, attendu in COTE.items():
            if not re.search(r'\b%s\b' % cle, c):
                continue
            autre = [x for x in LECTEURS if x != attendu][0]
            if autre in c:
                pb.append("%s:%d %s est lue par %s : elle est dans l'autre "
                          "classeur, et %s rendrait Empty sans rien dire"
                          % (f, no, cle, autre, autre))

# ---------------------------------------------------------------------------
# 13. CHANGER D'ANNEE PERIME TOUS LES CACHES
#
# Chaque module qui garde un tableau en memoire expose un invalidateur sans
# argument, dont le nom finit par Recharger. Datas_PerimerCaches les appelle
# tous ; en oublier un, c'est afficher les clients d'une annee au-dessus des
# interventions d'une autre -- sans erreur, et sans que rien ne se voie.
#
# Les procedures qui PRENNENT UN ARGUMENT ne sont pas des caches : elles
# rafraichissent un formulaire, qu'on leur passe.
# ---------------------------------------------------------------------------
mp = re.search(r'Sub Datas_PerimerCaches\(.*?\nEnd Sub', rd("modDatas_Classeur.bas"), re.S)
if not mp:
    pb.append("Datas_PerimerCaches a disparu : changer d'annee laisserait les "
              "caches sur le fichier precedent")
else:
    for f in FICHIERS:
        for no, l in lignes_logiques(rd(f)):
            c = sans_commentaire(l)
            ms = SIGNAT.match(c)
            if not ms or re.match(r'\s*Private\b', c, re.I):
                continue
            if not ms.group(1).lower().endswith("recharger"):
                continue
            if noms_parametres(c):
                continue
            if ms.group(1) not in mp.group(0):
                pb.append("%s:%d %s vide un cache mais Datas_PerimerCaches ne "
                          "l'appelle pas : apres un changement d'annee ce cache "
                          "rendrait les donnees de l'annee precedente"
                          % (f, no, ms.group(1)))

# ---------------------------------------------------------------------------
# 14. UNE VARIABLE DE MODULE EST DECLAREE DANS SON MODULE
#
# Option Explicit repond « Variable non definie », mais SEULEMENT a la
# compilation, module par module, et seulement une fois qu'on y arrive. Deplacer
# un etat d'un module a l'autre en laissant derriere soi la ligne qui l'affecte
# ne se voit donc pas en relisant le fichier -- c'est exactement ce qui est
# arrive a mChemin, parti dans modDatas_Classeur et reste dans Verrou_Rendre.
#
# Le classeur nomme ses variables de module mQuelqueChose : c'est a cette
# convention que le controle s'accroche. Un nom employe doit etre declare, au
# module ou dans la procedure qui l'emploie.
# ---------------------------------------------------------------------------
MVAR = re.compile(r'\bm[A-Z]\w*\b')
for f in FICHIERS:
    lignes = lignes_logiques(rd(f))

    # les declarations de module tiennent avant la premiere procedure
    module = set()
    for no, l in lignes:
        c = sans_commentaire(l)
        if SIGNAT.match(c):
            break
        module |= noms_module(c)

    # les locales de chaque procedure, AVANT de verifier : VBA accepte un Dim
    # place plus bas que le premier emploi
    ou, locales, proc = [], {}, None
    for no, l in lignes:
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            proc = no
            locales[proc] = noms_parametres(c)
        elif FIN_PROC.match(c):
            ou.append((no, c, proc))
            proc = None
            continue
        elif proc:
            locales[proc] |= noms_declares(c)
        ou.append((no, c, proc))

    for no, c, proc in ou:
        connus = module | locales.get(proc, set())
        for n in MVAR.findall(sans_chaine(c)):
            if n not in connus:
                pb.append("%s:%d %s n'est declaree ni au module ni dans la "
                          "procedure : « Variable non definie » a la compilation"
                          % (f, no, n))

# ---------------------------------------------------------------------------
# 15. UN FICHIER DE DONNEES NE SE PRESENTE JAMAIS EN KIOSQUE
#
# Trois des reglages du kiosque sont enregistres DANS le classeur : les feuilles
# tres masquees, les proprietes de la FENETRE, et la protection de structure. Le
# premier fichier de donnees, fabrique en copiant l'application verrouillee, les
# a donc emportes : ouvert a la main dans Excel, il ne montrait aucun onglet et
# refusait la moitie du ruban.
#
# Toute propriete de fenetre que PoserInterface eteint doit etre rallumee par
# NormaliserDonnees : ce fichier doit rester consultable SANS l'application.
# ---------------------------------------------------------------------------
datas = rd("modDatas_Classeur.bas")
mn = re.search(r'Sub NormaliserDonnees\(.*?\nEnd Sub', datas, re.S)
mp = re.search(r'Sub PoserInterface\(.*?\nEnd Sub', rd("modAccueil_Verrou.bas"), re.S)
corps_n = code_seul(mn.group(0)) if mn else ""
corps_p = code_seul(mp.group(0)) if mp else ""

if not mn:
    pb.append("NormaliserDonnees a disparu : un fichier de donnees garderait les "
              "reglages du kiosque, et s'ouvrirait sans onglets dans Excel")
elif len(re.findall(r'\bNormaliserDonnees\b', code_seul(datas))) < 2:
    pb.append("NormaliserDonnees n'est appelee par personne : elle ne sert a rien")
elif not mp:
    pb.append("PoserInterface est introuvable : impossible de verifier ce que "
              "NormaliserDonnees doit rallumer")
else:
    # celles posees sur la FENETRE, les seules qui s'enregistrent dans le fichier
    fen = sorted(set(re.findall(r'(?m)^\s*fen\.(Display\w+)\s*=', corps_p)))
    if not fen:
        pb.append("aucune propriete de fenetre reconnue dans PoserInterface : la "
                  "variable a du changer de nom, et le controle ne verifie plus rien")
    for prop in fen:
        if prop not in corps_n:
            pb.append("le kiosque eteint %s, que NormaliserDonnees ne rallume pas : "
                      "le fichier de donnees garderait ce reglage, et s'ouvrirait "
                      "nu dans Excel" % prop)
    for quoi, mal in (("Unprotect", "sa structure resterait protegee, ruban grise"),
                      ("xlSheetVisible", "ses feuilles resteraient tres masquees, "
                                         "et le menu Afficher ne les propose pas")):
        if quoi not in corps_n:
            pb.append("NormaliserDonnees ne fait rien de %s : %s" % (quoi, mal))

# ---------------------------------------------------------------------------
# 16. UN CLASSEUR NE S'ENREGISTRE JAMAIS FENETRE MASQUEE
#
# Window.Visible = False s'ecrit DANS le fichier, en clair :
#
#     <workbookView visibility="hidden" ... showSheetTabs="0" ...>
#
# Un classeur enregistre ainsi s'ouvre ensuite SANS AUCUNE FENETRE. Excel montre
# son fond gris, aucun onglet, et grise la quasi-totalite du ruban faute de
# classeur actif. Le fichier parait mort, et rien dans Excel ne dit pourquoi.
#
# On masque pourtant la fenetre du classeur de donnees, le temps de la seance,
# pour qu'elle ne tombe pas par-dessus le kiosque. Tout enregistrement doit donc
# etre precede de MontrerFenetre -- et ailleurs que dans modDatas_Classeur, un
# enregistrement du classeur de donnees n'a rien a faire.
# ---------------------------------------------------------------------------
COTE_APP = ("modAccueil_Verrou.bas", "modAccueil_Generateur.bas")
SAUVE = re.compile(r'\.Save\b')
datas_v = rd("modDatas_Classeur.bas")

for quoi in ("MasquerFenetre", "MontrerFenetre"):
    if not re.search(r'Sub %s\(' % quoi, datas_v):
        pb.append("%s a disparu de modDatas_Classeur : la fenetre du classeur de "
                  "donnees n'est plus pilotee, et il s'enregistrera masque" % quoi)

mnorm = re.search(r'Sub NormaliserDonnees\(.*?\nEnd Sub', datas_v, re.S)
if mnorm and "MontrerFenetre" not in code_seul(mnorm.group(0)):
    pb.append("NormaliserDonnees ne rend pas sa fenetre au classeur : un fichier "
              "deja enregistre masque resterait invisible a l'ouverture")

for f in FICHIERS:
    if f in COTE_APP:
        continue                      # ThisWorkbook.Save : l'application, jamais masquee
    proc, vu, debut = None, False, 0
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            proc, vu, debut = ms.group(1), False, no
            continue
        if FIN_PROC.match(c):
            proc = None
            continue
        if not proc:
            continue
        if "MontrerFenetre" in c:
            vu = True
        if SAUVE.search(c):
            if f != "modDatas_Classeur.bas":
                pb.append("%s:%d %s enregistre un classeur : seul "
                          "modDatas_Classeur sait que la fenetre des donnees est "
                          "masquee, et qu'il faut la rendre avant" % (f, no, proc))
            elif not vu:
                pb.append("%s:%d %s enregistre sans avoir appele MontrerFenetre : "
                          "le fichier garderait sa fenetre masquee et s'ouvrirait "
                          "vide, ruban grise" % (f, no, proc))

# ---------------------------------------------------------------------------
# 17. LE VERROU SE REND DANS TOUS LES CAS
#
# Un verrou qui ne se rend pas bloque les donnees QUATRE HEURES, et le poste
# bloque est d'abord le notre : au demarrage suivant, l'application demande a
# l'utilisateur s'il veut « prendre la main » sur lui-meme.
#
# Deux defauts l'ont provoque, et chacun a sa ligne ici :
#
#   Datas_Fermer sortait AVANT de rendre le verrou quand elle ne trouvait pas de
#   classeur a fermer -- or c'est exactement le cas ou le verrou reste ;
#
#   Verrou_Rendre n'effacait le fichier que si mMarque correspondait, et mMarque
#   est une variable de module : l'etat du projet remis a zero, plus personne ne
#   pouvait prouver que le verrou etait le notre. D'ou EstDeCePoste, qui le
#   reconnait a l'utilisateur et au poste.
# ---------------------------------------------------------------------------
verrou_v = rd("modDatas_Verrou.bas")
mf = re.search(r'Sub Datas_Fermer\(.*?\nEnd Sub', rd("modDatas_Classeur.bas"), re.S)
if not mf:
    pb.append("Datas_Fermer a disparu")
else:
    lf = code_seul(mf.group(0)).split("\n")
    rendu = next((i for i, l in enumerate(lf) if "Verrou_Rendre" in l), -1)
    if rendu < 0:
        pb.append("Datas_Fermer ne rend plus le verrou : il bloquerait les "
                  "donnees quatre heures, en commencant par ce poste-ci")
    else:
        for i, l in enumerate(lf[:rendu]):
            if re.match(r'\s*Exit\s+Sub\b', l, re.I):
                pb.append("modDatas_Classeur.bas Datas_Fermer sort (Exit Sub) AVANT "
                          "de rendre le verrou : sur ce chemin-la le verrou reste, "
                          "et c'est justement celui ou il n'y a plus de classeur a "
                          "fermer")
                break

if not re.search(r'Function EstDeCePoste\(', verrou_v):
    pb.append("EstDeCePoste a disparu : un verrou laisse par une seance precedente "
              "de ce poste ne serait plus reconnu, et bloquerait quatre heures")
else:
    for proc in ("Verrou_Rendre", "Saisir"):
        mb = re.search(r'(?:Sub|Function) %s\(.*?\nEnd (?:Sub|Function)' % proc,
                       verrou_v, re.S)
        if mb and "EstDeCePoste" not in code_seul(mb.group(0)):
            pb.append("%s ne consulte pas EstDeCePoste : un verrou laisse par une "
                      "seance precedente de ce poste y resterait %s"
                      % (proc, "sur le disque" if proc == "Verrou_Rendre"
                         else "et ferait demander de prendre la main sur soi-meme"))

mr = re.search(r'Sub Verrou_Rendre\(.*?\nEnd Sub', verrou_v, re.S)
if mr and not re.search(r'v\s*=\s*mVerrou\b', code_seul(mr.group(0))):
    pb.append("Verrou_Rendre ne retombe pas sur mVerrou : appelee sans chemin -- ce "
              "qui arrive quand l'appelant a perdu le sien -- elle ne rendrait rien")

# ---------------------------------------------------------------------------
# 18. LE KIOSQUE NE MARQUE PAS LE CLASSEUR MODIFIE
#
# Masquer des feuilles, proteger la structure, retirer le ruban, reecrire le
# bandeau : chacun de ces gestes touche le classeur ou sa fenetre, et Excel le
# marque aussitot « modifie ». En quittant, il demandait donc A CHAQUE FOIS s'il
# fallait l'enregistrer -- pour un ruban, dans un fichier qui ne contient plus
# aucune donnee.
#
# Et la question n'etait pas que du bruit : repondue « oui » par habitude, elle
# finit par graver dans un fichier un etat qu'on ne voulait pas. C'est ainsi que
# le classeur de donnees s'est retrouve avec sa fenetre masquee. Une question de
# trop apprend a cliquer sans lire.
#
# Toute procedure qui touche l'interface doit donc remettre le drapeau ou il
# etait, APRES l'avoir touchee -- le remettre avant ne servirait a rien.
# ---------------------------------------------------------------------------
POSE = re.compile(r'\bPoserInterface\b')
REMET = re.compile(r'\bAppli_PoserModifiee\b')

for quoi in ("Appli_Modifiee", "Appli_PoserModifiee"):
    if not re.search(r'Function %s\(|Sub %s\(' % (quoi, quoi), rd("modAccueil_Verrou.bas")):
        pb.append("%s a disparu : le kiosque remarquerait le classeur modifie, et "
                  "Excel demanderait a chaque fermeture s'il faut l'enregistrer"
                  % quoi)

for f in FICHIERS:
    proc, lignes_p, debut = None, [], 0
    for no, l in lignes_logiques(rd(f)):
        c = sans_commentaire(l)
        ms = SIGNAT.match(c)
        if ms:
            proc, lignes_p, debut = ms.group(1), [], no
            continue
        if FIN_PROC.match(c):
            if proc and proc != "PoserInterface":
                pose = [i for i, x in enumerate(lignes_p) if POSE.search(x)]
                remet = [i for i, x in enumerate(lignes_p) if REMET.search(x)]
                if pose and not remet:
                    pb.append("%s:%d %s touche l'interface sans remettre le drapeau "
                              "de modification : Excel demandera s'il faut "
                              "enregistrer, pour un ruban" % (f, debut, proc))
                elif pose and remet and max(remet) < max(pose):
                    pb.append("%s:%d %s remet le drapeau de modification AVANT de "
                              "toucher l'interface : celle-ci le remarque modifie "
                              "juste apres" % (f, debut, proc))
            proc, lignes_p = None, []
            continue
        if proc:
            lignes_p.append(c)

mb = re.search(r'Sub Accueil_MajBandeau\(.*?\nEnd Sub', rd("modAccueil_Generateur.bas"), re.S)
if not mb:
    pb.append("Accueil_MajBandeau a disparu")
elif "Appli_PoserModifiee" not in code_seul(mb.group(0)):
    pb.append("Accueil_MajBandeau ne remet pas le drapeau de modification : elle "
              "leve la protection de la feuille et y pose du texte a chaque "
              "demarrage, ce qui suffit a faire poser la question")

# ---------------------------------------------------------------------------
print("%d modules, %d procedures publiques"
      % (len(FICHIERS), len(PUB_OU)))
if pb:
    print("\n--- PROBLEMES ---")
    for p in pb:
        print("  ", p)
    sys.exit(1)
print("\nAucun probleme structurel detecte.")
