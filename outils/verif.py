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

def noms_declares(ligne):
    """une declaration porte souvent PLUSIEURS noms : Dim a As X, b As Y"""
    md = DECL.match(ligne)
    if not md:
        return set()
    out, prof, cur = set(), 0, ""
    for ch in md.group(1) + ",":
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
print("%d modules, %d procedures publiques"
      % (len(FICHIERS), len(PUB_OU)))
if pb:
    print("\n--- PROBLEMES ---")
    for p in pb:
        print("  ", p)
    sys.exit(1)
print("\nAucun probleme structurel detecte.")
