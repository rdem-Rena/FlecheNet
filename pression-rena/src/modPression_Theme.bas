Attribute VB_Name = "modPression_Theme"
Option Explicit
'==============================================================================
' modPression_Theme
'------------------------------------------------------------------------------
' Charte graphique et géométrie du formulaire UF_Pression.
'
' Ce classeur est INDÉPENDANT de FlècheNet : il ne partage ni ses modules ni
' ses feuilles, et ce fichier redit donc ce que modClients_Theme dit de son
' côté. La palette est la même, à dessein — les deux classeurs sont de la même
' main — mais elle est répétée plutôt qu'importée, sans quoi ce classeur ne
' fonctionnerait qu'à côté de l'autre.
'
' Les couleurs s'écrivent en hexadécimal VBA, &HBBGGRR& ; le commentaire de fin
' de ligne rappelle la notation web #RRGGBB.
'==============================================================================

'--- Constantes MSForms redéfinies localement ---------------------------------
' Elles évitent toute dépendance de compilation à la bibliothèque MSForms tant
' que le formulaire n'a pas encore été généré.
Public Const MSF_BorderStyleNone As Long = 0
Public Const MSF_BorderStyleSingle As Long = 1
Public Const MSF_SpecialEffectFlat As Long = 0
Public Const MSF_BackStyleTransparent As Long = 0
Public Const MSF_BackStyleOpaque As Long = 1
Public Const MSF_TextAlignLeft As Long = 1
Public Const MSF_TextAlignCenter As Long = 2
Public Const MSF_TextAlignRight As Long = 3
Public Const MSF_StyleDropDownCombo As Long = 0
Public Const MSF_StyleDropDownList As Long = 2
Public Const MSF_ScrollBarsNone As Long = 0

'--- Palette ------------------------------------------------------------------
Public Const PCOUL_FOND As Long = &HF8F3F0&          ' #F0F3F8  fond du formulaire
Public Const PCOUL_CARTE As Long = &HFFFFFF&         ' #FFFFFF  fond des cartes
Public Const PCOUL_BORDURE As Long = &HEEE6E0&       ' #E0E6EE  filet des cartes

Public Const PCOUL_BANDEAU As Long = &H5D361B&       ' #1B365D  bandeau de titre
Public Const PCOUL_BANDEAU_TXT As Long = &HFFFFFF&   ' #FFFFFF
Public Const PCOUL_BANDEAU_SOUS As Long = &HCDAF96&  ' #96AFCD

Public Const PCOUL_TEXTE As Long = &H3E2D20&         ' #202D3E  texte principal
Public Const PCOUL_TEXTE_DOUX As Long = &HA68C7A&    ' #7A8CA6  libellés de champs
Public Const PCOUL_SECTION As Long = &HA68C7A&       ' #7A8CA6  titres de section

Public Const PCOUL_CHAMP_FOND As Long = &HFFFDFC&    ' #FCFDFF  fond des saisies
Public Const PCOUL_CHAMP_BORD As Long = &HE9DED6&    ' #D6DEE9  filet des saisies
Public Const PCOUL_VERROU_FOND As Long = &HF7F2EE&   ' #EEF2F7  champ géré par le programme
Public Const PCOUL_VERROU_TXT As Long = &H99887C&    ' #7C8899

Public Const PCOUL_ENTETE_TBL As Long = &HF6EEE9&    ' #E9EEF6  en-tête du tableau
Public Const PCOUL_ENTETE_TXT As Long = &H6C5644&    ' #44566C

Public Const PCOUL_ENREG As Long = &H658F00&         ' #008F65  vert
Public Const PCOUL_EFFACER As Long = &H8C7A6C&       ' #6C7A8C  gris
Public Const PCOUL_QUITTER As Long = &H5A483A&       ' #3A485A  anthracite
Public Const PCOUL_AUJOURD As Long = &HB56917&       ' #1769B5  bleu
Public Const PCOUL_BOUTON_TXT As Long = &HFFFFFF&    ' #FFFFFF

' Les deux teintes des bandes du graphique lissé, reprises ici pour l'état de
' la mesure saisie : au-dessus du seuil, en dessous.
Public Const PCOUL_ALERTE As Long = &H3636C7&        ' #C73636  rouge
Public Const PCOUL_BON As Long = &H658F00&           ' #008F65  vert

'--- Typographie --------------------------------------------------------------
Public Const PPOLICE As String = "Segoe UI"
Public Const PT_PAR_PIXEL As Single = 0.75

Public Type StyleTexte
    Police As String
    Taille As Single
    Gras As Boolean
    Couleur As Long
End Type

'==============================================================================
' GÉOMÉTRIE (en points)
'==============================================================================
Public Const P_LARGEUR As Single = 620
Public Const P_HAUTEUR As Single = 412           ' surface UTILE, hors barre de titre
Public Const P_RESERVE_TITRE As Single = 24      ' Height est extérieure, pas la surface
Public Const P_MARGE As Single = 14
Public Const P_CARTE_LARG As Single = 592        ' P_LARGEUR - 2 * P_MARGE

'--- Zone 1 : le bandeau ------------------------------------------------------
Public Const P_Z1_TOP As Single = 12
Public Const P_Z1_HAUT As Single = 44

'--- Zone 2 : la fiche de saisie ----------------------------------------------
Public Const P_Z2_TOP As Single = 64
Public Const P_Z2_HAUT As Single = 176

Public Const P_NB_BLOCS As Long = 2
Public Const P_NB_LIGNES As Long = 4
Public Const P_GR_X As Single = 18               ' abscisse du 1er bloc, dans la carte
Public Const P_GR_Y As Single = 24               ' ordonnée de la 1re ligne
Public Const P_GR_BLOC As Single = 268           ' largeur d'un bloc
Public Const P_GR_GOUTTIERE As Single = 20       ' espace entre les deux blocs
Public Const P_GR_LIGNE As Single = 36           ' pas vertical entre deux lignes
Public Const P_LBL_HAUT As Single = 12           ' hauteur du libellé
Public Const P_CTL_HAUT As Single = 20           ' hauteur de la zone de saisie
Public Const P_ENTRE_CHAMPS As Single = 8        ' entre deux champs d'une ligne

'--- Zone 3 : les dernières mesures -------------------------------------------
Public Const P_Z3_TOP As Single = 248
Public Const P_Z3_HAUT As Single = 118
Public Const P_NB_DERNIERES As Long = 6
Public Const P_LIGNE_H As Single = 12.75         ' 17 pixels
Public Const P_PAD_X As Single = 4.5
Public Const P_TITRE_HAUT As Single = 15
Public Const P_ENTETE_HAUT As Single = 17

'--- Barre de boutons ---------------------------------------------------------
Public Const P_BT_TOP As Single = 374
Public Const P_BT_HAUT As Single = 28
Public Const P_BT_LARG As Single = 108
Public Const P_BT_GOUTTIERE As Single = 8

'==============================================================================
' CALAGE AU PIXEL
'------------------------------------------------------------------------------
' Un point ne vaut pas un nombre entier de pixels : à 96 ppp, 1 px = 0,75 pt.
' Une coordonnée tombée entre deux pixels fait rendre le texte à cheval, et
' Windows le dessine alors plus épais et légèrement décalé — d'une colonne à
' l'autre, sans régularité apparente. Toutes les coordonnées du tableau
' passent donc par ici.
'==============================================================================
Public Function AuPixel(ByVal v As Single) As Single
    AuPixel = Int(v / PT_PAR_PIXEL + 0.5) * PT_PAR_PIXEL
End Function

'------------------------------------------------------------------------------
' Abscisse du bord gauche d'un bloc de la fiche (1 ou 2, de gauche à droite).
'------------------------------------------------------------------------------
Public Function PGrilleX(ByVal bloc As Long) As Single
    PGrilleX = P_GR_X + (bloc - 1) * (P_GR_BLOC + P_GR_GOUTTIERE)
End Function

'------------------------------------------------------------------------------
' Ordonnée du couple « libellé + zone de saisie » d'une ligne (1 à P_NB_LIGNES).
'------------------------------------------------------------------------------
Public Function PGrilleY(ByVal ligne As Long) As Single
    PGrilleY = P_GR_Y + (ligne - 1) * P_GR_LIGNE
End Function

'------------------------------------------------------------------------------
' Étages du tableau des dernières mesures.
'------------------------------------------------------------------------------
Public Function PGrilleEnteteY() As Single
    PGrilleEnteteY = AuPixel(P_TITRE_HAUT)
End Function

Public Function PGrilleLignesY() As Single
    PGrilleLignesY = AuPixel(P_TITRE_HAUT + P_ENTETE_HAUT)
End Function

'==============================================================================
' LES COLONNES DU TABLEAU DES DERNIÈRES MESURES
'------------------------------------------------------------------------------
' Les quatre tableaux se lisent position par position et doivent rester de
' même longueur. Leur somme vaut 568 points, pour 590 disponibles.
'==============================================================================
Public Function PDernieresColonnes() As Variant
    PDernieresColonnes = Array(PC_DATE, PC_DEMI, PC_SYS, PC_DIA, PC_POULS, _
                               PC_POIDS, PC_ACTIVITE, PC_COMMENT)
End Function

Public Function PDernieresLibelles() As Variant
    PDernieresLibelles = Array("Date", "Demi", "Sys.", "Dia.", "Pouls", _
                               "Poids", "Activit" & ChrW(233), "Commentaire")
End Function

Public Function PDernieresLargeurs() As Variant
    PDernieresLargeurs = Array(74, 46, 40, 40, 44, 46, 110, 168)
End Function

Public Function PDernieresAlignements() As Variant
    PDernieresAlignements = Array(MSF_TextAlignLeft, MSF_TextAlignLeft, _
                                  MSF_TextAlignRight, MSF_TextAlignRight, _
                                  MSF_TextAlignRight, MSF_TextAlignRight, _
                                  MSF_TextAlignLeft, MSF_TextAlignLeft)
End Function

'==============================================================================
' TYPOGRAPHIE PAR RÔLE
'------------------------------------------------------------------------------
' Réunir les quatre caractéristiques d'un texte dans un seul objet les fait
' poser d'un bloc : impossible d'en changer une en oubliant les autres.
'==============================================================================
Private Function Style(ByVal police As String, ByVal taille As Single, _
                       ByVal gras As Boolean, ByVal couleur As Long) As StyleTexte
    Style.Police = police
    Style.Taille = taille
    Style.Gras = gras
    Style.Couleur = couleur
End Function

Public Function PSTitre() As StyleTexte
    PSTitre = Style(PPOLICE, 12, True, PCOUL_BANDEAU_TXT)
End Function

Public Function PSSousTitre() As StyleTexte
    PSSousTitre = Style(PPOLICE, 8, False, PCOUL_BANDEAU_SOUS)
End Function

Public Function PSSection() As StyleTexte
    PSSection = Style(PPOLICE, 7.5, True, PCOUL_SECTION)
End Function

Public Function PSLibelle() As StyleTexte
    PSLibelle = Style(PPOLICE, 7.5, False, PCOUL_TEXTE_DOUX)
End Function

Public Function PSEntete() As StyleTexte
    PSEntete = Style(PPOLICE, 7.5, True, PCOUL_ENTETE_TXT)
End Function

Public Function PSCase() As StyleTexte
    PSCase = Style(PPOLICE, 9, False, PCOUL_TEXTE)
End Function

Public Function PSEtat() As StyleTexte
    PSEtat = Style(PPOLICE, 9, True, PCOUL_TEXTE)
End Function
