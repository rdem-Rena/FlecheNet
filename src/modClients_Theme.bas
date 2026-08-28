Attribute VB_Name = "modClients_Theme"
Option Explicit
'==============================================================================
' modClients_Theme
'------------------------------------------------------------------------------
' Charte graphique et géométrie du formulaire UF_Clients.
' Tout l'aspect visuel du formulaire est pilote depuis ce module : modifier une
' constante ici puis relancer GenererFormulaireClients suffit à re-styler
' l'ensemble du formulaire.
'
' Les couleurs sont notées en hexadécimal VBA, c'est-à-dire &HBBGGRR& (bleu,
' vert, rouge) : le commentaire en fin de ligne rappelle la notation web #RRGGBB.
'==============================================================================

'--- Constantes MSForms redéfinies localement ---------------------------------
' Elles évitent toute dépendance de compilation à la bibliothèque MSForms tant
' que le UserForm n'a pas encore été généré.
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
Public Const MSF_MatchEntryFirstLetter As Long = 0
Public Const MSF_MatchEntryComplete As Long = 1
Public Const MSF_MultiSelectSingle As Long = 0
Public Const MSF_ScrollBarsNone As Long = 0
Public Const MSF_ScrollBarsVertical As Long = 2
Public Const MSF_ListStylePlain As Long = 0
Public Const MSF_ZOrderFront As Long = 0
Public Const MSF_ZOrderBack As Long = 1

'--- Palette ------------------------------------------------------------------
Public Const COUL_FOND As Long = &HF8F3F0&          ' #F0F3F8  fond du formulaire
Public Const COUL_CARTE As Long = &HFFFFFF&         ' #FFFFFF  fond des cartes
Public Const COUL_BORDURE As Long = &HEEE6E0&       ' #E0E6EE  filet des cartes

Public Const COUL_BANDEAU As Long = &H5D361B&       ' #1B365D  bandeau de titre
Public Const COUL_BANDEAU_TXT As Long = &HFFFFFF&   ' #FFFFFF
Public Const COUL_BANDEAU_SOUS As Long = &HCDAF96&  ' #96AFCD

Public Const COUL_TEXTE As Long = &H3E2D20&         ' #202D3E  texte principal
Public Const COUL_TEXTE_DOUX As Long = &HA68C7A&    ' #7A8CA6  libellés de champs
Public Const COUL_SECTION As Long = &HA68C7A&       ' #7A8CA6  titres de section

Public Const COUL_CHAMP_FOND As Long = &HFFFDFC&    ' #FCFDFF  fond des saisies
Public Const COUL_CHAMP_BORD As Long = &HE9DED6&    ' #D6DEE9  filet des saisies
Public Const COUL_CHAMP_FOCUS As Long = &HB56917&   ' #1769B5  filet au focus
Public Const COUL_VERROU_FOND As Long = &HF7F2EE&   ' #EEF2F7  champ géré par le programme
Public Const COUL_VERROU_TXT As Long = &H99887C&    ' #7C8899

Public Const COUL_ENTETE_TBL As Long = &HF6EEE9&    ' #E9EEF6  en-tête du tableau
Public Const COUL_ENTETE_TXT As Long = &H6C5644&    ' #44566C

Public Const COUL_AJOUTER As Long = &H658F00&       ' #008F65  vert
Public Const COUL_AJOUTER_H As Long = &H76A800&     ' #00A876
Public Const COUL_MODIFIER As Long = &HB56917&      ' #1769B5  bleu
Public Const COUL_MODIFIER_H As Long = &HD2802A&    ' #2A80D2
Public Const COUL_SUPPRIMER As Long = &H3636C7&     ' #C73636  rouge
Public Const COUL_SUPPRIMER_H As Long = &H4B4BDC&   ' #DC4B4B
Public Const COUL_EFFACER As Long = &H8C7A6C&       ' #6C7A8C  gris
Public Const COUL_EFFACER_H As Long = &HA89484&     ' #8494A8
Public Const COUL_QUITTER As Long = &H5A483A&       ' #3A485A  anthracite
Public Const COUL_QUITTER_H As Long = &H755F4E&     ' #4E5F75
Public Const COUL_BOUTON_TXT As Long = &HFFFFFF&    ' #FFFFFF
Public Const COUL_BOUTON_OFF As Long = &HD9CFC6&    ' #C6CFD9  bouton désactivé

Public Const COUL_FERMER_H As Long = &H4B4BE0&      ' #E04B4B  survol de la croix
Public Const COUL_LIEN As Long = &HB56917&          ' #1769B5  liens cliquables
Public Const COUL_LIEN_H As Long = &HD2802A&        ' #2A80D2

'--- Typographie --------------------------------------------------------------
Public Const POLICE As String = "Segoe UI"
Public Const TAILLE_TITRE As Single = 12
Public Const TAILLE_SOUSTITRE As Single = 8
Public Const TAILLE_SECTION As Single = 7.5
Public Const TAILLE_LIBELLE As Single = 7.5
Public Const TAILLE_CHAMP As Single = 9.5
Public Const TAILLE_BOUTON As Single = 9.5
Public Const TAILLE_ENTETE As Single = 7.5
Public Const TAILLE_LISTE As Single = 9
Public Const TAILLE_FILTRE As Single = 8.5

'--- Géométrie générale (en points) -------------------------------------------
Public Const F_LARGEUR As Single = 840
Public Const F_HAUTEUR As Single = 532
Public Const MARGE As Single = 16
Public Const CARTE_LARG As Single = 808

' Bandeau de titre
Public Const BAND_HAUT As Single = 48

' Carte "Fiche client" : grille de saisie 4 colonnes x 5 lignes
Public Const CS_TOP As Single = 58
Public Const CS_HAUT As Single = 200
Public Const GR_X As Single = 28            ' abscisse de la 1re colonne
Public Const GR_Y As Single = 86            ' ordonnée de la 1re ligne
Public Const GR_BLOC As Single = 184        ' largeur d'un bloc "libellé + champ"
Public Const GR_GOUTTIERE As Single = 16    ' espace entre deux blocs
Public Const GR_LIGNE As Single = 32        ' pas vertical entre deux lignes
Public Const CH_LBL_HAUT As Single = 11     ' hauteur du libellé
Public Const CH_CTL_HAUT As Single = 18     ' hauteur de la zone de saisie

' Carte de filtrage
Public Const CF_TOP As Single = 266
Public Const CF_HAUT As Single = 40

' Carte tableau des enregistrements
Public Const CT_TOP As Single = 314
Public Const CT_HAUT As Single = 168
Public Const CT_ENTETE As Single = 22

' Barre de boutons
Public Const BT_TOP As Single = 492
Public Const BT_HAUT As Single = 30
Public Const BT_LARG As Single = 118
Public Const BT_GOUTTIERE As Single = 6

'==============================================================================
' Position d'un bloc de la grille de saisie
'==============================================================================
'------------------------------------------------------------------------------
' Abscisse du bloc « libellé + zone de saisie » d'une colonne de la grille.
'   colonne : 1 à 4, de gauche à droite
'   renvoie : la position en points, mesurée depuis le bord gauche du formulaire
'------------------------------------------------------------------------------
Public Function GrilleX(ByVal colonne As Long) As Single
    GrilleX = GR_X + (colonne - 1) * (GR_BLOC + GR_GOUTTIERE)
End Function

'------------------------------------------------------------------------------
' Ordonnée du bloc « libellé + zone de saisie » d'une ligne de la grille.
'   ligne   : 1 à 5, de haut en bas
'   renvoie : la position en points, mesurée depuis le haut du formulaire
' Le libellé occupe CH_LBL_HAUT points, la zone de saisie commence juste dessous.
'------------------------------------------------------------------------------
Public Function GrilleY(ByVal ligne As Long) As Single
    GrilleY = GR_Y + (ligne - 1) * GR_LIGNE
End Function

'==============================================================================
' Couleur de fond d'un bouton d'action (état normal ou survol)
'==============================================================================
'------------------------------------------------------------------------------
' Couleur de fond d'un bouton d'action.
'   nomBouton : btnAjouter, btnModifier, btnSupprimer, btnEffacer ou btnQuitter
'   survol    : True pour la teinte claire appliquée au passage de la souris
' Utilisée à la génération (état normal) et à chaque MouseMove (survol).
'------------------------------------------------------------------------------
Public Function CouleurBouton(ByVal nomBouton As String, ByVal survol As Boolean) As Long
    Select Case nomBouton
        Case "btnAjouter":   CouleurBouton = IIf(survol, COUL_AJOUTER_H, COUL_AJOUTER)
        Case "btnModifier":  CouleurBouton = IIf(survol, COUL_MODIFIER_H, COUL_MODIFIER)
        Case "btnSupprimer": CouleurBouton = IIf(survol, COUL_SUPPRIMER_H, COUL_SUPPRIMER)
        Case "btnEffacer":   CouleurBouton = IIf(survol, COUL_EFFACER_H, COUL_EFFACER)
        Case "btnQuitter":   CouleurBouton = IIf(survol, COUL_QUITTER_H, COUL_QUITTER)
        Case Else:           CouleurBouton = IIf(survol, COUL_EFFACER_H, COUL_EFFACER)
    End Select
End Function

'==============================================================================
' Glyphes
'==============================================================================
'------------------------------------------------------------------------------
' Croix de fermeture du bandeau.
' Construite avec ChrW plutôt qu'écrite en clair : le caractère reste correct
' quel que soit l'encodage sous lequel ce fichier est ré-enregistré.
'------------------------------------------------------------------------------
Public Function GlypheFermer() As String
    GlypheFermer = ChrW(215)                ' croix de fermeture
End Function

'------------------------------------------------------------------------------
' Triangle vers le haut, accolé à l'en-tête de la colonne triée.
'------------------------------------------------------------------------------
Public Function GlypheTriCroissant() As String
    GlypheTriCroissant = " " & ChrW(9650)   ' triangle vers le haut
End Function

'------------------------------------------------------------------------------
' Triangle vers le bas, accolé à l'en-tête de la colonne triée.
'------------------------------------------------------------------------------
Public Function GlypheTriDecroissant() As String
    GlypheTriDecroissant = " " & ChrW(9660) ' triangle vers le bas
End Function
