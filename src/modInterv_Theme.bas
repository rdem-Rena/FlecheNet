Attribute VB_Name = "modInterv_Theme"
Option Explicit
'==============================================================================
' modInterv_Theme
'------------------------------------------------------------------------------
' Géométrie du formulaire UF_Interventions.
'
' La PALETTE et la TYPOGRAPHIE ne sont pas redéfinies ici : elles viennent de
' modClients_Theme, qui porte la charte de tout le classeur. Une couleur changée
' là-bas se répercute donc sur les deux formulaires. Ce module ne contient que
' ce qui est propre à la mise en page des interventions.
'
' modClients_Theme doit être présent dans le projet pour que ce module compile.
'==============================================================================

'--- Couleurs propres à ce formulaire -----------------------------------------
' Les deux boutons qui ouvrent un autre formulaire ont leur teinte à eux, pour
' se distinguer du groupe de saisie. Notation &HBBGGRR&, comme dans
' modClients_Theme : le commentaire rappelle l'équivalent web #RRGGBB.
Public Const COUL_FACTURER As Long = &HB79C2&      ' #C2790B  ambre
Public Const COUL_FACTURER_H As Long = &H1E8FDE&   ' #DE8F1E
Public Const COUL_INFO As Long = &H867C0E&         ' #0E7C86  bleu-vert
Public Const COUL_INFO_H As Long = &HA49812&       ' #1298A4

'--- Graphique et calendrier --------------------------------------------------
' Une seule série : une seule teinte, celle qui sert déjà d'accent au formulaire.
' Vérifiée sur la surface du graphique — bande de luminosité, chroma et contrast
' au-dessus de 3:1 — donc lisible sans dépendre de la couleur seule.
Public Const COUL_GR_BARRE As Long = &HB56917&     ' #1769B5  barres
Public Const COUL_GR_GRILLE As Long = &HEEE6E0&    ' #E0E6EE  lignes de repère
Public Const COUL_GR_BASE As Long = &HE9DED6&      ' #D6DEE9  ligne de base
Public Const COUL_CAL_WEEKEND As Long = &HF8F3F0&  ' #F0F3F8  colonnes samedi et dimanche
Public Const COUL_CAL_SURVOL As Long = &HF6EEE9&   ' #E9EEF6  case survolée

'--- Fenêtre ------------------------------------------------------------------
' Dimensions de la SURFACE UTILE, barre de titre exclue. Le formulaire est
' agrandi à l'ouverture pour que son intérieur mesure exactement cela : Width et
' Height d'un UserForm désignent l'extérieur, barre de titre comprise, et le bas
' du formulaire serait rogné sans cet ajustement.
Public Const I_LARGEUR As Single = 960
Public Const I_HAUTEUR As Single = 748

Public Const I_MARGE As Single = 16
Public Const I_CARTE_LARG As Single = 928     ' I_LARGEUR - 2 * I_MARGE

'--- Fiche 1 : intitulé -------------------------------------------------------
Public Const F1_TOP As Single = 12
Public Const F1_HAUT As Single = 46

'--- Fiche 2 : statistiques ---------------------------------------------------
Public Const F2_TOP As Single = 66
Public Const F2_HAUT As Single = 152
' Le graphique et le bloc de tuiles se partagent la largeur de la carte : ce que
' l'un prend, l'autre le perd. Les tuiles sont calées sur le bord droit, le
' graphique occupe tout ce qui reste à gauche.
Public Const F2_GRAPH_LARG As Single = 518    ' largeur de la zone du graphique

'--- Graphique dessiné --------------------------------------------------------
' Le graphique est tracé en contrôles MSForms plutôt qu'importé en image : net à
' toute taille, aux couleurs du formulaire, et sans fichier temporaire.
'
' GR_ORIGINE_* donne son coin haut-gauche dans le repère du FORMULAIRE. Le
' générateur y pose le décor et Graph_Tracer y place les barres : les deux
' doivent partir du même point, d'où ces deux constantes plutôt qu'un nombre
' écrit à un endroit et un calcul refait à l'autre. Les GR_ qui suivent se
' comptent, elles, depuis ce coin.
Public Const GR_ORIGINE_X As Single = 30
Public Const GR_ORIGINE_Y As Single = F2_TOP + 24

Public Const GR_LEGENDE_HAUT As Single = 12   ' bandeau du titre, et de la valeur survolée
Public Const GR_MARGE_G As Single = 48        ' colonne réservée à l'échelle
Public Const GR_TRACE_TOP As Single = 20      ' haut de l'aire de tracé
Public Const GR_TRACE_HAUT As Single = 80     ' hauteur de l'aire de tracé
Public Const GR_MOIS_HAUT As Single = 11      ' bandeau des noms de mois
Public Const GR_BARRE_LARG As Single = 26
Public Const GR_NB_MOIS As Long = 12
Public Const F2_TUILE_LARG As Single = 118    ' une tuile de statistique
Public Const F2_TUILE_HAUT As Single = 56
Public Const F2_TUILE_GX As Single = 8        ' gouttière horizontale entre tuiles
Public Const F2_TUILE_GY As Single = 6        ' gouttière verticale
Public Const F2_TUILES_COL As Long = 3        ' 3 colonnes sur 2 lignes = 6 tuiles
Public Const F2_PADDING As Single = 14        ' retrait intérieur de la carte
' Retrait des libellés dans une tuile : assez pour que le blanc des libellés ne
' touche pas les coins arrondis de l'image, assez peu pour laisser la place aux
' montants une fois les tuiles resserrées.
Public Const F2_TUILE_INSET As Single = 17    ' 118 - 2 x 17 = 84 pt de texte
Public Const TAILLE_STAT As Single = 11.5     ' montant affiché dans une tuile
Public Const TAILLE_TUILE_CAP As Single = 7.5 ' libellé de tuile, le plus long étant « MOIS NON FACTURÉ »

' Libellé et montant sont centrés dans la tuile, horizontalement par TextAlign
' et verticalement par le calcul ci-dessous : le bloc « libellé + écart +
' montant » est posé à mi-hauteur, de sorte qu'il reste centré même si l'on
' change F2_TUILE_HAUT. Un libellé MSForms dessine son texte en HAUT de son
' cadre, jamais au milieu : chaque cadre est donc taillé au plus près de sa
' police, sans quoi le texte se décalerait vers le haut.
Public Const F2_TUILE_CAP_HAUT As Single = 11
Public Const F2_TUILE_VAL_HAUT As Single = 15
Public Const F2_TUILE_ECART As Single = 3     ' entre le libellé et le montant

'--- Fiche 3 : saisie ---------------------------------------------------------
Public Const F3_TOP As Single = 226
Public Const F3_HAUT As Single = 168
' La grille se lit en TROIS RÉGIONS de deux colonnes chacune :
'
'   région 1 (col. 1-2)   région 2 (col. 3-4)   région 3 (col. 5-6)
'   le client             l'intervention        les textes libres
'
' Les six colonnes ont la même largeur et les trois régions aussi : c'est
' l'écart entre régions, plus large que la gouttière intérieure, qui les fait
' lire comme trois groupes plutôt que comme six colonnes alignées.
'
'   6 x 126 + 3 x 20 + 2 x 42 = 900 = I_CARTE_LARG
'
Public Const IG_X As Single = 30              ' abscisse de la 1re colonne
Public Const IG_Y As Single = 254             ' ordonnée de la 1re ligne
Public Const IG_BLOC As Single = 126          ' largeur d'un bloc libellé + champ
Public Const IG_GOUTTIERE As Single = 20      ' entre deux colonnes d'une même région
Public Const IG_ECART_REGION As Single = 42   ' entre deux régions
Public Const IG_COL_PAR_REGION As Long = 2
Public Const IG_LIGNE As Single = 32          ' pas vertical
Public Const ICH_LBL_HAUT As Single = 11
Public Const ICH_CTL_HAUT As Single = 18

'--- Barre de filtrage --------------------------------------------------------
Public Const IF_TOP As Single = 402
Public Const IF_HAUT As Single = 38

'--- Fiche 4 : tableau des enregistrements ------------------------------------
Public Const IT_TOP As Single = 448
Public Const IT_HAUT As Single = 246
Public Const IT_ENTETE As Single = 22

'--- Barre de boutons ---------------------------------------------------------
Public Const IB_TOP As Single = 704
Public Const IB_HAUT As Single = 30
Public Const IB_LARG As Single = 118
Public Const IB_GOUTTIERE As Single = 6
Public Const IB_ECART_GROUPE As Single = 24   ' entre le groupe CRUD et Facturer

'--- Sélecteur de date --------------------------------------------------------
Public Const CAL_LARGEUR As Single = 260
Public Const CAL_HAUTEUR As Single = 256
Public Const CAL_BANDEAU As Single = 40
Public Const CAL_JOUR_LARG As Single = 34     ' une case du calendrier
Public Const CAL_JOUR_HAUT As Single = 26
Public Const CAL_GRILLE_X As Single = 11      ' 7 x 34 + 2 x 11 = 260
Public Const CAL_GRILLE_Y As Single = 66
Public Const CAL_ENTETES_Y As Single = 48     ' ligne L M M J V S D
Public Const CAL_PIED_Y As Single = 232       ' raccourcis du bas, 10 pt sous la grille

'--- Ressources externes ------------------------------------------------------
' Image de fond des tuiles de statistiques, cherchée à côté du classeur.
' Si le fichier est absent, les tuiles retombent sur un aplat blanc : le
' formulaire reste utilisable, seuls les coins arrondis sont perdus.
Public Const DOSSIER_IMAGES As String = "Images"
Public Const IMAGE_TUILE As String = "CartePremium.jpg"

'==============================================================================
' Abscisse du bloc de tuiles, calé sur le bord droit de la carte.
'
' Calculée plutôt qu'écrite en dur : changer la largeur des tuiles, leur
' gouttière ou leur nombre de colonnes les repositionne toutes seules, et le
' graphique n'a qu'à s'arrêter F2_ECART_GRAPH points avant.
'==============================================================================
Public Function F2TuilesX() As Single
    F2TuilesX = I_MARGE + I_CARTE_LARG - F2_PADDING _
                - (F2_TUILES_COL * F2_TUILE_LARG + (F2_TUILES_COL - 1) * F2_TUILE_GX)
End Function

'==============================================================================
' Position d'un bloc de la grille de saisie
'==============================================================================
' Chaque frontière de région franchie remplace la gouttière ordinaire par
' l'écart de région : on ajoute donc la DIFFÉRENCE entre les deux, une fois par
' frontière déjà passée.
Public Function IGrilleX(ByVal colonne As Long) As Single
    IGrilleX = IG_X + (colonne - 1) * (IG_BLOC + IG_GOUTTIERE) _
               + IRegionDe(colonne) * (IG_ECART_REGION - IG_GOUTTIERE)
End Function

'------------------------------------------------------------------------------
' Numéro de région d'une colonne, à partir de 0 : colonnes 1-2 -> 0,
' colonnes 3-4 -> 1, colonnes 5-6 -> 2.
'------------------------------------------------------------------------------
Public Function IRegionDe(ByVal colonne As Long) As Long
    If colonne < 1 Then Exit Function
    IRegionDe = (colonne - 1) \ IG_COL_PAR_REGION
End Function

'------------------------------------------------------------------------------
' Hauteur d'un champ qui occupe plusieurs lignes de la grille : sa zone de
' saisie descend jusqu'au bas de la dernière ligne couverte.
'------------------------------------------------------------------------------
Public Function IHauteurLignes(ByVal nbLignes As Long) As Single
    If nbLignes < 2 Then
        IHauteurLignes = ICH_CTL_HAUT
    Else
        IHauteurLignes = ICH_CTL_HAUT + (nbLignes - 1) * IG_LIGNE
    End If
End Function

'------------------------------------------------------------------------------
' Ordonnée du bloc « libellé + zone de saisie » d'une ligne de la grille.
'   ligne   : 1 à 4, de haut en bas
'   renvoie : la position en points, mesurée depuis le haut du formulaire
'------------------------------------------------------------------------------
Public Function IGrilleY(ByVal ligne As Long) As Single
    IGrilleY = IG_Y + (ligne - 1) * IG_LIGNE
End Function

'------------------------------------------------------------------------------
' Largeur d'un champ qui s'étale sur plusieurs blocs, gouttières comprises.
'------------------------------------------------------------------------------
Public Function ILargeurBlocs(ByVal nbBlocs As Long) As Single
    ILargeurBlocs = nbBlocs * IG_BLOC + (nbBlocs - 1) * IG_GOUTTIERE
End Function

'==============================================================================
' Couleur de fond d'un bouton d'action du formulaire des interventions.
'   nomBouton : nom EXACT du contrôle, tel que ConstruireBoutonsInterv le crée
'   survol    : True pour la teinte claire du passage de souris
'
' Les sept boutons sont énumérés ici, y compris ceux dont la teinte vient de
' modClients_Theme. Déléguer à CouleurBouton pour ces derniers ne marchait pas :
' les contrôles s'appellent btnIAjouter, btnIModifier… avec un I, que
' CouleurBouton ne reconnaît pas — les cinq tombaient dans son cas par défaut et
' viraient au gris dès le premier déplacement de souris.
'==============================================================================
Public Function ICouleurBouton(ByVal nomBouton As String, ByVal survol As Boolean) As Long
    Select Case nomBouton
        Case "btnIAjouter":   ICouleurBouton = IIf(survol, COUL_AJOUTER_H, COUL_AJOUTER)
        Case "btnIModifier":  ICouleurBouton = IIf(survol, COUL_MODIFIER_H, COUL_MODIFIER)
        Case "btnISupprimer": ICouleurBouton = IIf(survol, COUL_SUPPRIMER_H, COUL_SUPPRIMER)
        Case "btnIEffacer":   ICouleurBouton = IIf(survol, COUL_EFFACER_H, COUL_EFFACER)
        Case "btnIQuitter":   ICouleurBouton = IIf(survol, COUL_QUITTER_H, COUL_QUITTER)
        Case "btnFacturer":   ICouleurBouton = IIf(survol, COUL_FACTURER_H, COUL_FACTURER)
        Case "btnInfo":       ICouleurBouton = IIf(survol, COUL_INFO_H, COUL_INFO)
        Case Else:            ICouleurBouton = IIf(survol, COUL_EFFACER_H, COUL_EFFACER)
    End Select
End Function
