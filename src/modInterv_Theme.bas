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
Public Const F2_GRAPH_LARG As Single = 420    ' largeur de l'image du graphique
Public Const F2_TUILE_LARG As Single = 150    ' une tuile de statistique
Public Const F2_TUILE_HAUT As Single = 56
Public Const F2_TUILE_GX As Single = 9        ' gouttière horizontale entre tuiles
Public Const F2_TUILE_GY As Single = 6        ' gouttière verticale
Public Const F2_TUILES_COL As Long = 3        ' 3 colonnes sur 2 lignes = 6 tuiles

'--- Fiche 3 : saisie ---------------------------------------------------------
Public Const F3_TOP As Single = 226
Public Const F3_HAUT As Single = 168
Public Const IG_X As Single = 30              ' abscisse de la 1re colonne
Public Const IG_Y As Single = 254             ' ordonnée de la 1re ligne
Public Const IG_BLOC As Single = 210          ' largeur d'un bloc libellé + champ
Public Const IG_GOUTTIERE As Single = 20
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
Public Const CAL_LARGEUR As Single = 224
Public Const CAL_HAUTEUR As Single = 216
Public Const CAL_BANDEAU As Single = 34
Public Const CAL_JOUR_LARG As Single = 30     ' une case du calendrier
Public Const CAL_JOUR_HAUT As Single = 22
Public Const CAL_GRILLE_X As Single = 7
Public Const CAL_GRILLE_Y As Single = 62

'--- Ressources externes ------------------------------------------------------
' Image de fond des tuiles de statistiques, cherchée à côté du classeur.
' Si le fichier est absent, les tuiles retombent sur un aplat blanc : le
' formulaire reste utilisable, seuls les coins arrondis sont perdus.
Public Const DOSSIER_IMAGES As String = "Images"
Public Const IMAGE_TUILE As String = "CartePremium.jpg"

'==============================================================================
' Position d'un bloc de la grille de saisie
'==============================================================================
Public Function IGrilleX(ByVal colonne As Long) As Single
    IGrilleX = IG_X + (colonne - 1) * (IG_BLOC + IG_GOUTTIERE)
End Function

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
' Les quatre premiers reprennent les couleurs du formulaire des clients ;
' Facturer et Info ont les leurs, pour se distinguer du groupe de saisie.
'==============================================================================
Public Function ICouleurBouton(ByVal nomBouton As String, ByVal survol As Boolean) As Long
    Select Case nomBouton
        Case "btnFacturer": ICouleurBouton = IIf(survol, COUL_FACTURER_H, COUL_FACTURER)
        Case "btnInfo":     ICouleurBouton = IIf(survol, COUL_INFO_H, COUL_INFO)
        Case Else:          ICouleurBouton = CouleurBouton(nomBouton, survol)
    End Select
End Function
