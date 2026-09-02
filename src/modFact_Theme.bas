Attribute VB_Name = "modFact_Theme"
Option Explicit
'==============================================================================
' modFact_Theme
'------------------------------------------------------------------------------
' Géométrie et typographie du formulaire UF_Facture.
'
' La PALETTE vient de modClients_Theme et les OUTILS de modInterv_Theme :
' AuPixel, qui cale une coordonnée sur le pixel, et le type StyleTexte. Rien
' n'est redéfini ici de ce qui existe déjà ailleurs — une couleur changée dans
' la charte suit sur les trois formulaires.
'
' Les cinq zones sont portées par de vrais CADRES, comme les quatre du
' formulaire des interventions : les coordonnées de leurs enfants comptent donc
' depuis le cadre, et non depuis le formulaire.
'==============================================================================

'--- Fenêtre et cartes --------------------------------------------------------
Public Const FA_LARGEUR As Single = 960
Public Const FA_HAUTEUR As Single = 636
Public Const FA_MARGE As Single = 16
Public Const FA_CARTE_LARG As Single = 928      ' FA_LARGEUR - 2 * FA_MARGE
Public Const FA_PADDING As Single = 14

'--- Zone 1 : l'intitulé ------------------------------------------------------
Public Const FA_Z1_TOP As Single = 12
Public Const FA_Z1_HAUT As Single = 46

'--- Zone 2 : le tableau des clients ------------------------------------------
Public Const FA_Z2_TOP As Single = 66
Public Const FA_Z2_HAUT As Single = 160
Public Const FA_Z2_LIGNES As Long = 9

'--- Zone 3 : la barre de filtrage --------------------------------------------
Public Const FA_Z3_TOP As Single = 234
Public Const FA_Z3_HAUT As Single = 40

'--- Zone 4 : le tableau des travaux ------------------------------------------
Public Const FA_Z4_TOP As Single = 282
Public Const FA_Z4_HAUT As Single = 220
Public Const FA_Z4_LIGNES As Long = 12

'--- Zone 5 : le texte de facture et les commentaires -------------------------
Public Const FA_Z5_TOP As Single = 510
Public Const FA_Z5_HAUT As Single = 76

'--- Barre de boutons ---------------------------------------------------------
Public Const FA_BT_TOP As Single = 596
Public Const FA_BT_HAUT As Single = 28
Public Const FA_BT_LARG As Single = 120

'--- Grilles ------------------------------------------------------------------
' Mêmes conventions que la grille des interventions : la hauteur de ligne vaut
' un nombre entier de pixels (12,75 pt = 17 px), sans quoi une ligne sur trois
' se décalerait, et le texte des cases est en neuf points, soit douze pixels.
Public Const FA_TITRE_HAUT As Single = 20       ' bandeau de titre d'un tableau
Public Const FA_ENTETE_HAUT As Single = 20
Public Const FA_LIGNE_H As Single = 12.75
Public Const FA_BARRE_L As Single = 16
Public Const FA_PAD_X As Single = 4.5
Public Const FA_TOTAUX_HAUT As Single = 18

'--- Contrôles de la barre de filtrage ----------------------------------------
Public Const FA_CTL_HAUT As Single = 18
Public Const FA_MOIS_LBL_X As Single = 30
Public Const FA_MOIS_X As Single = 68
Public Const FA_MOIS_L As Single = 140
Public Const FA_TOUTES_X As Single = 236
Public Const FA_TOUTES_L As Single = 160
Public Const FA_NUM_LBL_X As Single = 640
Public Const FA_NUM_X As Single = 740
Public Const FA_NUM_L As Single = 180

'==============================================================================
' Position d'un tableau DANS SON CADRE
'------------------------------------------------------------------------------
' Le bandeau de titre occupe le haut du cadre, l'en-tête vient dessous, puis
' les lignes. Ces trois fonctions donnent l'ordonnée de chaque étage, calée sur
' le pixel : c'est le seul endroit qui connaisse cet empilement.
'==============================================================================
Public Function FGrilleEnteteY() As Single
    FGrilleEnteteY = AuPixel(FA_TITRE_HAUT + 1)
End Function

Public Function FGrilleLignesY() As Single
    FGrilleLignesY = AuPixel(FA_TITRE_HAUT + 1 + FA_ENTETE_HAUT)
End Function

'------------------------------------------------------------------------------
' Largeur utile d'une grille : la carte, moins ses deux filets, moins la barre.
'------------------------------------------------------------------------------
Public Function FGrilleLargeur() As Single
    FGrilleLargeur = FA_CARTE_LARG - 2
End Function

'==============================================================================
' TYPOGRAPHIE PAR ZONE
'------------------------------------------------------------------------------
' Même principe que pour les interventions : chaque rôle de libellé a sa ligne,
' portant famille, taille, graisse et couleur. Pour changer l'aspect d'une
' zone, tout est ici.
'
' Les tailles tombent sur un nombre entier de pixels — 9 pt = 12 px, 9,75 pt =
' 13 px — parce qu'un corps bancal fait arrondir Windows et dessine le texte
' plus épais que demandé.
'==============================================================================

'--- Zone 1 : l'intitulé ------------------------------------------------------
Public Function ZF1Titre() As StyleTexte
    ZF1Titre = StyleF(POLICE, 14, True, COUL_BANDEAU)
End Function
Public Function ZF1Annee() As StyleTexte
    ZF1Annee = StyleF(POLICE, 20, True, COUL_MODIFIER)
End Function

'--- Bandeau de titre d'un tableau --------------------------------------------
Public Function ZFSection() As StyleTexte
    ZFSection = StyleF(POLICE, 9, True, COUL_SECTION)
End Function

'--- En-tête de colonne, commun aux deux grilles ------------------------------
Public Function ZFEntete() As StyleTexte
    ZFEntete = StyleF(POLICE_ENTETE, 9, False, COUL_GRILLE_ENTETE)
End Function

'--- Intérieur des grilles ----------------------------------------------------
Public Function ZFCase() As StyleTexte
    ZFCase = StyleF(POLICE, 9, False, COUL_GRILLE_TXT)
End Function

'--- Ligne des totaux, sous le tableau des travaux ----------------------------
Public Function ZFTotal() As StyleTexte
    ZFTotal = StyleF(POLICE, 9.75, True, COUL_BANDEAU)
End Function

'--- Zone 3 : libellés de la barre de filtrage --------------------------------
Public Function ZF3Libelle() As StyleTexte
    ZF3Libelle = StyleF(POLICE, 9, True, COUL_ENTETE_TXT)
End Function

'--- Zone 5 : libellés des deux zones de texte --------------------------------
Public Function ZF5Libelle() As StyleTexte
    ZF5Libelle = StyleF(POLICE, 7.5, True, COUL_TEXTE_DOUX)
End Function

'------------------------------------------------------------------------------
' Assemble les quatre caractéristiques. Porte le suffixe F pour ne pas entrer
' en conflit avec la fonction de même rôle de modInterv_Theme, qui est privée
' à ce module-là.
'------------------------------------------------------------------------------
Private Function StyleF(ByVal police As String, ByVal taille As Single, _
                        ByVal gras As Boolean, ByVal couleur As Long) As StyleTexte
    StyleF.Police = police
    StyleF.Taille = taille
    StyleF.Gras = gras
    StyleF.Couleur = couleur
End Function
