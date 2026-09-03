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
Public Const FA_HAUTEUR As Single = 636         ' surface UTILE, hors barre de titre

' La propriété Height d'un UserForm est sa hauteur EXTÉRIEURE : posée à
' FA_HAUTEUR, elle laisserait la barre de titre manger le bas, et le dernier
' bouton disparaîtrait sous le bord. Le générateur pose donc FA_HAUTEUR plus
' cette réserve, et Fact_Activer affine ensuite au point près en lisant
' InsideHeight — la réserve n'a pas besoin d'être exacte, seulement suffisante.
Public Const FA_RESERVE_TITRE As Single = 24
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
' LA CONFIGURATION DU TABLEAU DES INTERVENTIONS EST REPRISE TELLE QUELLE, et
' non recopiée : hauteur de ligne, retrait du texte et largeur de la barre
' viennent de IGR_LIGNE_H, IGR_PAD_X et IGR_BARRE_L, dans modInterv_Theme. Les
' trois tableaux du classeur bougent donc ensemble, et aucun ne peut dériver.
'
' Seuls le bandeau de titre et la ligne des totaux sont propres à ce
' formulaire : le tableau des interventions n'en a pas.
Public Const FA_TITRE_HAUT As Single = 20       ' bandeau de titre d'un tableau
Public Const FA_ENTETE_HAUT As Single = 20
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
' Délégués au bandeau commun : voir Z1Titre dans modInterv_Theme.
Public Function ZF1Titre() As StyleTexte
    ZF1Titre = Z1Titre()
End Function
Public Function ZF1Annee() As StyleTexte
    ZF1Annee = Z1Annee()
End Function

'--- Bandeau de titre d'un tableau --------------------------------------------
Public Function ZFSection() As StyleTexte
    ZFSection = StyleF(POLICE, 9, True, COUL_SECTION)
End Function

'--- En-tête de colonne, commun aux deux grilles ------------------------------
' DÉLÉGUÉ au tableau des interventions plutôt que redéfini : c'est la même
' configuration, elle ne doit donc exister qu'une fois. Changer Z5Entete
' change les trois en-têtes du classeur d'un coup.
Public Function ZFEntete() As StyleTexte
    ZFEntete = Z5Entete()
End Function

'--- Intérieur des grilles ----------------------------------------------------
' Délégué de même à Z6Case : famille, taille, graisse et couleur des cases
' viennent du tableau des interventions.
Public Function ZFCase() As StyleTexte
    ZFCase = Z6Case()
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
