Attribute VB_Name = "modStat_Theme"
Option Explicit
'==============================================================================
' modStat_Theme
'------------------------------------------------------------------------------
' Géométrie du formulaire UF_Statistiques.
'
' TOUT CE QUI EXISTE DÉJÀ EST REPRIS, JAMAIS RECOPIÉ : la palette vient de
' modClients_Theme, AuPixel et le type StyleTexte de modInterv_Theme, et les
' tuiles comme le graphique gardent les dimensions du formulaire des
' interventions — F2_TUILE_*, GR_*. Le tableau reprend la configuration de
' celui des interventions par délégation, comme celui de la facturation.
'
' Les six zones sont portées par de vrais cadres.
'==============================================================================

'--- Fenêtre et cartes --------------------------------------------------------
Public Const ST_LARGEUR As Single = 960
Public Const ST_HAUTEUR As Single = 680         ' surface UTILE, hors barre de titre
Public Const ST_RESERVE_TITRE As Single = 24    ' voir modFact_Theme : Height est extérieure
Public Const ST_MARGE As Single = 16
Public Const ST_CARTE_LARG As Single = 928      ' ST_LARGEUR - 2 * ST_MARGE
Public Const ST_PADDING As Single = 14

'--- Zone 1 : l'intitulé ------------------------------------------------------
Public Const ST_Z1_TOP As Single = 12
Public Const ST_Z1_HAUT As Single = 46

'--- Zone 2 : les cinq tuiles, sur une seule ligne ----------------------------
Public Const ST_Z2_TOP As Single = 66
Public Const ST_Z2_HAUT As Single = 68
Public Const ST_NB_TUILES As Long = 5

'--- Zone 3 : le graphique mensuel et la barre d'objectif ---------------------
Public Const ST_Z3_TOP As Single = 142
Public Const ST_Z3_HAUT As Single = 132
Public Const ST_GR_X As Single = 24             ' origine du graphique DANS son cadre
Public Const ST_GR_Y As Single = 8

' Largeur PROPRE à ce graphique, et non F2_GRAPH_LARG : celui du formulaire des
' interventions partage une carte avec six tuiles et n'a que 518 points ; ici la
' carte lui est presque entière, autant qu'il en profite.
Public Const ST_GR_LARG As Single = 640

' La barre de progression, à droite du graphique. Verticale : elle se remplit
' du bas vers le haut, comme un thermomètre.
Public Const ST_OBJ_X As Single = 690
Public Const ST_OBJ_LARG As Single = 54
Public Const ST_OBJ_TOP As Single = 24
Public Const ST_OBJ_HAUT As Single = 84
' Le libellé et les montants tiennent à droite de la cuve : 690 + 54 + 12 + 160
' = 916, sous les 926 points utiles de la carte. Le graphique s'arrête à 664,
' ce qui laisse 26 points entre les deux.
Public Const ST_OBJ_TXT_L As Single = 160

'--- Zone 4 : la barre de filtrage --------------------------------------------
Public Const ST_Z4_TOP As Single = 282
Public Const ST_Z4_HAUT As Single = 40
Public Const ST_CTL_HAUT As Single = 18

' Abscisses des sept filtres, dans leur cadre. Tout tient sur une ligne.
Public Const ST_F_MOIS_LBL As Single = 14
Public Const ST_F_MOIS As Single = 46
Public Const ST_F_MOIS_L As Single = 92
Public Const ST_F_ENT_LBL As Single = 144
Public Const ST_F_ENT As Single = 204
Public Const ST_F_ENT_L As Single = 110
Public Const ST_F_NOM_LBL As Single = 320
Public Const ST_F_NOM As Single = 354
Public Const ST_F_NOM_L As Single = 92
Public Const ST_F_TVA As Single = 452
Public Const ST_F_TVA_L As Single = 48
Public Const ST_F_FORF As Single = 504
Public Const ST_F_FORF_L As Single = 58
Public Const ST_F_FACT As Single = 566
Public Const ST_F_FACT_L As Single = 66
' Le bouton de remise à zéro, entre « Facturé » et « N° facture ».
Public Const ST_F_RAZ As Single = 638
Public Const ST_F_RAZ_L As Single = 80
Public Const ST_F_NUM_LBL As Single = 724
Public Const ST_F_NUM As Single = 792
Public Const ST_F_NUM_L As Single = 96

'--- Zone 5 : le tableau ------------------------------------------------------
Public Const ST_Z5_TOP As Single = 330
Public Const ST_Z5_HAUT As Single = 300
Public Const ST_Z5_LIGNES As Long = 20

'--- Barre de boutons ---------------------------------------------------------
Public Const ST_BT_TOP As Single = 640
Public Const ST_BT_HAUT As Single = 28
Public Const ST_BT_LARG As Single = 120

'--- Étages d'un tableau ------------------------------------------------------
Public Const ST_TITRE_HAUT As Single = 20
Public Const ST_ENTETE_HAUT As Single = 20

Public Function SGrilleEnteteY() As Single
    SGrilleEnteteY = AuPixel(ST_TITRE_HAUT + 1)
End Function

Public Function SGrilleLignesY() As Single
    SGrilleLignesY = AuPixel(ST_TITRE_HAUT + 1 + ST_ENTETE_HAUT)
End Function

Public Function SGrilleLargeur() As Single
    SGrilleLargeur = ST_CARTE_LARG - 2
End Function

' Abscisse de la première tuile : les cinq sont centrées dans leur cadre.
Public Function STuilesX() As Single
    STuilesX = AuPixel((ST_CARTE_LARG - (ST_NB_TUILES * F2_TUILE_LARG + _
                        (ST_NB_TUILES - 1) * F2_TUILE_GX)) / 2)
End Function

'==============================================================================
' TYPOGRAPHIE
'------------------------------------------------------------------------------
' Les rôles communs sont DÉLÉGUÉS aux styles des interventions plutôt que
' redéfinis : le tableau, les tuiles et le graphique doivent avoir exactement
' le même aspect d'un formulaire à l'autre, et deux définitions finiraient par
' diverger.
'==============================================================================
Public Function ZS1Titre() As StyleTexte
    ZS1Titre = Z1Titre()
End Function

Public Function ZS1Annee() As StyleTexte
    ZS1Annee = Z1Annee()
End Function

Public Function ZSTuileCap() As StyleTexte
    ZSTuileCap = Z2TuileCap()
End Function

Public Function ZSTuileVal() As StyleTexte
    ZSTuileVal = Z2TuileVal()
End Function

Public Function ZSGraphTitre() As StyleTexte
    ZSGraphTitre = Z2GraphTitre()
End Function

Public Function ZSGraphAxe() As StyleTexte
    ZSGraphAxe = Z2GraphAxe()
End Function

Public Function ZSSection() As StyleTexte
    ZSSection = Z2Section()
End Function

Public Function ZSFiltre() As StyleTexte
    ZSFiltre = Z4Libelle()
End Function

Public Function ZSEntete() As StyleTexte
    ZSEntete = Z5Entete()
End Function

Public Function ZSCase() As StyleTexte
    ZSCase = Z6Case()
End Function
