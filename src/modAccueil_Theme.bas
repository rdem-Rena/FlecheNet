Attribute VB_Name = "modAccueil_Theme"
Option Explicit
'==============================================================================
' modAccueil_Theme - LA FEUILLE D'ACCUEIL : COULEURS, GÉOMÉTRIE, CARTES
'------------------------------------------------------------------------------
' Tout ce qui se règle sans toucher au dessin est ici : modAccueil_Generateur
' ne fait que poser des formes aux endroits que ce module lui indique.
'
' LA PALETTE N'EST PAS REDÉFINIE. Elle vient de modClients_Theme, comme pour les
' quatre formulaires : le bandeau bleu foncé, le blanc des cartes, le filet, les
' deux gris de texte. Deux accents manquaient seulement au classeur — l'ambre de
' la facturation et le violet des statistiques — et ils sont posés ici.
'
' POURQUOI UNE FEUILLE ET NON UN CINQUIÈME FORMULAIRE. MSForms ne sait pas
' arrondir un coin, dégrader un fond ni porter une ombre : un menu y aurait
' l'aspect d'un panneau de contrôle de 1998. Les formes d'une feuille savent
' tout cela, et la feuille s'ouvre avec le classeur, sans macro à lancer.
'==============================================================================

'==============================================================================
' CONSTANTES DE L'OBJET SHAPE
'------------------------------------------------------------------------------
' Déclarées ici plutôt que reprises de la bibliothèque Office : le classeur ne
' dépend ainsi d'aucune référence, exactement comme les constantes MSF_* le font
' pour les formulaires.
'==============================================================================
Public Const MSO_VRAI As Long = -1
Public Const MSO_FAUX As Long = 0

Public Const MSO_RECT As Long = 1               ' msoShapeRectangle
Public Const MSO_RECT_ARRONDI As Long = 5       ' msoShapeRoundedRectangle
Public Const MSO_OVALE As Long = 9              ' msoShapeOval
Public Const MSO_TEXTE_HORIZONTAL As Long = 1   ' msoTextOrientationHorizontal

Public Const MSO_ALIGN_GAUCHE As Long = 1
Public Const MSO_ALIGN_CENTRE As Long = 2
Public Const MSO_ALIGN_DROITE As Long = 3

Public Const MSO_ANCRE_HAUT As Long = 1
Public Const MSO_ANCRE_MILIEU As Long = 3

Public Const MSO_AUTOSIZE_AUCUN As Long = 0
Public Const MSO_OMBRE_EXTERIEURE As Long = 1   ' msoShadowStyleOuterShadow

Public Const XL_FLOTTANT As Long = 3            ' xlFreeFloating
Public Const XL_AUCUNE_SELECTION As Long = -4142 ' xlNoSelection

'==============================================================================
' LA FEUILLE
'==============================================================================
Public Const NOM_FEUILLE_ACCUEIL As String = "Accueil"

'==============================================================================
' GÉOMÉTRIE, EN POINTS
'------------------------------------------------------------------------------
' Même largeur utile que les quatre formulaires — 960 points, soit 1280 pixels
' à 96 ppp — pour que l'accueil et ce qu'il ouvre aient la même empreinte.
'
'   +----------------------------------------------------------------+  0
'   |  BANDEAU : titre, sous-titre, année                            |
'   +----------------------------------------------------------------+  104
'      MODULES                                                          128
'   +----------+  +----------+  +----------+  +----------+              152
'   | pastille |  |          |  |          |  |          |
'   | Titre    |  |          |  |          |  |          |
'   | détail   |  |          |  |          |  |          |
'   | Ouvrir > |  |          |  |          |  |          |
'   +----------+  +----------+  +----------+  +----------+              376
'   ------------------------------------------------------------        404
'   pied de page                                                        414
'==============================================================================
Public Const AC_LARGEUR As Single = 960
Public Const AC_MARGE As Single = 30
Public Const AC_GOUTTIERE As Single = 20
Public Const AC_NB_CARTES As Long = 4

' Bandeau
Public Const AC_BAND_HAUT As Single = 104
Public Const AC_TITRE_TOP As Single = 26
Public Const AC_SOUS_TOP As Single = 63
Public Const AC_ANNEE_X As Single = 700
Public Const AC_ANNEE_TOP As Single = 28
Public Const AC_ANNEE_LARG As Single = 230

' Intitulé de section
Public Const AC_SECTION_TOP As Single = 128

' Cartes
Public Const AC_CARTE_TOP As Single = 152
Public Const AC_CARTE_HAUT As Single = 224
Public Const AC_CARTE_LARG As Single = 210      ' 30 + 4x210 + 3x20 + 30 = 960
Public Const AC_CARTE_PAD As Single = 22        ' marge intérieure d'une carte

' Contenu d'une carte, ordonnées comptées depuis son coin haut-gauche
Public Const AC_PASTILLE_TOP As Single = 24
Public Const AC_PASTILLE_D As Single = 46       ' diamètre
Public Const AC_TRAIT_TOP As Single = 84
Public Const AC_TRAIT_LARG As Single = 30
Public Const AC_TRAIT_HAUT As Single = 3
Public Const AC_TITRE_C_TOP As Single = 96
Public Const AC_DETAIL_TOP As Single = 126
Public Const AC_DETAIL_HAUT As Single = 62
Public Const AC_LIEN_TOP As Single = 192

' Pied de page
Public Const AC_FILET_TOP As Single = 404
Public Const AC_PIED_TOP As Single = 414
Public Const AC_HAUTEUR As Single = 448         ' hauteur totale dessinée

'==============================================================================
' TYPOGRAPHIE
'------------------------------------------------------------------------------
' POLICE et les couleurs viennent de modClients_Theme. Segoe UI Light n'est
' utilisée que pour l'année, où sa finesse porte : absente, Windows retombe sur
' Segoe UI sans rien casser.
'==============================================================================
Public Const POLICE_DEMI As String = "Segoe UI Semibold"
Public Const POLICE_LEGERE As String = "Segoe UI Light"

Public Const AC_T_TITRE As Single = 24
Public Const AC_T_SOUS As Single = 10.5
Public Const AC_T_ANNEE As Single = 32
Public Const AC_T_SECTION As Single = 8.5
Public Const AC_T_PASTILLE As Single = 19
Public Const AC_T_TITRE_C As Single = 14.5
Public Const AC_T_DETAIL As Single = 9.5
Public Const AC_T_LIEN As Single = 9.5
Public Const AC_T_PIED As Single = 8.5

' Interlettrage de l'intitulé de section : c'est lui qui fait lire « MODULES »
' comme une étiquette et non comme un mot.
Public Const AC_SECTION_ESPACE As Single = 1.6

'==============================================================================
' LES DEUX ACCENTS QUE LE CLASSEUR N'AVAIT PAS
'------------------------------------------------------------------------------
' Le bleu et le vert des deux premières cartes sont ceux des boutons Modifier
' et Ajouter. Le rouge et l'anthracite restants disent « supprimer » et
' « quitter » : les reprendre ici aurait mal parlé.
'==============================================================================
Public Const COUL_MENU_FACTURE As Long = &H1E7AC7&   ' #C77A1E  ambre
Public Const COUL_MENU_STAT As Long = &HA34F6B&      ' #6B4FA3  violet

'==============================================================================
' L'OMBRE PORTÉE DES CARTES
'------------------------------------------------------------------------------
' Très diffuse et très transparente : elle doit se sentir sans se voir.
'==============================================================================
Public Const AC_OMBRE_FLOU As Single = 14
Public Const AC_OMBRE_DY As Single = 4
Public Const AC_OMBRE_TRANSP As Single = 0.86

'==============================================================================
' LES QUATRE CARTES
'------------------------------------------------------------------------------
' C'EST LA TABLE À MODIFIER pour renommer un module, changer son texte, sa
' couleur ou la macro qu'il lance. Le générateur s'y adapte seul, y compris si
' le nombre de cartes change — à condition de corriger AC_NB_CARTES et la
' largeur, que simulate_accueil.py vérifie.
'==============================================================================
Public Type CarteMenu
    Lettre As String        ' l'initiale posée dans la pastille
    Titre As String
    Detail As String
    Macro As String         ' Sub publique sans argument, lancée au clic
    Couleur As Long
End Type

Private mCartes() As CarteMenu
Private mChargees As Boolean

Public Function ObtenirCartesMenu() As CarteMenu()
    If Not mChargees Then ConstruireCartesMenu
    ObtenirCartesMenu = mCartes
End Function

Private Sub ConstruireCartesMenu()
    ReDim mCartes(1 To AC_NB_CARTES)

    DefCarte mCartes, 1, "C", "Clients", _
             "Les fiches clients : coordonnées, adresse, taux horaire ou forfait.", _
             "OuvrirGestionClients", COUL_MODIFIER
    DefCarte mCartes, 2, "I", "Interventions", _
             "Saisir et retrouver les interventions. Tableau, calendrier et graphique du mois.", _
             "OuvrirGestionInterventions", COUL_AJOUTER
    DefCarte mCartes, 3, "F", "Facturation", _
             "Attribuer un numéro de facture aux travaux qui n'en ont pas encore.", _
             "OuvrirFacturation", COUL_MENU_FACTURE
    DefCarte mCartes, 4, "S", "Statistiques", _
             "Chiffre d'affaires, heures, objectif annuel et liste filtrable des travaux.", _
             "OuvrirStatistiques", COUL_MENU_STAT

    mChargees = True
End Sub

Private Sub DefCarte(ByRef tb() As CarteMenu, ByVal idx As Long, ByVal lettre As String, _
                     ByVal titre As String, ByVal detail As String, _
                     ByVal nomMacro As String, ByVal couleur As Long)
    tb(idx).Lettre = lettre
    tb(idx).Titre = titre
    tb(idx).Detail = detail
    tb(idx).Macro = nomMacro
    tb(idx).Couleur = couleur
End Sub

'------------------------------------------------------------------------------
' Abscisse de la carte n° idx, comptée depuis le bord gauche de la feuille.
'------------------------------------------------------------------------------
Public Function AcCarteX(ByVal idx As Long) As Single
    AcCarteX = AC_MARGE + (idx - 1) * (AC_CARTE_LARG + AC_GOUTTIERE)
End Function
