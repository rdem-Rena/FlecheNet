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

'--- Les quatre caractéristiques d'un texte -----------------------------------
' Réunies dans un seul objet, elles se posent d'un bloc : impossible d'en
' changer une en oubliant les autres. Voir la TYPOGRAPHIE PAR ZONE, tout en bas.
Public Type StyleTexte
    Police As String
    Taille As Single
    Gras As Boolean
    Couleur As Long
End Type

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
' ESSAI : la fiche 1 portée par un vrai CADRE au lieu d'un libellé de fond.
'
' Un cadre est le conteneur natif de MSForms. Ses enfants sont toujours devant
' lui, sans qu'on ait à les repousser après coup, et la zone entière se
' déplace, se masque ou se désactive d'un seul contrôle. En contrepartie les
' coordonnées des enfants deviennent RELATIVES au cadre.
'
' L'essai ne porte que sur cette fiche, la plus simple : trois libellés, aucun
' événement. Mettre False rend exactement la version en libellé de fond, sans
' rien d'autre à toucher — c'est le filet de sécurité tant que le comportement
' du cadre n'a pas été vu dans Excel.
Public Const F1_EN_CADRE As Boolean = True
Public Const F2_EN_CADRE As Boolean = True
Public Const F3_EN_CADRE As Boolean = True
Public Const F4_EN_CADRE As Boolean = True

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

' Abscisses de la barre de filtrage, de gauche à droite : le libellé, la colonne
' sur laquelle porter le filtre, le texte cherché, le mois, puis le lien de
' remise à zéro et le compteur calé à droite.
Public Const IFB_TITRE_X As Single = 30
Public Const IFB_CHAMP_X As Single = 100
Public Const IFB_CHAMP_L As Single = 124
Public Const IFB_TEXTE_X As Single = 232
Public Const IFB_TEXTE_L As Single = 250
Public Const IFB_MOIS_LBL_X As Single = 496
Public Const IFB_MOIS_X As Single = 528
Public Const IFB_MOIS_L As Single = 112
Public Const IFB_RESET_X As Single = 652
Public Const IFB_COMPTEUR_L As Single = 190

'--- Fiche 4 : tableau des enregistrements ------------------------------------
' Une ligne sur deux est teintée : sur douze colonnes, l'oeil suit une ligne
' bien mieux qu'avec un survol, qui n'éclaire que là où pointe la souris.
Public Const COUL_GRILLE_ZEBRE As Long = &HFAF5F1&   ' #F1F5FA

' L'INTÉRIEUR DU TABLEAU FAIT EXCEPTION AU DÉFAUT DU FORMULAIRE : neuf points
' et MAIGRE, là où le reste du formulaire est en huit points gras. Dix-sept
' lignes de gras pèseraient lourd, alors qu'un libellé n'a que quelques mots.
' C'est la couleur, plus claire que celle des fiches, qui allège encore le
' trait — MSForms ne proposant pas de graisse intermédiaire.
'
' Neuf points, et pas huit et demi : un point vaut 0,75 pixel, donc 9 pt font
' exactement 12 pixels de corps là où 8,5 pt en font 11,33. Windows arrondit un
' corps bancal, et le texte se dessine alors un peu plus épais que demandé.
Public Const COUL_GRILLE_TXT As Long = &H6E5A4A&     ' #4A5A6E
Public Const TAILLE_GRILLE_TXT As Single = 9
Public Const GRAS_GRILLE_TXT As Boolean = False

' L'en-tête, lui, est plus sombre et nettement plus grand que le corps : c'est
' ce qui le détache des lignes sans avoir à le souligner.
Public Const COUL_GRILLE_ENTETE As Long = &H523E2C&  ' #2C3E52
Public Const TAILLE_GRILLE_ENTETE As Single = 9.5

' Demi-gras pour l'en-tête. MSForms ne connaît que gras ou pas gras : la demi-
' graisse s'obtient en nommant la FAMILLE qui la porte. Segoe UI Semibold est
' livrée avec Windows depuis Vista ; si elle manquait, Windows retomberait sur
' Segoe UI ordinaire — l'en-tête serait maigre, jamais illisible.
' Pour un en-tête franchement gras : mettre POLICE ici et passer les libellés
' en gras dans ConstruireFiche4.
Public Const POLICE_ENTETE As String = "Segoe UI Semibold"

'--- Police par défaut du formulaire ------------------------------------------
' Celle que reçoit tout contrôle auquel on ne pose rien d'autre. Sans elle,
' MSForms donne aux contrôles MS Sans Serif 8 gras, sa valeur d'usine.
'
' Les contrôles qui posent leur propre police — libellés, zones de saisie,
' listes, cases à cocher, boutons — ne sont pas concernés : ils partent de ce
' défaut et s'en écartent. L'intérieur du tableau, lui, est explicitement en
' TAILLE_GRILLE_TXT et MAIGRE : voir GRAS_GRILLE_TXT, plus haut.
Public Const TAILLE_DEFAUT As Single = 8

' LE DÉFAUT N'EST PAS EN GRAS, ET C'EST VOULU.
'
' Les libellés sont bel et bien en gras : les vingt-deux styles de zone posent
' tous leur graisse eux-mêmes, aucun ne s'en remet à ce défaut. La valeur
' ci-dessous ne décide donc de l'aspect d'aucun texte affiché — elle ne fait
' que se répandre dans les contrôles qui ne demandent rien, zones de saisie et
' menus déroulants compris, où elle mettait le texte tapé en gras.
'
' L'override posé dans Zone ne suffisait pas : MSForms ne reprend pas toujours
' une graisse remise à False sur un contrôle créé par le concepteur. Plutôt que
' de lutter contre le défaut, on ne lui donne plus rien à imposer.
'
' Repasser à True remettrait le gras dans les zones de saisie.
Public Const GRAS_DEFAUT As Boolean = False

Public Const IT_TOP As Single = 448
Public Const IT_HAUT As Single = 246
Public Const IT_ENTETE As Single = 22

' Le tableau n'est plus une ListBox mais une GRILLE DE LIBELLÉS : une case par
' cellule. Une ListBox MSForms plafonne à dix colonnes, aligne tout à gauche et
' ne sait colorer ni une ligne ni une cellule. Ici, autant de colonnes qu'on
' veut, un alignement par colonne, un zébrage et la ligne choisie en couleur.
'
' Seules IGR_NB_LIGNES lignes de libellés existent. La barre de défilement ne
' déplace rien : elle change la première ligne affichée, et les mêmes libellés
' sont repeints. Un tableau de plusieurs milliers de lignes tient donc dans
' deux cents contrôles.
'
' HAUTEUR DE LIGNE : elle doit valoir un nombre ENTIER de pixels, sinon les
' lignes ne tombent pas toutes au même endroit. À 96 ppp, 1 pixel vaut 0,75
' point ; 13 pt font 17,33 px, et Windows arrondit alors les positions à 17, 18,
' 17, 17, 18... Une ligne sur trois se décalait d'un pixel et son texte paraissait
' plus gros et plus gras. 12,75 pt font exactement 17 px — c'est d'ailleurs la
' hauteur de ligne d'une ListBox, et ce n'est pas un hasard.
'
'   IGR_NB_LIGNES x IGR_LIGNE_H = 17 x 12,75 = 216,75, pour 221 disponibles
'
Public Const IGR_LIGNE_H As Single = 12.75    ' pas vertical d'une ligne = 17 px
Public Const IGR_NB_LIGNES As Long = 17       ' lignes affichées à la fois
' Barre de défilement : 16 points font 21 pixels, de quoi attraper le curseur
' sans viser. En dessous de 14, MSForms n'a plus la place de dessiner autre
' chose que les deux flèches.
Public Const IGR_BARRE_L As Single = 16
' Retrait du texte de CHAQUE CÔTÉ de sa colonne : une colonne alignée à droite
' viendrait sinon coller la colonne suivante, alignée à gauche.
Public Const IGR_PAD_X As Single = 4.5

'--- Calage sur le pixel ------------------------------------------------------
' À 96 ppp, un pixel vaut 0,75 point : voir AuPixel, plus bas.
Public Const PT_PAR_PIXEL As Single = 0.75

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
' CALAGE SUR LE PIXEL
'------------------------------------------------------------------------------
' Un point ne vaut pas un nombre entier de pixels : à 96 ppp, 1 px = 0,75 pt.
' Une coordonnée qui tombe entre deux pixels fait rendre le texte à cheval, et
' Windows le dessine alors plus épais et légèrement décalé — d'une colonne à
' l'autre, sans régularité apparente.
'
' Toutes les coordonnées de la grille passent donc par AuPixel. C'est ce qui
' permet de choisir les largeurs de colonnes librement, sans avoir à vérifier
' à la main que chacune tombe juste.
'
' Sur un écran à 120 ou 144 ppp, Windows met le formulaire entier à l'échelle :
' le calage reste bon, puisqu'il porte sur la grille logique.
'==============================================================================
Public Function AuPixel(ByVal v As Single) As Single
    AuPixel = Int(v / PT_PAR_PIXEL + 0.5) * PT_PAR_PIXEL
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

'==============================================================================
' TYPOGRAPHIE PAR ZONE
'------------------------------------------------------------------------------
' Une zone = une carte du formulaire, désignée par son fond. Chaque rôle de
' libellé qui s'y trouve a sa ligne, et cette ligne porte les quatre
' caractéristiques du texte : famille, taille, graisse, couleur.
'
'   Zone 1   lblCarte1         l'intitulé global
'   Zone 2   lblCarte2         les statistiques et le graphique
'   Zone 3   lblCarte3         la fiche intervention
'   Zone 4   lblCarteFiltreI   la barre de filtrage
'   Zone 5   lblEnteteTableI   la ligne d'en-tête du tableau
'   Zone 6   lblG_1_1 … lblG_17_12   l'intérieur du tableau
'   Zone 7   UF_Calendrier     le sélecteur de date, autre formulaire
'
' POUR CHANGER L'ASPECT D'UNE ZONE, tout est ici : il n'y a pas à ouvrir le
' générateur. Une taille écrite en clair ne vaut que pour cette zone ; une
' taille nommée (TAILLE_LIBELLE…) est partagée avec le formulaire des clients,
' et la changer déplace les deux — pour n'en bouger qu'une, remplacer le nom
' par un nombre.
'
' Les COULEURS, elles, sont nommées à dessein : c'est la palette du classeur,
' et une teinte doit rester la même partout où elle veut dire la même chose.
'==============================================================================

'--- Zone 1 : lblCarte1, l'intitulé global ------------------------------------
Public Function Z1Titre() As StyleTexte
    Z1Titre = Style(POLICE, 14, True, COUL_BANDEAU)
End Function
Public Function Z1Etat() As StyleTexte
    Z1Etat = Style(POLICE, TAILLE_SOUSTITRE, False, COUL_TEXTE_DOUX)
End Function
Public Function Z1Annee() As StyleTexte
    Z1Annee = Style(POLICE, 20, True, COUL_MODIFIER)
End Function

'--- Zone 2 : lblCarte2, statistiques et graphique ----------------------------
Public Function Z2Section() As StyleTexte
    Z2Section = Style(POLICE, TAILLE_SECTION, True, COUL_SECTION)
End Function
Public Function Z2TuileCap() As StyleTexte
    Z2TuileCap = Style(POLICE, TAILLE_TUILE_CAP, True, COUL_TEXTE_DOUX)
End Function
Public Function Z2TuileVal() As StyleTexte
    Z2TuileVal = Style(POLICE, TAILLE_STAT, True, COUL_BANDEAU)
End Function
Public Function Z2GraphTitre() As StyleTexte
    Z2GraphTitre = Style(POLICE, TAILLE_FILTRE, False, COUL_TEXTE_DOUX)
End Function
Public Function Z2GraphAxe() As StyleTexte
    Z2GraphAxe = Style(POLICE, TAILLE_LIBELLE, False, COUL_TEXTE_DOUX)
End Function

'--- Zone 3 : lblCarte3, la fiche intervention --------------------------------
Public Function Z3Section() As StyleTexte
    Z3Section = Style(POLICE, TAILLE_SECTION, True, COUL_SECTION)
End Function
Public Function Z3Libelle() As StyleTexte
    Z3Libelle = Style(POLICE, TAILLE_LIBELLE, True, COUL_TEXTE_DOUX)
End Function
Public Function Z3Chevron() As StyleTexte
    Z3Chevron = Style(POLICE, 9, True, COUL_BOUTON_TXT)
End Function

'--- Zone 4 : lblCarteFiltreI, la barre de filtrage ---------------------------
Public Function Z4Libelle() As StyleTexte
    Z4Libelle = Style(POLICE, TAILLE_FILTRE, True, COUL_ENTETE_TXT)
End Function
Public Function Z4Lien() As StyleTexte
    Z4Lien = Style(POLICE, TAILLE_FILTRE, False, COUL_LIEN)
End Function
Public Function Z4Compteur() As StyleTexte
    Z4Compteur = Style(POLICE, TAILLE_FILTRE, False, COUL_TEXTE_DOUX)
End Function

'--- Zone 5 : lblEnteteTableI, l'en-tête du tableau ---------------------------
' Maigre, mais dans une famille demi-grasse : MSForms ne connaît que gras ou
' pas gras, c'est donc la FAMILLE qui porte la graisse intermédiaire.
Public Function Z5Entete() As StyleTexte
    Z5Entete = Style(POLICE_ENTETE, TAILLE_GRILLE_ENTETE, False, COUL_GRILLE_ENTETE)
End Function

'--- Zone 6 : lblG_1_1 … lblG_17_12, l'intérieur du tableau -------------------
' Les 204 cases partagent ce seul style, et les bandes de fond des lignes le
' portent aussi. C'est l'exception au défaut du formulaire : neuf points
' maigres là où le reste est en huit points gras.
' La famille est écrite en clair, et non POLICE_ENTETE qui vaut pourtant la
' même chose : l'en-tête et les enregistrements doivent pouvoir changer
' séparément. GRAS_GRILLE_TXT reste False — épaissir une police déjà
' demi-grasse la ferait simuler par Windows, plus lourde et plus sale.
Public Function Z6Case() As StyleTexte
    Z6Case = Style("Segoe UI Semibold", TAILLE_GRILLE_TXT, GRAS_GRILLE_TXT, COUL_GRILLE_TXT)
End Function

'--- Zone 7 : UF_Calendrier ---------------------------------------------------
Public Function Z7Fleche() As StyleTexte
    Z7Fleche = Style(POLICE, 15, True, COUL_BANDEAU_SOUS)
End Function
Public Function Z7Mois() As StyleTexte
    Z7Mois = Style(POLICE, 11, True, COUL_BANDEAU_TXT)
End Function
Public Function Z7JourSem() As StyleTexte
    Z7JourSem = Style(POLICE, TAILLE_LIBELLE, True, COUL_TEXTE_DOUX)
End Function
Public Function Z7Jour() As StyleTexte
    Z7Jour = Style(POLICE, TAILLE_CHAMP, False, COUL_TEXTE)
End Function
Public Function Z7Lien() As StyleTexte
    Z7Lien = Style(POLICE, TAILLE_FILTRE, False, COUL_LIEN)
End Function
Public Function Z7Annuler() As StyleTexte
    Z7Annuler = Style(POLICE, TAILLE_FILTRE, False, COUL_TEXTE_DOUX)
End Function

'------------------------------------------------------------------------------
' Assemble les quatre caractéristiques. Un Const ne conviendrait pas : il ne
' peut porter qu'une valeur simple, et il faudrait quatre noms par rôle.
'------------------------------------------------------------------------------
Private Function Style(ByVal police As String, ByVal taille As Single, _
                       ByVal gras As Boolean, ByVal couleur As Long) As StyleTexte
    Style.Police = police
    Style.Taille = taille
    Style.Gras = gras
    Style.Couleur = couleur
End Function

'==============================================================================
' LES ZONES EN CADRE
'------------------------------------------------------------------------------
' Une carte portée par un cadre ne s'appelle plus lblCarteN mais fraCarteN : le
' générateur, le code des événements et la remise en arrière-plan doivent tous
' désigner le même contrôle, d'où ces deux fonctions plutôt que le nom écrit
' trois fois.
'==============================================================================
Public Function NomCarte1() As String
    NomCarte1 = IIf(F1_EN_CADRE, "fraCarte1", "lblCarte1")
End Function

Public Function NomCarte2() As String
    NomCarte2 = IIf(F2_EN_CADRE, "fraCarte2", "lblCarte2")
End Function

Public Function NomCarte3() As String
    NomCarte3 = IIf(F3_EN_CADRE, "fraCarte3", "lblCarte3")
End Function

Public Function NomCarteFiltre() As String
    NomCarteFiltre = IIf(F4_EN_CADRE, "fraCarteFiltreI", "lblCarteFiltreI")
End Function

'==============================================================================
' ORIGINE DU GRAPHIQUE DANS SON CONTENEUR
'------------------------------------------------------------------------------
' Left et Top se comptent depuis le coin du PARENT. Tant que le graphique était
' posé sur le formulaire, son origine valait GR_ORIGINE_X / GR_ORIGINE_Y ; dans
' un cadre, il faut en retrancher le coin du cadre.
'
' Le décor est posé par le générateur, les barres sont placées à l'exécution
' par Graph_Tracer : les deux DOIVENT partir du même point. C'est pour cela que
' l'origine est calculée ici, en un seul endroit, et non recopiée de part et
' d'autre — un tracé qui repartirait de l'origine absolue se retrouverait
' I_MARGE et F2_TOP trop loin, en bas à droite de son décor.
'==============================================================================
Public Function GrOrigineX() As Single
    GrOrigineX = GR_ORIGINE_X - IIf(F2_EN_CADRE, I_MARGE, 0)
End Function

Public Function GrOrigineY() As Single
    GrOrigineY = GR_ORIGINE_Y - IIf(F2_EN_CADRE, F2_TOP, 0)
End Function
