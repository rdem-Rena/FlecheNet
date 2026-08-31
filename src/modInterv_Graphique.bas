Attribute VB_Name = "modInterv_Graphique"
Option Explicit
'==============================================================================
' modInterv_Graphique
'------------------------------------------------------------------------------
' Graphique du chiffre d'affaires mensuel, tracé directement en contrôles
' MSForms plutôt qu'importé en image.
'
' Pourquoi le dessiner. L'export d'un graphique Excel produit une image à la
' résolution qu'il occupe sur la feuille : agrandie dans le formulaire, elle
' devient floue. Le tracé, lui, est net à toute taille, prend les couleurs du
' formulaire, ne laisse aucun fichier temporaire et ne dépend ni de Chart.Export
' ni de LoadPicture — les deux points qui avaient déjà posé problème.
'
' En contrepartie, le formulaire ne suit plus la mise en forme donnée au
' graphique CAGraphique dans Excel : seules ses DONNÉES sont reprises, lues dans
' le tableau Tableau7 de la feuille Statistiques.
'
' Choix de représentation : douze valeurs mensuelles, une seule série, à
' comparer entre elles — des barres verticales. Une seule teinte, donc pas de
' légende : le titre nomme la série. La grille reste discrète, et seul le
' maximum de l'échelle est écrit ; les autres valeurs se lisent au survol.
'==============================================================================

' Le nom de la table source est déclaré avec les autres noms d'objets du
' classeur, dans modInterv_Schema.

Private mMois(1 To 12) As Date      ' premier jour de chaque mois affiché
Private mCA(1 To 12) As Double
Private mNbMois As Long
Private mEchelle As Double          ' valeur au sommet de l'aire de tracé
Private mSurvol As Long             ' barre sous la souris, 0 si aucune

'==============================================================================
' Trace le graphique : lit les données, calcule l'échelle, place les barres.
'   f : le formulaire, reçu en Object comme partout ailleurs
'
' Appelée à l'ouverture puis après chaque ajout, modification ou suppression.
'==============================================================================
Public Sub Graph_Tracer(f As Object)
    Dim i As Long, maxi As Double, c As Object
    Dim gauche As Single, largeur As Single, base As Single, pas As Single

    LireDonnees

    Set c = GCtl(f, "lblGrLegende")
    If Not c Is Nothing Then
        If mNbMois = 0 Then
            c.Caption = "Chiffre d'affaires par mois  -  données indisponibles (" & _
                        NOM_TABLE_GRAPH & " introuvable)"
        Else
            c.Caption = "Chiffre d'affaires par mois"
        End If
    End If

    ' --- échelle --------------------------------------------------------------
    For i = 1 To mNbMois
        If mCA(i) > maxi Then maxi = mCA(i)
    Next i
    mEchelle = EchelleHaute(maxi)

    Set c = GCtl(f, "lblGrMax")
    If Not c Is Nothing Then c.Caption = Format$(mEchelle, "#,##0")

    ' --- barres ---------------------------------------------------------------
    ' Left et Top d'un contrôle se comptent depuis le coin du FORMULAIRE, pas
    ' depuis celui du graphique : les positions calculées ici repartent donc de
    ' GR_ORIGINE_X / GR_ORIGINE_Y, comme le décor posé par le générateur.
    '
    ' L'aire de tracé commence après la colonne de l'échelle ; chaque mois
    ' dispose d'une tranche égale, la barre est centrée dedans.
    gauche = GR_ORIGINE_X + GR_MARGE_G
    largeur = F2_GRAPH_LARG - GR_MARGE_G
    base = GR_ORIGINE_Y + GR_TRACE_TOP + GR_TRACE_HAUT
    pas = largeur / GR_NB_MOIS

    For i = 1 To GR_NB_MOIS
        PlacerBarre f, i, gauche + (i - 1) * pas + (pas - GR_BARRE_LARG) / 2, base
        EtiqueterMois f, i, gauche + (i - 1) * pas, pas
    Next i

    mSurvol = 0
End Sub

'------------------------------------------------------------------------------
' Positionne une barre : hauteur proportionnelle à la valeur, ancrée sur la
' ligne de base. Une valeur nulle donne une barre de hauteur nulle, donc
' invisible — ce qui est exactement ce qu'il faut dire d'un mois sans activité.
'------------------------------------------------------------------------------
Private Sub PlacerBarre(f As Object, ByVal i As Long, ByVal x As Single, ByVal base As Single)
    Dim c As Object, h As Single

    Set c = GCtl(f, "lblGrBarre_" & CStr(i))
    If c Is Nothing Then Exit Sub

    If i > mNbMois Or mEchelle <= 0 Then
        c.Visible = False
        Exit Sub
    End If

    h = GR_TRACE_HAUT * (mCA(i) / mEchelle)
    If h < 0 Then h = 0
    If h > GR_TRACE_HAUT Then h = GR_TRACE_HAUT

    c.Left = x
    c.Width = GR_BARRE_LARG
    c.Height = IIf(h < 1, 0, h)
    c.Top = base - c.Height
    c.Visible = (h >= 1)
    c.ControlTipText = TexteValeur(i)
End Sub

'------------------------------------------------------------------------------
' Écrit le nom abrégé du mois sous sa barre. Le mois en cours est mis en gras :
' c'est une distinction d'état, pas une série de plus, donc pas de teinte
' différente.
'------------------------------------------------------------------------------
Private Sub EtiqueterMois(f As Object, ByVal i As Long, ByVal x As Single, ByVal pas As Single)
    Dim c As Object, courant As Boolean

    Set c = GCtl(f, "lblGrMois_" & CStr(i))
    If c Is Nothing Then Exit Sub

    c.Left = x
    c.Width = pas
    If i > mNbMois Then
        c.Caption = vbNullString
        Exit Sub
    End If

    courant = (Year(mMois(i)) = Year(Date) And Month(mMois(i)) = Month(Date))
    c.Caption = MoisAbrege(Month(mMois(i)))
    c.Font.Bold = courant
    c.ForeColor = IIf(courant, COUL_TEXTE, COUL_TEXTE_DOUX)
End Sub

'==============================================================================
' Survol d'une barre : le mois et son montant s'écrivent dans le bandeau, à la
' place du titre.
'   indice : numéro de barre, ou 0 pour revenir au titre
'
' Une seule barre change d'aspect à la fois : on ne repeint que celle qu'on
' quitte et celle qu'on prend, jamais les douze.
'
' Fond ET bordure changent ensemble : le générateur pose les barres avec une
' bordure de la même teinte que leur remplissage, ne toucher qu'au fond
' laisserait un liseré foncé autour d'une barre éclaircie.
'==============================================================================
Public Sub Graph_Survol(f As Object, ByVal indice As Long)
    Dim c As Object

    If indice = mSurvol Then Exit Sub

    If mSurvol >= 1 And mSurvol <= GR_NB_MOIS Then
        Set c = GCtl(f, "lblGrBarre_" & CStr(mSurvol))
        If Not c Is Nothing Then
            c.BackColor = COUL_GR_BARRE
            c.BorderColor = COUL_GR_BARRE
        End If
    End If

    mSurvol = indice

    Set c = GCtl(f, "lblGrLegende")
    If Not c Is Nothing Then
        If indice >= 1 And indice <= mNbMois Then
            c.Caption = TexteValeur(indice)
            c.ForeColor = COUL_TEXTE
            c.Font.Bold = True
        Else
            c.Caption = "Chiffre d'affaires par mois"
            c.ForeColor = COUL_TEXTE_DOUX
            c.Font.Bold = False
        End If
    End If

    If indice >= 1 And indice <= GR_NB_MOIS Then
        Set c = GCtl(f, "lblGrBarre_" & CStr(indice))
        If Not c Is Nothing Then
            c.BackColor = COUL_MODIFIER_H
            c.BorderColor = COUL_MODIFIER_H
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' « Mai 2026  -  11'560 CHF »
'------------------------------------------------------------------------------
Private Function TexteValeur(ByVal i As Long) As String
    If i < 1 Or i > mNbMois Then Exit Function
    TexteValeur = NomDuMois(Month(mMois(i))) & " " & Format$(mMois(i), "yyyy") & _
                  "  -  " & Format$(mCA(i), "#,##0") & " CHF"
End Function

'==============================================================================
' Lecture des douze valeurs mensuelles dans le tableau Tableau7 de la feuille
' Statistiques — celui-là même qui alimente le graphique Excel.
'==============================================================================
Private Sub LireDonnees()
    Dim lo As ListObject, v As Variant, i As Long
    Dim icMois As Long, icCA As Long, n As Long

    mNbMois = 0
    For i = 1 To 12
        mCA(i) = 0
    Next i

    Set lo = ObtenirTable(NOM_TABLE_GRAPH)
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    icMois = IndexColonne(lo, "Mois")
    icCA = IndexColonne(lo, "CA")
    If icMois = 0 Or icCA = 0 Then Exit Sub

    v = lo.DataBodyRange.Value
    If Not IsArray(v) Then Exit Sub

    For i = LBound(v, 1) To UBound(v, 1)
        If n >= 12 Then Exit For
        If IsDate(v(i, icMois)) Then
            n = n + 1
            mMois(n) = CDate(v(i, icMois))
            If IsNumeric(v(i, icCA)) Then mCA(n) = CDbl(v(i, icCA))
        End If
    Next i
    mNbMois = n
End Sub

'==============================================================================
' Sommet d'échelle « rond » immédiatement au-dessus du maximum : 13'260 donne
' 15'000, 21'918 donne 25'000, 8'400 donne 10'000. Une échelle calée pile sur
' le maximum ferait toucher le haut de l'aire de tracé à la plus grande barre.
'
' Les six échelons (1 - 1,5 - 2 - 2,5 - 5 - 10) sont assez serrés pour que la
' plus grande barre occupe rarement moins des deux tiers de la hauteur
' disponible : une échelle trop large écraserait tout le graphique en bas.
'==============================================================================
Public Function EchelleHaute(ByVal maxi As Double) As Double
    Dim p As Double, m As Double

    If maxi <= 0 Then
        EchelleHaute = 1
        Exit Function
    End If

    p = 10 ^ Int(Log(maxi) / Log(10#))      ' puissance de 10 juste en dessous
    m = maxi / p
    If m <= 1 Then
        m = 1
    ElseIf m <= 1.5 Then
        m = 1.5
    ElseIf m <= 2 Then
        m = 2
    ElseIf m <= 2.5 Then
        m = 2.5
    ElseIf m <= 5 Then
        m = 5
    Else
        m = 10
    End If
    EchelleHaute = m * p
End Function

'------------------------------------------------------------------------------
' Accès sûr à un contrôle du graphique.
'------------------------------------------------------------------------------
Private Function GCtl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set GCtl = f.Controls(nom)
    On Error GoTo 0
End Function
