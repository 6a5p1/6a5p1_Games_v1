Attribute VB_Name = "Module1"
Public ccc As String
Public Language As Integer
Public WhichAbout As Integer
Public WhichHelp As Integer
Public WhichQuit As Integer
Public WhichWell As Integer
Public opnd As Boolean

Public NumFicConf As Integer
Public NomFicConf As String

Public nil(81), k, aux As Integer
Public BitMask(9) As Integer
Public TabVal(9, 9) As Integer
Public TabInit(9, 9) As Integer
Public TabPossible(9, 9) As Integer
Public TabHypothese(9, 9) As Integer
Public TabInvHypothese(9, 9) As Integer
Public MasqueHypothese As Integer
Public MasqueInvHypothese As Integer
Public TabPaires(36, 2) As Integer
Public TabTrios(84, 3) As Integer
Public TabQuartets(126, 4) As Integer
Public NbMaxGrp(4) As Integer

Public Const MASK_HYPOTHESE As Integer = &H400&
Public Const MASK_ISOLE As Integer = &H200&
Public Const MASK123456789 As Integer = &H1FF&

Public AucuneAide As Boolean
Public AideValPossible As Boolean
Public AideIndIsole As Boolean
Public AideValIsole As Boolean
Public AideRecherchePaires As Boolean
Public AideRechJumeauxTriplets As Boolean
Public AidePlacementAuto As Boolean
Public AideRechercheTrios As Boolean
Public AideRechercheQuartets As Boolean
Public AidePropagation As Boolean

Public NbGrilleOK As Integer

Public Const TAILLE_LIGNE = 9

Public Const LRG_CASEGRILLE = 12
Public Const HTR_CASEGRILLE = 12
Public ImprimeValeurDispo As Boolean

Public NbCasesPlacees As Integer
Public NbCasesInitiales As Integer

Public Const Level_Easy = 1
Public Const Level_Medium = 2
Public Const Level_Hard = 3

Function ligne(l As Integer) As Integer
Dim c As Integer
For c = 1 To 9
    ligne = ligne + (TabVal(l, c) And MASK123456789)
Next c
End Function

Function col(c As Integer) As Integer
Dim l As Integer
For l = 1 To 9
    col = col + (TabVal(l, c) And MASK123456789)
Next l
End Function

Function reg(L_reg As Integer, C_reg As Integer) As Integer
Dim l As Integer
Dim c As Integer
For l = L_reg To L_reg + 2
    For c = C_reg To C_reg + 2
        reg = reg + (TabVal(l, c) And MASK123456789)
    Next c
Next l
End Function

Sub InitGrille()
Dim i As Integer
Dim l As Integer
Dim c As Integer
Dim L_reg As Integer
Dim C_reg As Integer
BitMask(0) = 0
For i = 0 To 8
    BitMask(i + 1) = 2 ^ i
Next i
For l = 1 To 9
    For c = 1 To 9
        TabPossible(l, c) = MASK123456789
        TabHypothese(l, c) = 0
        TabInvHypothese(l, c) = MASK123456789
        TabInit(l, c) = 0
        TabVal(l, c) = 0
    Next c
Next l
MasqueHypothese = 0
MasqueInvHypothese = MASK123456789
InitialiseTabPaires
InitialiseTabTrios
InitialiseTabQuartets
NbMaxGrp(0) = 0
NbMaxGrp(1) = 0
NbMaxGrp(2) = 36
NbMaxGrp(3) = 84
NbMaxGrp(4) = 126
End Sub

Sub RAZGrille()
Dim i As Integer
Dim l As Integer
Dim c As Integer
For l = 1 To 9
    For c = 1 To 9
        TabPossible(l, c) = MASK123456789
        TabHypothese(l, c) = 0
        TabInvHypothese(l, c) = MASK123456789
        TabVal(l, c) = 0
    Next c
Next l
End Sub

Function MiseAJourValeursPossibles() As Boolean
Dim i As Integer
Dim l As Integer
Dim c As Integer
Dim L_reg As Integer
Dim C_reg As Integer
Dim ValeurIsole As Integer
Dim ValeurPossible As Integer
Dim ValeurImpossible As Integer
Dim svg_ValeurPossible As Integer
Dim IndexPaires As Integer
MiseAJourValeursPossibles = False
For l = 1 To 9
    For c = 1 To 9
        If TabVal(l, c) <> MASK_HYPOTHESE Then
            If (TabVal(l, c) And MASK123456789) = 0 Then
                svg_ValeurPossible = TabPossible(l, c) And MASK123456789
                If AideValPossible Then
                    ValeurPossible = MASK123456789
                Else
                    ValeurPossible = svg_ValeurPossible
                End If
                L_reg = 1 + ((l - 1) \ 3) * 3
                C_reg = 1 + ((c - 1) \ 3) * 3
                ValeurImpossible = ligne(l) Or col(c) Or reg(L_reg, C_reg)
                TabPossible(l, c) = (ValeurPossible Or ValeurImpossible) - ValeurImpossible
                If (TabPossible(l, c) And MASK123456789) <> svg_ValeurPossible Then
                    MiseAJourValeursPossibles = True
                End If
            Else
                TabPossible(l, c) = TabVal(l, c) And MASK123456789
            End If
        Else
            i = 0
        End If
    Next c
Next l
End Function

Function RechercheIsole(Nettoyage As Boolean) As Boolean
Dim l As Integer, c As Integer
Dim ValeurIsole As Integer
Dim svg_ValeurPossible As Integer
RechercheIsole = False
For l = 1 To 9
    For c = 1 To 9
        If TabVal(l, c) <> MASK_HYPOTHESE Then
            If (TabVal(l, c) And MASK123456789) = 0 Then
                svg_ValeurPossible = TabPossible(l, c) And MASK123456789
                ValeurIsole = RechercheIsoleLigne(l, c)
                If ValeurIsole Then
                    If Nettoyage Then
                        TabPossible(l, c) = BitMask(ValeurIsole) + MASK_ISOLE
                    Else
                        TabPossible(l, c) = TabPossible(l, c) + MASK_ISOLE
                    End If
                Else
                    ValeurIsole = RechercheIsoleColonne(l, c)
                    If ValeurIsole Then
                        If Nettoyage Then
                            TabPossible(l, c) = BitMask(ValeurIsole) + MASK_ISOLE
                        Else
                            TabPossible(l, c) = TabPossible(l, c) + MASK_ISOLE
                        End If
                    Else
                        ValeurIsole = RechercheIsoleRegion(l, c)
                        If ValeurIsole Then
                            If Nettoyage Then
                                TabPossible(l, c) = BitMask(ValeurIsole) + MASK_ISOLE
                            Else
                                TabPossible(l, c) = TabPossible(l, c) + MASK_ISOLE
                            End If
                        End If
                    End If
                End If
                If (TabPossible(l, c) And MASK123456789) <> svg_ValeurPossible Then RechercheIsole = True
            End If
        End If
    Next c
Next l
End Function

Function RechercheIsoleLigne(l As Integer, c As Integer) As Integer
Dim isole As Boolean
Dim col As Integer, i As Integer
Dim MasquePossible As Integer
RechercheIsoleLigne = 0
MasquePossible = TabPossible(l, c) And MASK123456789
For i = 1 To 9
    If MasquePossible And BitMask(i) Then
        isole = True
        For col = 1 To 9
            If col <> c Then
                If TabPossible(l, col) And BitMask(i) Then
                    isole = False
                    Exit For
                End If
            End If
        Next col
        If isole Then
            RechercheIsoleLigne = i
            Exit For
        End If
    End If
Next i
End Function

Function RechercheIsoleColonne(l As Integer, c As Integer) As Integer
Dim isole As Boolean
Dim ligne As Integer, i As Integer
Dim MasquePossible As Integer
RechercheIsoleColonne = 0
MasquePossible = TabPossible(l, c) And MASK123456789
For i = 1 To 9
    If MasquePossible And BitMask(i) Then
        isole = True
        For ligne = 1 To 9
            If ligne <> l Then
                If TabPossible(ligne, c) And BitMask(i) Then
                    isole = False
                    Exit For
                End If
            End If
        Next ligne
        If isole Then
            RechercheIsoleColonne = i
            Exit For
        End If
    End If
Next i
End Function

Function RechercheIsoleRegion(l As Integer, c As Integer) As Integer
Dim isole As Boolean
Dim i As Integer
Dim ligne As Integer
Dim col As Integer
Dim L_reg As Integer
Dim C_reg As Integer
Dim MasquePossible As Integer
L_reg = 1 + ((l - 1) \ 3) * 3
C_reg = 1 + ((c - 1) \ 3) * 3
RechercheIsoleRegion = 0
MasquePossible = TabPossible(l, c) And MASK123456789
For i = 1 To 9
    If MasquePossible And BitMask(i) Then
        isole = True
        For ligne = L_reg To L_reg + 2
            For col = C_reg To C_reg + 2
                If Not (ligne = l And col = c) Then
                    If TabPossible(ligne, col) And BitMask(i) Then
                        isole = False
                        Exit For
                    End If
                End If
            Next col
            If Not isole Then Exit For
        Next ligne
        If isole Then
            RechercheIsoleRegion = i
            Exit For
        End If
    End If
Next i
End Function

Function ControleLigneOK(l As Integer) As Boolean
Dim col As Integer, c As Integer
ControleLigneOK = True
For c = 1 To 9
    For col = 1 To 9
        If col <> c Then
            If (TabVal(l, c) And MASK123456789) <> 0 Then
                If TabVal(l, c) = TabVal(l, col) Then
                    ControleLigneOK = False
                    Exit Function
                End If
            End If
        End If
    Next col
Next c
End Function

Function ControleColonneOK(c As Integer) As Boolean
Dim ligne As Integer, l As Integer
ControleColonneOK = True
For l = 1 To 9
    For ligne = 1 To 9
        If ligne <> l Then
            If (TabVal(l, c) And MASK123456789) <> 0 Then
                If TabVal(l, c) = TabVal(ligne, c) Then
                    ControleColonneOK = False
                    Exit Function
                End If
            End If
        End If
    Next ligne
Next l
End Function

Function ControleRegOK(L_reg As Integer, C_reg As Integer) As Boolean
Dim ligne As Integer, l As Integer
Dim col As Integer, c As Integer
ControleRegOK = True
For l = L_reg To L_reg + 2
    For c = C_reg To C_reg + 2
        For ligne = L_reg To L_reg + 2
            For col = C_reg To C_reg + 2
                If ligne <> l And col <> c Then
                    If (TabVal(l, c) And MASK123456789) <> 0 Then
                        If TabVal(l, c) = TabVal(ligne, col) Then
                            ControleRegOK = False
                            Exit Function
                        End If
                    End If
                End If
            Next col
        Next ligne
    Next c
Next l
End Function

Function PlacementCaseDefinie() As Boolean
Dim l As Integer, c As Integer, i As Integer
Dim NbPossible As Integer, Masque As Integer, MasquePossible As Integer
PlacementCaseDefinie = False
For l = 1 To 9
    For c = 1 To 9
        If TabVal(l, c) = 0 Then
            MasquePossible = TabPossible(l, c) And MASK123456789
            If MasquePossible <> 0 Then
                NbPossible = 0
                For i = 1 To 9
                    If MasquePossible And BitMask(i) Then
                        NbPossible = NbPossible + 1
                        Masque = BitMask(i)
                    End If
                Next i
                If NbPossible = 1 Then
                    PlacementCaseDefinie = True
                    TabVal(l, c) = Masque
                    NbCasesPlacees = NbCasesPlacees + 1
                End If
            End If
        End If
    Next c
Next l
End Function

Sub MiseAJourGrille()
Dim Optimisation As Boolean
Do
    MiseAJourValeursPossibles
    Optimisation = OptimiserValeursPossibles()
Loop While (AidePropagation And Optimisation)
End Sub

Function OptimiserValeursPossibles() As Boolean
Dim Changement As Boolean
Dim Continue As Boolean
Dim ChangementIsole As Boolean
Dim ChangementJ_T As Boolean
Dim ChangementPaires As Boolean
OptimiserValeursPossibles = False
If AideIndIsole Or AideValIsole Then
    ChangementIsole = RechercheIsole(AideValIsole)
    OptimiserValeursPossibles = OptimiserValeursPossibles + ChangementIsole
End If
If AideRechJumeauxTriplets Then
    ChangementJ_T = NettoyerMonoLigneOuCol()
    OptimiserValeursPossibles = OptimiserValeursPossibles + ChangementJ_T
End If
If AideRecherchePaires Then
    ChangementPaires = NettoyerGroupe(2)
    OptimiserValeursPossibles = OptimiserValeursPossibles + ChangementPaires
End If
If AideRechercheTrios Then
    ChangementPaires = NettoyerGroupe(3)
    OptimiserValeursPossibles = OptimiserValeursPossibles + ChangementPaires
End If
If AideRechercheQuartets Then
    ChangementPaires = NettoyerGroupe(4)
    OptimiserValeursPossibles = OptimiserValeursPossibles + ChangementPaires
End If
If AidePlacementAuto Then
    OptimiserValeursPossibles = OptimiserValeursPossibles + PlacementCaseDefinie()
End If
End Function

Sub InitialiseTabPaires()
Dim i As Integer, j As Integer
Dim Index As Integer
Index = 0
For i = 1 To 8
    For j = i + 1 To 9
        Index = Index + 1
        TabPaires(Index, 0) = BitMask(i) Or BitMask(j)
        TabPaires(Index, 1) = i
        TabPaires(Index, 2) = j
    Next j
Next i
End Sub

Sub InitialiseTabTrios()
Dim i As Integer, j As Integer, k As Integer
Dim Index As Integer
Index = 0
For i = 1 To 7
    For j = i + 1 To 8
        For k = j + 1 To 9
            Index = Index + 1
            TabTrios(Index, 0) = BitMask(i) Or BitMask(j) Or BitMask(k)
            TabTrios(Index, 1) = i
            TabTrios(Index, 2) = j
            TabTrios(Index, 3) = k
        Next k
    Next j
Next i
End Sub

Sub InitialiseTabQuartets()
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim Index As Integer
Index = 0
For i = 1 To 6
    For j = i + 1 To 7
        For k = j + 1 To 8
            For l = k + 1 To 9
                Index = Index + 1
                TabQuartets(Index, 0) = BitMask(i) Or BitMask(j) Or BitMask(k) Or BitMask(l)
                TabQuartets(Index, 1) = i
                TabQuartets(Index, 2) = j
                TabQuartets(Index, 3) = k
                TabQuartets(Index, 4) = l
            Next l
        Next k
    Next j
Next i
End Sub

Function RechercheValeurMonoLigneRegion(L_reg As Integer, C_reg As Integer, ligne As Integer) As Integer
Dim i As Integer
Dim col As Integer
Dim Svg_ligne As Integer
Dim MasquePossible As Integer
RechercheValeurMonoLigneRegion = 0
For i = 1 To 9
    Svg_ligne = 0
    For ligne = L_reg To L_reg + 2
        For col = C_reg To C_reg + 2
            If (TabVal(ligne, col) And MASK123456789) = 0 Then
                MasquePossible = TabPossible(ligne, col) And MASK123456789
                If MasquePossible And BitMask(i) Then
                    If Svg_ligne = 0 Then
                        Svg_ligne = ligne
                    ElseIf Svg_ligne <> ligne Then
                        Svg_ligne = -1
                    End If
                End If
            End If
        Next col
    Next ligne
    If Svg_ligne > 0 Then
        RechercheValeurMonoLigneRegion = i
        ligne = Svg_ligne
        Exit For
    End If
Next i
End Function

Function RechercheValeurMonoColRegion(L_reg As Integer, C_reg As Integer, col As Integer) As Integer
Dim i As Integer
Dim ligne As Integer
Dim Svg_col As Integer
Dim MasquePossible As Integer
RechercheValeurMonoColRegion = 0
For i = 1 To 9
    Svg_col = 0
    For ligne = L_reg To L_reg + 2
        For col = C_reg To C_reg + 2
            If (TabVal(ligne, col) And MASK123456789) = 0 Then
                MasquePossible = TabPossible(ligne, col) And MASK123456789
                If MasquePossible And BitMask(i) Then
                    If Svg_col = 0 Then
                        Svg_col = col
                    ElseIf Svg_col <> col Then
                        Svg_col = -1
                    End If
                End If
            End If
        Next col
    Next ligne
    If Svg_col > 0 Then
        RechercheValeurMonoColRegion = i
        col = Svg_col
        Exit For
    End If
Next i
End Function

Function NettoyageValeurMonoLigneRegion(C_reg As Integer, ligne As Integer, val As Integer) As Boolean
Dim C_reg_loc As Integer
Dim MasquePossible As Integer
Dim col As Integer
NettoyageValeurMonoLigneRegion = False
For col = 1 To 9
    C_reg_loc = 1 + ((col - 1) \ 3) * 3
    If C_reg_loc <> C_reg Then
        MasquePossible = TabPossible(ligne, col) And MASK123456789
        If MasquePossible And BitMask(val) Then
            TabPossible(ligne, col) = TabPossible(ligne, col) Xor BitMask(val)
            NettoyageValeurMonoLigneRegion = True
        End If
    End If
Next col
End Function

Function NettoyageValeurMonoColRegion(L_reg As Integer, col As Integer, val As Integer) As Boolean
Dim L_reg_loc As Integer
Dim MasquePossible As Integer
Dim ligne As Integer
NettoyageValeurMonoColRegion = False
For ligne = 1 To 9
    L_reg_loc = 1 + ((ligne - 1) \ 3) * 3
    If L_reg_loc <> L_reg Then
        MasquePossible = TabPossible(ligne, col) And MASK123456789
        If MasquePossible And BitMask(val) Then
            TabPossible(ligne, col) = TabPossible(ligne, col) Xor BitMask(val)
            NettoyageValeurMonoColRegion = True
        End If
    End If
Next ligne
End Function

Function NettoyerMonoLigneOuCol() As Boolean
Dim L_reg As Integer
Dim C_reg As Integer
Dim ligne As Integer
Dim col As Integer
Dim val As Integer
Dim Changement As Boolean
NettoyerMonoLigneOuCol = False
For L_reg = 1 To 7 Step 3
    For C_reg = 1 To 7 Step 3
        val = RechercheValeurMonoLigneRegion(L_reg, C_reg, ligne)
        If val > 0 Then
            Changement = NettoyageValeurMonoLigneRegion(C_reg, ligne, val)
            If Changement Then
                NettoyerMonoLigneOuCol = True
            End If
        End If
        val = RechercheValeurMonoColRegion(L_reg, C_reg, col)
        If val > 0 Then
            Changement = NettoyageValeurMonoColRegion(L_reg, col, val)
            If Changement Then
                NettoyerMonoLigneOuCol = True
            End If
        End If
    Next C_reg
Next L_reg
End Function

Function RechercheGroupeLigne(l As Integer, Nb As Integer) As Integer
Dim j As Integer, c As Integer
Dim NbCase As Integer
RechercheGroupeLigne = 0
For j = 1 To NbMaxGrp(Nb)
    NbCase = 0
    Select Case Nb
        Case 2
            ValGroupe = TabPaires(j, 0)
        Case 3
            ValGroupe = TabTrios(j, 0)
        Case 4
            ValGroupe = TabQuartets(j, 0)
        Case Else
            Exit Function
    End Select
    For c = 1 To 9
        If (TabVal(l, c) And MASK123456789) = 0 Then
            If (TabPossible(l, c) Or ValGroupe) = ValGroupe Then
                NbCase = NbCase + 1
            End If
        End If
    Next c
    If NbCase = Nb Then
        RechercheGroupeLigne = j
        Exit For
    End If
Next j
End Function

Function RechercheGroupeCol(c As Integer, Nb As Integer) As Integer
Dim j As Integer, l As Integer
Dim NbCase As Integer
RechercheGroupeCol = 0
For j = 1 To NbMaxGrp(Nb)
    NbCase = 0
    Select Case Nb
        Case 2
            ValGroupe = TabPaires(j, 0)
        Case 3
            ValGroupe = TabTrios(j, 0)
        Case 4
            ValGroupe = TabQuartets(j, 0)
        Case Else
            Exit Function
    End Select
    For l = 1 To 9
        If (TabVal(l, c) And MASK123456789) = 0 Then
            If (TabPossible(l, c) Or ValGroupe) = ValGroupe Then
                NbCase = NbCase + 1
            End If
        End If
    Next l
    If NbCase = Nb Then
        RechercheGroupeCol = j
        Exit For
    End If
Next j
End Function

Function RechercheGroupeReg(L_reg As Integer, C_reg As Integer, Nb As Integer) As Integer
Dim l As Integer
Dim c As Integer
Dim j As Integer
Dim NbCase As Integer
RechercheGroupeReg = 0
For j = 1 To NbMaxGrp(Nb)
    NbCase = 0
    Select Case Nb
        Case 2
            ValGroupe = TabPaires(j, 0)
        Case 3
            ValGroupe = TabTrios(j, 0)
        Case 4
            ValGroupe = TabQuartets(j, 0)
        Case Else
            Exit Function
    End Select
    For l = L_reg To L_reg + 2
        For c = C_reg To C_reg + 2
            If (TabVal(l, c) And MASK123456789) = 0 Then
                If (TabPossible(l, c) Or ValGroupe) = ValGroupe Then
                    NbCase = NbCase + 1
                End If
            End If
        Next c
    Next l
    If NbCase = Nb Then
        RechercheGroupeReg = j
        Exit For
    End If
Next j
End Function

Function NettoyerGroupeLigne(l As Integer, j As Integer, Nb As Integer) As Boolean
Dim c As Integer, val As Integer
Dim ValPossible As Integer
Dim ValPossibleIsole As Integer
Dim BitVal(4) As Integer
NettoyerGroupeLigne = False
Select Case Nb
    Case 2
        ValGroupe = TabPaires(j, 0)
        BitVal(1) = BitMask(TabPaires(j, 1))
        BitVal(2) = BitMask(TabPaires(j, 2))
    Case 3
        ValGroupe = TabTrios(j, 0)
        BitVal(1) = BitMask(TabTrios(j, 1))
        BitVal(2) = BitMask(TabTrios(j, 2))
        BitVal(3) = BitMask(TabTrios(j, 3))
    Case 4
        ValGroupe = TabQuartets(j, 0)
        BitVal(1) = BitMask(TabQuartets(j, 1))
        BitVal(2) = BitMask(TabQuartets(j, 2))
        BitVal(3) = BitMask(TabQuartets(j, 3))
        BitVal(4) = BitMask(TabQuartets(j, 4))
    Case Else
        Exit Function
End Select
For c = 1 To 9
    ValPossible = TabPossible(l, c) And MASK123456789
    If (ValPossible Or ValGroupe) <> ValGroupe Then
        For val = 1 To Nb
            If ValPossible And BitVal(val) Then
                TabPossible(l, c) = TabPossible(l, c) Xor BitVal(val)
                NettoyerGroupeLigne = True
            End If
        Next val
    End If
Next c
End Function

Function NettoyerGroupeCol(c As Integer, j As Integer, Nb As Integer) As Boolean
Dim l As Integer, val As Integer
Dim ValPossible As Integer
Dim ValPossibleIsole As Integer
Dim BitVal(4) As Integer
NettoyerGroupeCol = False
Select Case Nb
    Case 2
        ValGroupe = TabPaires(j, 0)
        BitVal(1) = BitMask(TabPaires(j, 1))
        BitVal(2) = BitMask(TabPaires(j, 2))
    Case 3
        ValGroupe = TabTrios(j, 0)
        BitVal(1) = BitMask(TabTrios(j, 1))
        BitVal(2) = BitMask(TabTrios(j, 2))
        BitVal(3) = BitMask(TabTrios(j, 3))
    Case 4
        ValGroupe = TabQuartets(j, 0)
        BitVal(1) = BitMask(TabQuartets(j, 1))
        BitVal(2) = BitMask(TabQuartets(j, 2))
        BitVal(3) = BitMask(TabQuartets(j, 3))
        BitVal(4) = BitMask(TabQuartets(j, 4))
    Case Else
        Exit Function
End Select
For l = 1 To 9
    ValPossible = TabPossible(l, c) And MASK123456789
    If (ValPossible Or ValGroupe) <> ValGroupe Then
        For val = 1 To Nb
            If ValPossible And BitVal(val) Then
                TabPossible(l, c) = TabPossible(l, c) Xor BitVal(val)
                NettoyerGroupeCol = True
            End If
        Next val
    End If
Next l
End Function

Function NettoyerGroupeReg(L_reg As Integer, C_reg As Integer, j As Integer, Nb As Integer) As Boolean
Dim l As Integer, val As Integer
Dim ValPossible As Integer
Dim ValPossibleIsole As Integer
Dim BitVal(4) As Integer
NettoyerGroupeReg = False
Select Case Nb
    Case 2
        ValGroupe = TabPaires(j, 0)
        BitVal(1) = BitMask(TabPaires(j, 1))
        BitVal(2) = BitMask(TabPaires(j, 2))
    Case 3
        ValGroupe = TabTrios(j, 0)
        BitVal(1) = BitMask(TabTrios(j, 1))
        BitVal(2) = BitMask(TabTrios(j, 2))
        BitVal(3) = BitMask(TabTrios(j, 3))
    Case 4
        ValGroupe = TabQuartets(j, 0)
        BitVal(1) = BitMask(TabQuartets(j, 1))
        BitVal(2) = BitMask(TabQuartets(j, 2))
        BitVal(3) = BitMask(TabQuartets(j, 3))
        BitVal(4) = BitMask(TabQuartets(j, 4))
    Case Else
        Exit Function
End Select
For l = L_reg To L_reg + 2
    For c = C_reg To C_reg + 2
        ValPossible = TabPossible(l, c) And MASK123456789
        If (ValPossible Or ValGroupe) <> ValGroupe Then
            For val = 1 To Nb
                If ValPossible And BitVal(val) Then
                    TabPossible(l, c) = TabPossible(l, c) Xor BitVal(val)
                    NettoyerGroupeReg = True
                End If
            Next val
        End If
    Next c
Next l
End Function

Function NettoyerGroupe(Nb As Integer) As Boolean
Dim l As Integer
Dim c As Integer
Dim ValPossible As Integer
Dim IndexGroupe As Integer
NettoyerGroupe = False
For l = 1 To 9
    IndexGroupe = RechercheGroupeLigne(l, Nb)
    If IndexGroupe > 0 Then
        NettoyerGroupe = NettoyerGroupeLigne(l, IndexGroupe, Nb)
    End If
Next
For c = 1 To 9
    IndexGroupe = RechercheGroupeCol(c, Nb)
    If IndexGroupe > 0 Then
        NettoyerGroupe = NettoyerGroupeCol(c, IndexGroupe, Nb)
    End If
Next
For l = 1 To 7 Step 3
    For c = 1 To 7 Step 3
        IndexGroupe = RechercheGroupeReg(l, c, Nb)
        If IndexGroupe > 0 Then
            NettoyerGroupe = NettoyerGroupeReg(l, c, IndexGroupe, Nb)
        End If
    Next c
Next l
End Function

Public Function ControleGrille() As Boolean
Dim ligne As Integer, l As Integer, L_reg As Integer
Dim col As Integer, c As Integer, C_reg As Integer
ControleGrille = True
For l = 1 To 9
    ControleGrille = ControleLigneOK(l)
    If Not ControleGrille Then Exit Function
Next
For c = 1 To 9
    ControleGrille = ControleColonneOK(c)
    If Not ControleGrille Then Exit Function
Next
For l = 1 To 7 Step 3
    For c = 1 To 7 Step 3
        ControleGrille = ControleRegOK(l, c)
        If Not ControleGrille Then Exit Function
    Next c
Next l
End Function

Sub CreerGrille(niveau As Integer)
Dim i As Integer, j As Integer
Dim ligne As Integer, col As Integer
Dim SvgFlagAide As Boolean
Dim Bcl As Integer
Dim StringNiveau As String
SvgFlagAide = AucuneAide
InitGrille
RAZGrille
While (Not ConstruireGrille(niveau))
    Bcl = Bcl + 1
    DoEvents
Wend
Sudoku.Init = False
RAZGrille
End Sub

Function ConstruireGrille(niveau As Integer) As Boolean
Dim i As Integer, j As Integer, lr As Integer, cr As Integer
Dim ligne As Integer, col As Integer
Dim CaseOK As Boolean
Dim valeur As Integer
Dim ValeurOK As Boolean
Dim Tps As Integer
Dim ErrReg As Integer
Dim ErrGrille As Integer
Dim GrilleOK As Boolean
Randomize Timer
InitGrille
PositionneNiveau (niveau)
ligne = 1 + Int(Rnd * 9)
col = 1 + Int(Rnd * 9)
valeur = 1 + Int(Rnd * 9)
TabInit(ligne, col) = BitMask(valeur)
TabVal(ligne, col) = BitMask(valeur)
MiseAJourValeursPossibles
NbCasesPlacees = 1
NbCasesInitiales = 1
ConstruireGrille = True
While (NbCasesPlacees < 81 And ConstruireGrille)
    DoEvents
    For lr = 1 To 7 Step 3
        For cr = 1 To 7 Step 3
            CaseOK = False
            ligne = lr + Int(Rnd * 3)
            col = cr + Int(Rnd * 3)
            If TabVal(ligne, col) = 0 Then
                CaseOK = True
                If (TabPossible(ligne, col) And MASK123456789) <> 0 Then
                    valeur = PlaceValeur(ligne, col, valeur)
                Else
                    ConstruireGrille = False
                End If
                MiseAJourGrille
            End If
            If Not (NbCasesPlacees < 81 And ConstruireGrille) Then Exit For
        Next cr
        If Not (NbCasesPlacees < 81 And ConstruireGrille) Then Exit For
    Next lr
Wend
If ConstruireGrille Then
    If ControleGrille() Then
        PositionneNiveau (niveau)
        ConstruireGrille = ReConstruireGrille(niveau)
        If niveau = Level_Easy Then
            If NbCasesInitiales < 30 Then ConstruireGrille = False
        ElseIf niveau = Level_Medium Then
            PositionneNiveau (Level_Easy)
            RAZGrille
            NbCasesPlacees = 0
            For l = 1 To 9
                For c = 1 To 9
                    TabVal(l, c) = TabInit(l, c)
                    If TabInit(l, c) <> 0 Then NbCasesPlacees = NbCasesPlacees + 1
                Next c
            Next l
            MiseAJourGrille
            If NbCasesPlacees > 70 Then ConstruireGrille = False
            If NbCasesInitiales < 28 Then ConstruireGrille = False
        ElseIf niveau = Level_Hard Then
            PositionneNiveau (Level_Medium)
            RAZGrille
            NbCasesPlacees = 0
            For l = 1 To 9
                For c = 1 To 9
                    TabVal(l, c) = TabInit(l, c)
                    If TabInit(l, c) <> 0 Then NbCasesPlacees = NbCasesPlacees + 1
                Next c
            Next l
            MiseAJourGrille
            If NbCasesPlacees > 70 Then ConstruireGrille = False
            End If
        Else
            ConstruireGrille = False
        End If
    End If
k = 0
For i = 1 To 9
    For j = 1 To 9
        If (TabInit(i, j)) <> 0 Then
            aux = Log(TabInit(i, j)) / Log(2) + 1
        Else
            aux = 0
        End If
        nil(k) = aux
        k = k + 1
    Next j
Next i
End Function

Function PlaceValeur(ligne As Integer, col As Integer, Optional valeur As Variant) As Integer
Dim i As Integer
If IsMissing(valeur) Then
    PlaceValeur = 1 + Int(Rnd * 9)
Else
    PlaceValeur = valeur
    If PlaceValeur = 10 Then PlaceValeur = 1
End If
For i = 1 To 9
    If TabPossible(ligne, col) And BitMask(PlaceValeur) Then
        TabInit(ligne, col) = BitMask(PlaceValeur)
        TabVal(ligne, col) = BitMask(PlaceValeur)
        NbCasesPlacees = NbCasesPlacees + 1
        NbCasesInitiales = NbCasesInitiales + 1
        Exit For
    Else
        PlaceValeur = PlaceValeur + 1
        If PlaceValeur = 10 Then PlaceValeur = 1
    End If
Next i
End Function

Function ReConstruireGrille(niveau As Integer) As Boolean
Dim i As Integer, j As Integer, lr As Integer, cr As Integer
Dim ligne As Integer, col As Integer
Dim CaseOK As Boolean
Dim valeur As Integer
Dim ValeurOK As Boolean
Dim Tps As Integer
Dim ErrReg As Integer
Dim ErrGrille As Integer
Dim GrilleOK As Boolean
GrilleOK = False
ErrGrille = 0
ReConstruireGrille = False
While (Not GrilleOK)
    ligne = 1 + Int(Rnd * 9)
    col = 1 + Int(Rnd * 9)
    Select Case niveau
        Case Level_Easy
            If ReConstruireGrille = True And NbCasesInitiales < 37 Then Exit Function
        Case Level_Medium
            If ReConstruireGrille = True And NbCasesInitiales < 31 Then Exit Function
        Case Level_Hard
            If ReConstruireGrille = True And NbCasesInitiales < 25 Then Exit Function
    End Select
    If TabInit(ligne, col) <> 0 Then
        svgVal = TabInit(ligne, col)
        TabInit(ligne, col) = 0
        TabVal(ligne, col) = 0
        NbCasesInitiales = NbCasesInitiales - 1
        CaseOK = True
        RAZGrille
        NbCasesPlacees = 0
        For l = 1 To 9
            For c = 1 To 9
                TabVal(l, c) = TabInit(l, c)
                If TabInit(l, c) <> 0 Then NbCasesPlacees = NbCasesPlacees + 1
            Next c
        Next l
        MiseAJourGrille
        If NbCasesPlacees <> 81 Then
            ErrGrille = ErrGrille + 1
            TabInit(ligne, col) = svgVal
            TabVal(ligne, col) = svgVal
            NbCasesInitiales = NbCasesInitiales + 1
            If ErrGrille > 20 Then
                Exit Function
            End If
        Else
            ReConstruireGrille = True
            ErrGrille = 0
        End If
    End If
Wend
End Function

Sub PositionneNiveau(niveau As Integer)
AucuneAide = False
AideValPossible = False
AidePlacementAuto = True
AidePropagation = True
If niveau = Level_Easy Then
    AideValPossible = True
    AideValIsole = False
    AideRechJumeauxTriplets = False
    AideRecherchePaires = False
    AideRechercheTrios = False
    AideRechercheQuartets = False
ElseIf niveau = Level_Medium Then
    AideValIsole = True
    AideRechJumeauxTriplets = True
    AideRecherchePaires = True
    AideRechercheTrios = False
    AideRechercheQuartets = False
ElseIf niveau = Level_Hard Then
    AideValIsole = True
    AideRechJumeauxTriplets = True
    AideRecherchePaires = True
    AideRechercheTrios = True
    AideRechercheQuartets = True
End If
End Sub
