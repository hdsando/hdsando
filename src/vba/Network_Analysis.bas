Attribute VB_Name = "Network_Analysis"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: NETWORK_ANALYSIS v10.0
' ==============================================================================
' Description: Analyse de réseau pour détection de fraude avancée
' Fonctionnalités:
'   - Construction de graphe de transactions
'   - Détection de circuits fermés (round-tripping)
'   - Identification des comptes pivots (hubs)
'   - Analyse de communautés
'   - Détection de patterns de layering
' ==============================================================================

' --- CONSTANTES ---
Private Const MODULE_NAME As String = "Network_Analysis"
Private Const MAX_DEPTH As Integer = 5          ' Profondeur max recherche circuits
Private Const MIN_CIRCUIT_AMOUNT As Double = 100000  ' Montant min pour circuit suspect
Private Const HUB_THRESHOLD As Integer = 10     ' Nb connexions pour être hub

' --- TYPES ---
Private Type NetworkNode
    AccountID As String
    AccountName As String
    InDegree As Long        ' Nb connexions entrantes
    OutDegree As Long       ' Nb connexions sortantes
    TotalDegree As Long     ' Total connexions
    InVolume As Double      ' Volume entrant
    OutVolume As Double     ' Volume sortant
    NetFlow As Double       ' Flux net
    IsHub As Boolean
    ClusterID As Long
    RiskScore As Double
End Type

Private Type NetworkEdge
    SourceAccount As String
    TargetAccount As String
    TransactionCount As Long
    TotalAmount As Double
    AvgAmount As Double
    FirstDate As Date
    LastDate As Date
    IsSuspect As Boolean
End Type

Private Type CircuitPath
    PathLength As Integer
    Accounts(1 To 10) As String
    TotalAmount As Double
    AverageAmount As Double
    Frequency As Integer
    RiskLevel As String
End Type

' --- VARIABLES MODULE ---
Private mNodes As Collection
Private mEdges As Collection
Private mAdjacencyDict As Object

' ==============================================================================
' 1. POINT D'ENTRÉE PRINCIPAL
' ==============================================================================

Public Sub Lancer_Analyse_Reseau()
    Dim wsOut As Worksheet
    Dim startTime As Double

    On Error GoTo NetworkError

    startTime = Timer

    ' Initialiser collections
    Set mNodes = New Collection
    Set mEdges = New Collection
    Set mAdjacencyDict = CreateObject("Scripting.Dictionary")

    ' Créer feuille de résultats
    Set wsOut = SAFA_Common.GetOrCreateSheet("NETWORK_ANALYSIS", True)

    ' Titre
    wsOut.Range("A1").Value = "ANALYSE DE RÉSEAU TRANSACTIONNEL"
    wsOut.Range("A1").Font.Size = 16
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1:H1").Merge
    wsOut.Range("A1:H1").Interior.Color = RGB(0, 100, 100)
    wsOut.Range("A1:H1").Font.Color = vbWhite

    wsOut.Range("A2").Value = "Date génération: " & Format(Now, "dd/mm/yyyy hh:nn")

    ' Étape 1: Construire le graphe
    Call ConstruireGraphe

    ' Étape 2: Analyser les hubs
    Call AnalyserHubs(wsOut, 4)

    ' Étape 3: Détecter les circuits
    Call DetecterCircuits(wsOut, 30)

    ' Étape 4: Analyser le layering
    Call AnalyserLayering(wsOut, 55)

    ' Étape 5: Identifier communautés suspectes
    Call IdentifierCommunautes(wsOut, 80)

    ' Étape 6: Générer rapport de risque réseau
    Call GenererRapportRisqueReseau(wsOut, 105)

    wsOut.Columns("A:J").AutoFit

    SAFA_Common.WriteAuditLog "PROCESS", "Analyse réseau terminée", "Durée: " & Format(Timer - startTime, "0.00") & "s"

    MsgBox "Analyse de réseau terminée." & vbCrLf & _
           "Noeuds: " & mNodes.Count & vbCrLf & _
           "Arêtes: " & mEdges.Count & vbCrLf & _
           "Durée: " & Format(Timer - startTime, "0.0") & " secondes", _
           vbInformation, "Network Analysis"
    Exit Sub

NetworkError:
    SAFA_Common.LogError MODULE_NAME, "Lancer_Analyse_Reseau", Err.Number, Err.Description
    MsgBox "Erreur lors de l'analyse réseau: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 2. CONSTRUCTION DU GRAPHE
' ==============================================================================

Private Sub ConstruireGraphe()
    Dim wsTrans As Worksheet
    Dim dictNodes As Object
    Dim dictEdges As Object
    Dim i As Long, lr As Long

    On Error GoTo BuildError

    On Error Resume Next
    Set wsTrans = ThisWorkbook.Sheets("TRANSACTION_DATA")
    On Error GoTo BuildError

    If wsTrans Is Nothing Then Exit Sub

    Set dictNodes = CreateObject("Scripting.Dictionary")
    Set dictEdges = CreateObject("Scripting.Dictionary")

    lr = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row

    ' Parcourir les transactions
    For i = 2 To lr
        Dim sourceAcct As String
        Dim targetAcct As String
        Dim amount As Double
        Dim transDate As Date
        Dim edgeKey As String

        ' Supposons colonnes: 1=Compte, 2=Date, 3=Contre-partie, 4=Montant, 5=Type (D/C)
        sourceAcct = SAFA_Common.SafeText(wsTrans.Cells(i, 1).Value)
        targetAcct = SAFA_Common.SafeText(wsTrans.Cells(i, 3).Value)  ' Contre-partie
        amount = Abs(SAFA_Common.SafeVal(wsTrans.Cells(i, 4).Value))
        transDate = SAFA_Common.SafeDate(wsTrans.Cells(i, 2).Value)

        If sourceAcct = "" Or targetAcct = "" Or sourceAcct = targetAcct Then GoTo NextTrans

        ' Créer/mettre à jour noeud source
        If Not dictNodes.Exists(sourceAcct) Then
            Dim srcNode As NetworkNode
            srcNode.AccountID = sourceAcct
            srcNode.AccountName = SAFA_Common.SafeText(wsTrans.Cells(i, 1).Value)
            dictNodes.Add sourceAcct, srcNode
        End If

        ' Créer/mettre à jour noeud cible
        If Not dictNodes.Exists(targetAcct) Then
            Dim tgtNode As NetworkNode
            tgtNode.AccountID = targetAcct
            tgtNode.AccountName = targetAcct
            dictNodes.Add targetAcct, tgtNode
        End If

        ' Mettre à jour statistiques noeud
        Dim tempNode As NetworkNode
        tempNode = dictNodes(sourceAcct)
        tempNode.OutDegree = tempNode.OutDegree + 1
        tempNode.OutVolume = tempNode.OutVolume + amount
        tempNode.TotalDegree = tempNode.InDegree + tempNode.OutDegree
        dictNodes(sourceAcct) = tempNode

        tempNode = dictNodes(targetAcct)
        tempNode.InDegree = tempNode.InDegree + 1
        tempNode.InVolume = tempNode.InVolume + amount
        tempNode.TotalDegree = tempNode.InDegree + tempNode.OutDegree
        dictNodes(targetAcct) = tempNode

        ' Créer/mettre à jour arête
        edgeKey = sourceAcct & "|" & targetAcct

        If Not dictEdges.Exists(edgeKey) Then
            Dim newEdge As NetworkEdge
            newEdge.SourceAccount = sourceAcct
            newEdge.TargetAccount = targetAcct
            newEdge.TransactionCount = 1
            newEdge.TotalAmount = amount
            newEdge.FirstDate = transDate
            newEdge.LastDate = transDate
            dictEdges.Add edgeKey, newEdge
        Else
            Dim tempEdge As NetworkEdge
            tempEdge = dictEdges(edgeKey)
            tempEdge.TransactionCount = tempEdge.TransactionCount + 1
            tempEdge.TotalAmount = tempEdge.TotalAmount + amount
            If transDate < tempEdge.FirstDate Then tempEdge.FirstDate = transDate
            If transDate > tempEdge.LastDate Then tempEdge.LastDate = transDate
            dictEdges(edgeKey) = tempEdge
        End If

        ' Construire liste d'adjacence pour recherche de chemins
        If Not mAdjacencyDict.Exists(sourceAcct) Then
            mAdjacencyDict.Add sourceAcct, CreateObject("Scripting.Dictionary")
        End If
        mAdjacencyDict(sourceAcct)(targetAcct) = True

NextTrans:
    Next i

    ' Copier dans collections et calculer métriques finales
    Dim key As Variant

    For Each key In dictNodes.keys
        Dim finalNode As NetworkNode
        finalNode = dictNodes(key)
        finalNode.NetFlow = finalNode.InVolume - finalNode.OutVolume
        finalNode.IsHub = (finalNode.TotalDegree >= HUB_THRESHOLD)
        mNodes.Add finalNode, CStr(key)
    Next key

    For Each key In dictEdges.keys
        Dim finalEdge As NetworkEdge
        finalEdge = dictEdges(key)
        finalEdge.AvgAmount = IIf(finalEdge.TransactionCount > 0, _
                                  finalEdge.TotalAmount / finalEdge.TransactionCount, 0)
        mEdges.Add finalEdge, CStr(key)
    Next key

    Exit Sub

BuildError:
    SAFA_Common.LogError MODULE_NAME, "ConstruireGraphe", Err.Number, Err.Description
End Sub

' ==============================================================================
' 3. ANALYSE DES HUBS (COMPTES PIVOTS)
' ==============================================================================

Private Sub AnalyserHubs(ws As Worksheet, startRow As Long)
    Dim rowNum As Long
    Dim hubCount As Long
    Dim node As Variant

    On Error GoTo HubError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "1. COMPTES PIVOTS (HUBS)"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ws.Cells(rowNum, 1).Value = "Note: Un hub est un compte avec " & HUB_THRESHOLD & "+ connexions. " & _
                                "Les hubs peuvent être des comptes légitimes (trésorerie) ou suspects (intermédiaires frauduleux)."
    ws.Cells(rowNum, 1).Font.Italic = True
    rowNum = rowNum + 2

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Compte"
    ws.Cells(rowNum, 2).Value = "Connexions In"
    ws.Cells(rowNum, 3).Value = "Connexions Out"
    ws.Cells(rowNum, 4).Value = "Total"
    ws.Cells(rowNum, 5).Value = "Volume In"
    ws.Cells(rowNum, 6).Value = "Volume Out"
    ws.Cells(rowNum, 7).Value = "Flux Net"
    ws.Cells(rowNum, 8).Value = "Classification"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8))
    rowNum = rowNum + 1

    hubCount = 0

    ' Trier les noeuds par degré (simple tri par insertion)
    Dim sortedNodes() As NetworkNode
    Dim n As NetworkNode
    Dim i As Long, j As Long
    Dim nodeCount As Long

    nodeCount = mNodes.Count
    If nodeCount = 0 Then
        ws.Cells(rowNum, 1).Value = "Aucun noeud dans le graphe."
        Exit Sub
    End If

    ReDim sortedNodes(1 To nodeCount)
    i = 1
    For Each node In mNodes
        sortedNodes(i) = node
        i = i + 1
    Next node

    ' Tri décroissant par TotalDegree
    For i = 1 To nodeCount - 1
        For j = i + 1 To nodeCount
            If sortedNodes(j).TotalDegree > sortedNodes(i).TotalDegree Then
                Dim temp As NetworkNode
                temp = sortedNodes(i)
                sortedNodes(i) = sortedNodes(j)
                sortedNodes(j) = temp
            End If
        Next j
    Next i

    ' Afficher top 20 hubs
    For i = 1 To WorksheetFunction.Min(20, nodeCount)
        n = sortedNodes(i)

        If n.TotalDegree < HUB_THRESHOLD Then Exit For

        hubCount = hubCount + 1

        ws.Cells(rowNum, 1).Value = n.AccountID
        ws.Cells(rowNum, 2).Value = n.InDegree
        ws.Cells(rowNum, 3).Value = n.OutDegree
        ws.Cells(rowNum, 4).Value = n.TotalDegree
        ws.Cells(rowNum, 5).Value = n.InVolume
        ws.Cells(rowNum, 5).NumberFormat = "#,##0"
        ws.Cells(rowNum, 6).Value = n.OutVolume
        ws.Cells(rowNum, 6).NumberFormat = "#,##0"
        ws.Cells(rowNum, 7).Value = n.NetFlow
        ws.Cells(rowNum, 7).NumberFormat = "#,##0"

        ' Classification
        Dim inOutRatio As Double
        If n.OutDegree > 0 Then
            inOutRatio = n.InDegree / n.OutDegree
        Else
            inOutRatio = 999
        End If

        If inOutRatio > 3 Then
            ws.Cells(rowNum, 8).Value = "COLLECTEUR"
            ws.Cells(rowNum, 8).Interior.Color = RGB(255, 200, 200)
        ElseIf inOutRatio < 0.33 Then
            ws.Cells(rowNum, 8).Value = "DISTRIBUTEUR"
            ws.Cells(rowNum, 8).Interior.Color = RGB(200, 200, 255)
        ElseIf Abs(inOutRatio - 1) < 0.3 And n.TotalDegree > 20 Then
            ws.Cells(rowNum, 8).Value = "PASSEUR"
            ws.Cells(rowNum, 8).Interior.Color = RGB(255, 255, 0)
        Else
            ws.Cells(rowNum, 8).Value = "Normal"
        End If

        rowNum = rowNum + 1
    Next i

    ws.Cells(rowNum + 1, 1).Value = "Total hubs détectés: " & hubCount
    ws.Cells(rowNum + 1, 1).Font.Bold = True

    Exit Sub

HubError:
    SAFA_Common.LogError MODULE_NAME, "AnalyserHubs", Err.Number, Err.Description
End Sub

' ==============================================================================
' 4. DÉTECTION DE CIRCUITS (ROUND-TRIPPING)
' ==============================================================================

Private Sub DetecterCircuits(ws As Worksheet, startRow As Long)
    Dim rowNum As Long
    Dim circuitCount As Long
    Dim node As Variant

    On Error GoTo CircuitError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "2. CIRCUITS DÉTECTÉS (ROUND-TRIPPING)"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ws.Cells(rowNum, 1).Value = "Note: Un circuit est un chemin A→B→...→A. Indicateur potentiel de fraude par circulation."
    ws.Cells(rowNum, 1).Font.Italic = True
    rowNum = rowNum + 2

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Longueur"
    ws.Cells(rowNum, 2).Value = "Circuit (Comptes)"
    ws.Cells(rowNum, 3).Value = "Nb Occurrences"
    ws.Cells(rowNum, 4).Value = "Volume Total"
    ws.Cells(rowNum, 5).Value = "Risque"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 5))
    rowNum = rowNum + 1

    circuitCount = 0

    ' Rechercher circuits depuis chaque noeud hub
    For Each node In mNodes
        Dim n As NetworkNode
        n = node

        If n.IsHub Then
            Dim path(1 To 10) As String
            Dim visited As Object
            Set visited = CreateObject("Scripting.Dictionary")

            path(1) = n.AccountID
            visited.Add n.AccountID, True

            Call FindCircuitsDFS(n.AccountID, n.AccountID, 1, path, visited, ws, rowNum, circuitCount)

            ' Limiter le nombre de circuits trouvés
            If circuitCount >= 15 Then Exit For
        End If
    Next node

    If circuitCount = 0 Then
        ws.Cells(rowNum, 1).Value = "Aucun circuit suspect détecté."
        ws.Cells(rowNum, 1).Font.Color = RGB(0, 150, 0)
    Else
        ws.Cells(rowNum + 1, 1).Value = "Total circuits: " & circuitCount
        ws.Cells(rowNum + 1, 1).Font.Bold = True
        ws.Cells(rowNum + 1, 1).Font.Color = RGB(200, 0, 0)
    End If

    Exit Sub

CircuitError:
    SAFA_Common.LogError MODULE_NAME, "DetecterCircuits", Err.Number, Err.Description
End Sub

Private Sub FindCircuitsDFS(startNode As String, currentNode As String, depth As Integer, _
                            path() As String, visited As Object, ws As Worksheet, _
                            ByRef rowNum As Long, ByRef circuitCount As Long)
    ' Recherche en profondeur de circuits
    Dim neighbor As Variant

    If depth > MAX_DEPTH Then Exit Sub
    If circuitCount >= 15 Then Exit Sub

    If Not mAdjacencyDict.Exists(currentNode) Then Exit Sub

    For Each neighbor In mAdjacencyDict(currentNode).keys
        If CStr(neighbor) = startNode And depth >= 2 Then
            ' Circuit trouvé!
            circuitCount = circuitCount + 1

            ' Écrire le circuit
            ws.Cells(rowNum, 1).Value = depth + 1
            ws.Cells(rowNum, 2).Value = BuildPathString(path, depth) & " → " & startNode

            ' Estimer volume (simplifié)
            ws.Cells(rowNum, 3).Value = "1+"
            ws.Cells(rowNum, 4).Value = "À analyser"

            If depth <= 3 Then
                ws.Cells(rowNum, 5).Value = "CRITICAL"
                ws.Cells(rowNum, 5).Interior.Color = RGB(255, 0, 0)
                ws.Cells(rowNum, 5).Font.Color = vbWhite
            Else
                ws.Cells(rowNum, 5).Value = "HIGH"
                ws.Cells(rowNum, 5).Interior.Color = RGB(255, 165, 0)
            End If

            rowNum = rowNum + 1

        ElseIf Not visited.Exists(CStr(neighbor)) And depth < MAX_DEPTH Then
            ' Continuer exploration
            visited.Add CStr(neighbor), True
            path(depth + 1) = CStr(neighbor)

            Call FindCircuitsDFS(startNode, CStr(neighbor), depth + 1, path, visited, ws, rowNum, circuitCount)

            visited.Remove CStr(neighbor)
        End If
    Next neighbor
End Sub

Private Function BuildPathString(path() As String, length As Integer) As String
    Dim result As String
    Dim i As Integer

    result = path(1)
    For i = 2 To length
        result = result & " → " & path(i)
    Next i

    BuildPathString = result
End Function

' ==============================================================================
' 5. ANALYSE DU LAYERING
' ==============================================================================

Private Sub AnalyserLayering(ws As Worksheet, startRow As Long)
    Dim rowNum As Long
    Dim edge As Variant
    Dim layeringCount As Long

    On Error GoTo LayeringError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "3. PATTERNS DE LAYERING"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ws.Cells(rowNum, 1).Value = "Note: Le layering consiste à faire passer des fonds par plusieurs comptes rapidement pour masquer l'origine."
    ws.Cells(rowNum, 1).Font.Italic = True
    rowNum = rowNum + 2

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Source"
    ws.Cells(rowNum, 2).Value = "Cible"
    ws.Cells(rowNum, 3).Value = "Nb Trans"
    ws.Cells(rowNum, 4).Value = "Volume"
    ws.Cells(rowNum, 5).Value = "Moy/Trans"
    ws.Cells(rowNum, 6).Value = "Durée (jours)"
    ws.Cells(rowNum, 7).Value = "Fréquence/jour"
    ws.Cells(rowNum, 8).Value = "Risque"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8))
    rowNum = rowNum + 1

    layeringCount = 0

    For Each edge In mEdges
        Dim e As NetworkEdge
        e = edge

        ' Critères de layering suspect:
        ' - Nombreuses transactions (>4)
        ' - Sur courte période
        ' - Montants relativement constants
        Dim durationDays As Long
        Dim freqPerDay As Double

        durationDays = DateDiff("d", e.FirstDate, e.LastDate)
        If durationDays < 1 Then durationDays = 1

        freqPerDay = e.TransactionCount / durationDays

        If e.TransactionCount >= 4 And freqPerDay >= 0.5 Then
            layeringCount = layeringCount + 1

            ws.Cells(rowNum, 1).Value = e.SourceAccount
            ws.Cells(rowNum, 2).Value = e.TargetAccount
            ws.Cells(rowNum, 3).Value = e.TransactionCount
            ws.Cells(rowNum, 4).Value = e.TotalAmount
            ws.Cells(rowNum, 4).NumberFormat = "#,##0"
            ws.Cells(rowNum, 5).Value = e.AvgAmount
            ws.Cells(rowNum, 5).NumberFormat = "#,##0"
            ws.Cells(rowNum, 6).Value = durationDays
            ws.Cells(rowNum, 7).Value = freqPerDay
            ws.Cells(rowNum, 7).NumberFormat = "0.00"

            ' Évaluer risque
            If e.TransactionCount >= 10 And freqPerDay >= 2 Then
                ws.Cells(rowNum, 8).Value = "CRITICAL"
                ws.Cells(rowNum, 8).Interior.Color = RGB(255, 0, 0)
                ws.Cells(rowNum, 8).Font.Color = vbWhite
            ElseIf e.TransactionCount >= 6 Or freqPerDay >= 1 Then
                ws.Cells(rowNum, 8).Value = "HIGH"
                ws.Cells(rowNum, 8).Interior.Color = RGB(255, 165, 0)
            Else
                ws.Cells(rowNum, 8).Value = "MEDIUM"
                ws.Cells(rowNum, 8).Interior.Color = RGB(255, 255, 0)
            End If

            rowNum = rowNum + 1

            ' Limiter à 20 résultats
            If layeringCount >= 20 Then Exit For
        End If
    Next edge

    If layeringCount = 0 Then
        ws.Cells(rowNum, 1).Value = "Aucun pattern de layering suspect détecté."
        ws.Cells(rowNum, 1).Font.Color = RGB(0, 150, 0)
    End If

    Exit Sub

LayeringError:
    SAFA_Common.LogError MODULE_NAME, "AnalyserLayering", Err.Number, Err.Description
End Sub

' ==============================================================================
' 6. IDENTIFICATION DE COMMUNAUTÉS
' ==============================================================================

Private Sub IdentifierCommunautes(ws As Worksheet, startRow As Long)
    Dim rowNum As Long
    Dim dictCommunities As Object
    Dim node As Variant

    On Error GoTo CommunityError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "4. CLUSTERS DE COMPTES"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ws.Cells(rowNum, 1).Value = "Note: Groupes de comptes fortement interconnectés peuvent indiquer activité coordonnée."
    ws.Cells(rowNum, 1).Font.Italic = True
    rowNum = rowNum + 2

    ' Algorithme simplifié de clustering basé sur les hubs
    ' Chaque hub définit une communauté avec ses voisins directs

    Set dictCommunities = CreateObject("Scripting.Dictionary")

    Dim communityID As Long
    communityID = 0

    For Each node In mNodes
        Dim n As NetworkNode
        n = node

        If n.IsHub And Not IsNodeInCommunity(dictCommunities, n.AccountID) Then
            communityID = communityID + 1

            Dim members As New Collection
            members.Add n.AccountID

            ' Ajouter voisins
            If mAdjacencyDict.Exists(n.AccountID) Then
                Dim neighbor As Variant
                For Each neighbor In mAdjacencyDict(n.AccountID).keys
                    If Not IsNodeInCommunity(dictCommunities, CStr(neighbor)) Then
                        members.Add CStr(neighbor)
                    End If
                Next neighbor
            End If

            dictCommunities.Add communityID, members
        End If
    Next node

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Cluster ID"
    ws.Cells(rowNum, 2).Value = "Nb Membres"
    ws.Cells(rowNum, 3).Value = "Compte Central"
    ws.Cells(rowNum, 4).Value = "Membres"
    ws.Cells(rowNum, 5).Value = "Risque"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 5))
    rowNum = rowNum + 1

    ' Afficher communautés
    Dim key As Variant
    For Each key In dictCommunities.keys
        Dim members As Collection
        Set members = dictCommunities(key)

        ws.Cells(rowNum, 1).Value = key
        ws.Cells(rowNum, 2).Value = members.Count

        If members.Count > 0 Then
            ws.Cells(rowNum, 3).Value = members(1)

            ' Liste des membres (limité à 50 chars)
            Dim memberList As String
            Dim m As Variant
            memberList = ""
            For Each m In members
                memberList = memberList & CStr(m) & ", "
            Next m
            ws.Cells(rowNum, 4).Value = Left(memberList, 50) & IIf(Len(memberList) > 50, "...", "")

            ' Risque basé sur taille
            If members.Count >= 10 Then
                ws.Cells(rowNum, 5).Value = "HIGH"
                ws.Cells(rowNum, 5).Interior.Color = RGB(255, 165, 0)
            ElseIf members.Count >= 5 Then
                ws.Cells(rowNum, 5).Value = "MEDIUM"
                ws.Cells(rowNum, 5).Interior.Color = RGB(255, 255, 0)
            Else
                ws.Cells(rowNum, 5).Value = "LOW"
            End If
        End If

        rowNum = rowNum + 1
    Next key

    ws.Cells(rowNum + 1, 1).Value = "Total clusters: " & dictCommunities.Count
    ws.Cells(rowNum + 1, 1).Font.Bold = True

    Exit Sub

CommunityError:
    SAFA_Common.LogError MODULE_NAME, "IdentifierCommunautes", Err.Number, Err.Description
End Sub

Private Function IsNodeInCommunity(dictComm As Object, nodeID As String) As Boolean
    Dim key As Variant
    Dim members As Collection

    For Each key In dictComm.keys
        Set members = dictComm(key)
        Dim m As Variant
        For Each m In members
            If CStr(m) = nodeID Then
                IsNodeInCommunity = True
                Exit Function
            End If
        Next m
    Next key

    IsNodeInCommunity = False
End Function

' ==============================================================================
' 7. RAPPORT DE RISQUE RÉSEAU
' ==============================================================================

Private Sub GenererRapportRisqueReseau(ws As Worksheet, startRow As Long)
    Dim rowNum As Long

    On Error GoTo ReportError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "5. SYNTHÈSE RISQUE RÉSEAU"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ' Statistiques globales
    ws.Cells(rowNum, 1).Value = "Métrique"
    ws.Cells(rowNum, 2).Value = "Valeur"
    ws.Cells(rowNum, 3).Value = "Seuil"
    ws.Cells(rowNum, 4).Value = "Statut"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 4))
    rowNum = rowNum + 1

    ' Nombre de noeuds
    ws.Cells(rowNum, 1).Value = "Nombre de comptes dans le réseau"
    ws.Cells(rowNum, 2).Value = mNodes.Count
    rowNum = rowNum + 1

    ' Nombre d'arêtes
    ws.Cells(rowNum, 1).Value = "Nombre de connexions"
    ws.Cells(rowNum, 2).Value = mEdges.Count
    rowNum = rowNum + 1

    ' Densité du graphe
    Dim density As Double
    If mNodes.Count > 1 Then
        density = mEdges.Count / (mNodes.Count * (mNodes.Count - 1))
    Else
        density = 0
    End If
    ws.Cells(rowNum, 1).Value = "Densité du réseau"
    ws.Cells(rowNum, 2).Value = density
    ws.Cells(rowNum, 2).NumberFormat = "0.00%"
    ws.Cells(rowNum, 3).Value = "> 10%"
    If density > 0.1 Then
        ws.Cells(rowNum, 4).Value = "Dense"
        ws.Cells(rowNum, 4).Interior.Color = RGB(255, 255, 0)
    Else
        ws.Cells(rowNum, 4).Value = "Normal"
    End If
    rowNum = rowNum + 1

    ' Nombre de hubs
    Dim hubCount As Long
    Dim node As Variant
    hubCount = 0
    For Each node In mNodes
        Dim n As NetworkNode
        n = node
        If n.IsHub Then hubCount = hubCount + 1
    Next node

    ws.Cells(rowNum, 1).Value = "Nombre de hubs"
    ws.Cells(rowNum, 2).Value = hubCount
    ws.Cells(rowNum, 3).Value = "< 5"
    If hubCount > 5 Then
        ws.Cells(rowNum, 4).Value = "Élevé"
        ws.Cells(rowNum, 4).Interior.Color = RGB(255, 165, 0)
    Else
        ws.Cells(rowNum, 4).Value = "Normal"
    End If
    rowNum = rowNum + 2

    ' Score de risque global
    Dim networkRiskScore As Integer
    networkRiskScore = 0

    If density > 0.2 Then networkRiskScore = networkRiskScore + 30
    If hubCount > 10 Then networkRiskScore = networkRiskScore + 30
    If mEdges.Count > 1000 Then networkRiskScore = networkRiskScore + 20

    ws.Cells(rowNum, 1).Value = "SCORE DE RISQUE RÉSEAU GLOBAL"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 2).Value = networkRiskScore & "/100"
    ws.Cells(rowNum, 2).Font.Bold = True
    ws.Cells(rowNum, 2).Font.Size = 14

    If networkRiskScore >= 70 Then
        ws.Cells(rowNum, 2).Interior.Color = RGB(255, 0, 0)
        ws.Cells(rowNum, 2).Font.Color = vbWhite
        ws.Cells(rowNum, 3).Value = "CRITIQUE - Investigation urgente recommandée"
    ElseIf networkRiskScore >= 50 Then
        ws.Cells(rowNum, 2).Interior.Color = RGB(255, 165, 0)
        ws.Cells(rowNum, 3).Value = "ÉLEVÉ - Surveillance renforcée"
    ElseIf networkRiskScore >= 30 Then
        ws.Cells(rowNum, 2).Interior.Color = RGB(255, 255, 0)
        ws.Cells(rowNum, 3).Value = "MOYEN - Suivi normal"
    Else
        ws.Cells(rowNum, 2).Interior.Color = RGB(0, 200, 0)
        ws.Cells(rowNum, 2).Font.Color = vbWhite
        ws.Cells(rowNum, 3).Value = "FAIBLE - Pas d'anomalie réseau majeure"
    End If

    Exit Sub

ReportError:
    SAFA_Common.LogError MODULE_NAME, "GenererRapportRisqueReseau", Err.Number, Err.Description
End Sub
