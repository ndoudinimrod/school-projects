Attribute VB_Name = "AuditIT_Macro"
' ================================================================
' PROGRAMME D'AUDIT IT — NIS2 / DORA
' Macro de génération de rapport Word
' Nimrod Ndoudi | Portfolio IT Audit
' ================================================================
' INSTRUCTIONS D'IMPORT :
'   1. Ouvrir ce fichier .xlsm dans Excel
'   2. Alt+F11 → Outils → Importer un fichier → sélectionner ce .bas
'   3. Fermer l'éditeur VBA et sauvegarder en .xlsm
'   4. Sur la feuille Accueil : cliquer sur le bouton "Générer rapport"
'      OU Alt+F8 → GenererRapportWord → Exécuter
' ================================================================

Sub GenererRapportWord()
    Dim wb As Workbook
    Dim wsControles As Worksheet
    Dim wsSynthese As Worksheet
    Dim wsAccueil As Worksheet
    
    Set wb = ThisWorkbook
    Set wsControles = wb.Sheets("Contrôles")
    Set wsSynthese = wb.Sheets("Synthèse")
    Set wsAccueil = wb.Sheets("Accueil")
    
    Dim choix As Integer
    choix = MsgBox("Choisir le type de rapport :" & Chr(13) & Chr(13) & _
                   "OUI  = Rapport Simple (Synthèse + Tableau)" & Chr(13) & _
                   "NON = Rapport Big 4 (Couverture, Executive Summary, Fiches)", _
                   vbYesNoCancel + vbQuestion, "Type de rapport")
    If choix = vbCancel Then Exit Sub
    
    Dim wdApp As Object
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    If wdApp Is Nothing Then
        MsgBox "Microsoft Word est requis pour générer le rapport.", vbCritical
        Exit Sub
    End If
    
    wdApp.Visible = False
    
    Dim entite As String:     entite = Nz(wsAccueil.Range("C8").Value, "[Entité]")
    Dim responsable As String: responsable = Nz(wsAccueil.Range("C9").Value, "[Responsable]")
    Dim dateDebut As String:  dateDebut = Nz(wsAccueil.Range("C10").Value, Format(Now(), "dd/mm/yyyy"))
    Dim dateFin As String:    dateFin = Nz(wsAccueil.Range("C11").Value, "")
    Dim perimetre As String:  perimetre = Nz(wsAccueil.Range("C12").Value, "")
    
    If entite = "" Then entite = "[Entité non renseignée]"
    If responsable = "" Then responsable = "[Responsable non renseigné]"
    If dateDebut = "" Then dateDebut = Format(Now(), "dd/mm/yyyy")
    
    Dim wdDoc As Object
    Set wdDoc = wdApp.Documents.Add
    
    Dim savePath As String
    If choix = vbYes Then
        Call GenererRapportSimple(wdDoc, wsControles, wsSynthese, entite, responsable, dateDebut, dateFin, perimetre)
        savePath = wb.Path & "\Rapport_Audit_IT_Simple_" & Format(Now(), "yyyymmdd") & ".docx"
    Else
        Call GenererRapportBig4(wdDoc, wsControles, wsSynthese, entite, responsable, dateDebut, dateFin, perimetre)
        savePath = wb.Path & "\Rapport_Audit_IT_Big4_" & Format(Now(), "yyyymmdd") & ".docx"
    End If
    
    wdDoc.SaveAs2 savePath, 16
    wdDoc.Close
    wdApp.Quit
    
    MsgBox "Rapport généré avec succès !" & Chr(13) & Chr(13) & savePath, vbInformation, "Rapport Word"
End Sub

' ============================================================
' HELPERS
' ============================================================
Function Nz(val As Variant, default As String) As String
    If IsNull(val) Or val = "" Then
        Nz = default
    Else
        Nz = CStr(val)
    End If
End Function

Sub AddPara(wdDoc As Object, txt As String, sz As Integer, bold As Boolean, col As Long)
    With wdDoc.Content
        .Collapse 0
        .InsertParagraphAfter
        .Collapse 0
        .InsertAfter txt
        .Font.Name = "Arial"
        .Font.Size = sz
        .Font.Bold = bold
        If col <> 0 Then .Font.Color = col
    End With
End Sub

Sub AddHeading(wdDoc As Object, txt As String, lvl As Integer)
    With wdDoc.Content
        .Collapse 0
        .InsertParagraphAfter
        .Collapse 0
        .InsertAfter txt
        .Style = wdDoc.Styles("Titre " & lvl)
    End With
End Sub

Sub AddPageBreak(wdDoc As Object)
    With wdDoc.Content
        .Collapse 0
        .InsertBreak 7
    End With
End Sub

' ============================================================
' RAPPORT SIMPLE
' ============================================================
Sub GenererRapportSimple(wdDoc As Object, wsCtrl As Worksheet, wsSyn As Worksheet, _
    entite As String, resp As String, dDeb As String, dFin As String, perim As String)
    
    ' En-tête
    With wdDoc.Sections(1).Headers(1).Range
        .Text = "Rapport d'Audit IT — " & entite & "   |   " & dDeb
        .Font.Name = "Arial": .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)
    End With
    
    ' Titre
    With wdDoc.Content
        .InsertAfter "RAPPORT D'AUDIT IT"
        .Font.Name = "Arial": .Font.Size = 22: .Font.Bold = True
        .Font.Color = RGB(31, 56, 100)
        .ParagraphFormat.Alignment = 1
    End With
    
    Call AddPara(wdDoc, "Conformité NIS2 / DORA", 13, False, RGB(46, 117, 182))
    Call AddPara(wdDoc, "Entité : " & entite & "  |  Responsable : " & resp, 10, False, 0)
    Call AddPara(wdDoc, "Période : " & dDeb & " — " & dFin, 10, False, 0)
    Call AddPara(wdDoc, "Périmètre : " & perim, 10, False, 0)
    
    Call AddPageBreak(wdDoc)
    Call AddHeading(wdDoc, "1. Synthèse par domaine", 1)
    Call AddTableSynthese(wdDoc, wsSyn)
    Call AddPara(wdDoc, "", 8, False, 0)
    Call AddPageBreak(wdDoc)
    Call AddHeading(wdDoc, "2. Résultats détaillés des contrôles", 1)
    Call AddTableControles(wdDoc, wsCtrl)
End Sub

' ============================================================
' RAPPORT BIG 4
' ============================================================
Sub GenererRapportBig4(wdDoc As Object, wsCtrl As Worksheet, wsSyn As Worksheet, _
    entite As String, resp As String, dDeb As String, dFin As String, perim As String)
    
    ' ── Page de couverture ──────────────────────────────────
    With wdDoc.Content
        .InsertAfter Chr(13) & Chr(13)
        .InsertAfter "AUDIT IT" & Chr(13)
        .Font.Name = "Arial": .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(31, 56, 100): .ParagraphFormat.Alignment = 1
    End With
    Call AddPara(wdDoc, "Rapport de mission — Conformité NIS2 / DORA", 16, False, RGB(46, 117, 182))
    wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Alignment = 1
    Call AddPara(wdDoc, Chr(13) & "CONFIDENTIEL — USAGE RESTREINT", 9, True, RGB(192, 0, 0))
    wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Alignment = 1
    Call AddPara(wdDoc, Chr(13), 10, False, 0)
    
    ' Table de couverture
    Dim tCov As Object
    Set tCov = wdDoc.Tables.Add( _
        wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1), 5, 2)
    tCov.Style = "Tableau de grille"
    On Error Resume Next
    tCov.Style = "Table Grid"
    On Error GoTo 0
    
    Dim covD(4, 1) As String
    covD(0, 0) = "Entité auditée":    covD(0, 1) = entite
    covD(1, 0) = "Responsable":       covD(1, 1) = resp
    covD(2, 0) = "Date de début":     covD(2, 1) = dDeb
    covD(3, 0) = "Date de fin":       covD(3, 1) = dFin
    covD(4, 0) = "Périmètre":         covD(4, 1) = perim
    Dim ir As Integer, ic As Integer
    For ir = 0 To 4
        For ic = 0 To 1
            With tCov.Cell(ir + 1, ic + 1)
                .Range.Text = covD(ir, ic)
                .Range.Font.Name = "Arial": .Range.Font.Size = 10
                If ic = 0 Then
                    .Range.Font.Bold = True
                    .Shading.BackgroundPatternColor = RGB(214, 228, 240)
                End If
            End With
        Next ic
    Next ir
    
    Call AddPageBreak(wdDoc)
    
    ' ── Sommaire ────────────────────────────────────────────
    Call AddHeading(wdDoc, "Sommaire", 1)
    Dim somLines(6) As String
    somLines(0) = "1.  Contexte et objectifs de la mission"
    somLines(1) = "2.  Executive Summary"
    somLines(2) = "3.  Synthèse des résultats par domaine"
    somLines(3) = "4.  Fiches de contrôles — Access Management"
    somLines(4) = "5.  Fiches de contrôles — Sécurité Opérationnelle"
    somLines(5) = "6.  Fiches de contrôles — Cybersécurité"
    somLines(6) = "7.  Plan d'actions recommandées"
    Dim sl As Integer
    For sl = 0 To 6
        Call AddPara(wdDoc, somLines(sl), 10, False, 0)
    Next sl
    Call AddPageBreak(wdDoc)
    
    ' ── Section 1 ───────────────────────────────────────────
    Call AddHeading(wdDoc, "1.  Contexte et objectifs de la mission", 1)
    Call AddPara(wdDoc, "La présente mission d'audit IT a été conduite conformément aux référentiels NIS2 " & _
        "(Directive (UE) 2022/2555) et DORA (Règlement (UE) 2022/2554). Elle vise à évaluer le niveau " & _
        "de maturité de l'entité " & entite & " sur trois domaines : gestion des accès, sécurité " & _
        "opérationnelle et cybersécurité.", 10, False, 0)
    Call AddPara(wdDoc, "Mission conduite par " & resp & ", du " & dDeb & " au " & dFin & _
        ". Périmètre : " & perim & ".", 10, False, 0)
    Call AddPageBreak(wdDoc)
    
    ' ── Section 2 : Executive Summary ───────────────────────
    Call AddHeading(wdDoc, "2.  Executive Summary", 1)
    
    Dim cConf As Long, cNC As Long, cObs As Long, cNA As Long, cTot As Long
    cConf = 0: cNC = 0: cObs = 0: cNA = 0: cTot = 0
    Dim rr As Long
    For rr = 5 To wsCtrl.UsedRange.Rows.Count + wsCtrl.UsedRange.Row - 1
        Dim sv As String: sv = wsCtrl.Cells(rr, 6).Value
        Select Case sv
            Case "Conforme":     cConf = cConf + 1: cTot = cTot + 1
            Case "Non-Conforme": cNC = cNC + 1:     cTot = cTot + 1
            Case "Observation":  cObs = cObs + 1:   cTot = cTot + 1
            Case "N/A":          cNA = cNA + 1:      cTot = cTot + 1
        End Select
    Next rr
    
    Dim txC As String
    If cTot > 0 Then txC = Format(cConf / cTot * 100, "0") & "%" Else txC = "N/D"
    
    Call AddPara(wdDoc, cTot & " contrôles évalués — Taux de conformité : " & txC, 12, True, RGB(31, 56, 100))
    Call AddPara(wdDoc, "", 8, False, 0)
    Call AddPara(wdDoc, "Conformes        : " & cConf, 10, False, RGB(0, 160, 0))
    Call AddPara(wdDoc, "Non-Conformes  : " & cNC,   10, False, RGB(200, 0, 0))
    Call AddPara(wdDoc, "Observations     : " & cObs, 10, False, RGB(200, 130, 0))
    Call AddPara(wdDoc, "Non Applicables : " & cNA,   10, False, RGB(100, 100, 100))
    If cNC > 0 Then
        Call AddPara(wdDoc, Chr(13) & _
            "Des non-conformités ont été identifiées. Des recommandations sont formulées en section 7.", _
            10, False, RGB(192, 0, 0))
    End If
    Call AddPageBreak(wdDoc)
    
    ' ── Section 3 : Synthèse ────────────────────────────────
    Call AddHeading(wdDoc, "3.  Synthèse des résultats par domaine", 1)
    Call AddTableSynthese(wdDoc, wsSyn)
    Call AddPageBreak(wdDoc)
    
    ' ── Sections 4-6 : Fiches ───────────────────────────────
    Dim doms(2) As String
    doms(0) = "Access Management"
    doms(1) = "Sécurité Opérationnelle"
    doms(2) = "Cybersécurité"
    Dim secNums(2) As String
    secNums(0) = "4": secNums(1) = "5": secNums(2) = "6"
    
    Dim dd As Integer
    For dd = 0 To 2
        Call AddHeading(wdDoc, secNums(dd) & ".  Fiches de contrôles — " & doms(dd), 1)
        Call AddFichesParDomaine(wdDoc, wsCtrl, doms(dd))
        If dd < 2 Then Call AddPageBreak(wdDoc)
    Next dd
    Call AddPageBreak(wdDoc)
    
    ' ── Section 7 : Plan d'actions ──────────────────────────
    Call AddHeading(wdDoc, "7.  Plan d'actions recommandées", 1)
    Call AddPara(wdDoc, "Recensement des non-conformités et observations avec recommandations associées.", 10, False, 0)
    Call AddPara(wdDoc, "", 8, False, 0)
    Call AddTablePlanActions(wdDoc, wsCtrl)
End Sub

' ============================================================
' TABLEAUX
' ============================================================
Sub AddTableSynthese(wdDoc As Object, wsSyn As Worksheet)
    Dim tbl As Object
    Set tbl = wdDoc.Tables.Add( _
        wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1), 5, 6)
    On Error Resume Next
    tbl.Style = "Table Grid"
    On Error GoTo 0
    
    Dim hdrs(5) As String
    hdrs(0) = "Domaine": hdrs(1) = "Conforme": hdrs(2) = "Non-Conforme"
    hdrs(3) = "Observation": hdrs(4) = "N/A": hdrs(5) = "Total"
    
    Dim ci As Integer
    For ci = 0 To 5
        With tbl.Cell(1, ci + 1)
            .Range.Text = hdrs(ci)
            .Range.Font.Bold = True: .Range.Font.Name = "Arial": .Range.Font.Size = 9
            .Range.Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(46, 117, 182)
        End With
    Next ci
    
    Dim ri As Integer
    For ri = 0 To 2
        For ci = 0 To 5
            With tbl.Cell(ri + 2, ci + 1)
                .Range.Text = CStr(wsSyn.Cells(4 + ri, 2 + ci).Value)
                .Range.Font.Name = "Arial": .Range.Font.Size = 9
                If ri Mod 2 = 0 Then
                    .Shading.BackgroundPatternColor = RGB(214, 228, 240)
                End If
            End With
        Next ci
    Next ri
    
    For ci = 0 To 5
        With tbl.Cell(5, ci + 1)
            .Range.Text = CStr(wsSyn.Cells(7, 2 + ci).Value)
            .Range.Font.Bold = True: .Range.Font.Name = "Arial": .Range.Font.Size = 9
            .Range.Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(31, 56, 100)
        End With
    Next ci
End Sub

Sub AddTableControles(wdDoc As Object, wsCtrl As Worksheet)
    Dim lastR As Long: lastR = wsCtrl.UsedRange.Rows.Count + wsCtrl.UsedRange.Row - 1
    
    Dim validR(200) As Long: Dim vCnt As Long: vCnt = 0
    Dim rr As Long
    For rr = 5 To lastR
        If wsCtrl.Cells(rr, 3).Value <> "" And IsNumeric(wsCtrl.Cells(rr, 2).Value) Then
            vCnt = vCnt + 1: validR(vCnt) = rr
        End If
    Next rr
    If vCnt = 0 Then Call AddPara(wdDoc, "Aucun contrôle.", 10, False, 0): Exit Sub
    
    Dim tbl As Object
    Set tbl = wdDoc.Tables.Add( _
        wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1), vCnt + 1, 5)
    On Error Resume Next: tbl.Style = "Table Grid": On Error GoTo 0
    
    Dim hdrs(4) As String
    hdrs(0) = "Réf.": hdrs(1) = "Intitulé": hdrs(2) = "Statut"
    hdrs(3) = "Constatations": hdrs(4) = "Référence réglementaire"
    
    Dim ci As Integer
    For ci = 0 To 4
        With tbl.Cell(1, ci + 1)
            .Range.Text = hdrs(ci)
            .Range.Font.Bold = True: .Range.Font.Name = "Arial": .Range.Font.Size = 8
            .Range.Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(46, 117, 182)
        End With
    Next ci
    
    Dim ii As Long
    For ii = 1 To vCnt
        rr = validR(ii)
        Dim sv As String: sv = wsCtrl.Cells(rr, 6).Value
        
        tbl.Cell(ii + 1, 1).Range.Text = wsCtrl.Cells(rr, 3).Value
        tbl.Cell(ii + 1, 2).Range.Text = wsCtrl.Cells(rr, 4).Value
        tbl.Cell(ii + 1, 3).Range.Text = sv
        tbl.Cell(ii + 1, 4).Range.Text = wsCtrl.Cells(rr, 7).Value
        tbl.Cell(ii + 1, 5).Range.Text = wsCtrl.Cells(rr, 8).Value
        
        For ci = 1 To 5
            tbl.Cell(ii + 1, ci).Range.Font.Name = "Arial"
            tbl.Cell(ii + 1, ci).Range.Font.Size = 8
        Next ci
        
        Select Case sv
            Case "Conforme":     tbl.Cell(ii+1,3).Shading.BackgroundPatternColor = RGB(198,239,206): tbl.Cell(ii+1,3).Range.Font.Bold = True
            Case "Non-Conforme": tbl.Cell(ii+1,3).Shading.BackgroundPatternColor = RGB(255,199,206): tbl.Cell(ii+1,3).Range.Font.Bold = True
            Case "Observation":  tbl.Cell(ii+1,3).Shading.BackgroundPatternColor = RGB(255,235,156): tbl.Cell(ii+1,3).Range.Font.Bold = True
            Case "N/A":          tbl.Cell(ii+1,3).Shading.BackgroundPatternColor = RGB(220,220,220): tbl.Cell(ii+1,3).Range.Font.Bold = True
        End Select
        
        If ii Mod 2 = 0 Then
            For ci = 1 To 5
                If tbl.Cell(ii+1,ci).Shading.BackgroundPatternColor = -16777216 Then
                    tbl.Cell(ii+1,ci).Shading.BackgroundPatternColor = RGB(240, 246, 252)
                End If
            Next ci
        End If
    Next ii
End Sub

Sub AddFichesParDomaine(wdDoc As Object, wsCtrl As Worksheet, dom As String)
    Dim lastR As Long: lastR = wsCtrl.UsedRange.Rows.Count + wsCtrl.UsedRange.Row - 1
    Dim found As Boolean: found = False
    Dim rr As Long
    For rr = 5 To lastR
        Dim rv As String: rv = wsCtrl.Cells(rr, 3).Value
        If rv <> "" And IsNumeric(wsCtrl.Cells(rr, 2).Value) Then
            Dim rd As String
            If Left(rv, 7) = "NIS2-AM" Or Left(rv, 7) = "DORA-AM" Then
                rd = "Access Management"
            ElseIf Left(rv, 7) = "NIS2-SO" Or Left(rv, 7) = "DORA-SO" Then
                rd = "Sécurité Opérationnelle"
            ElseIf Left(rv, 7) = "NIS2-CY" Or Left(rv, 7) = "DORA-CY" Then
                rd = "Cybersécurité"
            Else
                rd = ""
            End If
            If rd = dom Then
                found = True
                Call AddFicheUnitaire(wdDoc, wsCtrl, rr)
                Call AddPara(wdDoc, "", 6, False, 0)
            End If
        End If
    Next rr
    If Not found Then Call AddPara(wdDoc, "Aucun contrôle dans ce domaine.", 10, False, 0)
End Sub

Sub AddFicheUnitaire(wdDoc As Object, wsCtrl As Worksheet, rr As Long)
    Dim tbl As Object
    Set tbl = wdDoc.Tables.Add( _
        wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1), 5, 2)
    On Error Resume Next: tbl.Style = "Table Grid": On Error GoTo 0
    tbl.Columns(1).Width = 108  ' ~1.5 inch
    tbl.Columns(2).Width = 360  ' ~5 inch
    
    Dim sv As String: sv = wsCtrl.Cells(rr, 6).Value
    
    ' Titre
    tbl.Cell(1, 1).Merge tbl.Cell(1, 2)
    With tbl.Cell(1, 1)
        .Range.Text = wsCtrl.Cells(rr, 3).Value & "  —  " & wsCtrl.Cells(rr, 4).Value
        .Range.Font.Bold = True: .Range.Font.Name = "Arial": .Range.Font.Size = 9
        .Range.Font.Color = RGB(255, 255, 255)
        .Shading.BackgroundPatternColor = RGB(31, 56, 100)
    End With
    
    Dim labels(3) As String
    labels(0) = "Référence réglementaire"
    labels(1) = "Description / Procédure de test"
    labels(2) = "Statut"
    labels(3) = "Constatations / Preuves collectées"
    
    Dim vals(3) As String
    vals(0) = wsCtrl.Cells(rr, 8).Value
    vals(1) = wsCtrl.Cells(rr, 5).Value
    vals(2) = sv
    vals(3) = wsCtrl.Cells(rr, 7).Value
    
    Dim ri As Integer
    For ri = 0 To 3
        With tbl.Cell(ri + 2, 1)
            .Range.Text = labels(ri)
            .Range.Font.Bold = True: .Range.Font.Name = "Arial": .Range.Font.Size = 8
            .Shading.BackgroundPatternColor = RGB(214, 228, 240)
        End With
        With tbl.Cell(ri + 2, 2)
            .Range.Text = vals(ri)
            .Range.Font.Name = "Arial": .Range.Font.Size = 8
            If ri = 2 Then
                .Range.Font.Bold = True
                Select Case sv
                    Case "Conforme":     .Shading.BackgroundPatternColor = RGB(198, 239, 206)
                    Case "Non-Conforme": .Shading.BackgroundPatternColor = RGB(255, 199, 206)
                    Case "Observation":  .Shading.BackgroundPatternColor = RGB(255, 235, 156)
                    Case "N/A":          .Shading.BackgroundPatternColor = RGB(220, 220, 220)
                End Select
            End If
        End With
    Next ri
End Sub

Sub AddTablePlanActions(wdDoc As Object, wsCtrl As Worksheet)
    Dim lastR As Long: lastR = wsCtrl.UsedRange.Rows.Count + wsCtrl.UsedRange.Row - 1
    Dim aRows(100) As Long: Dim aCnt As Long: aCnt = 0
    Dim rr As Long
    For rr = 5 To lastR
        Dim sv As String: sv = wsCtrl.Cells(rr, 6).Value
        If sv = "Non-Conforme" Or sv = "Observation" Then
            aCnt = aCnt + 1: aRows(aCnt) = rr
        End If
    Next rr
    
    If aCnt = 0 Then
        Call AddPara(wdDoc, "Aucune non-conformité ou observation identifiée.", 10, False, RGB(0, 128, 0))
        Exit Sub
    End If
    
    Dim tbl As Object
    Set tbl = wdDoc.Tables.Add( _
        wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1), aCnt + 1, 4)
    On Error Resume Next: tbl.Style = "Table Grid": On Error GoTo 0
    
    Dim hdrs(3) As String
    hdrs(0) = "Réf.": hdrs(1) = "Intitulé": hdrs(2) = "Statut": hdrs(3) = "Recommandation / Plan d'action"
    
    Dim ci As Integer
    For ci = 0 To 3
        With tbl.Cell(1, ci + 1)
            .Range.Text = hdrs(ci)
            .Range.Font.Bold = True: .Range.Font.Name = "Arial": .Range.Font.Size = 9
            .Range.Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(31, 56, 100)
        End With
    Next ci
    
    Dim ii As Long
    For ii = 1 To aCnt
        rr = aRows(ii)
        sv = wsCtrl.Cells(rr, 6).Value
        tbl.Cell(ii+1,1).Range.Text = wsCtrl.Cells(rr, 3).Value
        tbl.Cell(ii+1,2).Range.Text = wsCtrl.Cells(rr, 4).Value
        tbl.Cell(ii+1,3).Range.Text = sv
        tbl.Cell(ii+1,4).Range.Text = wsCtrl.Cells(rr, 7).Value
        For ci = 1 To 4
            tbl.Cell(ii+1,ci).Range.Font.Name = "Arial"
            tbl.Cell(ii+1,ci).Range.Font.Size = 9
        Next ci
        If sv = "Non-Conforme" Then
            tbl.Cell(ii+1,3).Shading.BackgroundPatternColor = RGB(255,199,206)
            tbl.Cell(ii+1,3).Range.Font.Bold = True
        Else
            tbl.Cell(ii+1,3).Shading.BackgroundPatternColor = RGB(255,235,156)
            tbl.Cell(ii+1,3).Range.Font.Bold = True
        End If
    Next ii
End Sub
