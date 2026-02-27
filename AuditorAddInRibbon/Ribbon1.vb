'-------------------------------------------------------------------------
' Created by : Liao Yan Ying
'Date :  03.25.2017
'Purpose: Tick marks For auditors
'Version: 1.0.0
'Copyright:This addin I created are licensed under 
'          the Creative Commons Attribution 3.0 License, 
'          which means you can use them for personal stuff,
'          use them for commercial stuff, And change them 
'          however you Like. In exchange, just give Me credit 
'          For the design And tell your friends/colleagues about it:)

'-------------------------------------------------------------------------
'=========================================================
' TICKMARK_Pt — PADRÃO OFICIAL PARA TODOS OS _Click (CONTRATO)
'
' SELEÇÃO / ALVOS
' 1) Seleção múltipla: suportar sel.Areas (múltiplas áreas) e inserir em cada célula.
' 2) Mescladas: se MergeCells=True, atuar só em MergeArea.Cells(1,1) (topo-esquerda).
'    Usar HashSet para evitar duplicar inserção em merge/áreas.
'
' NÃO EMPILHAR
' 3) Antes de inserir, remover tick anterior DO MESMO TIPO na mesma célula,
'    identificado por AlternativeText = "TM|<TickId>|<CellAddr>".
'
' ANCORAGEM / POSICIONAMENTO
' 4) Padrão: ancorado à ESQUERDA da célula + centralizado verticalmente.
'    Left = cell.Left + innerOffset; Top = cell.Top + (cell.Height - shp.Height)/2.
' 5) Placement obrigatório: shp.Placement = xlMoveAndSize (acompanhar célula).
' 6) Não usar ancoragem à direita como padrão (instável com redimensionamento de coluna).
'
' TAMANHO
' 7) Imagem: altura ~75% da célula; preservar proporção; limitar largura ~45% da célula.
'
' TAGGING OBRIGATÓRIO (PARA REMOÇÃO)
' 8) AlternativeText: "TM|<TickId>|<CellAddr>" (obrigatório).
' 9) Name (backup): "TM_<TickId>_..." (se falhar por colisão, ignorar; não quebrar).
'
' REMOÇÃO CENTRAL (ClearAllShape / Remove all)
' 10) Apagar SOMENTE objetos do add-in:
'     - TickMarks: AlternativeText começa com "TM|" OU Name começa com "TM_"
'     - Group9: AlternativeText começa com "GP9|" OU Name começa com "GP9_"
'     Nunca apagar shapes do usuário.
'
' RENDERIZAÇÃO (PREFERÊNCIA)
' 11) Preferencial: ticks devem ser inseridos como IMAGEM do próprio botão do Ribbon
'     (Ribbon1.resx) via ComponentResourceManager("<ControlName>.Image") + AddPicture (PNG temporário).
' 12) Não escrever símbolos diretamente no conteúdo da célula (ex.: "✔") e não depender de fonte.
' 13) TextBox (texto) só é permitido quando NÃO houver ícone próprio no Ribbon (exceção justificada).
'
' PARÂMETROS VISUAIS PADRÃO
' 14) innerOffset padrão:
'     - IMAGEM: Const innerOffset = 0.0R  (mais colado possível à borda esquerda)
'     - TEXTO (TextBox): Const innerOffset = 1.0R (ajuste fino; pode variar por estética)
' 15) TextBox (quando permitido): alinhar texto à esquerda dentro do box:
'     box.TextFrame.HorizontalAlignment = xlHAlignLeft
'     e usar box estreito (largura controlada) para evitar “folga” visual.
'
' REGRA EXTRA — Group9 / Work Paper
' 16) Todo _Click do Group9 que crie linhas/formas/conectores deve taguear cada shape criado
'     com GP9| e/ou GP9_ usando groupId único por clique (p/ ClearAllShape apagar corretamente).
'=========================================================

Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub CapP_Click(sender As Object, e As RibbonControlEventArgs) Handles CapP.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "P"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .Font.Italic = True
        End With
    End Sub
    Private Sub footed_Click(sender As Object, e As RibbonControlEventArgs) Handles footed.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "f"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .Font.Italic = True
        End With
    End Sub
    Private Sub Tick_Click(sender As Object, e As RibbonControlEventArgs) Handles tick.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim sel As Excel.Range = TryCast(app.Selection, Excel.Range)
        If sel Is Nothing Then Exit Sub

        Dim ws As Excel.Worksheet = TryCast(app.ActiveSheet, Excel.Worksheet)
        If ws Is Nothing Then Exit Sub

        'Imagem do próprio botão tick (Ribbon1.resx)
        Dim img As System.Drawing.Image = Nothing
        Try
            Dim crm As New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
            img = TryCast(crm.GetObject("tick.Image"), System.Drawing.Image)
        Catch
            img = Nothing
        End Try
        If img Is Nothing Then Exit Sub

        Dim ratio As Double = CDbl(img.Width) / Math.Max(1.0R, CDbl(img.Height))

        ' Alvos (multi-seleção e merge-aware)
        Dim targetCells As New List(Of Excel.Range)()
        Dim keySet As New HashSet(Of String)(StringComparer.Ordinal)

        Try
            For Each area As Excel.Range In sel.Areas
                For Each cellObj As Object In area.Cells
                    Dim c As Excel.Range = TryCast(cellObj, Excel.Range)
                    If c Is Nothing Then Continue For
                    If CBool(c.MergeCells) Then c = c.MergeArea.Cells(1, 1)

                    Dim addr As String = c.Address(False, False)
                    Dim cellKey As String = "TM|tick|" & addr
                    If keySet.Add(cellKey) Then targetCells.Add(c)
                Next
            Next
        Catch
            Exit Sub
        End Try

        If targetCells.Count = 0 Then Exit Sub

        ' Remover qualquer Tick / Tick2 / X existente na célula (FS não é removido)
        Try
            For Each c As Excel.Range In targetCells
                For i As Integer = ws.Shapes.Count To 1 Step -1
                    Try
                        Dim shp As Excel.Shape = ws.Shapes.Item(i)
                        If shp Is Nothing Then Continue For

                        Dim tag As String = ""
                        Dim nm As String = ""
                        Try : tag = shp.AlternativeText : Catch : End Try
                        Try : nm = shp.Name : Catch : End Try

                        ' Apaga Tick, Tick2 ou X na célula
                        If ((Not String.IsNullOrEmpty(tag) AndAlso
                         (tag.StartsWith("TM|tick|") OrElse tag.StartsWith("TM|Tick2|") OrElse tag.StartsWith("TM|X|"))) _
                        OrElse (Not String.IsNullOrEmpty(nm) AndAlso
                                (nm.StartsWith("TM_tick_") OrElse nm.StartsWith("TM_Tick2_") OrElse nm.StartsWith("TM_X_")))) _
                        AndAlso shp.TopLeftCell.Address(False, False) = c.Address(False, False) Then
                            shp.Delete()
                        End If
                    Catch
                    End Try
                Next
            Next
        Catch
        End Try

        'Salvar PNG temporário
        Dim tmpPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                "tickmark_tick_" & Guid.NewGuid().ToString("N") & ".png")
        Try
            Try
                img.Save(tmpPath, System.Drawing.Imaging.ImageFormat.Png)
            Catch
                Using bmp As New System.Drawing.Bitmap(img)
                    bmp.Save(tmpPath, System.Drawing.Imaging.ImageFormat.Png)
                End Using
            End Try

            Dim prevScreenUpdating As Boolean = app.ScreenUpdating
            app.ScreenUpdating = False

            Try
                Const innerOffset As Double = 0.0R

                For Each c As Excel.Range In targetCells
                    Dim targetH As Double = Math.Max(8.0R, c.Height * 0.75R)
                    Dim targetW As Double = targetH * ratio
                    Dim maxW As Double = Math.Max(8.0R, c.Width * 0.45R)
                    If targetW > maxW Then
                        targetW = maxW
                        targetH = targetW / ratio
                    End If

                    Dim addr As String = c.Address(False, False)
                    Dim cellKey As String = "TM|tick|" & addr

                    Dim pic As Excel.Shape = ws.Shapes.AddPicture(
                Filename:=tmpPath,
                LinkToFile:=Microsoft.Office.Core.MsoTriState.msoFalse,
                SaveWithDocument:=Microsoft.Office.Core.MsoTriState.msoCTrue,
                Left:=CSng(c.Left),
                Top:=CSng(c.Top),
                Width:=CSng(targetW),
                Height:=CSng(targetH)
            )

                    pic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                    pic.Placement = Excel.XlPlacement.xlMoveAndSize
                    pic.AlternativeText = cellKey
                    Try
                        pic.Name = "TM_tick_" & addr.Replace("$", "").Replace(":", "_") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6)
                    Catch
                    End Try

                    pic.Left = CSng(c.Left + innerOffset)
                    pic.Top = CSng(c.Top + (c.Height - pic.Height) / 2.0R)
                Next

            Finally
                app.ScreenUpdating = prevScreenUpdating
            End Try

        Finally
            Try
                If System.IO.File.Exists(tmpPath) Then System.IO.File.Delete(tmpPath)
            Catch
            End Try
        End Try
        Try
            ' Pega a última célula da seleção
            Dim lastCell As Excel.Range = sel.Cells(sel.Cells.Count)
            ' Move o cursor para a célula imediatamente abaixo
            Dim nextCell As Excel.Range = lastCell.Offset(1, 0)
            nextCell.Select()
        Catch
            ' Ignora qualquer erro (por exemplo, se estiver na última linha)
        End Try
    End Sub
    Private Sub X_Click(sender As Object, e As RibbonControlEventArgs) Handles X.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim sel As Excel.Range = TryCast(app.Selection, Excel.Range)
        If sel Is Nothing Then Exit Sub

        Dim ws As Excel.Worksheet = TryCast(app.ActiveSheet, Excel.Worksheet)
        If ws Is Nothing Then Exit Sub

        Dim tickId As String = "X"

        ' Multi-seleção e merge-aware
        Dim targetCells As New List(Of Excel.Range)()
        Dim keySet As New HashSet(Of String)(StringComparer.Ordinal)

        For Each area As Excel.Range In sel.Areas
            For Each cellObj As Object In area.Cells
                Dim c As Excel.Range = TryCast(cellObj, Excel.Range)
                If c Is Nothing Then Continue For
                If CBool(c.MergeCells) Then c = c.MergeArea.Cells(1, 1)
                Dim cellKey As String = "TM|" & tickId & "|" & c.Address(False, False)
                If keySet.Add(cellKey) Then targetCells.Add(c)
            Next
        Next

        If targetCells.Count = 0 Then Exit Sub

        ' Remover qualquer Tick / Tick2 / X existente na célula (FS não é removido)
        Try
            For Each c As Excel.Range In targetCells
                For i As Integer = ws.Shapes.Count To 1 Step -1
                    Try
                        Dim shp As Excel.Shape = ws.Shapes.Item(i)
                        If shp Is Nothing Then Continue For

                        Dim tag As String = ""
                        Dim nm As String = ""
                        Try : tag = shp.AlternativeText : Catch : End Try
                        Try : nm = shp.Name : Catch : End Try

                        ' Verifica TM|tick, TM|Tick2, TM|X e os nomes TM_tick_, TM_Tick2_, TM_X_
                        If ((Not String.IsNullOrEmpty(tag) AndAlso
                         (tag.StartsWith("TM|tick|") OrElse tag.StartsWith("TM|Tick2|") OrElse tag.StartsWith("TM|X|"))) _
                        OrElse (Not String.IsNullOrEmpty(nm) AndAlso
                                (nm.StartsWith("TM_tick_") OrElse nm.StartsWith("TM_Tick2_") OrElse nm.StartsWith("TM_X_")))) _
                        AndAlso shp.TopLeftCell.Address(False, False) = c.Address(False, False) Then
                            shp.Delete()
                        End If
                    Catch
                        ' Ignorar erro
                    End Try
                Next
            Next
        Catch
        End Try

        ' Inserir X como TextBox
        Dim prevScreenUpdating As Boolean = app.ScreenUpdating
        app.ScreenUpdating = False

        Try
            Const innerOffset As Double = 0.5R

            For Each c As Excel.Range In targetCells
                Dim addr As String = c.Address(False, False)
                Dim cellKey As String = "TM|" & tickId & "|" & addr

                Dim targetH As Double = Math.Max(12.0R, c.Height * 0.75R)
                Dim targetW As Double = targetH ' quadrado para o X
                Dim maxW As Double = Math.Max(12.0R, c.Width * 0.45R)
                If targetW > maxW Then targetW = maxW

                Dim box As Excel.Shape = ws.Shapes.AddTextbox(
                Orientation:=Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                Left:=CSng(c.Left),
                Top:=CSng(c.Top),
                Width:=CSng(targetW),
                Height:=CSng(targetH)
            )

                Try
                    box.TextFrame.Characters.Text = "X"
                    box.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    box.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    box.TextFrame.AutoSize = False

                    ' Margens internas zeradas
                    Try
                        box.TextFrame.MarginLeft = 0
                        box.TextFrame.MarginRight = 0
                        box.TextFrame.MarginTop = 0
                        box.TextFrame.MarginBottom = 0
                    Catch
                    End Try

                    Try : box.TextFrame.WordWrap = False : Catch : End Try

                    ' Fonte / estilo
                    box.TextFrame.Characters.Font.Bold = True
                    box.TextFrame.Characters.Font.Color = RGB(255, 0, 0)
                Catch
                    Try : box.Delete() : Catch : End Try
                    Continue For
                End Try

                ' Sem borda / preenchimento
                Try : box.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse : Catch : End Try
                Try : box.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse : Catch : End Try

                ' Placement e tagging
                Try : box.Placement = Excel.XlPlacement.xlMoveAndSize : Catch : End Try
                Try : box.AlternativeText = cellKey : Catch : End Try
                Try
                    Dim safeAddr As String = addr.Replace("$", "").Replace(":", "_")
                    box.Name = "TM_X_" & safeAddr & "_" & Guid.NewGuid().ToString("N").Substring(0, 6)
                Catch
                End Try

                ' Posição final: esquerda + centro vertical
                box.Left = CSng(c.Left + innerOffset)
                box.Top = CSng(c.Top + (c.Height - box.Height) / 2.0R)
            Next

        Finally
            app.ScreenUpdating = prevScreenUpdating
        End Try
        Try
            ' Pega a última célula da seleção
            Dim lastCell As Excel.Range = sel.Cells(sel.Cells.Count)
            ' Move o cursor para a célula imediatamente abaixo
            Dim nextCell As Excel.Range = lastCell.Offset(1, 0)
            nextCell.Select()
        Catch
            ' Ignora qualquer erro (por exemplo, se estiver na última linha)
        End Try
    End Sub

    Private Sub FontName1_Click(sender As Object, e As RibbonControlEventArgs) Handles FontName1.Click
        Globals.ThisAddIn.Application.Selection.Font.Name = "標楷體"
    End Sub

    Private Sub FontName2_Click(sender As Object, e As RibbonControlEventArgs) Handles FontName2.Click
        Globals.ThisAddIn.Application.Selection.Font.Name = "微軟正黑體"
    End Sub

    Private Sub PBC_Click(sender As Object, e As RibbonControlEventArgs) Handles PBC.Click
        Dim shape As Object
        shape = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 50, 25)
        shape.TextFrame2.TextRange.text = "PBC"
        shape.line.forecolor.rgb = RGB(102, 178, 255)
        With shape.textframe2.textrange.font
            .size = 16
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
    End Sub

    Private Sub Calendar_Click(sender As Object, e As RibbonControlEventArgs)
        Dim ctpCal As Microsoft.Office.Tools.CustomTaskPane =
 Globals.ThisAddIn.CustomTaskPanes.Add(New CalendarTaskPaneControl, "Calendar")
        ctpCal.Width = 232
        ctpCal.Visible = True
    End Sub


    Private Sub CheckBox_Click(sender As Object, e As RibbonControlEventArgs) Handles CheckBox.Click
        Dim checkbox As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        checkbox = Globals.ThisAddIn.Application.ActiveSheet.checkboxes.add _
            (Left, Top, Width, Height)
        With checkbox
            .Caption = ""
        End With
    End Sub

    Private Sub ClearCheckBox_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearCheckBox.Click
        Dim sht As Excel.Worksheet
        sht = Globals.ThisAddIn.Application.ActiveSheet
        For Each cb In sht.CheckBoxes
            If cb.TopLeftCell.Address = Globals.ThisAddIn.Application.ActiveCell.Address Then cb.Delete
        Next
    End Sub

    Private Sub Submit_Click(sender As Object, e As RibbonControlEventArgs)
        Dim username As String
        Dim submit As String
        Dim time As Date = Now

        username = Environment.UserName
        submit = "Prepared by" & " " & username & " " & "on " & time.Month _
                                                              & "/" & time.Day _
                                                              & "/" & time.Year

        With Globals.ThisAddIn.Application.ActiveCell
            .Value = submit
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub Review_Click(sender As Object, e As RibbonControlEventArgs)
        Dim username As String
        Dim review As String
        Dim time As Date = Now

        username = Environment.UserName

        review = "Reviewed by" & " " & username & " " & "on " & time.Month _
                                                             & "/" & time.Day _
                                                             & "/" & time.Year
        With Globals.ThisAddIn.Application.ActiveCell
            .Value = review
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub RedPen_Click(sender As Object, e As RibbonControlEventArgs) Handles RedPen.Click
        Dim connector As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        connector = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 15 + Width, Top + 2, Left - 2.5 + Width, Top + 9.25)
        With connector
            .Name = "RedPen"
            .line.weight = 2
            .line.ForeColor.RGB = RGB(255, 0, 0)
        End With
    End Sub

    Private Sub BluePen_Click(sender As Object, e As RibbonControlEventArgs) Handles BluePen.Click
        Dim connector As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        connector = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 15 + Width, Top - 1, Left - 2.5 + Width, Top + 6.25)
        With connector
            .Name = "BluePen"
            .line.weight = 2
            .line.ForeColor.RGB = RGB(0, 0, 255)
        End With
    End Sub

    Private Sub BlackPen_Click(sender As Object, e As RibbonControlEventArgs) Handles BlackPen.Click
        Dim connector As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        connector = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 15 + Width, Top - 6, Left - 2.5 + Width, Top + 1.25)
        With connector
            .Name = "BlackPen"
            .line.weight = 2
            .line.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End Sub

    Private Sub ClearAllShape_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearAllShape.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim ws As Excel.Worksheet = TryCast(app.ActiveSheet, Excel.Worksheet)
        If ws Is Nothing Then Exit Sub

        Dim resp As Integer = MsgBox(
        "ATENÇÃO: isto irá apagar TODOS os TickMarks do grupo Work Paper nesta planilha (" & ws.Name & ")." & vbCrLf &
        "Essa ação não pode ser desfeita." & vbCrLf & vbCrLf &
        "Deseja continuar?",
        vbYesNo + vbExclamation,
        "TickMark - Remover todos"
    )
        If resp <> vbYes Then Exit Sub

        Try
            For i As Integer = ws.Shapes.Count To 1 Step -1
                Try
                    Dim shp As Excel.Shape = ws.Shapes.Item(i)
                    If shp Is Nothing Then Continue For

                    Dim nm As String = ""
                    Dim tag As String = ""

                    Try : nm = shp.Name : Catch : nm = "" : End Try
                    Try : tag = shp.AlternativeText : Catch : tag = "" : End Try

                    Dim isTickMark As Boolean =
                    (Not String.IsNullOrEmpty(tag) AndAlso tag.StartsWith("TM|", StringComparison.Ordinal)) OrElse
                    (Not String.IsNullOrEmpty(nm) AndAlso nm.StartsWith("TM_", StringComparison.Ordinal))

                    Dim isGroup9 As Boolean =
                    (Not String.IsNullOrEmpty(tag) AndAlso tag.StartsWith("GP9|", StringComparison.Ordinal)) OrElse
                    (Not String.IsNullOrEmpty(nm) AndAlso nm.StartsWith("GP9_", StringComparison.Ordinal))

                    If isTickMark OrElse isGroup9 Then
                        shp.Delete()
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    Private Sub TagGroup9Shape(shp As Excel.Shape, Optional groupId As String = Nothing)
        If shp Is Nothing Then Exit Sub

        If String.IsNullOrEmpty(groupId) Then
            groupId = Guid.NewGuid().ToString("N")
        End If

        'marca por AlternativeText (não exige unicidade)
        shp.AlternativeText = "GP9|" & groupId

        'tenta marcar por Name também (pode falhar por colisão)
        Try
            shp.Name = "GP9_" & groupId
        Catch
        End Try
    End Sub
    Private Sub RedArrow_Click(sender As Object, e As RibbonControlEventArgs) Handles RedArrow.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connector As Object
        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top

        Dim location As Object
        location = app.inputbox("Selecione uma célula", "Desenhando uma seta", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If
        Dim TargetLeft As Double = location.left
        Dim TargetTop As Double = location.Top
        Dim TargetHeight As Double = location.Height


        connector = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 4, Top, TargetLeft + 4, TargetTop + TargetHeight)
        With connector
            .line.weight = 1
            .line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
        End With
    End Sub

    Private Sub HyperLink_Click(sender As Object, e As RibbonControlEventArgs) Handles HyperLink.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        'Origin Sheet
        Dim strBeginningSheetName As String = app.Activesheet.name
        Dim BeginningSheet As Object = Globals.ThisAddIn.Application.ActiveSheet
        'Store the Original sheet & cell location
        Dim OriSheetCellAddress As String = appCell.address

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height
        Dim connector1 As Excel.Shape
        Dim connector2 As Excel.Shape

        Dim location As Object
        location = app.inputbox("Select a cell", "Creating a HyperLink", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If
        Dim TargetLeft As Double = location.left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height

        'Get the Target's sheet name
        Dim strDestinationSheetName As String = location.Parent.Name

        'Create the textbox in the beginning sheet
        Dim shapeBeg As Object
        shapeBeg = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width * 0.796, Top + 7, 100, Height)
        shapeBeg.TextFrame2.TextRange.text = strDestinationSheetName
        With shapeBeg.textframe2.textrange.font
            .size = 9
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        With shapeBeg
            .Name = "box1"
            .fill.visible = Microsoft.Office.Core.MsoTriState.msoFalse
            .line.visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End With
        shapeBeg.TextFrame.autosize = True

        'Add the target address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=shapeBeg,
             Address:="", SubAddress:=strDestinationSheetName & "!" & location.address,
             TextToDisplay:=strDestinationSheetName)

        'Add the connector1 to point to the textbox
        connector1 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width * 0.796, Top + 12, Left + Width * 0.796 + 6, Top + 14)
        With connector1
            .Name = "line1"
            .Line.Weight = 1
            .Line.ForeColor.RGB = RGB(255, 0, 0)
        End With

        'Create the hyperlink next to the Target Cell in Target Sheet
        'Select the target sheet first
        app.Sheets(strDestinationSheetName).Activate
        Dim shape As Object
        shape = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 40, TargetTop - 10, 100, TargetHeight)
        shape.TextFrame2.TextRange.text = strBeginningSheetName
        With shape.textframe2.textrange.font
            .size = 9
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        With shape
            .Name = "box2"
            .fill.visible = Microsoft.Office.Core.MsoTriState.msoFalse
            .line.visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End With
        shape.TextFrame.autosize = True
        'Add the origin address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=shape,
             Address:="", SubAddress:=strBeginningSheetName & "!" & OriSheetCellAddress,
             TextToDisplay:=strBeginningSheetName)

        'Add the connector2 to point to the textbox
        connector2 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + 10, TargetTop + 5, TargetLeft + 1, TargetTop + 1)
        With connector2
            .Name = "line2"
            .Line.Weight = 1
            .Line.ForeColor.RGB = RGB(255, 0, 0)
        End With

        app.Sheets(strBeginningSheetName).Activate
    End Sub

    Private Sub SheetArrow_Click(sender As Object, e As RibbonControlEventArgs) Handles SheetArrow.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting arrows", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height


        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left, Top + Height, Left - 5, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop, TargetLeft + TargetWidth + 5, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left, Top, Left - 5, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight, TargetLeft + TargetWidth + 5, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left, Top + Height, Left - 13, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight, TargetLeft + TargetWidth + 13, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height, Left + Width + 13, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft, TargetTop + TargetHeight, TargetLeft - 13, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width, Top + Height, Left + Width + 5, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft, TargetTop, TargetLeft - 5, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width, Top, Left + Width + 5, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft, TargetTop + TargetHeight, TargetLeft - 5, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top, Left + Width, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight, TargetLeft + TargetWidth, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height, Left + Width, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop, TargetLeft + TargetWidth, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        End If

    End Sub

    Private Sub DoubleLine_Click(sender As Object, e As RibbonControlEventArgs) Handles DoubleLine.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        With app.selection.borders(Excel.XlBordersIndex.xlEdgeBottom)
            .linestyle = Excel.XlLineStyle.xlDouble
        End With
    End Sub

    Private Sub a_Click(sender As Object, e As RibbonControlEventArgs) Handles a.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "a"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub b_Click(sender As Object, e As RibbonControlEventArgs) Handles b.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "b"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub c_Click(sender As Object, e As RibbonControlEventArgs) Handles c.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "c"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub siga_Click(sender As Object, e As RibbonControlEventArgs) Handles siga.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "∑a="
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub sigb_Click(sender As Object, e As RibbonControlEventArgs) Handles sigb.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "∑b="
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub sigc_Click(sender As Object, e As RibbonControlEventArgs) Handles sigc.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "∑c="
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub AJE_Click(sender As Object, e As RibbonControlEventArgs) Handles AJE.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "AJE<>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 10
            .font.bold = True
        End With
    End Sub

    Private Sub RJE_Click(sender As Object, e As RibbonControlEventArgs) Handles RJE.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "RJE<>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 10
            .font.bold = True
        End With
    End Sub

    Private Sub o_Click(sender As Object, e As RibbonControlEventArgs) Handles o.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "○"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub OO_Click(sender As Object, e As RibbonControlEventArgs) Handles OO.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "◎"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub star1_Click(sender As Object, e As RibbonControlEventArgs) Handles star1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "☆"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub star2_Click(sender As Object, e As RibbonControlEventArgs) Handles star2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "★"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub tri1_Click(sender As Object, e As RibbonControlEventArgs) Handles tri1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "△"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub tri2_Click(sender As Object, e As RibbonControlEventArgs) Handles tri2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "▲"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub dia1_Click(sender As Object, e As RibbonControlEventArgs) Handles dia1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "◇"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub dia2_Click(sender As Object, e As RibbonControlEventArgs) Handles dia2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "◆"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub squ1_Click(sender As Object, e As RibbonControlEventArgs) Handles squ1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "□"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub squ2_Click(sender As Object, e As RibbonControlEventArgs) Handles squ2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "■"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub sigma_Click(sender As Object, e As RibbonControlEventArgs) Handles sigma.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "∑"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub divided_Click(sender As Object, e As RibbonControlEventArgs) Handles divided.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "/"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub alpha_Click(sender As Object, e As RibbonControlEventArgs) Handles alpha.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "α"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub beta_Click(sender As Object, e As RibbonControlEventArgs) Handles beta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "β"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub gamma_Click(sender As Object, e As RibbonControlEventArgs) Handles gamma.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "γ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub CommaStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles CommaStyle.Click

        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim rng As Excel.Range = TryCast(app.Selection, Excel.Range)

        If rng Is Nothing Then Exit Sub

        rng.NumberFormatLocal = "#.##0,00"

    End Sub

    Private Sub InsertColumn_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertColumn.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell
        app.selection.entirecolumn.Offset(0, 1).insert
        appCell.Offset(0, 1).Select
        app.selection.columnwidth = 2

    End Sub

    Private Sub Preparer2_Click(sender As Object, e As RibbonControlEventArgs) Handles Preparer2.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim username As String
        Dim submit As String
        Dim time As Date = Now

        username = app.UserName
        submit = "Elaborado por" & " " & username & " " & "em " & time.Day _
                                                              & "/" & time.Month _
                                                              & "/" & time.Year

        With app.ActiveCell
            .Value = submit
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub Reviewer2_Click(sender As Object, e As RibbonControlEventArgs) Handles Reviewer2.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim username As String
        Dim review As String
        Dim time As Date = Now

        username = app.UserName
        review = "Revisado por" & " " & username & " " & "em " & time.Day _
                                                              & "/" & time.Month _
                                                              & "/" & time.Year

        With app.ActiveCell
            .Value = review
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub Note_Click(sender As Object, e As RibbonControlEventArgs) Handles Note.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<Note>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note1_Click(sender As Object, e As RibbonControlEventArgs) Handles Note1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N1>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note2_Click(sender As Object, e As RibbonControlEventArgs) Handles Note2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N2>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note3_Click(sender As Object, e As RibbonControlEventArgs) Handles Note3.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N3>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note4_Click(sender As Object, e As RibbonControlEventArgs) Handles Note4.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N4>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note5_Click(sender As Object, e As RibbonControlEventArgs) Handles Note5.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N5>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Copyright_ToggleButton_Click(sender As Object, e As RibbonControlEventArgs) Handles Copyright_ToggleButton.Click
        Globals.ThisAddIn.TaskPane.Visible =
    TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
    End Sub

    Private Sub ClearHyperLink_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearHyperLink.Click

        Dim app As Object = Globals.ThisAddIn.Application
        Dim sht As Excel.Worksheet = Nothing
        Dim shp As Excel.Shape = Nothing

        sht = app.ActiveWorkbook.ActiveSheet
        For Each shp In sht.Shapes
            If shp.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox And (shp.Name = "box1" Or shp.Name = "box2") Then
                shp.Select()
                shp.Delete()
            ElseIf shp.Type = Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight And (shp.Name = "line1" Or shp.Name = "line2") Then
                shp.Select()
                shp.Delete()
            End If
        Next

    End Sub

    Private Sub Quote_Click(sender As Object, e As RibbonControlEventArgs) Handles Quote.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        Dim Quote As String = Nothing
        Randomize()
        ' Generate random value between 1 and 900.
        Dim value As Integer = Int((900 * Rnd()) + 1)
        Select Case value
            Case 1
                Quote = "O pássaro que se arrisca a cair é o pássaro que aprende a voar."
                Exit Select
            Case 2
                Quote = "Trabalhe duro em silêncio. Deixe o sucesso fazer barulho."
                Exit Select
            Case 3
                Quote = "Uma das decisões mais difíceis da vida é escolher entre ir embora ou tentar mais uma vez. - Ziad K. Abdelnour"
                Exit Select
            Case 4
                Quote = "Enquanto você achar que seu passado foi ruim, é sinal de que você está melhorando. - Louis C.K."
                Exit Select
            Case 5
                Quote = "E se eu te pedisse para listar tudo o que você ama… quanto tempo levaria até você se incluir?"
                Exit Select
            Case 6
                Quote = "Daqui a um ano, você vai desejar ter começado hoje."
                Exit Select
            Case 7
                Quote = "Pare de ter medo do que pode dar errado e comece a se empolgar com o que pode dar certo."
                Exit Select
            Case 8
                Quote = "Você não precisa ver a escada inteira; apenas dê o primeiro passo. - Martin Luther King Jr."
                Exit Select
            Case 9
                Quote = "Acredite que você pode — e você já está no meio do caminho. - Theodore Roosevelt"
                Exit Select
            Case 10
                Quote = "O sucesso é a soma de pequenos esforços repetidos dia após dia. - Robert Collier"
                Exit Select
            Case 11
                Quote = "Disciplina é escolher entre o que você quer agora e o que você quer mais. - Abraham Lincoln"
                Exit Select
            Case 12
                Quote = "Não espere por oportunidade. Crie a oportunidade."
                Exit Select
            Case 13
                Quote = "Sonhos não funcionam a menos que você trabalhe por eles."
                Exit Select
            Case 14
                Quote = "Faça hoje o que os outros não fazem, para amanhã fazer o que os outros não conseguem."
                Exit Select
            Case 15
                Quote = "Se você cansar, descanse. Só não desista."
                Exit Select
            Case 16
                Quote = "Pequenos progressos todos os dias geram grandes resultados."
                Exit Select
            Case 17
                Quote = "A única pessoa que você deve tentar superar é quem você foi ontem."
                Exit Select
            Case 18
                Quote = "A dor que você sente hoje será a força que você terá amanhã."
                Exit Select
            Case 19
                Quote = "O que você faz hoje pode melhorar todos os seus amanhãs."
                Exit Select
            Case 20
                Quote = "Você não precisa ser ótimo para começar, mas precisa começar para ser ótimo."
                Exit Select
            Case 21
                Quote = "Seja mais forte do que suas desculpas."
                Exit Select
            Case 22
                Quote = "Não pare quando estiver cansado. Pare quando terminar."
                Exit Select
            Case 23
                Quote = "A persistência transforma o impossível em possível."
                Exit Select
            Case 24
                Quote = "Acredite no processo."
                Exit Select
            Case 25
                Quote = "Não desista só porque está difícil. Difícil é o que faz valer a pena."
                Exit Select
            Case 26
                Quote = "O medo é temporário. O arrependimento dura."
                Exit Select
            Case 27
                Quote = "Você é capaz de mais do que imagina."
                Exit Select
            Case 28
                Quote = "Tudo o que você quer está do outro lado do esforço."
                Exit Select
            Case 29
                Quote = "Ação cura o medo."
                Exit Select
            Case 30
                Quote = "Faça acontecer."
                Exit Select
            Case 31
                Quote = "Você não precisa de motivação; precisa de compromisso."
                Exit Select
            Case 32
                Quote = "Comece onde você está. Use o que você tem. Faça o que você pode. - Arthur Ashe"
                Exit Select
            Case 33
                Quote = "A melhor forma de prever o futuro é criá-lo. - Peter Drucker"
                Exit Select
            Case 34
                Quote = "Um dia ou dia um. Você escolhe."
                Exit Select
            Case 35
                Quote = "Foque no progresso, não na perfeição."
                Exit Select
            Case 36
                Quote = "Sua zona de conforto é uma ótima lugar, mas nada cresce lá."
                Exit Select
            Case 37
                Quote = "Se você quer resultados diferentes, faça coisas diferentes."
                Exit Select
            Case 38
                Quote = "Não compare seus bastidores com o palco de alguém."
                Exit Select
            Case 39
                Quote = "Continue. Mesmo devagar, continue."
                Exit Select
            Case 40
                Quote = "O que importa é a constância."
                Exit Select
            Case 41
                Quote = "Você vence quando não desiste."
                Exit Select
            Case 42
                Quote = "Você não falha; você aprende."
                Exit Select
            Case 43
                Quote = "A cada dia, um pouco melhor."
                Exit Select
            Case 44
                Quote = "A mente acredita no que você repete."
                Exit Select
            Case 45
                Quote = "Você não é suas circunstâncias; você é suas escolhas."
                Exit Select
            Case 46
                Quote = "Treine até ficar inevitável."
                Exit Select
            Case 47
                Quote = "A disciplina vai te levar onde a motivação não alcança."
                Exit Select
            Case 48
                Quote = "Não deixe para amanhã o passo que você pode dar hoje."
                Exit Select
            Case 49
                Quote = "O esforço de hoje é o resultado de amanhã."
                Exit Select
            Case 50
                Quote = "Você só perde quando desiste."
                Exit Select
            Case 51
                Quote = "Seja paciente. Grandes coisas levam tempo."
                Exit Select
            Case 52
                Quote = "Se for importante pra você, você vai dar um jeito. Se não for, vai arrumar uma desculpa."
                Exit Select
            Case 53
                Quote = "Seja consistente, não perfeito."
                Exit Select
            Case 54
                Quote = "Faça o básico bem feito — todos os dias."
                Exit Select
            Case 55
                Quote = "Você está mais perto do que pensa."
                Exit Select
            Case 56
                Quote = "A coragem não é ausência de medo; é agir apesar do medo."
                Exit Select
            Case 57
                Quote = "A diferença entre quem consegue e quem não consegue é a persistência."
                Exit Select
            Case 58
                Quote = "Cada repetição conta."
                Exit Select
            Case 59
                Quote = "Sem pressão, não há crescimento."
                Exit Select
            Case 60
                Quote = "Você não precisa de mais tempo; precisa de mais prioridade."
                Exit Select
            Case 61
                Quote = "A motivação te coloca em movimento; o hábito te mantém em movimento."
                Exit Select
            Case 62
                Quote = "Faça o que tem que ser feito, mesmo quando você não estiver a fim."
                Exit Select
            Case 63
                Quote = "Grandes resultados vêm de pequenas decisões diárias."
                Exit Select
            Case 64
                Quote = "Você não tem que ver o final para dar o primeiro passo."
                Exit Select
            Case 65
                Quote = "O fracasso é apenas um passo a caminho do sucesso."
                Exit Select
            Case 66
                Quote = "Pense grande, comece pequeno, aja agora."
                Exit Select
            Case 67
                Quote = "Se você quer algo que nunca teve, precisa fazer algo que nunca fez."
                Exit Select
            Case 68
                Quote = "A disciplina é liberdade."
                Exit Select
            Case 69
                Quote = "Você se torna aquilo que pratica."
                Exit Select
            Case 70
                Quote = "Seja a energia que você quer atrair."
                Exit Select
            Case 71
                Quote = "Você consegue. Só continua."
                Exit Select
            Case 72
                Quote = "O seu futuro precisa de você hoje."
                Exit Select
            Case 73
                Quote = "O melhor investimento é em você."
                Exit Select
            Case 74
                Quote = "Faça por você."
                Exit Select
            Case 75
                Quote = "O esforço sempre volta em resultado."
                Exit Select
            Case 76
                Quote = "É no desconforto que você evolui."
                Exit Select
            Case 77
                Quote = "Seja teimoso com o objetivo, flexível com o caminho."
                Exit Select
            Case 78
                Quote = "Fé e ação: as duas juntas."
                Exit Select
            Case 79
                Quote = "Não espere se sentir pronto. Comece e fique pronto no caminho."
                Exit Select
            Case 80
                Quote = "O que você repete vira padrão."
                Exit Select
            Case 81
                Quote = "Você não está atrasado; você está no seu tempo."
                Exit Select
            Case 82
                Quote = "Você é mais forte do que pensa."
                Exit Select
            Case 83
                Quote = "Faça o que é certo, não o que é fácil."
                Exit Select
            Case 84
                Quote = "Todo dia conta."
                Exit Select
            Case 85
                Quote = "A melhor hora para começar foi ontem. A segunda melhor é agora."
                Exit Select
            Case 86
                Quote = "A constância vence o talento quando o talento não é constante."
                Exit Select
            Case 87
                Quote = "Seja disciplina, não desculpa."
                Exit Select
            Case 88
                Quote = "Se o plano não funcionar, mude o plano — não a meta."
                Exit Select
            Case 89
                Quote = "Continue mesmo quando ninguém estiver vendo."
                Exit Select
            Case 90
                Quote = "Você não precisa de sorte; precisa de rotina."
                Exit Select
            Case 91
                Quote = "Aprenda a ser consistente com você."
                Exit Select
            Case 92
                Quote = "Não negocie com a preguiça."
                Exit Select
            Case 93
                Quote = "O que você tolera, você mantém."
                Exit Select
            Case 94
                Quote = "Quem quer dá um jeito; quem não quer, dá uma desculpa."
                Exit Select
            Case 95
                Quote = "Sem esforço, sem resultado."
                Exit Select
            Case 96
                Quote = "O corpo conquista o que a mente decide."
                Exit Select
            Case 97
                Quote = "Você não precisa ser perfeito. Precisa ser constante."
                Exit Select
            Case 98
                Quote = "Foco. Força. Fé."
                Exit Select
            Case 99
                Quote = "Acredite, trabalhe, conquiste."
                Exit Select
            Case 100
                Quote = "O difícil de hoje é o forte de amanhã."
                Exit Select
            Case 101
                Quote = "Se você quer mudar, comece com um passo."
                Exit Select
            Case 102
                Quote = "A disciplina te coloca onde o sonho sozinho não chega."
                Exit Select
            Case 103
                Quote = "Menos desculpas, mais ação."
                Exit Select
            Case 104
                Quote = "Você não precisa de permissão para crescer."
                Exit Select
            Case 105
                Quote = "Você não é fraco; você só está começando."
                Exit Select
            Case 106
                Quote = "Chega de adiar: faça."
                Exit Select
            Case 107
                Quote = "O melhor projeto é você."
                Exit Select
            Case 108
                Quote = "Você vai se agradecer por não ter parado."
                Exit Select
            Case 109
                Quote = "Cada dia é uma chance de recomeçar."
                Exit Select
            Case 110
                Quote = "O sucesso é treinável."
                Exit Select
            Case 111
                Quote = "Sua melhor versão exige consistência."
                Exit Select
            Case 112
                Quote = "Faça o que precisa ser feito, e o resultado vem."
                Exit Select
            Case 113
                Quote = "A vitória começa na decisão."
                Exit Select
            Case 114
                Quote = "Um passo de cada vez — mas sem parar."
                Exit Select
            Case 115
                Quote = "Trabalhe no silêncio; deixe os resultados falarem."
                Exit Select
            Case 116
                Quote = "Não é sobre velocidade; é sobre direção."
                Exit Select
            Case 117
                Quote = "Quando doer, lembre: está funcionando."
                Exit Select
            Case 118
                Quote = "Você não está sem tempo; está sem prioridade."
                Exit Select
            Case 119
                Quote = "A disciplina começa quando a motivação acaba."
                Exit Select
            Case 120
                Quote = "Eu não conto abdominais. Só começo a contar quando começa a doer. Quando sinto dor, é aí que começo a contar — porque é aí que realmente conta. - Muhammad Ali"
                Exit Select
            Case 121
                Quote = "Apenas continue. O que parece tão difícil agora um dia será só o seu aquecimento."
                Exit Select
            Case 122
                Quote = "Se você não construir o seu sonho, alguém vai te contratar para ajudar a construir o sonho dela. - Tony Gaskins"
                Exit Select
            Case 123
                Quote = "O que você faria se não tivesse medo? - Sheryl Sandberg"
                Exit Select
            Case 124
                Quote = "Coragem nem sempre ruge. Às vezes, coragem é aquela voz baixinha no fim do dia que diz: amanhã eu tento de novo."
                Exit Select
            Case 125
                Quote = "Faça o que você pode, com o que você tem, onde você está. - Theodore Roosevelt"
                Exit Select
            Case 126
                Quote = "Se você acha que é pequeno demais para fazer diferença, tente dormir com um mosquito. - Dalai Lama XIV"
                Exit Select
            Case 127
                Quote = "A forma como passamos nossos dias é, claro, a forma como passamos nossas vidas. - Annie Dillard"
                Exit Select
            Case 128
                Quote = "Eu já sei como é desistir. Quero ver o que acontece se eu não desistir."
                Exit Select
            Case 129
                Quote = "Pare de deixar pessoas que fazem tão pouco por você controlarem tanto da sua mente, sentimentos e emoções. - Will Smith"
                Exit Select
            Case 130
                Quote = "Hábitos que levaram anos para ser construídos não mudam em um dia. - Susan Powter"
                Exit Select
            Case 131
                Quote = "Preocupe-se um pouco todos os dias e, ao longo da vida, você perde alguns anos. Se algo estiver errado, conserte se puder. Mas treine-se para não se preocupar: preocupação nunca conserta nada. - Ernest Hemingway"
                Exit Select
            Case 132
                Quote = "Não é a carga que te derruba; é a forma como você a carrega. - Lou Holtz"
                Exit Select
            Case 133
                Quote = "Seja sempre uma versão de primeira linha de você mesmo, e não uma versão de segunda linha de outra pessoa. - Judy Garland"
                Exit Select
            Case 134
                Quote = "Melhor ser quem sorriu do que quem não sorriu de volta. - Mari Gayatri Stein"
                Exit Select
            Case 135
                Quote = "Estou ocupado demais cuidando do meu próprio jardim para notar se o seu é mais verde."
                Exit Select
            Case 136
                Quote = "Disciplina é a ponte entre metas e conquistas. - Jim Rohn"
                Exit Select
            Case 137
                Quote = "Se você não pode fazer grandes coisas, faça pequenas coisas de um jeito grandioso. - Napoleon Hill"
                Exit Select
            Case 138
                Quote = "Uma meta é um sonho com prazo. - Napoleon Hill"
                Exit Select
            Case 139
                Quote = "Motivação é o que faz você começar. Hábito é o que faz você continuar."
                Exit Select
            Case 140
                Quote = "Sucesso é a soma de pequenos esforços, repetidos dia após dia. - Robert Collier"
                Exit Select
            Case 141
                Quote = "Quem desiste nunca vence, e quem vence nunca desiste. - Napoleon Hill"
                Exit Select
            Case 142
                Quote = "Você erra 100% dos arremessos que não faz. - Wayne Gretzky"
                Exit Select
            Case 143
                Quote = "Caia sete vezes, levante-se oito. - Provérbio japonês"
                Exit Select
            Case 144
                Quote = "Não deseje que fosse mais fácil. Deseje ser melhor. - Jim Rohn"
                Exit Select
            Case 145
                Quote = "Dane-se. Vamos fazer. - Richard Branson"
                Exit Select
            Case 146
                Quote = "Nada é particularmente difícil se você dividir em pequenas tarefas. - Henry Ford"
                Exit Select
            Case 147
                Quote = "Educação formal te dá sustento. Autoeducação te dá fortuna. - Jim Rohn"
                Exit Select
            Case 148
                Quote = "O que não se começa hoje, nunca se termina amanhã. - Johann Wolfgang von Goethe"
                Exit Select
            Case 149
                Quote = "Eu não falhei. Apenas descobri 10.000 maneiras que não funcionam. - Thomas A. Edison"
                Exit Select
            Case 150
                Quote = "O segredo para avançar é começar. - Mark Twain"
                Exit Select
            Case 151
                Quote = "A melhor forma de prever o futuro é criá-lo. - Peter Drucker"
                Exit Select
            Case 152
                Quote = "Pense grande — ou fique em casa. - Eliza Dushku"
                Exit Select
            Case 153
                Quote = "Qualidade é fazer certo quando ninguém está olhando. - Henry Ford"
                Exit Select
            Case 154
                Quote = "Sucesso não é definitivo, fracasso não é fatal: o que conta é a coragem de continuar. - Winston Churchill"
                Exit Select
            Case 155
                Quote = "Aonde quer que você vá, vá com todo o seu coração. - Confúcio"
                Exit Select
            Case 156
                Quote = "Se você está passando pelo inferno, continue. - Winston Churchill"
                Exit Select
            Case 157
                Quote = "Se você não está fazendo ondas, é porque não está chutando forte o bastante."
                Exit Select
            Case 158
                Quote = "Encontre o que você ama e deixe isso te consumir. - Charles Bukowski"
                Exit Select
            Case 159
                Quote = "O verdadeiro empreendedor faz; não apenas sonha. - Nolan Bushnell"
                Exit Select
            Case 160
                Quote = "Quando você esgotar todas as possibilidades, lembre-se: você ainda não esgotou. - Thomas A. Edison"
                Exit Select
            Case 161
                Quote = "A maneira de começar é parar de falar e começar a fazer. - Walt Disney"
                Exit Select
            Case 162
                Quote = "Ideias são fáceis. Execução é difícil. - Guy Kawasaki"
                Exit Select
            Case 163
                Quote = "É difícil vencer alguém que nunca desiste. - Babe Ruth"
                Exit Select
            Case 164
                Quote = "Se tudo parece sob controle, é porque você não está indo rápido o bastante. - Mario Andretti"
                Exit Select
            Case 165
                Quote = "Sempre entregue mais do que esperam. - Larry Page"
                Exit Select
            Case 166
                Quote = "Fracasso é apenas uma oportunidade de recomeçar — desta vez, de forma mais inteligente. - Henry Ford"
                Exit Select
            Case 167
                Quote = "Acorde sempre com um sorriso, sabendo que hoje você vai se divertir realizando o que outros têm medo de fazer."
                Exit Select
            Case 168
                Quote = "Quando você encontra uma ideia na qual não consegue parar de pensar, provavelmente é uma boa para perseguir. - Josh James"
                Exit Select
            Case 169
                Quote = "Esqueça os concorrentes; foque nos seus clientes. - Jack Ma"
                Exit Select
            Case 170
                Quote = "Continue faminto. Continue tolo. - Steve Jobs"
                Exit Select
            Case 171
                Quote = "Altas expectativas são a chave para tudo. - Sam Walton"
                Exit Select
            Case 172
                Quote = "Muitas vezes, as pessoas não sabem o que querem até você mostrar a elas. - Steve Jobs"
                Exit Select
            Case 173
                Quote = "Não se preocupe com o fracasso; você só precisa acertar uma vez. - Drew Houston"
                Exit Select
            Case 174
                Quote = "O pessimista vê dificuldade em toda oportunidade; o otimista vê oportunidade em toda dificuldade. - Winston Churchill"
                Exit Select
            Case 175
                Quote = "Tudo é difícil antes de ficar fácil. - Thomas Fuller"
                Exit Select
            Case 176
                Quote = "A vida é 10% o que acontece com você e 90% como você reage. - Charles R. Swindoll"
                Exit Select
            Case 177
                Quote = "Ou você comanda o dia, ou o dia comanda você. - Jim Rohn"
                Exit Select
            Case 178
                Quote = "Se a oportunidade não bater à porta, construa uma porta. - Milton Berle"
                Exit Select
            Case 179
                Quote = "Você não pode ter uma vida positiva com uma mente negativa. - Joyce Meyer"
                Exit Select
            Case 180
                Quote = "Transforme sempre uma situação negativa em uma situação positiva. - Michael Jordan"
                Exit Select
            Case 181
                Quote = "Seu tempo é limitado, então não o desperdice vivendo a vida de outra pessoa. - Steve Jobs"
                Exit Select
            Case 182
                Quote = "Se você quer vencer, não pode ter medo de perder. - Floyd Mayweather"
                Exit Select
            Case 183
                Quote = "Sua vida só melhora quando você melhora. - Brian Tracy"
                Exit Select
            Case 184
                Quote = "Você não se torna o melhor por querer; você se torna o melhor por fazer."
                Exit Select
            Case 185
                Quote = "O sucesso normalmente vem para quem está ocupado demais para procurá-lo. - Henry David Thoreau"
                Exit Select
            Case 186
                Quote = "Você não encontra vontade; você cria vontade."
                Exit Select
            Case 187
                Quote = "Não deixe a opinião dos outros virar a sua realidade. - Les Brown"
                Exit Select
            Case 188
                Quote = "Se você pode sonhar, você pode realizar. - Walt Disney"
                Exit Select
            Case 189
                Quote = "A diferença entre o possível e o impossível está na determinação."
                Exit Select
            Case 190
                Quote = "A vida começa no fim da sua zona de conforto. - Neale Donald Walsch"
                Exit Select
            Case 191
                Quote = "Não é sobre ter tempo. É sobre fazer tempo."
                Exit Select
            Case 192
                Quote = "Seja tão bom que não possam te ignorar. - Steve Martin"
                Exit Select
            Case 193
                Quote = "A vitória pertence ao mais perseverante. - Napoleão Bonaparte"
                Exit Select
            Case 194
                Quote = "Você não precisa ser motivado; precisa ser disciplinado."
                Exit Select
            Case 195
                Quote = "Acredite em você e todo o resto se encaixa. - Norman Vincent Peale"
                Exit Select
            Case 196
                Quote = "O que você faz em silêncio, grita nos resultados."
                Exit Select
            Case 197
                Quote = "Desistir não é uma opção."
                Exit Select
            Case 198
                Quote = "Você não precisa estar pronto. Você só precisa começar."
                Exit Select
            Case 199
                Quote = "Quando você quer muito, você encontra um caminho."
                Exit Select
            Case 200
                Quote = "Seja a pessoa que você precisava quando era mais novo."
                Exit Select
            Case 201
                Quote = "Se você quer algo que nunca teve, faça o que nunca fez."
                Exit Select
            Case 202
                Quote = "Foque no que você pode controlar."
                Exit Select
            Case 203
                Quote = "O sucesso é construído no que você faz repetidamente."
                Exit Select
            Case 204
                Quote = "Corra o risco ou perca a chance."
                Exit Select
            Case 205
                Quote = "Você só cresce quando se desafia."
                Exit Select
            Case 206
                Quote = "A consistência vence."
                Exit Select
            Case 207
                Quote = "Quem tem um porquê enfrenta quase qualquer como. - Friedrich Nietzsche"
                Exit Select
            Case 208
                Quote = "Progresso, não perfeição."
                Exit Select
            Case 209
                Quote = "A disciplina é o melhor amigo do seu futuro."
                Exit Select
            Case 210
                Quote = "Você é o seu maior projeto."
                Exit Select
            Case 211
                Quote = "O que você faz hoje define quem você será amanhã."
                Exit Select
            Case 212
                Quote = "Grandes coisas nunca vêm da zona de conforto."
                Exit Select
            Case 213
                Quote = "Aja como se fosse impossível falhar."
                Exit Select
            Case 214
                Quote = "Não conte os dias; faça os dias contarem. - Muhammad Ali"
                Exit Select
            Case 215
                Quote = "A sua única limitação é você."
                Exit Select
            Case 216
                Quote = "Trabalhe duro até que seu ídolo vire seu rival."
                Exit Select
            Case 217
                Quote = "A dor é temporária. O orgulho é para sempre."
                Exit Select
            Case 218
                Quote = "Você não se arrependerá do esforço."
                Exit Select
            Case 219
                Quote = "Se você falhar, pelo menos falhou tentando."
                Exit Select
            Case 220
                Quote = "O esforço de hoje é o seu resultado de amanhã."
                Exit Select
            Case 221
                Quote = "Treine sua mente para ver o bem em tudo."
                Exit Select
            Case 222
                Quote = "Não espere. O momento nunca será perfeito. - Napoleon Hill"
                Exit Select
            Case 223
                Quote = "Você é mais forte do que qualquer desculpa."
                Exit Select
            Case 224
                Quote = "Foco e disciplina: o resto é ruído."
                Exit Select
            Case 225
                Quote = "A mudança começa quando você decide."
                Exit Select
            Case 226
                Quote = "Seja a prova viva de que dá para virar o jogo."
                Exit Select
            Case 227
                Quote = "Não é sorte. É trabalho."
                Exit Select
            Case 228
                Quote = "Quanto mais você sua no treino, menos sangra na guerra."
                Exit Select
            Case 229
                Quote = "Faça o que é difícil agora e viva o que é fácil depois."
                Exit Select
            Case 230
                Quote = "Você não precisa ser excelente; precisa ser constante."
                Exit Select
            Case 231
                Quote = "Quando você quer parar, é aí que você deve continuar."
                Exit Select
            Case 232
                Quote = "Continue. Você está mais perto do que imagina."
                Exit Select
            Case 233
                Quote = "A disciplina te dá aquilo que a motivação promete."
                Exit Select
            Case 234
                Quote = "Se você tem um sonho, proteja-o. - À Procura da Felicidade"
                Exit Select
            Case 235
                Quote = "Você pode até ser lento. Só não pare."
                Exit Select
            Case 236
                Quote = "Não existe atalho. Existe processo."
                Exit Select
            Case 237
                Quote = "Cada dia é um novo começo."
                Exit Select
            Case 238
                Quote = "Seja melhor do que ontem."
                Exit Select
            Case 239
                Quote = "Persistência: continue quando quiser desistir."
                Exit Select
            Case 240
                Quote = "Nada muda se nada mudar."
                Exit Select
            Case 241
                Quote = "Aprendemos muito mais sabedoria com o fracasso do que com o sucesso. Muitas vezes descobrimos o que vamos fazer ao descobrir o que não vamos fazer. - Samuel Smiles"
                Exit Select
            Case 242
                Quote = "Se você muda a forma como vê as coisas, as coisas que você vê mudam. - Dr. Wayne Dyer"
                Exit Select
            Case 243
                Quote = "Nossa mente pode moldar como algo será, porque agimos de acordo com nossas expectativas. - Federico Fellini"
                Exit Select
            Case 244
                Quote = "O que você se torna é muito mais importante do que o que você obtém. O que você obtém é influenciado pelo que você se torna. - Jim Rohn"
                Exit Select
            Case 245
                Quote = "Podemos fazer qualquer coisa que queiramos, se nos mantivermos nisso tempo suficiente. - Helen Keller"
                Exit Select
            Case 246
                Quote = "O tolo precisa de companhia; o sábio, de solitude. - Desconhecido"
                Exit Select
            Case 247
                Quote = "Uma luta constante, uma batalha sem cessar para fazer nascer o sucesso em meio a condições hostis, é o preço de todas as grandes conquistas. - Marden"
                Exit Select
            Case 248
                Quote = "O universo está cheio de coisas mágicas esperando pacientemente que nossa inteligência fique mais afiada. - Eden Phillpotts"
                Exit Select
            Case 249
                Quote = "Você precisa fazer a coisa que acha que não consegue fazer. - Eleanor Roosevelt"
                Exit Select
            Case 250
                Quote = "Quando seus desejos são fortes o bastante, você parece possuir poderes sobre-humanos para realizar. - Napoleon Hill"
                Exit Select
            Case 251
                Quote = "Você é a única pessoa na Terra que pode usar a sua capacidade. - Zig Ziglar"
                Exit Select
            Case 252
                Quote = "Lembre-se: nenhum esforço que fazemos para alcançar algo belo é jamais perdido. - Helen Keller"
                Exit Select
            Case 253
                Quote = "Tornar-se uma estrela pode não ser o seu destino, mas ser o melhor que você pode ser é uma meta que você pode definir para si. - Bryan Lindsay"
                Exit Select
            Case 254
                Quote = "Uma alma sem um grande objetivo é como um navio sem leme. - Eileen Caddy"
                Exit Select
            Case 255
                Quote = "Não antecipe problemas nem se preocupe com o que talvez nunca aconteça. Permaneça na luz do sol. - Benjamin Franklin"
                Exit Select
            Case 256
                Quote = "Perder a paciência é perder a batalha. - Mahatma Gandhi"
                Exit Select
            Case 257
                Quote = "Gentileza é difícil de se dar embora: ela sempre volta para quem deu. - Ralph Scott"
                Exit Select
            Case 258
                Quote = "Entregue ao mundo o melhor que você tem, e o melhor voltará para você. - Madeline Bridges"
                Exit Select
            Case 259
                Quote = "A oportunidade é perdida pela maioria das pessoas porque vem vestida de macacão e parece trabalho. - Thomas Edison"
                Exit Select
            Case 260
                Quote = "O segredo da disciplina é a motivação. Quando alguém está suficientemente motivado, a disciplina cuida de si mesma. - Alexander Paterson"
                Exit Select
            Case 261
                Quote = "Conhecer os outros é inteligência; conhecer a si mesmo é verdadeira sabedoria. Dominar os outros é força; dominar a si mesmo é verdadeiro poder. - Lao-Tzu"
                Exit Select
            Case 262
                Quote = "Fracassar não significa que você é um fracasso... só significa que você ainda não teve sucesso. - Robert Schuller"
                Exit Select
            Case 263
                Quote = "Tudo o que acontece, acontece como deveria; e, se você observar com cuidado, verá que é assim. - Marcus Antoninus"
                Exit Select
            Case 264
                Quote = "É preciso um grande homem para dar um conselho sensato com tato, mas um ainda maior para aceitá-lo com elegância. - J.C. Macaulay"
                Exit Select
            Case 265
                Quote = "Sua meta deve estar fora do alcance, mas não fora da vista. - Anita DeFrantz"
                Exit Select
            Case 266
                Quote = "Você não se torna extremamente bem-sucedido sem enfrentar e superar vários problemas extremamente desafiadores. - Mark Victor Hansen"
                Exit Select
            Case 267
                Quote = "Comece fazendo o que é necessário; depois, o que é possível; e, de repente, você estará fazendo o impossível. - São Francisco de Assis"
                Exit Select
            Case 268
                Quote = "Erros e falhas são a disciplina pela qual avançamos. - William Ellery Channing"
                Exit Select
            Case 269
                Quote = "SUCESSO NÃO É ACIDENTE. É trabalho duro, perseverança, aprendizado, estudo, sacrifício e, acima de tudo, amor pelo que você faz. - Pelé"
                Exit Select
            Case 270
                Quote = "Hoje é a primeira página em branco de um livro de 365 páginas. Faça valer a pena."
                Exit Select
            Case 271
                Quote = "É uma coisa terrível, eu acho, esperar até estar pronto. Tenho a sensação de que ninguém nunca está pronto para fazer nada. Quase não existe ""estar pronto"". Existe apenas o agora. E você pode muito bem fazer agora. Em geral, agora é tão bom quanto qualquer momento. - Hugh Laurie"
                Exit Select
            Case 272
                Quote = "Já amou alguém tanto que faria qualquer coisa por ela? Então faça dessa pessoa você mesmo e faça o que quiser. - Harvey Specter"
                Exit Select
            Case 273
                Quote = "Às vezes, as pessoas com maior potencial demoram mais para encontrar o próprio caminho, porque a sensibilidade é uma espada de dois gumes: está no coração do brilho, mas também torna mais vulnerável às dores da vida. Ainda bem que não somos punidos por descobrir nosso propósito mais tarde. A alma não sabe nada sobre prazos. - Jeff Brown"
                Exit Select
            Case 274
                Quote = "Infelizmente, a única forma de algumas pessoas aprenderem a te valorizar é te perdendo."
                Exit Select
            Case 275
                Quote = "Ninguém é superior, ninguém é inferior, mas ninguém é igual também. As pessoas são únicas, incomparáveis. VOCÊ É VOCÊ. EU SOU EU. - Osho"
                Exit Select
            Case 276
                Quote = "Siga o seu coração, mas leve o seu cérebro junto."
                Exit Select
            Case 277
                Quote = "Por que as pessoas escondem o amor e demonstram o ódio tão abertamente?"
                Exit Select
            Case 278
                Quote = "Uma das melhores lições da vida é aprender a permanecer calmo."
                Exit Select
            Case 279
                Quote = "Eu não penso dentro da caixa, nem fora dela. Eu nem sei onde essa caixa está."
                Exit Select
            Case 280
                Quote = "Muita gente se preocupa com aparência e bens materiais e acaba ignorando totalmente a própria personalidade."
                Exit Select
            Case 281
                Quote = "Conheci poucas pessoas que realmente me entendem. O resto acha que eu estou sempre bravo, sarcástico ou sendo grosso."
                Exit Select
            Case 282
                Quote = "Me diga para não fazer algo e eu faço duas vezes e ainda tiro foto."
                Exit Select
            Case 283
                Quote = "Se você não consegue lidar com estresse, não consegue lidar com sucesso."
                Exit Select
            Case 284
                Quote = "Tudo bem ter a cabeça um pouco bagunçada — todo mundo tem. O problema é quando o coração fica amargo."
                Exit Select
            Case 285
                Quote = "Não desista. Você já está com dor; já está machucado. Tire uma recompensa disso."
                Exit Select
            Case 286
                Quote = "Uma pessoa pode fazer a diferença — e todos deveriam tentar. - JFK"
                Exit Select
            Case 287
                Quote = "Se isso te empolga e te assusta ao mesmo tempo, talvez valha a pena tentar."
                Exit Select
            Case 288
                Quote = "Eu ainda amo as pessoas que já amei, mesmo que eu atravesse a rua para evitá-las. - Uma Thurman"
                Exit Select
            Case 289
                Quote = "Amor não é motivo para tolerar desrespeito."
                Exit Select
            Case 290
                Quote = "Seja o motivo de alguém acreditar na bondade das pessoas. - Lori Deschene"
                Exit Select
            Case 291
                Quote = "Relaxe e confie no tempo da sua vida. Você vai encontrar sua carreira. Vai encontrar o relacionamento certo. Vai se tornar a pessoa que sempre quis ser. Só não esqueça de valorizar quem você é agora. - Ruben Chavez"
                Exit Select
            Case 292
                Quote = "Você precisa chegar a um ponto em que seu humor não muda por causa das ações insignificantes de outra pessoa."
                Exit Select
            Case 293
                Quote = "Nada se compara à dor na barriga de tanto rir com seus melhores amigos."
                Exit Select
            Case 294
                Quote = "Eu gosto mais da minha cama do que da maioria das pessoas."
                Exit Select
            Case 295
                Quote = "A vida é uma sopa e eu sou um garfo."
                Exit Select
            Case 296
                Quote = "Nunca assuma que o barulhento é forte e o quieto é fraco."
                Exit Select
            Case 297
                Quote = "Melhor um ""ops"" do que um ""e se...?""."
                Exit Select
            Case 298
                Quote = "Seja discreto. Nem todo mundo precisa saber tudo sobre você."
                Exit Select
            Case 299
                Quote = "Quando você vir algo bonito em alguém, diga. Leva segundos para falar, mas para a pessoa pode durar a vida inteira."
                Exit Select
            Case 300
                Quote = "É mais fácil criar crianças fortes do que consertar adultos quebrados. - F. Douglass"
                Exit Select
            Case 301
                Quote = "Estou em construção no momento. Obrigado pela sua paciência."
                Exit Select
            Case 302
                Quote = "Três C's na vida: Escolhas, Chances, Mudanças. Você precisa escolher arriscar, ou sua vida nunca vai mudar."
                Exit Select
            Case 303
                Quote = "Lembre-se do porquê você começou."
                Exit Select
            Case 304
                Quote = "Amor de verdade não te encontra no seu melhor. Ele te encontra no seu caos. - J.S. PARK"
                Exit Select
            Case 305
                Quote = "Você nem sempre vai ter um bom dia, mas sempre pode enfrentar um dia ruim com uma boa atitude."
                Exit Select
            Case 306
                Quote = "Pessoas que atacam repetidamente sua confiança e autoestima sabem muito bem do seu potencial, mesmo que você ainda não saiba. - Wayne Gerard Trotman"
                Exit Select
            Case 307
                Quote = "Eu sou muito paciente e dou muitas segundas chances, mas não sou santo: eu tenho meus limites."
                Exit Select
            Case 308
                Quote = "Nem todo dia é um bom dia — viva mesmo assim. Nem todo mundo vai te dizer a verdade — seja honesto mesmo assim. Nem todo mundo que você ama vai te amar de volta — ame mesmo assim. Nem todos os acordos são justos — seja justo mesmo assim."
                Exit Select
            Case 309
                Quote = "As pessoas falam muito. Então eu observo o que elas fazem. - via Quotes 'nd Notes"
                Exit Select
            Case 310
                Quote = "Escolha pessoas que escolhem você."
                Exit Select
            Case 311
                Quote = "Eu parei de me explicar quando percebi que as pessoas só entendem a partir do próprio nível de percepção."
                Exit Select
            Case 312
                Quote = "Acredite: existe alguém por aí que vai se apaixonar pelo seu tipo de loucura."
                Exit Select
            Case 313
                Quote = "Treine sua mente para ver o bem em toda situação."
                Exit Select
            Case 314
                Quote = "Deixe que te julguem. Deixe que te entendam errado. Deixe que falem de você. A opinião deles não é problema seu. Você siga gentil, comprometido com o amor e livre na sua autenticidade. Não importa o que façam ou digam, não ouse duvidar do seu valor ou da beleza da sua verdade. Continue brilhando do seu jeito. - Scott Stabile"
                Exit Select
            Case 315
                Quote = "Você pode conhecer amanhã alguém que tem melhores intenções por você do que alguém que você conhece há anos. Tempo não significa nada; caráter, sim."
                Exit Select
            Case 316
                Quote = "Tudo é temporário."
                Exit Select
            Case 317
                Quote = "Eu percebo tudo, mas fico calado."
                Exit Select
            Case 318
                Quote = "Quando doer — observe. A vida está tentando te ensinar alguma coisa. - Anita Krizzan"
                Exit Select
            Case 319
                Quote = "Não me subestime porque eu sou quieto. Eu sei mais do que digo, penso mais do que falo e observo mais do que você imagina. - Michaela Chung"
                Exit Select
            Case 320
                Quote = "Se alguém te corrige e você se ofende, então você tem um problema de ego. - Nouman Ali Khan"
                Exit Select
            Case 321
                Quote = "Se você ajuda alguém esperando algo em troca, você está fazendo negócio — não gentileza."
                Exit Select
            Case 322
                Quote = "A vida é como um livro. Alguns capítulos são tristes, outros são felizes e outros são empolgantes. Mas, se você nunca virar a página, nunca vai saber o que o próximo capítulo reserva."
                Exit Select
            Case 323
                Quote = "Não importa quanto tempo eu leve: eu vou chegar a um lugar bonito."
                Exit Select
            Case 324
                Quote = "Há uma hora de ser gentil e há uma hora de dizer: ""já chega da sua besteira."""
                Exit Select
            Case 325
                Quote = "Dica de vida: quando nada dá certo, vá dormir."
                Exit Select
            Case 326
                Quote = "Ame sempre seus amigos com o coração, não de acordo com seu humor ou necessidade."
                Exit Select
            Case 327
                Quote = "A alma sempre sabe o que fazer para se curar. O desafio é silenciar a mente. - Caroline Myss"
                Exit Select
            Case 328
                Quote = "Nunca implore para alguém ficar na sua vida. Se você manda mensagem, liga, visita e ainda é ignorado, vá embora. Isso se chama ""autorrespeito"". - Steve Wentworth"
                Exit Select
            Case 329
                Quote = "Eu gosto de pessoas diretas. Mesmo quando vocês discordam, pelo menos você sabe onde elas estão e elas não fazem joguinhos."
                Exit Select
            Case 330
                Quote = "Aos 20, a gente se preocupa com o que os outros pensam da gente. Aos 40, a gente não liga. Aos 60, a gente descobre que eles nem estavam pensando na gente. - Ann Landers"
                Exit Select
            Case 331
                Quote = "Sempre lembre: o esforço de alguém reflete o interesse que ela tem por você."
                Exit Select
            Case 332
                Quote = "Converse com pessoas que te fazem ver o mundo de um jeito diferente."
                Exit Select
            Case 333
                Quote = "A todas as pessoas que são amorosas e gentis comigo: obrigado pela luz que vocês colocam na minha vida. - Brigitte Nicole"
                Exit Select
            Case 334
                Quote = "Quando você estiver errado, admita. Quando estiver certo, fique em silêncio. - Ogden Nash"
                Exit Select
            Case 335
                Quote = "Alguns de vocês secretamente não gostam de mim. E eu sei disso — e não tô nem aí."
                Exit Select
            Case 336
                Quote = "O maior problema de comunicação é que não ouvimos para entender; ouvimos para responder."
                Exit Select
            Case 337
                Quote = "Minhas palavras ou atraem uma mente forte, ou ofendem uma mente fraca."
                Exit Select
            Case 338
                Quote = "A solidão é perigosa. Vicia. Quando você percebe o quanto ela é tranquila, não quer mais lidar com pessoas. - Hedonist Poet"
                Exit Select
            Case 339
                Quote = "Você precisa de você mais do que precisa deles. Confia em mim."
                Exit Select
            Case 340
                Quote = "Se existe a menor chance de conseguir algo que te faça feliz, ARRISQUE. A vida é curta demais e a felicidade é rara demais. - A.R. Lucas"
                Exit Select
            Case 341
                Quote = """ocupado demais"" é um mito. As pessoas arrumam tempo para aquilo que é realmente importante para elas."
                Exit Select
            Case 342
                Quote = "Algumas pessoas são só caixas de presente bonitas cheias de mentira."
                Exit Select
            Case 343
                Quote = "Felicidade é… reencontrar um amigo antigo depois de muito tempo e sentir que nada mudou."
                Exit Select
            Case 344
                Quote = "Quando algo bom acontecer, viaje para comemorar. Se algo ruim acontecer, viaje para esquecer. Se nada acontecer, viaje para fazer algo acontecer."
                Exit Select
            Case 345
                Quote = "Estranhos podem virar melhores amigos tão facilmente quanto melhores amigos podem virar estranhos."
                Exit Select
            Case 346
                Quote = "A doença do coração mais perigosa: memória forte. - Nizar Qabbani"
                Exit Select
            Case 347
                Quote = "Eu sou uma pessoa muito reservada, mas também sou um livro aberto. Se você não pergunta… eu não conto."
                Exit Select
            Case 348
                Quote = "Não acredite tão rápido no que você ouve, porque mentiras se espalham mais rápido do que a verdade."
                Exit Select
            Case 349
                Quote = "Eu já não procuro o ""bom"" nas pessoas; eu procuro o real… porque o bom muitas vezes vem vestido de falsidade, mas o real é nu e orgulhoso, não importa as cicatrizes. - Chishala Lishomwa"
                Exit Select
            Case 350
                Quote = "Eu gosto de pessoas com quem eu consigo ter silêncios confortáveis."
                Exit Select
            Case 351
                Quote = "A tecnologia conectou o mundo e desconectou os humanos."
                Exit Select
            Case 352
                Quote = "Você não precisa ser positivo o tempo todo. Está tudo bem se sentir triste, com raiva, irritado, frustrado, com medo ou ansioso. Ter sentimentos não faz de você uma ""pessoa negativa"". Faz de você humano. - Lori Deschene"
                Exit Select
            Case 353
                Quote = "Às vezes você precisa lembrar: nem todo mundo foi criado do jeito que você foi."
                Exit Select
            Case 354
                Quote = "Metade do mundo passa fome e a outra metade tenta emagrecer. - Roseanne Barr"
                Exit Select
            Case 355
                Quote = "Todos nós precisamos de alguém para conversar — alguém que ouça, alguém que entenda."
                Exit Select
            Case 356
                Quote = "Eu não tenho tempo para brigar com egos e mentes pequenas."
                Exit Select
            Case 357
                Quote = "Às vezes… quando você dá a mínima demais, isso acaba te derrubando."
                Exit Select
            Case 358
                Quote = "Não ter motivo para ficar é um bom motivo para ir embora."
                Exit Select
            Case 359
                Quote = "Sorria. Ninguém se importa com como você se sente."
                Exit Select
            Case 360
                Quote = "Ando ficando mais calado esses dias. Falo cada vez menos em público. Mas meus olhos… meus olhos veem tudo."
                Exit Select
            Case 361
                Quote = "A magia acontece quando você não desiste, mesmo querendo desistir. O universo sempre se apaixona por um coração teimoso. - JmStorm"
                Exit Select
            Case 362
                Quote = "Um dia vamos encontrar o que estamos procurando. Ou talvez não. Talvez encontremos algo muito maior do que isso."
                Exit Select
            Case 363
                Quote = "Não seja o motivo de alguém se sentir inseguro. Seja o motivo de alguém se sentir visto, ouvido e apoiado pelo universo inteiro. - Cleo Wade"
                Exit Select
            Case 364
                Quote = "O mundo vai te julgar, não importa o que você faça. Então viva sua vida do jeito que você realmente quer."
                Exit Select
            Case 365
                Quote = "Eu amo a versão das pessoas às 3 da manhã: vulnerável, honesta, real."
                Exit Select
            Case 366
                Quote = "Eu sou frio como gelo. Mas, nas mãos certas, eu derreto."
                Exit Select
            Case 367
                Quote = "Tudo chega até você, no momento certo. Tenha paciência. Seja grato."
                Exit Select
            Case 368
                Quote = "Tenha esperança, mas não espere demais. Olhe para frente, mas não fique parado esperando."
                Exit Select
            Case 369
                Quote = "Quando você começa a olhar para o coração das pessoas em vez do rosto, a vida fica mais clara."
                Exit Select
            Case 370
                Quote = "Entre o que é dito sem intenção e o que é sentido mas não é dito, muito do amor se perde. - Kahlil Gibran"
                Exit Select
            Case 371
                Quote = "Paciência é quando você tinha motivo para ficar bravo, mas escolhe entender."
                Exit Select
            Case 372
                Quote = "A gente conhece todo mundo por um motivo. Ou é uma bênção, ou é uma lição."
                Exit Select
            Case 373
                Quote = "Gente falsa tem uma imagem para manter. Gente de verdade simplesmente não se importa."
                Exit Select
            Case 374
                Quote = "Fique sozinho. Coma sozinho, saia consigo mesmo, durma sozinho. No meio disso, você aprende sobre você, cresce, descobre o que te inspira, constrói seus sonhos e crenças, ganha clareza. E quando encontrar alguém que faça seu corpo vibrar, você vai ter certeza — porque você tem certeza de si. - Bianca Sparacino"
                Exit Select
            Case 375
                Quote = "Você nem sempre precisa de uma razão lógica para tudo. Faça porque você quer, porque é divertido, porque te faz FELIZ."
                Exit Select
            Case 376
                Quote = "Não se preocupe se você ainda não está onde queria. Coisas grandes levam tempo."
                Exit Select
            Case 377
                Quote = "Diga o que você sente. Isso não é grosseria — é ser real."
                Exit Select
            Case 378
                Quote = "Eu não tenho pavio curto; eu só reajo rápido a bobagem."
                Exit Select
            Case 379
                Quote = "Não escale montanhas para que as pessoas vejam você. Escale para que você veja o mundo."
                Exit Select
            Case 380
                Quote = "O problema do mundo é que as pessoas inteligentes ficam cheias de dúvidas, enquanto as pessoas tolas ficam cheias de certezas. - Charles Bukowski"
                Exit Select
            Case 381
                Quote = "Quando você morrer, não vão lembrar do seu carro ou da sua casa. Vão lembrar de quem você foi. Seja uma boa pessoa, não apenas alguém preso a coisas materiais."
                Exit Select
            Case 382
                Quote = "Eu adoro quando a risada de alguém é mais engraçada do que a piada."
                Exit Select
            Case 383
                Quote = "Você nunca sabe o que as pessoas estão passando. Às vezes, quem sorri mais é quem mais luta. Seja gentil."
                Exit Select
            Case 384
                Quote = "Não confunda minha personalidade com minha atitude… Minha personalidade é quem eu sou. Minha atitude depende de quem você é…"
                Exit Select
            Case 385
                Quote = "A gente amadurece com os danos, não com os anos."
                Exit Select
            Case 386
                Quote = "Quem é você? 'Demônio para alguns. Anjo para outros.'"
                Exit Select
            Case 387
                Quote = "Alguns acham que eu sou infeliz, mas não sou. Eu só valorizo o silêncio num mundo que não para de falar."
                Exit Select
            Case 388
                Quote = "Se todo mundo gosta de você, você tem um problema sério."
                Exit Select
            Case 389
                Quote = "Não deixe alguém te tratar mal só porque você ama essa pessoa."
                Exit Select
            Case 390
                Quote = "Por trás de toda mulher bem-sucedida, está ela mesma."
                Exit Select
            Case 391
                Quote = "Eu era quieta, mas não era cega. - Jane Austen"
                Exit Select
            Case 392
                Quote = "Então, se você estiver cansado demais para falar, sente ao meu lado, porque eu também sou fluente em silêncio. - R. Arnold"
                Exit Select
            Case 393
                Quote = "Não me estude… você não vai se formar."
                Exit Select
            Case 394
                Quote = "Quando eu tinha 17 anos, eu admirava pessoas com luxo. Hoje, eu admiro pessoas com paz interior."
                Exit Select
            Case 395
                Quote = "Ser ao mesmo tempo sensível e forte é uma combinação que poucos dominam."
                Exit Select
            Case 396
                Quote = "Fique perto de tudo o que te faz feliz por estar vivo."
                Exit Select
            Case 397
                Quote = "Rostos bonitos estão por toda parte, mas mentes bonitas são difíceis de encontrar."
                Exit Select
            Case 398
                Quote = "Não recebemos uma vida boa ou ruim. Recebemos uma vida. Cabe a nós torná-la boa ou ruim. - Ward Foley"
                Exit Select
            Case 399
                Quote = "Quero lembrar que ninguém vai realizar meus sonhos por mim. É meu trabalho levantar todo dia e caminhar na direção do que é mais profundo no meu coração, aproveitando cada passo da jornada em vez de desejar já estar no fim."
                Exit Select
            Case 400
                Quote = "O que ocupa sua mente controla sua vida."
                Exit Select
            Case 401
                Quote = "Sobre mim: posso ser bem duro. Doce como bala. Frio como gelo. Ou leal como um soldado. Tudo depende de você."
                Exit Select
            Case 402
                Quote = "Às vezes, a melhor terapia é pegar a estrada e ouvir música."
                Exit Select
            Case 403
                Quote = "Espero nunca me cansar do céu noturno, das tempestades, de ver o creme formar galáxias no café. Espero nunca me tornar alguém que não enxerga mais as pequenas coisas bonitas."
                Exit Select
            Case 404
                Quote = "Eu amo lugares que fazem você perceber o quão pequeno você — e seus problemas — são."
                Exit Select
            Case 405
                Quote = "Às vezes as pessoas te machucam e ainda agem como se você tivesse machucado elas."
                Exit Select
            Case 406
                Quote = "Eu sou uma pessoa diferente para pessoas diferentes: irritante para um, talentoso para outro, quieto para alguns, desconhecido para muitos. Mas quem eu sou para mim? - dream-jackson"
                Exit Select
            Case 407
                Quote = "Deixe as coisas irem e virem. O que for para ficar, vai ficar."
                Exit Select
            Case 408
                Quote = "Qualquer um pode dizer que se importa… observe as ações, não as palavras."
                Exit Select
            Case 409
                Quote = "Às vezes, as coisas precisam dar muito errado antes de dar certo."
                Exit Select
            Case 410
                Quote = "A melhor coisa do pior momento da sua vida é que você vê as verdadeiras cores de todo mundo."
                Exit Select
            Case 411
                Quote = "O melhor ainda está por vir. Tenha paciência."
                Exit Select
            Case 412
                Quote = "Nem todo mundo merece conhecer o seu eu verdadeiro. Deixe que critiquem quem eles acham que você é."
                Exit Select
            Case 413
                Quote = "Nosso caráter não é definido pelas batalhas que vencemos ou perdemos, mas pelas batalhas que temos coragem de lutar."
                Exit Select
            Case 414
                Quote = "Você nunca vai entender até acontecer com você."
                Exit Select
            Case 415
                Quote = "Quando você sentir que alguém está te evitando, não insista mais."
                Exit Select
            Case 416
                Quote = "Às vezes, coisas ruins precisam acontecer para te inspirar a mudar e crescer."
                Exit Select
            Case 417
                Quote = "Não mexa com alguém que não tem medo de ficar sozinho. Você vai perder todas as vezes."
                Exit Select
            Case 418
                Quote = "Seja a mesma pessoa no privado, no público e no pessoal."
                Exit Select
            Case 419
                Quote = "Quem já se quebrou sempre consegue amar mais forte do que a maioria. Depois de estar no escuro, você aprende a valorizar tudo o que brilha."
                Exit Select
            Case 420
                Quote = "É difícil achar um amigo fofo, amoroso, generoso, atraente, cuidadoso e inteligente. Meu conselho a todos os meus amigos: não me percam."
                Exit Select
            Case 421
                Quote = "Nunca deixe alguém desperdiçar o seu tempo duas vezes."
                Exit Select
            Case 422
                Quote = "Não dar a mínima é melhor do que vingança."
                Exit Select
            Case 423
                Quote = "Há amigos, há família… e há amigos que viram família."
                Exit Select
            Case 424
                Quote = "No fundo do meu coração, eu sei que sou um solitário. Eu tentei me misturar ao mundo e ser sociável, mas quanto mais pessoas eu conheço, mais eu me decepciono. Então aprendi a aproveitar a mim mesmo, minha família e alguns bons amigos. - Steven Aitchison"
                Exit Select
            Case 425
                Quote = "Que se dane sua coleção de sapatos. Me mostre sua coleção de livros."
                Exit Select
            Case 426
                Quote = "Pare de pensar demais. Você está arruinando sua felicidade."
                Exit Select
            Case 427
                Quote = "O problema não é que eu não tenho tempo. O problema é que eu não tenho energia."
                Exit Select
            Case 428
                Quote = "Pessoas que se importam de verdade não te fazem questionar isso."
                Exit Select
            Case 429
                Quote = "Se alguém quer você na vida dela, vai abrir espaço. Sem desculpas. Sem esforço pela metade."
                Exit Select
            Case 430
                Quote = "Você pode sentir falta de alguém e ainda assim reconhecer que ela não é boa para você."
                Exit Select
            Case 431
                Quote = "Aprenda a se desapegar do que te faz mal."
                Exit Select
            Case 432
                Quote = "Seja maduro o suficiente para pedir desculpas, inteligente o suficiente para aprender com isso e forte o bastante para não repetir."
                Exit Select
            Case 433
                Quote = "Às vezes, a melhor resposta é o silêncio."
                Exit Select
            Case 434
                Quote = "Se você se importa demais, vai se machucar demais."
                Exit Select
            Case 435
                Quote = "A sua paz é mais importante do que provar seu ponto."
                Exit Select
            Case 436
                Quote = "Quanto menos você espera, menos você se decepciona."
                Exit Select
            Case 437
                Quote = "Faça o que é certo, mesmo quando ninguém estiver olhando."
                Exit Select
            Case 438
                Quote = "Seja gentil, mas não seja bobo."
                Exit Select
            Case 439
                Quote = "Não mude por ninguém. Mude por você."
                Exit Select
            Case 440
                Quote = "O universo sempre dá um jeito de colocar as pessoas certas no seu caminho."
                Exit Select
            Case 441
                Quote = "Você não precisa convencer ninguém do seu valor."
                Exit Select
            Case 442
                Quote = "A solidão não é vazia; ela é cheia de respostas."
                Exit Select
            Case 443
                Quote = "Não seja disponível para quem só te procura quando convém."
                Exit Select
            Case 444
                Quote = "Algumas pessoas entram na sua vida só para te ensinar a sair dela."
                Exit Select
            Case 445
                Quote = "Se você não se curar do que te feriu, vai sangrar em quem não te cortou."
                Exit Select
            Case 446
                Quote = "Nem todo mundo merece acesso à sua energia."
                Exit Select
            Case 447
                Quote = "Não confunda atenção com amor."
                Exit Select
            Case 448
                Quote = "Você merece alguém que fique — não alguém que vai e volta."
                Exit Select
            Case 449
                Quote = "Você não precisa ter tudo resolvido para seguir em frente."
                Exit Select
            Case 450
                Quote = "Faça as pazes com o fato de que nem todo mundo vai gostar de você."
                Exit Select
            Case 451
                Quote = "Você não perde pessoas; você se livra de lições."
                Exit Select
            Case 452
                Quote = "Não leve tudo para o lado pessoal. Nem tudo é sobre você."
                Exit Select
            Case 453
                Quote = "Você não deve explicações para quem não te respeita."
                Exit Select
            Case 454
                Quote = "Se alguém te quiser, vai te tratar como prioridade."
                Exit Select
            Case 455
                Quote = "A felicidade é uma escolha diária."
                Exit Select
            Case 456
                Quote = "Você não precisa agradar todo mundo. Você precisa se respeitar."
                Exit Select
            Case 457
                Quote = "Não tenha medo de recomeçar. Dessa vez, você não está começando do zero — está começando com experiência."
                Exit Select
            Case 458
                Quote = "Às vezes, dizer ""não"" é autocuidado."
                Exit Select
            Case 459
                Quote = "O que é para você não passa por você."
                Exit Select
            Case 460
                Quote = "Você é mais forte do que pensa e mais capaz do que imagina."
                Exit Select
            Case 461
                Quote = "Não confie palavras bonitas. Confie consistência."
                Exit Select
            Case 462
                Quote = "Crescer dói. Mudar dói. Mas não mudar dói ainda mais."
                Exit Select
            Case 463
                Quote = "A vida é curta demais para correr atrás de quem não corre por você."
                Exit Select
            Case 464
                Quote = "Algumas pessoas não são ruins; só não são para você."
                Exit Select
            Case 465
                Quote = "Não é egoísmo escolher você."
                Exit Select
            Case 466
                Quote = "Você não precisa se justificar para se proteger."
                Exit Select
            Case 467
                Quote = "Não se apegue ao que te custa paz."
                Exit Select
            Case 468
                Quote = "O tempo revela intenções."
                Exit Select
            Case 469
                Quote = "Siga em frente. O que é seu te encontra."
                Exit Select
            Case 470
                Quote = "A maior prova de amor-próprio é ir embora do que te diminui."
                Exit Select
            Case 471
                Quote = "Você não pode salvar alguém que não quer ser salvo."
                Exit Select
            Case 472
                Quote = "Uma mente em paz vale mais do que qualquer coisa."
                Exit Select
            Case 473
                Quote = "Não espere que as pessoas sejam como você."
                Exit Select
            Case 474
                Quote = "Aprenda a dizer ""chega"" sem se sentir culpado."
                Exit Select
            Case 475
                Quote = "Às vezes, o fechamento que você precisa é aceitar e seguir."
                Exit Select
            Case 476
                Quote = "Cuide de você como você cuida dos outros."
                Exit Select
            Case 477
                Quote = "Quem te respeita não te coloca em dúvida."
                Exit Select
            Case 478
                Quote = "Você não precisa de permissão para se afastar."
                Exit Select
            Case 479
                Quote = "Não confunda dependência com amor."
                Exit Select
            Case 480
                Quote = "Faça o bem aos outros. Isso volta de formas inesperadas. - Karen Salmansohn"
                Exit Select
            Case 481
                Quote = "Mantenha distância de pessoas que nunca admitem que estão erradas e sempre tentam fazer você sentir que a culpa é sua."
                Exit Select
            Case 482
                Quote = "Se alguém bêbado te manda mensagem, valorize: mesmo mal conseguindo pensar direito, a pessoa pensou em você."
                Exit Select
            Case 483
                Quote = "O sol vê o que eu faço, mas a lua conhece todos os meus segredos. - J.M. Wonderland"
                Exit Select
            Case 484
                Quote = "Tenho mais conversas na minha cabeça do que na vida real."
                Exit Select
            Case 485
                Quote = "Se pensar demais queimasse calorias, eu já estaria morto."
                Exit Select
            Case 486
                Quote = "Às vezes você tem que fazer o papel de bobo para enganar o bobo que acha que está te enganando."
                Exit Select
            Case 487
                Quote = "Você vê as verdadeiras cores de alguém quando você deixa de ser útil para a vida dela."
                Exit Select
            Case 488
                Quote = "Não me impressiono com dinheiro, status social ou cargo. Eu me impressiono com a forma como alguém trata outros seres humanos."
                Exit Select
            Case 489
                Quote = "Pessoas falsas não me surpreendem mais; pessoas leais, sim."
                Exit Select
            Case 490
                Quote = "Crescer dói. Mudar dói. Mas nada dói tanto quanto ficar preso em um lugar onde você não pertence."
                Exit Select
            Case 491
                Quote = "Sempre admire pessoas que mudam para melhor."
                Exit Select
            Case 492
                Quote = "Onde há amor, há vida. - Mahatma Gandhi"
                Exit Select
            Case 493
                Quote = "Eu queria voltar no tempo. Não para mudar as coisas, mas para sentir algumas coisas duas vezes."
                Exit Select
            Case 494
                Quote = "Quando alguém faz você se sentir mal por se importar, essa pessoa não merece você."
                Exit Select
            Case 495
                Quote = "Posso ser teimoso, impaciente e, às vezes, inseguro. Eu cometo erros, saio do controle e posso ser difícil. Mas, apesar de tudo isso, eu sei que, no fundo, eu tenho um coração enorme."
                Exit Select
            Case 496
                Quote = "Me conte seus sonhos e eu te digo se são reais."
                Exit Select
            Case 497
                Quote = "Você não ganha pontos extras por fazer certas coisas até um prazo. Vá no seu ritmo. Não é uma corrida."
                Exit Select
            Case 498
                Quote = "Ações provam quem alguém é; palavras só provam quem ela finge ser."
                Exit Select
            Case 499
                Quote = "Ações provam quem alguém é; palavras só provam quem ela finge ser."
                Exit Select
            Case 500
                Quote = "Às vezes, tomar um café com seu melhor amigo é toda a terapia de que você precisa."
                Exit Select
            Case 501
                Quote = "Caráter é como você trata alguém que não pode fazer nada por você. - Johann Wolfgang von Goethe"
                Exit Select
            Case 502
                Quote = "Às vezes, a maior aventura é simplesmente uma conversa."
                Exit Select
            Case 503
                Quote = "Peço desculpas se você não gosta da minha honestidade dura. Mas eu também não gosto da sua falsidade açucarada."
                Exit Select
            Case 504
                Quote = "A forma como ele te trata é como ele se sente sobre você."
                Exit Select
            Case 505
                Quote = "Não arrume desculpas para pessoas ruins. Você não coloca flores num babaca e chama de vaso."
                Exit Select
            Case 506
                Quote = "Mas, meu bem… no fim, você precisa ser seu próprio herói, porque todo mundo está ocupado tentando se salvar."
                Exit Select
            Case 507
                Quote = "Amanhã é a primeira página em branco de um livro de 365 páginas. Escreva uma boa história. - Brad Paisley"
                Exit Select
            Case 508
                Quote = "Viaje e não conte a ninguém; viva uma história de amor de verdade e não conte a ninguém; seja feliz e não conte a ninguém — as pessoas estragam coisas bonitas. - Kahlil Gibran"
                Exit Select
            Case 509
                Quote = "Prove para você mesmo, não para os outros."
                Exit Select
            Case 510
                Quote = "Sem mais expectativas; vou seguir o fluxo e, o que acontecer, aconteceu."
                Exit Select
            Case 511
                Quote = "Queria que minha vida tivesse música de fundo para eu entender o que diabos está acontecendo."
                Exit Select
            Case 512
                Quote = """Por que você só tem tipo 5 amigos?"" Eu: ""qualidade, não quantidade."""
                Exit Select
            Case 513
                Quote = "Uma alma gêmea é alguém que aprecia o seu nível de esquisitice."
                Exit Select
            Case 514
                Quote = "Apaixonar-se é fácil. Fazer sexo é mais fácil. Mas trombar com alguém que acende a sua alma… isso é raro."
                Exit Select
            Case 515
                Quote = "Pare de pensar demais e de inventar problemas que não existem."
                Exit Select
            Case 516
                Quote = "Eu nem fico mais ""decepcionado""; eu só penso: ah, de novo? ok kkk"
                Exit Select
            Case 517
                Quote = "Às vezes, as lembranças felizes são as que mais doem."
                Exit Select
            Case 518
                Quote = "Um capítulo ruim não significa que sua história acabou."
                Exit Select
            Case 519
                Quote = "Ser bom demais às vezes pode ser perigoso."
                Exit Select
            Case 520
                Quote = "As pessoas vão tentar te dizer quem você é. Não acredite nisso."
                Exit Select
            Case 521
                Quote = "Pare de correr atrás de quem não quer. A pessoa certa não foge. - Alfa"
                Exit Select
            Case 522
                Quote = "Os maiores presentes que você pode dar a alguém são seu tempo, seu amor e sua atenção. - Joel Osteen."
                Exit Select
            Case 523
                Quote = "Nunca zombe de alguém que fala inglês ''quebrado''. Isso significa que essa pessoa sabe outra língua. - H. Jackson Brown"
                Exit Select
            Case 524
                Quote = "O problema é que você acha que tem tempo."
                Exit Select
            Case 525
                Quote = "Estamos todos no mesmo jogo, só em níveis diferentes; passando pelo mesmo inferno, só com demônios diferentes."
                Exit Select
            Case 526
                Quote = "Sou a pessoa grosseira mais gentil que existe."
                Exit Select
            Case 527
                Quote = "Existe uma mensagem no jeito como alguém te trata… é só ouvir. - r.h. Sin"
                Exit Select
            Case 528
                Quote = "Valorize aquele amigo rude/sem papas na língua… ele sempre é o mais realista."
                Exit Select
            Case 529
                Quote = "Tudo bem se você não gosta de mim; nem todo mundo tem bom gosto."
                Exit Select
            Case 530
                Quote = "Não mate as pessoas com gentileza, porque nem todo mundo merece sua gentileza. Mate com silêncio, porque nem todo mundo merece sua atenção."
                Exit Select
            Case 531
                Quote = "A vida fica muito mais simples quando você para de se explicar para as pessoas e só faz o que funciona para você."
                Exit Select
            Case 532
                Quote = "Faça a coisa certa, mesmo quando ninguém estiver olhando. Isso se chama integridade."
                Exit Select
            Case 533
                Quote = "Para todas as garotas que já não acreditam em contos de fadas ou finais felizes: você é a autora desta história. Erga a cabeça e endireite sua coroa; você é a rainha deste reino e só você sabe como governá-lo. - B. Devine"
                Exit Select
            Case 534
                Quote = "Lembre-se de que, às vezes, não conseguir o que você quer é uma maravilhosa sorte. - Dalai Lama"
                Exit Select
            Case 535
                Quote = "Três fases da vida: 1. Nascimento 2. Que diabos é isso 3. Morte"
                Exit Select
            Case 536
                Quote = "Minha meta em 2017 é cumprir as metas que eu defini em 2016, que eu deveria ter cumprido em 2015, porque eu fiz uma promessa em 2014, que eu planejei em 2013."
                Exit Select
            Case 537
                Quote = "Eu sou forte, mas estou cansado."
                Exit Select
            Case 538
                Quote = """Um ano muda você demais."""
                Exit Select
            Case 539
                Quote = "Se você não gosta de mim, por favor, não finja que gosta. Nunca."
                Exit Select
            Case 540
                Quote = "Quando você descobre o gosto do respeito, você sempre vai escolher isso em vez de atenção. - Pink"
                Exit Select
            Case 541
                Quote = "E, no fim, tudo o que aprendi foi a ser forte sozinho. - Highway Heart"
                Exit Select
            Case 542
                Quote = "Pessoas que te mostram músicas novas são importantes."
                Exit Select
            Case 543
                Quote = """Qual foi a coisa mais importante que você fez este ano?"" ""SOBREVIVI"""
                Exit Select
            Case 544
                Quote = "Eu me importo. Eu sempre me importo. Esse é o meu problema."
                Exit Select
            Case 545
                Quote = "O tamanho do seu público não importa. Continue com o bom trabalho."
                Exit Select
            Case 546
                Quote = "e às vezes, só às vezes, quando as pessoas dizem 'pra sempre'… elas falam sério."
                Exit Select
            Case 547
                Quote = "Eu ajo diferente com certas pessoas. Não é porque sou falso. É porque eu tenho uma zona de conforto diferente com algumas pessoas."
                Exit Select
            Case 548
                Quote = "Cresça com o que você passa."
                Exit Select
            Case 549
                Quote = "Você precisa morrer algumas vezes antes de conseguir viver de verdade. - Charles Bukowski"
                Exit Select
            Case 550
                Quote = "Não deixe as pessoas saberem demais sobre você."
                Exit Select
            Case 551
                Quote = "Fique tão ocupado melhorando a si mesmo que não tenha tempo para criticar os outros."
                Exit Select
            Case 552
                Quote = "Então me diga: para onde eu devo ir? Para a esquerda, onde nada é certo? Ou para a direita, onde nada sobra?"
                Exit Select
            Case 553
                Quote = "SE VOCÊ QUER SER PODEROSO, EDUQUE-SE."
                Exit Select
            Case 554
                Quote = "Conheça o seu valor… e depois acrescente imposto."
                Exit Select
            Case 555
                Quote = "Algumas pessoas são a versão humana da enxaqueca."
                Exit Select
            Case 556
                Quote = "Eu definitivamente não sou a mesma pessoa que eu era quando este ano começou."
                Exit Select
            Case 557
                Quote = "Decida o que você quer. Anote isso. Faça um plano de verdade. E… trabalhe nele. Todo. Santo. Dia."
                Exit Select
            Case 558
                Quote = "Quanto mais eu envelheço, mais eu entendo que está tudo bem viver uma vida que os outros não entendem."
                Exit Select
            Case 559
                Quote = "Pecadores julgando pecadores por pecarem de um jeito diferente."
                Exit Select
            Case 560
                Quote = "Se você gosta de alguém, liberte. Se voltar, é porque ninguém gostou. Libere de novo."
                Exit Select
            Case 561
                Quote = "Por melhor que seu coração seja, uma hora você tem que começar a tratar as pessoas do jeito que elas te tratam…"
                Exit Select
            Case 562
                Quote = "A vida me ensinou que você não controla a lealdade de ninguém. Por melhor que você seja com alguém, isso não significa que vão te tratar igual. - Trent Shelton"
                Exit Select
            Case 563
                Quote = "Eu acordei um dia e decidi que não queria mais me sentir assim… então eu mudei."
                Exit Select
            Case 564
                Quote = "Uma rainha sempre transforma dor em poder."
                Exit Select
            Case 565
                Quote = "Menos amigos, menos enrolação. Mantenha seu círculo pequeno."
                Exit Select
            Case 566
                Quote = "Quanto mais gentil você é, mais fácil é te ferirem. Então… seja um babaca."
                Exit Select
            Case 567
                Quote = "Eu te dei $10, ele te deu $20. Você achou que ele era melhor só porque deu mais. Mas ele tinha $200, e eu só tinha $10."
                Exit Select
            Case 568
                Quote = "Minha coisa favorita é quando as pessoas lembram de pequenas coisas que eu contei. Tipo, sério? Você realmente me ouviu? Obrigado."
                Exit Select
            Case 569
                Quote = "Seja uma boa pessoa, mas não perca tempo tentando provar isso."
                Exit Select
            Case 570
                Quote = "Eu perdoo as pessoas esquecendo delas."
                Exit Select
            Case 571
                Quote = "Quem sou eu para julgar outra pessoa, se eu caminho imperfeitamente?"
                Exit Select
            Case 572
                Quote = "@melhor amigo(a), eu não te digo todo dia o quanto sou grato por você. Mas saiba: lá no fundo, eu sou muito abençoado por ter você na minha vida e por dividir tantas memórias com você."
                Exit Select
            Case 573
                Quote = "Você merece alguém que tenha pavor de te perder. - r.h. Sin"
                Exit Select
            Case 574
                Quote = "A vida é a prova mais difícil. Muita gente reprova porque tenta copiar os outros, sem perceber que cada um tem uma prova diferente."
                Exit Select
            Case 575
                Quote = "O que vem é melhor do que o que já foi."
                Exit Select
            Case 576
                Quote = "Trabalhar duro por algo que não nos importa se chama estresse; trabalhar duro por algo que amamos se chama paixão. - Simon Sinek"
                Exit Select
            Case 577
                Quote = "No fim, a gente só quer alguém que nos escolha… acima de todo mundo, em qualquer circunstância."
                Exit Select
            Case 578
                Quote = "Nunca tenha medo de dizer o que você realmente sente."
                Exit Select
            Case 579
                Quote = "Qual a diferença entre 'gosto de você' e 'eu te amo'? Buda respondeu lindamente: quando você gosta de uma flor, você a arranca. Mas quando você ama uma flor, você a rega todos os dias. Quem entende isso, entende a vida..."
                Exit Select
            Case 580
                Quote = "Perguntei ao meu coração por que não consigo dormir à noite. Ele respondeu: 'porque você dormiu à tarde — não finja que está apaixonado(a).'"
                Exit Select
            Case 581
                Quote = "As pessoas vivem me dizendo que a pessoa certa vai aparecer. Mas acho que a minha foi atropelada por um ônibus ou algo assim."
                Exit Select
            Case 582
                Quote = "Um pássaro sentado numa árvore nunca tem medo de o galho quebrar, porque a confiança dele não está no galho, e sim nas próprias asas. Acredite sempre em você."
                Exit Select
            Case 583
                Quote = "Antes de se diagnosticar com depressão ou baixa autoestima, primeiro confirme se você não está, na verdade, cercado de idiotas. - Sigmund Freud"
                Exit Select
            Case 584
                Quote = "Eu já vi almas feias escondidas atrás de rostos bonitos; então não me diga que beleza é só física."
                Exit Select
            Case 585
                Quote = "Não importa o quão bom você seja, sempre dá para te substituir."
                Exit Select
            Case 586
                Quote = "Aprenda a: se divertir sem beber; conversar sem celular; sonhar sem drogas; sorrir sem selfies; amar sem condições."
                Exit Select
            Case 587
                Quote = "Em breve, quando tudo estiver bem, você vai olhar para trás e ficar muito feliz por nunca ter desistido."
                Exit Select
            Case 588
                Quote = "Coisas que ficam sem ser ditas ficam com a gente para sempre."
                Exit Select
            Case 589
                Quote = "Você chega em casa, faz um chá, senta na poltrona e, ao redor, só silêncio. Cada um decide se isso é solidão… ou liberdade."
                Exit Select
            Case 590
                Quote = "Notas não medem inteligência e idade não define maturidade."
                Exit Select
            Case 591
                Quote = "Às vezes não existe próxima vez, nem segunda chance, nem tempo. Às vezes é agora ou nunca."
                Exit Select
            Case 592
                Quote = "A maior lição que aprendi este ano é que ninguém é realmente seu amigo ou te ama de verdade até ver toda a sua escuridão… e ficar."
                Exit Select
            Case 593
                Quote = "Eduque-se. Quando surgir uma dúvida sobre um tema, pesquise. Assista filmes e documentários. Quando algo te interessar, leia sobre isso. Leia, leia, leia. Estude, aprenda, estimule seu cérebro. Não dependa só da escola: eduque essa mente linda que você tem."
                Exit Select
            Case 594
                Quote = "Não é grosseria, é só ser real."
                Exit Select
            Case 595
                Quote = "O outono mostra como é bonito deixar as coisas irem."
                Exit Select
            Case 596
                Quote = "Quando a confiança é quebrada, 'desculpa' não significa nada."
                Exit Select
            Case 597
                Quote = "Você não está velho demais e não é tarde demais. - R. M. Rilke"
                Exit Select
            Case 598
                Quote = "Seja uma versão melhor de você… por você."
                Exit Select
            Case 599
                Quote = "A arte de saber é saber o que ignorar. - Rumi"
                Exit Select
            Case 600
                Quote = "Precisamos nos aventurar para saber onde realmente pertencemos."
                Exit Select
            Case 601
                Quote = "Seu coração sabe coisas que sua mente não consegue explicar."
                Exit Select
            Case 602
                Quote = "Quando foi a última vez que você fez algo pela primeira vez?"
                Exit Select
            Case 603
                Quote = "Aprendi mais com a dor do que eu jamais poderia aprender com o prazer."
                Exit Select
            Case 604
                Quote = "Você nunca vai ser feliz se estiver sempre com medo de soltar o que é confortável, familiar. Às vezes, são justamente essas coisas que nos machucam."
                Exit Select
            Case 605
                Quote = "Não eduque seus filhos para serem ricos. Eduque-os para serem felizes, assim eles saberão o valor das coisas, não o preço."
                Exit Select
            Case 606
                Quote = "É melhor olhar para trás e dizer: ""Não acredito que eu fiz isso."" do que olhar para trás e dizer: ""Eu queria ter feito isso."""
                Exit Select
            Case 607
                Quote = "Eu te amo leva 3 segundos para dizer, 3 horas para explicar e uma vida inteira para provar."
                Exit Select
            Case 608
                Quote = "Sempre ajude alguém. Você pode ser a única pessoa que ajuda."
                Exit Select
            Case 609
                Quote = "Ver alguém lendo um livro que você ama é como ver um livro recomendando uma pessoa."
                Exit Select
            Case 610
                Quote = "Nunca deseje mais do que você trabalha para conquistar."
                Exit Select
            Case 611
                Quote = "Não tenha medo de abrir mão do BOM para buscar o ÓTIMO. - John D. Rockefeller"
                Exit Select
            Case 612
                Quote = "Tenho orgulho do meu coração. Ele foi apunhalado, quebrado e despedaçado… e ainda assim continua batendo."
                Exit Select
            Case 613
                Quote = "Você nunca atravessa o oceano se não tiver coragem de perder de vista a margem."
                Exit Select
            Case 614
                Quote = "Melhor ter um inimigo que te dá um tapa na cara do que um amigo que te apunhala pelas costas."
                Exit Select
            Case 615
                Quote = "COLECIONE MOMENTOS, NÃO COISAS."
                Exit Select
            Case 616
                Quote = "Não deixe o medo de cair te impedir de voar."
                Exit Select
            Case 617
                Quote = "Seja teimoso com seus objetivos, mas flexível com seus métodos."
                Exit Select
            Case 618
                Quote = "Você vai querer desistir. Não desista."
                Exit Select
            Case 619
                Quote = "É só um dia ruim, não uma vida ruim."
                Exit Select
            Case 620
                Quote = "Por favor, não espere que eu seja sempre bom, gentil e carinhoso. Há momentos em que eu vou ser frio, sem pensar e difícil de entender."
                Exit Select
            Case 621
                Quote = "Nunca desista. Coisas grandes levam tempo."
                Exit Select
            Case 622
                Quote = "Fale o que pensa, mesmo que sua voz trema. - Maggie Kuhn"
                Exit Select
            Case 623
                Quote = "Decidi ser feliz, porque isso faz bem à minha saúde."
                Exit Select
            Case 624
                Quote = "CORAGEM não é ausência de medo, mas o julgamento de que algo é mais importante do que o medo. - Ambrose Redmoon"
                Exit Select
            Case 625
                Quote = "Um coração partido é o que muda as pessoas."
                Exit Select
            Case 626
                Quote = "Como um dia pode ser bonito quando a bondade o toca."
                Exit Select
            Case 627
                Quote = "Por sua causa, eu rio um pouco mais alto, choro um pouco menos e sorrio muito mais."
                Exit Select
            Case 628
                Quote = "Você nunca sabe o quanto é forte até que ser forte seja a única escolha que você tem."
                Exit Select
            Case 629
                Quote = "Que se danem as pessoas que brincam com os sentimentos dos outros."
                Exit Select
            Case 630
                Quote = "Pergunte a si mesmo se o que você está fazendo hoje está te deixando mais perto de onde você quer estar amanhã."
                Exit Select
            Case 631
                Quote = "Quando a gente cresce, percebe que é menos importante ter muitos amigos e mais importante ter amigos de verdade."
                Exit Select
            Case 632
                Quote = "Eu nasci para cometer erros, não para fingir perfeição."
                Exit Select
            Case 633
                Quote = "Ela transformou seus 'não consigo' em 'eu consigo' e seus sonhos em planos."
                Exit Select
            Case 634
                Quote = "A felicidade pode ser encontrada nos tempos mais sombrios, se a gente apenas se lembrar de acender a luz."
                Exit Select
            Case 635
                Quote = "Não apresse nada. Quando for a hora certa, vai acontecer."
                Exit Select
            Case 636
                Quote = "Seja sempre grato. A vida poderia ser pior."
                Exit Select
            Case 637
                Quote = "O custo de não seguir seu coração é passar o resto da vida desejando ter seguido."
                Exit Select
            Case 638
                Quote = "O que vem fácil não dura. O que dura não vem fácil."
                Exit Select
            Case 639
                Quote = "EU POSSO E EU VOU. Me observe."
                Exit Select
            Case 640
                Quote = "Nunca deixe seu medo decidir seu futuro."
                Exit Select
            Case 641
                Quote = "Seja uma voz, não um eco."
                Exit Select
            Case 642
                Quote = "As pessoas boas são as que passam pelo pior tipo de merda."
                Exit Select
            Case 643
                Quote = "Você está realmente feliz ou só realmente confortável?"
                Exit Select
            Case 644
                Quote = "Você já sentiu que nem é amigo de alguns dos seus amigos?"
                Exit Select
            Case 645
                Quote = "Você só falha quando para de tentar."
                Exit Select
            Case 646
                Quote = "Às vezes alguém diz uma coisa bem pequena e ela encaixa exatamente naquele espaço vazio do seu coração."
                Exit Select
            Case 647
                Quote = "Eu não quero a atenção do mundo. A sua já é suficiente."
                Exit Select
            Case 648
                Quote = "Eu encontro pedaços de você em cada música que eu escuto."
                Exit Select
            Case 649
                Quote = "Ela acreditou que podia, então ela fez."
                Exit Select
            Case 650
                Quote = "Algumas pessoas fazem sua risada ficar um pouco mais alta, seu sorriso um pouco mais brilhante e sua vida um pouco melhor."
                Exit Select
            Case 651
                Quote = "Nossas conversas pequenas e bobas significam mais para mim do que você imagina."
                Exit Select
            Case 652
                Quote = "Que você sempre faça aquilo que tem medo de fazer."
                Exit Select
            Case 653
                Quote = "Rir com frequência e muito; conquistar o respeito de pessoas inteligentes e o carinho das crianças; ganhar a apreciação de críticos honestos e suportar a traição de falsos amigos; apreciar a beleza, ver o melhor nos outros; deixar o mundo um pouco melhor, seja por uma criança saudável, um pedaço de jardim ou uma condição social recuperada; saber que pelo menos uma vida respirou mais leve porque você viveu. Isso é ter sucesso. - Ralph Waldo Emerson"
                Exit Select
            Case 654
                Quote = "Nem todo dia é bom, mas há algo de bom em todo dia."
                Exit Select
            Case 655
                Quote = "No fim, tudo vai ficar bem. Se ainda não está bem, então ainda não é o fim. - John Lennon"
                Exit Select
            Case 656
                Quote = "O para-brisa da vida é grande e o retrovisor é pequeno. Usar o retrovisor de vez em quando ajuda você a não se envolver em outro acidente, mas ficar focado nele faz você bater no que estiver à sua frente."
                Exit Select
            Case 657
                Quote = "Talento é um interesse cultivado. Em outras palavras, tudo o que você estiver disposto a praticar, você consegue fazer."
                Exit Select
            Case 658
                Quote = "Não importa o quão devagar você vá: você ainda está ultrapassando todo mundo que ficou no sofá."
                Exit Select
            Case 659
                Quote = "Ideias sem ação nunca ficam maiores do que as células do cérebro que elas ocupam."
                Exit Select
            Case 660
                Quote = "Saber o que você quer, entender por que está fazendo, dedicar cada fôlego do seu corpo para alcançar… Se você sente que tem algo a oferecer, se sente que seu talento vale a pena ser desenvolvido e cuidado, então não existe nada que você não possa alcançar. - Kevin Spacey"
                Exit Select
            Case 661
                Quote = "Escuta: eu queria poder dizer que melhora. Mas não melhora. Você é que melhora. - Joan Rivers"
                Exit Select
            Case 662
                Quote = "Trabalhe até que você não precise mais se apresentar."
                Exit Select
            Case 663
                Quote = "Ninguém está realmente contando quantas vezes você vacila… então relaxa a porra toda."
                Exit Select
            Case 664
                Quote = "Não existe tempo sobrando, não existe tempo livre, não existe tempo morto. Tudo o que você tem é tempo de vida. Vá. - Henry Rollins"
                Exit Select
            Case 665
                Quote = "Um hábito não pode ser jogado pela janela; ele precisa ser convencido a descer as escadas, degrau por degrau. - Mark Twain"
                Exit Select
            Case 666
                Quote = "Às vezes você nunca vai saber o valor de um momento até que ele vire uma memória. - Dr. Seuss"
                Exit Select
            Case 667
                Quote = "Quanto menos você reage a pessoas negativas, mais tranquila sua vida fica."
                Exit Select
            Case 668
                Quote = "Só porque você não se parece com o que alguém acha atraente, não significa que você não seja atraente. Flores são bonitas, mas luzes de Natal também, e não se parecem em nada."
                Exit Select
            Case 669
                Quote = "1,01^365 = 37,8; 0,99^365 = 0,03"
                Exit Select
            Case 670
                Quote = "Se você deixar a percepção das pessoas sobre você ditar seu comportamento, você nunca vai crescer como pessoa."
                Exit Select
            Case 671
                Quote = "Que se dane a motivação. Ela é instável e pouco confiável — e não vale seu tempo. Melhor cultivar disciplina do que depender de motivação. Force-se a fazer as coisas. Force-se a levantar da cama e praticar. Force-se a trabalhar. Motivação é passageira e é fácil se apoiar nela porque não exige esforço concentrado para aparecer. A motivação vem até você; você não precisa correr atrás. Disciplina é confiável; motivação passa. A pergunta não é como se manter motivado. É como treinar para trabalhar mesmo sem motivação."
                Exit Select
            Case 672
                Quote = "Perdoe os outros não porque eles merecem perdão, mas porque você merece paz."
                Exit Select
            Case 673
                Quote = "A ideia de motivação é uma armadilha. Esqueça a motivação. Apenas faça. Exercite-se, emagreça, controle sua glicose, ou o que for. Faça sem motivação. E sabe o que acontece? Depois que você começa a fazer, é aí que a motivação aparece e fica mais fácil continuar. - John Maxwell"
                Exit Select
            Case 674
                Quote = "É melhor se cagar nas calças do que morrer de prisão de ventre. - Meu pai, sobre eu chamar uma garota pra sair quando eu era mais novo."
                Exit Select
            Case 675
                Quote = "Tentaram nos enterrar. Não sabiam que éramos sementes."
                Exit Select
            Case 676
                Quote = "Seu foco determina sua realidade."
                Exit Select
            Case 677
                Quote = "O pessimista olha para baixo e bate a cabeça. O otimista olha para cima e perde o equilíbrio. O realista olha para frente e ajusta o caminho. - Rei Ezequiel"
                Exit Select
            Case 678
                Quote = "Sucesso é aceitar a si mesmo, gostar do que faz e amar a forma como faz."
                Exit Select
            Case 679
                Quote = "Se você ficar cansado, aprenda a descansar — não a desistir."
                Exit Select
            Case 680
                Quote = "Sucesso não é propriedade, é aluguel. E o aluguel vence todo dia. - J.J. Watt"
                Exit Select
            Case 681
                Quote = "A vida é curta. Sorria enquanto ainda tem dentes."
                Exit Select
            Case 682
                Quote = "Quando você quer alguma coisa, o universo inteiro conspira para que você realize seu desejo. - Paulo Coelho"
                Exit Select
            Case 683
                Quote = "Ser feliz nunca sai de moda."
                Exit Select
            Case 684
                Quote = "A vida tem um jeito engraçado de ensinar lições."
                Exit Select
            Case 685
                Quote = "Faça sua vida valer a pena."
                Exit Select
            Case 686
                Quote = "Onde quer que a vida plante você, floresça com graça."
                Exit Select
            Case 687
                Quote = "Nem todo mundo que se afasta te perde."
                Exit Select
            Case 688
                Quote = "Não se limite. Muitas pessoas se limitam ao que acham que conseguem. Você pode ir tão longe quanto sua mente permitir. O que você acredita, você consegue. - Mary Kay Ash"
                Exit Select
            Case 689
                Quote = "Às vezes, a pessoa mais forte é a que sorri em silêncio."
                Exit Select
            Case 690
                Quote = "Nunca deixe que o medo de cair te impeça de voar."
                Exit Select
            Case 691
                Quote = "Ninguém se sente realmente autoconfiante lá no fundo, porque isso é uma ideia artificial. Na real, as pessoas não estão tão preocupadas com o que você faz ou diz. Você pode andar pelo mundo quase anonimamente. Não se sinta perseguido e examinado. Liberte-se da ideia de que estão te observando. - Russell Brand"
                Exit Select
            Case 692
                Quote = "Seja uma voz, não um eco."
                Exit Select
            Case 693
                Quote = "O que é para ser seu, encontrará você."
                Exit Select
            Case 694
                Quote = "Um dia, tudo vai fazer sentido."
                Exit Select
            Case 695
                Quote = "Acredite no seu potencial."
                Exit Select
            Case 696
                Quote = "Cresça no que você atravessa."
                Exit Select
            Case 697
                Quote = "Você merece o amor que fica."
                Exit Select
            Case 698
                Quote = "Uma coisa de cada vez."
                Exit Select
            Case 699
                Quote = "Às vezes, a melhor resposta é o silêncio."
                Exit Select
            Case 700
                Quote = "Seja gentil com você mesmo."
                Exit Select
            Case 701
                Quote = "Você não é difícil de amar. Você só se acostumou com pouco."
                Exit Select
            Case 702
                Quote = "Algumas pessoas se vão para abrir espaço para coisas melhores."
                Exit Select
            Case 703
                Quote = "Aprenda a ser sua própria paz."
                Exit Select
            Case 704
                Quote = "Sem pressão, não há crescimento."
                Exit Select
            Case 705
                Quote = "Foque em você. O resto vem."
                Exit Select
            Case 706
                Quote = "Não confunda familiaridade com amor."
                Exit Select
            Case 707
                Quote = "O que é seu, não passa."
                Exit Select
            Case 708
                Quote = "Um tubarão bebê ainda é um tubarão."
                Exit Select
            Case 709
                Quote = "Seu valor não diminui porque alguém não consegue enxergar o seu valor."
                Exit Select
            Case 710
                Quote = "A verdade é que todo mundo vai te machucar. Você só precisa achar aqueles por quem vale a pena suportar. - Bob Marley"
                Exit Select
            Case 711
                Quote = "Só na escuridão você consegue ver as estrelas. - Martin Luther King"
                Exit Select
            Case 712
                Quote = "Nunca conte seus problemas a ninguém… 20% não ligam e os outros 80% ficam felizes por você tê-los. - Lou Holtz"
                Exit Select
            Case 713
                Quote = "Visão sem execução é só alucinação. - Henry Ford"
                Exit Select
            Case 714
                Quote = "Algumas pessoas sentem a chuva. Outras só se molham. - Bob Marley"
                Exit Select
            Case 715
                Quote = "Nem tudo que pode ser contado conta, e nem tudo que conta pode ser contado. - Albert Einstein"
                Exit Select
            Case 716
                Quote = "Um homem superior é modesto na fala, mas excede nas ações. - Confúcio"
                Exit Select
            Case 717
                Quote = "O que você pensa é o que você se torna. - Muhammad Ali"
                Exit Select
            Case 718
                Quote = "Só eu posso mudar minha vida. Ninguém pode fazer isso por mim. - Carol Burnett"
                Exit Select
            Case 719
                Quote = "Os homens, para ganhar a vida, esquecem de viver. - Margaret Fuller"
                Exit Select
            Case 720
                Quote = "Educação é a arma mais poderosa que você pode usar para mudar o mundo. - Nelson Mandela"
                Exit Select
            Case 721
                Quote = "Nada é impossível. A própria palavra diz: ""sou possível""! - Audrey Hepburn"
                Exit Select
            Case 722
                Quote = "Não leve a vida tão a sério. Você não vai sair vivo dela. - Elbert Hubbard"
                Exit Select
            Case 723
                Quote = "Qualidade não é um ato, é um hábito. - Aristóteles"
                Exit Select
            Case 724
                Quote = "Nossa maior fraqueza está em desistir. O jeito mais certo de vencer é tentar só mais uma vez. - Thomas A. Edison"
                Exit Select
            Case 725
                Quote = "Sempre parece impossível até ser feito. - Nelson Mandela"
                Exit Select
            Case 726
                Quote = "Aceite os desafios para sentir a emoção da vitória. - George S. Patton"
                Exit Select
            Case 727
                Quote = "Definir metas é o primeiro passo para transformar o invisível em visível. - Tony Robbins"
                Exit Select
            Case 728
                Quote = "O fracasso nunca vai me alcançar se a minha determinação de vencer for forte o bastante. - Og Mandino"
                Exit Select
            Case 729
                Quote = "Um homem criativo é motivado pelo desejo de realizar, não pelo desejo de superar os outros. - Ayn Rand"
                Exit Select
            Case 730
                Quote = "Persistência não é uma corrida longa; são várias corridas curtas, uma depois da outra. - Walter Elliot"
                Exit Select
            Case 731
                Quote = "Não fique olhando o relógio; faça o que ele faz. Continue. - Sam Levenson"
                Exit Select
            Case 732
                Quote = "Você nunca é velho demais para definir outra meta ou sonhar um novo sonho."
                Exit Select
            Case 733
                Quote = "Um bom plano executado com força agora é melhor do que um plano perfeito executado semana que vem. - George S. Patton"
                Exit Select
            Case 734
                Quote = "Miramos acima do alvo para acertar o alvo. - Ralph Waldo Emerson"
                Exit Select
            Case 735
                Quote = "As coisas não acontecem. As coisas são feitas acontecer. - John F. Kennedy"
                Exit Select
            Case 736
                Quote = "O objetivo final do ego não é ver algo, mas ser algo. - Muhammad Iqbal"
                Exit Select
            Case 737
                Quote = "Faça algo maravilhoso; as pessoas podem imitar. - Albert Schweitzer"
                Exit Select
            Case 738
                Quote = "Mereça o seu sonho. - Octavio Paz"
                Exit Select
            Case 739
                Quote = "Quando é preciso, é possível. - Charlotte Whitton"
                Exit Select
            Case 740
                Quote = "Você tem que fazer acontecer. - Denis Diderot"
                Exit Select
            Case 741
                Quote = "O quanto você quer isso?"
                Exit Select
            Case 742
                Quote = "Muitos são chamados, mas poucos se levantam. - Oliver Herford"
                Exit Select
            Case 743
                Quote = "Quem busca, encontra. - Sófocles"
                Exit Select
            Case 744
                Quote = "Não há nada lá no fundo dentro de nós além do que nós mesmos colocamos ali. - Richard Rorty"
                Exit Select
            Case 745
                Quote = "Perdedores desistem quando estão cansados. Vencedores desistem quando venceram."
                Exit Select
            Case 746
                Quote = "Erros são prova de que você está tentando."
                Exit Select
            Case 747
                Quote = "Um dia você vai agradecer por não ter desistido."
                Exit Select
            Case 748
                Quote = "Não deixe a opinião de alguém virar a sua verdade."
                Exit Select
            Case 749
                Quote = "Você está mais perto do que pensa."
                Exit Select
            Case 750
                Quote = "Seja constante."
                Exit Select
            Case 751
                Quote = "Não espere a motivação. Construa disciplina."
                Exit Select
            Case 752
                Quote = "Faça o que precisa ser feito."
                Exit Select
            Case 753
                Quote = "Foco cria resultado."
                Exit Select
            Case 754
                Quote = "Comece antes de estar pronto."
                Exit Select
            Case 755
                Quote = "O que você repete vira você."
                Exit Select
            Case 756
                Quote = "Seja o tipo de pessoa que você admira."
                Exit Select
            Case 757
                Quote = "Resultados não mentem."
                Exit Select
            Case 758
                Quote = "Paciência também é ação."
                Exit Select
            Case 759
                Quote = "O melhor momento é agora."
                Exit Select
            Case 760
                Quote = "Consistência vence."
                Exit Select
            Case 761
                Quote = "Você consegue mais do que imagina."
                Exit Select
            Case 762
                Quote = "Não pare no meio do caminho."
                Exit Select
            Case 763
                Quote = "Todo dia é treino."
                Exit Select
            Case 764
                Quote = "Você não precisa de sorte; precisa de processo."
                Exit Select
            Case 765
                Quote = "Seu futuro agradece seu esforço."
                Exit Select
            Case 766
                Quote = "É melhor ser tentado do que se arrepender."
                Exit Select
            Case 767
                Quote = "Pare de adiar sua vida."
                Exit Select
            Case 768
                Quote = "Disciplina é autocuidado."
                Exit Select
            Case 769
                Quote = "A cada dia, um pouco melhor."
                Exit Select
            Case 770
                Quote = "O caminho se abre andando."
                Exit Select
            Case 771
                Quote = "Tudo começa com uma decisão."
                Exit Select
            Case 772
                Quote = "Trabalhe em silêncio."
                Exit Select
            Case 773
                Quote = "Escolhas pequenas, resultados grandes."
                Exit Select
            Case 774
                Quote = "O progresso é a prova."
                Exit Select
            Case 775
                Quote = "Eu vou. Fim de papo."
                Exit Select
            Case 776
                Quote = "Coloque coração, mente e alma até nos menores atos. Esse é o segredo do sucesso. - Swami Sivananda"
                Exit Select
            Case 777
                Quote = "Bom, melhor, o melhor. Nunca se contente. Até o bom virar melhor e o melhor virar o melhor de verdade. - São Jerônimo"
                Exit Select
            Case 778
                Quote = "Pra que trair? Se não está feliz, vá embora!"
                Exit Select
            Case 779
                Quote = "Algumas caminhadas você tem que fazer sozinho."
                Exit Select
            Case 780
                Quote = "Nem todo mundo que você perde é uma perda."
                Exit Select
            Case 781
                Quote = "Se você perdeu alguém, mas se encontrou, você ganhou."
                Exit Select
            Case 782
                Quote = "As pessoas me perguntam por que é tão difícil confiar. Eu pergunto por que é tão difícil cumprir uma promessa."
                Exit Select
            Case 783
                Quote = "Ter uma alma gêmea nem sempre é sobre amor; você também pode encontrar uma em um amigo."
                Exit Select
            Case 784
                Quote = "Existem momentos que você acha que não vai sobreviver. E então você sobrevive. - David Levithan"
                Exit Select
            Case 785
                Quote = "Me julgue pelas pessoas que eu evito."
                Exit Select
            Case 786
                Quote = "Quando o seu passado ligar… não atenda. Ele não tem nada novo para dizer."
                Exit Select
            Case 787
                Quote = "Nem todo mundo foi feito para o seu futuro. Algumas pessoas só passam para te ensinar lições na vida."
                Exit Select
            Case 788
                Quote = "Se meus olhos mostrassem minha alma, todo mundo choraria ao me ver sorrir. - Kurt Cobain"
                Exit Select
            Case 789
                Quote = "Se você ainda fala disso, você ainda se importa com isso."
                Exit Select
            Case 790
                Quote = "E eles viveram felizes para sempre. Separadamente."
                Exit Select
            Case 791
                Quote = "Eu me afasto das pessoas por um motivo."
                Exit Select
            Case 792
                Quote = "Você não pode abrir a história da minha vida, pular para a página 738 e achar que me conhece."
                Exit Select
            Case 793
                Quote = "Algumas pessoas não mudam. Só acham novas maneiras de mentir."
                Exit Select
            Case 794
                Quote = "A dor mais profunda que já senti foi negar meus próprios sentimentos para deixar todo mundo confortável. - Nicole Lyons"
                Exit Select
            Case 795
                Quote = "Se você tem coragem de atravessar uma noite solitária com nada além dos seus pensamentos autodestrutivos te fazendo companhia, querido(a), você tem coragem de atravessar qualquer coisa."
                Exit Select
            Case 796
                Quote = "Lembre-se sempre: por trás de toda mulher forte e independente, existem dias em que ela esteve sozinha e sem amparo. Existem lições que a vida ensinou e histórias de batalhas e lutas que ela enfrentou sozinha. Sob o escudo de confiança e força, existe um mar de tristeza e dor que ela suportou. - Aarti Khurana"
                Exit Select
            Case 797
                Quote = "Eu não estou chateado(a) por você ter mentido; estou chateado(a) porque, a partir de agora, não posso mais acreditar em você. - Friedrich Nietzsche"
                Exit Select
            Case 798
                Quote = "Disseram que eu era perigoso(a)… perguntei por quê. Disseram: ""Porque você não precisa de ninguém."" Foi aí que eu sorri."
                Exit Select
            Case 799
                Quote = "Delicadeza não é fraqueza. É preciso coragem para continuar sensível em um mundo tão cruel. - Beau Taplin."
                Exit Select
            Case 800
                Quote = "A distância não separa as pessoas… o silêncio separa…"
                Exit Select
            Case 801
                Quote = "Pessoas feridas são perigosas. Elas sabem como fazer o inferno parecer lar."
                Exit Select
            Case 802
                Quote = "Eu não tenho medo de lobisomens, vampiros ou hotéis mal-assombrados. Eu tenho medo do que seres humanos reais fazem com outros seres humanos reais. - Walter Jon Williams"
                Exit Select
            Case 803
                Quote = "3 coisas para manter em segredo: 1) Vida amorosa. 2) Renda. 3) Próximo passo."
                Exit Select
            Case 804
                Quote = "Eu sou próximo de pouquíssimas pessoas, mas essas poucas significam tudo para mim."
                Exit Select
            Case 805
                Quote = "O primeiro passo da mudança é perceber a sua própria merda."
                Exit Select
            Case 806
                Quote = "Se você vê algo bonito em alguém, diga. - Ruthie Lindsey"
                Exit Select
            Case 807
                Quote = "Você está sempre a uma decisão de distância de uma vida totalmente diferente."
                Exit Select
            Case 808
                Quote = "Se está te destruindo, então não é amor, meu bem."
                Exit Select
            Case 809
                Quote = "As pessoas mais bonitas que você vai conhecer nem sempre são as que chamam atenção primeiro. As mais bonitas são as que não dá para decifrar. As que você conversaria por horas e ainda teria um milhão de perguntas. Pessoas com mentes tão lindas e especiais que você não consegue evitar se apaixonar…"
                Exit Select
            Case 810
                Quote = "Tudo se resume à primeira pessoa para quem você quer contar uma boa notícia."
                Exit Select
            Case 811
                Quote = "Eu tenho uma habilidade especial de sentir demais quando não devia, e não sentir nada quando eu devia."
                Exit Select
            Case 812
                Quote = "As coisas acabam. As pessoas mudam. E sabe de uma coisa? A vida segue. - Elizabeth Scott"
                Exit Select
            Case 813
                Quote = "Se eu estiver errado, me ensine. Não me diminua."
                Exit Select
            Case 814
                Quote = "Quanto mais eu envelheço, mais eu valorizo ficar em casa sem fazer absolutamente nada."
                Exit Select
            Case 815
                Quote = "A pessoa que você será em 5 anos depende dos livros que você lê e das pessoas de que você se cerca hoje."
                Exit Select
            Case 816
                Quote = "As coisas que você esconde no coração… te corroem por dentro…"
                Exit Select
            Case 817
                Quote = "Só porque eu fico em silêncio não quer dizer que eu concordo com você. Eu só estou sendo educado. Ou então estou tão chocado com a sua estupidez que nem consigo responder."
                Exit Select
            Case 818
                Quote = "Eu nunca conheci uma pessoa forte com um passado fácil."
                Exit Select
            Case 819
                Quote = "Um coração que sempre entende também se cansa."
                Exit Select
            Case 820
                Quote = "Você vai me procurar em outra pessoa. Eu prometo."
                Exit Select
            Case 821
                Quote = "Nunca se esqueça das pessoas que tiram um tempo do dia delas para saber como você está."
                Exit Select
            Case 822
                Quote = "Acontece coisa ruim. Todo dia. Com todo mundo. A diferença está em como as pessoas lidam com isso."
                Exit Select
            Case 823
                Quote = "Ela está presa entre quem ela é, quem ela quer ser e quem ela deveria ser."
                Exit Select
            Case 824
                Quote = "A cura para qualquer coisa é água salgada: suor, lágrimas ou o mar. - Isak Dinesen"
                Exit Select
            Case 825
                Quote = "Eu sei que você não precisa de mim, mas eu preciso de você."
                Exit Select
            Case 826
                Quote = "O amor não é algo que você encontra. O amor é algo que te encontra."
                Exit Select
            Case 827
                Quote = "Eu tenho vontade de ser feliz."
                Exit Select
            Case 828
                Quote = "Se você não correr atrás do que quer, nunca vai ter."
                Exit Select
            Case 829
                Quote = "Você não pode voltar e mudar o começo, mas pode começar onde está e mudar o final."
                Exit Select
            Case 830
                Quote = "Pare de se preocupar com o que pode dar errado e comece a pensar no que pode dar certo."
                Exit Select
            Case 831
                Quote = "A vida é sobre criar a si mesmo."
                Exit Select
            Case 832
                Quote = "Não perca tempo com explicações. As pessoas só entendem o que querem entender."
                Exit Select
            Case 833
                Quote = "Não espere. A hora nunca vai ser perfeita."
                Exit Select
            Case 834
                Quote = "Às vezes, você precisa esquecer o que sente e lembrar do que merece."
                Exit Select
            Case 835
                Quote = "Seja uma prioridade para você."
                Exit Select
            Case 836
                Quote = "O que é para você, não te deixa."
                Exit Select
            Case 837
                Quote = "Paz é poder respirar sem culpa."
                Exit Select
            Case 838
                Quote = "Nem todo mundo vai entender seu caminho — e tudo bem."
                Exit Select
            Case 839
                Quote = "Fique perto de quem soma."
                Exit Select
            Case 840
                Quote = "Seu futuro depende do que você faz hoje."
                Exit Select
            Case 841
                Quote = "Escolha a si mesmo todos os dias."
                Exit Select
            Case 842
                Quote = "Não se apaixone por potencial; apaixone-se por atitudes."
                Exit Select
            Case 843
                Quote = "Um dia, você vai olhar para trás e ver o quanto foi forte."
                Exit Select
            Case 844
                Quote = "A gratidão muda o que você tem em suficiente."
                Exit Select
            Case 845
                Quote = "Você não precisa provar nada para ninguém."
                Exit Select
            Case 846
                Quote = "Seja o motivo do seu orgulho."
                Exit Select
            Case 847
                Quote = "Nem todo mundo merece acesso a você."
                Exit Select
            Case 848
                Quote = "Não se diminua para caber no mundo de alguém."
                Exit Select
            Case 849
                Quote = "Se não for paz, não é seu."
                Exit Select
            Case 850
                Quote = "A vida é dura, meu bem, mas você também é."
                Exit Select
            Case 851
                Quote = "Deus, por favor, me dê paciência; se me der força, eu vou acabar dando um soco na cara deles."
                Exit Select
            Case 852
                Quote = "A vida não é medida pelo número de respirações que você dá, mas pelos momentos que tiram o seu fôlego."
                Exit Select
            Case 853
                Quote = "Quando meu chefe perguntou quem é o burro, eu ou ele, eu disse: todo mundo sabe que ele não contrata gente burra."
                Exit Select
            Case 854
                Quote = "Todo mundo quer ir para o céu; mas ninguém quer morrer."
                Exit Select
            Case 855
                Quote = "Algumas pessoas entram na nossa vida e deixam pegadas no nosso coração. Outras entram e dá vontade de deixar pegadas na cara delas!"
                Exit Select
            Case 856
                Quote = "Quem disse que nada é impossível? Eu não faço nada há anos."
                Exit Select
            Case 857
                Quote = "Dizem que amor é mais importante que dinheiro… mas você já tentou pagar as contas com um abraço?"
                Exit Select
            Case 858
                Quote = "A estrada para o sucesso está sempre em obras. - Lily Tomlin"
                Exit Select
            Case 859
                Quote = "Um bom discurso deve ser como a saia de uma mulher: longo o bastante para cobrir o assunto e curto o bastante para criar interesse. - Winston Churchill"
                Exit Select
            Case 860
                Quote = "Nunca leve a vida tão a sério. Ninguém sai vivo dela mesmo."
                Exit Select
            Case 861
                Quote = "Sorria hoje… amanhã pode ser pior."
                Exit Select
            Case 862
                Quote = "O homem ideal não fuma, não bebe, não usa drogas, não xinga, não se irrita… e não existe."
                Exit Select
            Case 863
                Quote = "Relacionamentos hoje em dia começam apertando LIKE na foto dela."
                Exit Select
            Case 864
                Quote = "Às vezes, quando eu fecho os olhos, eu não consigo ver."
                Exit Select
            Case 865
                Quote = "Não chore porque acabou; sorria porque aconteceu. - Dr. Seuss"
                Exit Select
            Case 866
                Quote = "Você sabe que está apaixonado quando não consegue dormir porque a realidade finalmente é melhor do que seus sonhos. - Dr. Seuss"
                Exit Select
            Case 867
                Quote = "Ser você mesmo em um mundo que tenta o tempo todo te transformar em outra coisa é a maior conquista. - Ralph Waldo Emerson"
                Exit Select
            Case 868
                Quote = "Sempre que você se encontrar do lado da maioria, é hora de parar e refletir. - Mark Twain"
                Exit Select
            Case 869
                Quote = "Quando você está feliz, você curte a música. Quando está triste, você entende a letra."
                Exit Select
            Case 870
                Quote = "Eu nunca cometo o mesmo erro duas vezes. Eu cometo umas cinco ou seis, só para ter certeza."
                Exit Select
            Case 871
                Quote = "Só peixe morto vai com a correnteza."
                Exit Select
            Case 872
                Quote = "Eu odeio quando me mandam mensagem: ""Me liga"". Vou começar a ligar… e quando atenderem, eu vou dizer: ""Me manda mensagem"" e desligar."
                Exit Select
            Case 873
                Quote = "Minha vida parece uma prova para a qual eu não estudei."
                Exit Select
            Case 874
                Quote = """Seja forte"", eu sussurrei para o meu sinal de wi-fi."
                Exit Select
            Case 875
                Quote = "Felicidade é… não precisar colocar despertador para o dia seguinte."
                Exit Select
            Case 876
                Quote = "Não confunda silêncio com fraqueza. Gente inteligente não anuncia em voz alta os grandes movimentos."
                Exit Select
            Case 877
                Quote = "Minha alma gêmea deve estar por aí em algum lugar, empurrando uma porta de puxar… eu tenho certeza."
                Exit Select

            Case 880
                Quote = "A vida é curta. Sorria enquanto ainda tem dentes."
                Exit Select
            Case 881
                Quote = "Quando você está estressado, você come sorvete, bolo, chocolate e doces. Por quê? Porque ""stressed"" ao contrário é ""desserts""."
                Exit Select
            Case 882
                Quote = "A derrota não é amarga, a menos que você a engula. - Joe Clark"
                Exit Select
            Case 883
                Quote = "Fique com quem te ouviu quando você não disse uma palavra."
                Exit Select
            Case 884
                Quote = "Daqui a dez anos, garanta que você possa dizer que escolheu sua vida. Que você não se conformou com ela. - Mandy Hale"
                Exit Select
            Case 885
                Quote = "Antes eu tinha medo de ficar sozinho. Agora eu tenho medo de ter as pessoas erradas como companhia."
                Exit Select
            Case 886
                Quote = "Não tente entender tudo. Às vezes não é para ser entendido, apenas aceito."
                Exit Select
            Case 887
                Quote = "Todo mundo tem um capítulo que não lê em voz alta."
                Exit Select
            Case 888
                Quote = "É incrível como alguém pode virar um estranho tão rápido."
                Exit Select
            Case 889
                Quote = "Não é o tamanho da vida, e sim a profundidade."
                Exit Select
            Case 890
                Quote = "Lágrimas são palavras que o coração não consegue dizer."
                Exit Select
            Case 891
                Quote = "Mesmo eu vendo que ia acontecer, ainda dói."
                Exit Select
            Case 892
                Quote = "E se eu cair? Ah, meu bem… e se você voar?"
                Exit Select
            Case 893
                Quote = "Fique bem com o desconforto. Porque um dia, todo esse esforço vai valer a pena."
                Exit Select
            Case 894
                Quote = "Você não precisa ver a escada inteira, apenas dê o primeiro passo. - Martin Luther King"
                Exit Select
            Case 895
                Quote = "A maioria de vocês não quer sucesso tanto quanto quer dormir! - Eric Thomas"
                Exit Select
            Case 896
                Quote = "Gosto de pensar que um dia você vai ser um velho como eu, falando sem parar no ouvido de um jovem, explicando como você pegou o limão mais azedo que a vida ofereceu e transformou em algo que parecia limonada. - This Is Us"
                Exit Select
            Case 897
                Quote = "Eu não cheguei até aqui só para chegar até aqui."
                Exit Select
            Case 898
                Quote = "Não tenha medo de fracassar. Tenha medo de não tentar."
                Exit Select
            Case 899
                Quote = "EU NÃO ESTOU AQUI PARA SER MEDIANO. EU ESTOU AQUI PARA SER INCRÍVEL."
                Exit Select
            Case 900
                Quote = "Um dia, eu quero dizer de verdade: ""Eu consegui."""
                Exit Select
        End Select

        With appCell
            .Value = begValue & Quote
            .Font.Name = "Comic Sans MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub ClearPen_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearPen.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim sht As Excel.Worksheet = Nothing
        Dim shp As Excel.Shape = Nothing

        sht = app.ActiveWorkbook.ActiveSheet
        For Each shp In sht.Shapes
            If shp.Type = Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight And (shp.Name = "RedPen" Or shp.Name = "BluePen" Or shp.Name = "BlackPen") Then
                shp.Select()
                shp.Delete()
            End If
        Next
    End Sub

    Private Sub Direction1_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If

    End Sub

    Private Sub Direction2_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub Direction3_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction3.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub Direction4_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction4.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub Direction5_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction5.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub hyperlinkCell_Click(sender As Object, e As RibbonControlEventArgs) Handles hyperlinkCell.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        'Origin Sheet
        Dim strBeginningSheetName As String = app.Activesheet.name
        'Store the Original sheet & cell location
        Dim OriSheetCellAddress As String = appCell.address


        Dim location As Object
        location = app.inputbox("Select a cell", "Creating a HyperLink", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If



        'Get the Target's sheet name
        Dim strDestinationSheetName As String = location.Parent.Name


        'Add the target address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=appCell,
             Address:="", SubAddress:=strDestinationSheetName & "!" & location.address,
             TextToDisplay:=appCell.Value.ToString)
        appCell.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
        appCell.font.color = RGB(0, 0, 0)

        'Add the origin address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=location,
             Address:="", SubAddress:=strBeginningSheetName & "!" & OriSheetCellAddress,
             TextToDisplay:=location.value.ToString)
        location.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
        location.font.color = RGB(0, 0, 0)
    End Sub

    Private Sub Calendar_ToggleButton_Click(sender As Object, e As RibbonControlEventArgs) Handles Calendar_ToggleButton.Click
        Globals.ThisAddIn.TaskPane2.Visible =
    TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
    End Sub

    Private Sub LeftBrace_Click(sender As Object, e As RibbonControlEventArgs) Handles LeftBrace.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        shp = app.ActiveSheet.Shapes.Addshape _
(Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftBrace, r.Left - 17 + r.Width, r.Top, 17, r.Height)
        shp.Select()
        app.Selection.ShapeRange.Adjustments.Item(1) = 0.4
        shp.Line.ForeColor.RGB = RGB(0, 0, 0)

    End Sub

    Private Sub DownApply_Click(sender As Object, e As RibbonControlEventArgs) Handles DownApply.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim arrow As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)
        shp = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 6, r.Top, r.Left + 6, r.Height + r.Top)
        shp.Line.ForeColor.RGB = RGB(255, 0, 0)

        arrow = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 6, r.Height + r.Top, r.Left + 13, r.Top + r.Height - 10)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {shp.Name, arrow.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()

    End Sub

    Private Sub RightApply_Click(sender As Object, e As RibbonControlEventArgs) Handles RightApply.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim arrow As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)
        shp = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left, r.Top + r.Height - 4.5, r.Left + r.Width, r.Top + r.Height - 4.5)
        shp.Line.ForeColor.RGB = RGB(255, 0, 0)

        arrow = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Height + r.Top - 4.5, r.Left + r.Width - 10, r.Top + r.Height - 10)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {shp.Name, arrow.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub PMingLiU_Click(sender As Object, e As RibbonControlEventArgs) Handles PMingLiU.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "新細明體"
    End Sub

    Private Sub Arial_Click(sender As Object, e As RibbonControlEventArgs) Handles Arial.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "Arial"
    End Sub

    Private Sub BookAntiqua_Click(sender As Object, e As RibbonControlEventArgs) Handles BookAntiqua.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "Book Antiqua"
    End Sub

    Private Sub Calibri_Click(sender As Object, e As RibbonControlEventArgs) Handles Calibri.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "Calibri"
    End Sub

    Private Sub MultiIllu_Click(sender As Object, e As RibbonControlEventArgs) Handles MultiIllu.Click
        multi()
    End Sub
    Private Sub multi()
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim sht As Excel.Worksheet = app.ActiveWorkbook.ActiveSheet
        Dim arrow As Excel.Shape
        Dim connectorref As Excel.Shape = Nothing
        Dim r As Excel.Range
        Dim i As Integer = Nothing

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)
        shp = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width + 8, r.Top + 8.25, r.Left + r.Width + 8, r.Height + r.Top + 8)
        shp.Line.ForeColor.RGB = RGB(255, 0, 0)


        arrow = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width + 8, r.Height + r.Top + 8, r.Left + r.Width + 22, r.Top + r.Height + 8)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)
        arrow.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle

        Dim MidObj() As Object
        ReDim Preserve MidObj(2)
        MidObj(0) = shp.Name
        MidObj(1) = arrow.Name
        For i = 0 To (r.Height / 16.5 - 1)
            connectorref = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Top + 8.25 + i * 16.5, r.Left + r.Width + 8, r.Top + 8.25 + i * 16.5)
            connectorref.Line.ForeColor.RGB = RGB(255, 0, 0)
            'connectorref.Name = "connectorref" & i
            ReDim Preserve MidObj(i + 2)
            MidObj(i + 2) = connectorref.Name
        Next

        'Group shps all together
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub delta_Click(sender As Object, e As RibbonControlEventArgs) Handles delta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "δ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub epsilon_Click(sender As Object, e As RibbonControlEventArgs) Handles epsilon.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ε"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub zeta_Click(sender As Object, e As RibbonControlEventArgs) Handles zeta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ζ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub eta_Click(sender As Object, e As RibbonControlEventArgs) Handles eta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "η"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub theta_Click(sender As Object, e As RibbonControlEventArgs) Handles theta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "θ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub mu_Click(sender As Object, e As RibbonControlEventArgs) Handles mu.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "μ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub pi_Click(sender As Object, e As RibbonControlEventArgs) Handles pi.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "π"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub rho_Click(sender As Object, e As RibbonControlEventArgs) Handles rho.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ρ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub phi_Click(sender As Object, e As RibbonControlEventArgs) Handles phi.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "φ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub psi_Click(sender As Object, e As RibbonControlEventArgs) Handles psi.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ψ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub omega_Click(sender As Object, e As RibbonControlEventArgs) Handles omega.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ω"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub one_Click(sender As Object, e As RibbonControlEventArgs) Handles one.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅰ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub two_Click(sender As Object, e As RibbonControlEventArgs) Handles two.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅱ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub three_Click(sender As Object, e As RibbonControlEventArgs) Handles three.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅲ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub four_Click(sender As Object, e As RibbonControlEventArgs) Handles four.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅳ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub five_Click(sender As Object, e As RibbonControlEventArgs) Handles five.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅴ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub six_Click(sender As Object, e As RibbonControlEventArgs) Handles six.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅵ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub seven_Click(sender As Object, e As RibbonControlEventArgs) Handles seven.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅶ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub eight_Click(sender As Object, e As RibbonControlEventArgs) Handles eight.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅷ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub nine_Click(sender As Object, e As RibbonControlEventArgs) Handles nine.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅸ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub ten_Click(sender As Object, e As RibbonControlEventArgs) Handles ten.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅹ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub ColorYellow_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorYellow.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(255, 255, 0)
    End Sub

    Private Sub ColorAqua_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorAqua.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(0, 255, 255)
    End Sub

    Private Sub ColorLime_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorLime.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(211, 112, 112)
    End Sub

    Private Sub ColorSilver_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorSilver.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(255, 255, 196)
    End Sub

    Private Sub ColorFuchsia_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorFuchsia.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(196, 196, 255)
    End Sub

    Private Sub ColorWhite_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorWhite.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.pattern = Microsoft.Office.Core.XlConstants.xlNone
    End Sub



    Private Sub RD_Click(sender As Object, e As RibbonControlEventArgs) Handles RD.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "RD"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 10
            .font.bold = True
        End With
    End Sub

    Private Sub btnCB_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCB.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "CB"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 10
            .font.bold = True
        End With
    End Sub

    Private Sub TB_Click(sender As Object, e As RibbonControlEventArgs) Handles TB.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = appCell.Value

        With appCell
            .Value = begValue & "TB"
            .Font.Color = RGB(220, 65, 54)
            .Font.Name = "Times New Roman"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlCenter
        End With
    End Sub

    Private Sub FS_Click(sender As Object, e As RibbonControlEventArgs) Handles FS.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = appCell.Value

        With appCell
            .Value = begValue & "FS"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlCenter
        End With
    End Sub

    Private Sub GL_Click(sender As Object, e As RibbonControlEventArgs) Handles GL.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = appCell.Value

        With appCell
            .Value = begValue & "GL"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = Excel.Constants.xlLeft
            .VerticalAlignment = Excel.Constants.xlCenter
        End With
    End Sub

    Private Sub Tick2_Click(sender As Object, e As RibbonControlEventArgs) Handles Tick2.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim sel As Excel.Range = TryCast(app.Selection, Excel.Range)
        If sel Is Nothing Then Exit Sub

        Dim ws As Excel.Worksheet = TryCast(app.ActiveSheet, Excel.Worksheet)
        If ws Is Nothing Then Exit Sub

        'Imagem do próprio botão Tick2 (Ribbon1.resx)
        Dim img As System.Drawing.Image = Nothing
        Try
            Dim crm As New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
            img = TryCast(crm.GetObject("Tick2.Image"), System.Drawing.Image)
        Catch
            img = Nothing
        End Try
        If img Is Nothing Then Exit Sub

        Dim ratio As Double = CDbl(img.Width) / Math.Max(1.0R, CDbl(img.Height))

        'Alvos (seleção múltipla; mescladas viram topo-esquerda)
        Dim targetCells As New List(Of Excel.Range)()
        Dim keySet As New HashSet(Of String)(StringComparer.Ordinal)

        Try
            For Each area As Excel.Range In sel.Areas
                For Each cellObj As Object In area.Cells
                    Dim c As Excel.Range = TryCast(cellObj, Excel.Range)
                    If c Is Nothing Then Continue For
                    If CBool(c.MergeCells) Then c = c.MergeArea.Cells(1, 1)

                    Dim addr As String = c.Address(False, False)
                    Dim cellKey As String = "TM|Tick2|" & addr
                    If keySet.Add(cellKey) Then targetCells.Add(c)
                Next
            Next
        Catch
            Exit Sub
        End Try

        If targetCells.Count = 0 Then Exit Sub

        'Remover qualquer tick existente na célula (Tick, Tick2, X) – FS não é removido
        Try
            For Each c As Excel.Range In targetCells
                For i As Integer = ws.Shapes.Count To 1 Step -1
                    Try
                        Dim shp As Excel.Shape = ws.Shapes.Item(i)
                        If shp Is Nothing Then Continue For

                        Dim tag As String = ""
                        Dim nm As String = ""
                        Try : tag = shp.AlternativeText : Catch : End Try
                        Try : nm = shp.Name : Catch : End Try

                        'Apaga Tick, Tick2 ou X na célula
                        If ((Not String.IsNullOrEmpty(tag) AndAlso
                         (tag.StartsWith("TM|tick|") OrElse tag.StartsWith("TM|Tick2|") OrElse tag.StartsWith("TM|X|"))) _
                        OrElse (Not String.IsNullOrEmpty(nm) AndAlso
                                (nm.StartsWith("TM_tick_") OrElse nm.StartsWith("TM_Tick2_") OrElse nm.StartsWith("TM_X_")))) _
                        AndAlso shp.TopLeftCell.Address(False, False) = c.Address(False, False) Then
                            shp.Delete()
                        End If
                    Catch
                    End Try
                Next
            Next
        Catch
        End Try

        'Salvar PNG temporário
        Dim tmpPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                 "tickmark_tick2_" & Guid.NewGuid().ToString("N") & ".png")
        Try
            Try
                img.Save(tmpPath, System.Drawing.Imaging.ImageFormat.Png)
            Catch
                Using bmp As New System.Drawing.Bitmap(img)
                    bmp.Save(tmpPath, System.Drawing.Imaging.ImageFormat.Png)
                End Using
            End Try

            Dim prevScreenUpdating As Boolean = app.ScreenUpdating
            app.ScreenUpdating = False

            Try
                Const innerOffset As Double = 0.5R

                For Each c As Excel.Range In targetCells
                    Dim targetH As Double = Math.Max(8.0R, c.Height * 0.75R)
                    Dim targetW As Double = targetH * ratio
                    Dim maxW As Double = Math.Max(8.0R, c.Width * 0.45R)
                    If targetW > maxW Then
                        targetW = maxW
                        targetH = targetW / ratio
                    End If

                    Dim addr As String = c.Address(False, False)
                    Dim cellKey As String = "TM|Tick2|" & addr

                    Dim pic As Excel.Shape = ws.Shapes.AddPicture(
                    Filename:=tmpPath,
                    LinkToFile:=Microsoft.Office.Core.MsoTriState.msoFalse,
                    SaveWithDocument:=Microsoft.Office.Core.MsoTriState.msoCTrue,
                    Left:=CSng(c.Left),
                    Top:=CSng(c.Top),
                    Width:=CSng(targetW),
                    Height:=CSng(targetH)
                )

                    pic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                    pic.Placement = Excel.XlPlacement.xlMoveAndSize
                    pic.AlternativeText = cellKey
                    Try
                        pic.Name = "TM_Tick2_" & addr.Replace("$", "").Replace(":", "_") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6)
                    Catch
                    End Try

                    'ESQUERDA + centro vertical
                    pic.Left = CSng(c.Left + innerOffset)
                    pic.Top = CSng(c.Top + (c.Height - pic.Height) / 2.0R)
                Next

            Finally
                app.ScreenUpdating = prevScreenUpdating
            End Try

        Finally
            Try
                If System.IO.File.Exists(tmpPath) Then System.IO.File.Delete(tmpPath)
            Catch
            End Try
        End Try
        Try
            ' Pega a última célula da seleção
            Dim lastCell As Excel.Range = sel.Cells(sel.Cells.Count)
            ' Move o cursor para a célula imediatamente abaixo
            Dim nextCell As Excel.Range = lastCell.Offset(1, 0)
            nextCell.Select()
        Catch
            ' Ignora qualquer erro (por exemplo, se estiver na última linha)
        End Try
    End Sub

    Private Sub Tick3_Click(sender As Object, e As RibbonControlEventArgs) Handles Tick3.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim ws As Excel.Worksheet = TryCast(app.ActiveSheet, Excel.Worksheet)
        If ws Is Nothing Then Exit Sub

        Dim r As Excel.Range
        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                                    RowAbsolute:=False, ColumnAbsolute:=False)
        r = app.Range(selection)

        'marcação do Group9 (um id por clique)
        Dim groupId As String = Guid.NewGuid().ToString("N")

        Dim connector1 As Excel.Shape = ws.Shapes.AddConnector(
        Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
        r.Left + 5.8F, r.Top + 2.7F, r.Left + 8.8F, r.Top + 8.0F
    )
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.15
        TagGroup9Shape(connector1, groupId)

        Dim connector4 As Excel.Shape = ws.Shapes.AddConnector(
        Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
        r.Left + 4.8F, r.Top + 5.2F, r.Left + 7.8F, r.Top + 10.5F
    )
        connector4.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector4.Line.Weight = 1.15
        TagGroup9Shape(connector4, groupId)

        Dim connector2 As Excel.Shape = ws.Shapes.AddConnector(
        Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
        r.Left + 4.7F, r.Top + r.Height - 1.0F, r.Left + 7.7F, r.Top + 3.0F
    )
        connector2.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector2.Line.Weight = 1.15
        TagGroup9Shape(connector2, groupId)

        Dim connector3 As Excel.Shape = ws.Shapes.AddConnector(
        Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
        r.Left + 2.3F, r.Top + r.Height - 6.5F, r.Left + 4.8F, r.Top + r.Height - 1.0F
    )
        connector3.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector3.Line.Weight = 1.15
        TagGroup9Shape(connector3, groupId)

        'Agrupar (opcional)
        Try
            Dim midObj As Object() = New Object() {connector1.Name, connector2.Name, connector3.Name, connector4.Name}
            ws.Shapes.Range(midObj).Group()
            'Se quiser marcar o grupo também, dá pra tentar (não é obrigatório):
            'Dim grp As Excel.Shape = ws.Shapes.Range(midObj).Group()
            'TagGroup9Shape(grp, groupId)
        Catch
            'se falhar o agrupamento, os shapes já estão marcados individualmente
        End Try
    End Sub

    Private Sub NotMaterial_Click(sender As Object, e As RibbonControlEventArgs) Handles NotMaterial.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape = Nothing
        Dim shapeBeg As Excel.Shape = Nothing
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        shapeBeg = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, r.Left, r.Top, 50, r.Height)
        With shapeBeg.TextFrame2.TextRange.Font
            .Size = 12
            .Name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        With shapeBeg
            .Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            .Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End With
        With shapeBeg.TextFrame2
            .TextRange.Text = "M"
            .MarginBottom = 0
            .MarginTop = 0
            .MarginRight = 0
            .MarginLeft = 3
        End With
        shapeBeg.TextFrame.AutoSize = True

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 3, r.Top + r.Height - 5, r.Left + 13, r.Top + 3)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.3

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, shapeBeg.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub Minor_Pass_Click(sender As Object, e As RibbonControlEventArgs) Handles Minor_Pass.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "/Desvio Menor/"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 10
            .font.bold = True
        End With
    End Sub

    Private Sub Conclusion_Click(sender As Object, e As RibbonControlEventArgs) Handles Conclusion.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Conclusão:"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub ArrowBox_Click(sender As Object, e As RibbonControlEventArgs) Handles ArrowBox.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape = Nothing
        Dim connector2 As Excel.Shape = Nothing
        Dim shapeBeg As Excel.Shape = Nothing
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        shapeBeg = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, r.Left + r.Width + 27, r.Top + r.Height + 2.5, 29, 14)
        With shapeBeg.TextFrame2.TextRange.Font
            .Size = 10
            .Name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With

        With shapeBeg.TextFrame2
            .TextRange.Text = ""
            .MarginBottom = 0
            .MarginTop = 1
            .MarginRight = 0
            .MarginLeft = 2
        End With
        shapeBeg.Line.ForeColor.RGB = RGB(0, 0, 0)

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Top + r.Height, r.Left + r.Width, r.Top + r.Height + 10)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.3

        connector2 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Top + r.Height + 10, r.Left + r.Width + 24, r.Top + r.Height + 10)
        connector2.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector2.Line.Weight = 1.3
        connector2.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, connector2.Name, shapeBeg.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub
End Class
