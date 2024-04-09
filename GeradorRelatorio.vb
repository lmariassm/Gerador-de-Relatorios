Public varMes As String
Public varM As Integer
Public tag As String
Public modelo As String
Public varRelAnt As String
Public varEstação As String
Public varClick As Integer

Sub Proteger() 'protege a aplicação para que apenas adms possam editar
                
            'atalho Ctrl+U
            Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
            Application.DisplayFormulaBar = False
            Application.DisplayStatusBar = False
            With ActiveWindow
                .DisplayHeadings = False
                .DisplayHorizontalScrollBar = False
                .DisplayVerticalScrollBar = False
                .DisplayWorkbookTabs = False
            End With
            
            ActiveSheet.Range("a100").Select  'seleciona célula longe do campo da interface
            
            ActiveSheet.ScrollArea = "a1"     
            ActiveSheet.Protect Password:="12345"
       
    End Sub

    Sub Desproteger() 'desprotege a aplicação para edição
        
            'atalho Ctrl+L
            Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",true)"
            Application.DisplayFormulaBar = True
            Application.DisplayStatusBar = True
            With ActiveWindow
                .DisplayHeadings = True
                .DisplayHorizontalScrollBar = True
                .DisplayVerticalScrollBar = True
                .DisplayWorkbookTabs = True
            End With
            
            ActiveSheet.ScrollArea = ""
            ActiveSheet.Unprotect Password:="12345"
    
    End Sub
Sub home() 'faz com que a Home Page apareça primeiro quando abrir o arquivo

    Plan9.Activate
    
End Sub
Sub iconeRAF() 'clique no icone de adequação de fornecimento. vai para a planilha principal de geração

    Application.ScreenUpdating = False
    Plan1.Activate 
    Call bt_RelRAF 'chama a sub que seleciona o relatório desejado
    Application.ScreenUpdating = True
    
End Sub
Sub iconePressoes() 'clique no pressoes diarias. vai para a planilha principal de geração

    Application.ScreenUpdating = False
    Plan1.Activate
    Call bt_RelPressao 'chama a sub que seleciona o relatório desejado
    Application.ScreenUpdating = True
    
End Sub
Sub iconeIndicevaz() 'clique no icone de indice de vazamento. vai para a planilha principal de geração

    Application.ScreenUpdating = False
    Plan1.Activate
    Call bt_RelVazamentos 'chama a sub que seleciona o relatório desejado
    Application.ScreenUpdating = True
    
End Sub
Sub iconeTAEvaz() 'clique no icone de tempo de vazamento. vai para a planilha principal de geração

    Application.ScreenUpdating = False
    Plan1.Activate
    Call bt_relTAEvaz 'chama a sub que seleciona o relatório desejado
    Application.ScreenUpdating = True
    
End Sub
Sub iconeTAEfalta() 'clique no icone de tempo de falta de gás. vai para a planilha principal de geração

    Application.ScreenUpdating = False
    Plan1.Activate
    Call bt_relTAEfalta 'chama a sub que seleciona o relatório desejado
    Application.ScreenUpdating = True
    
End Sub
Sub iconeReligação() 'clique no icone de religação de fornecimento. vai para a planilha principal de geração

    Application.ScreenUpdating = False
    Plan1.Activate
    Call bt_relReligações 'chama a sub que seleciona o relatório desejado
    Application.ScreenUpdating = True
    
End Sub

Sub bt_RelRAF() 'vai para a planilha principal de geração com o tipo de relatorio selecionado

    Application.ScreenUpdating = False
    Plan2.Range("c4").Value = "CI006"
	'mostra os botoes de escolher data e esconde outros botoes 
    ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    
    'exibir seletor de mês e ano
    
    Plan2.Range("g6").Value = Year(Now) - 1
    Call addAno
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    
    Plan2.Range("j4").Value = Month(Now)
    If Month(Now) = 1 Then
        Plan2.Range("j6").Value = 12
        Plan2.Range("j8").Value = 11
    Else
        Plan2.Range("j6").Value = Month(Now) - 1
        Plan2.Range("j8").Value = Month(Now) - 2
    End If

    Call addMês
    'deixa os botoes dos outros relatórios na cor branca, e pinta o selecionado de amarelo
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    ActiveSheet.Shapes.Range(Array("Retângulo 1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    ActiveSheet.Shapes.Range(Array("Retângulo 29")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 58")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 61")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 62")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 76")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    'ActiveSheet.Range("a100").Select
    Application.ScreenUpdating = True
End Sub
Sub bt_RelPressao() 'vai para a planilha principal de geração com o tipo de relatorio selecionado

    Application.ScreenUpdating = False
    Plan2.Range("c4").Value = "QP001"
	'mostra os botoes de escolher data e esconde outros botoes
    ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoTrue
   
    
    'exibir seletor de mês e ano

    Plan2.Range("g6").Value = Year(Now) - 1
    Call addAno
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    
    Plan2.Range("j4").Value = Month(Now)
    If Month(Now) = 1 Then
        Plan2.Range("j6").Value = 12
        Plan2.Range("j8").Value = 11
    Else
        Plan2.Range("j6").Value = Month(Now) - 1
        Plan2.Range("j8").Value = Month(Now) - 2
    End If

    Call addMês
    'deixa os botoes dos outros relatórios na cor branca, e pinta o selecionado de amarelo
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    ActiveSheet.Shapes.Range(Array("Retângulo 1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 29")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    ActiveSheet.Shapes.Range(Array("Retângulo 58")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 61")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 62")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 76")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    'ActiveSheet.Range("a100").Select
    Application.ScreenUpdating = True
End Sub
Sub bt_RelVazamentos() 'vai para a planilha principal de geração com o tipo de relatorio selecionado
    
    Application.ScreenUpdating = False

    Plan5.Range("c2").Value = "VAZAMENTO DE GÁS"
    Plan2.Range("c4").Value = "SG002"
    ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    
    'exibir seletor de mês e ano

    Plan2.Range("g6").Value = Year(Now) - 1
    Call addAno
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    
    Plan2.Range("j4").Value = Month(Now)
    If Month(Now) = 1 Then
        Plan2.Range("j6").Value = 12
        Plan2.Range("j8").Value = 11
    Else
        Plan2.Range("j6").Value = Month(Now) - 1
        Plan2.Range("j8").Value = Month(Now) - 2
    End If

    Call addMês
     'deixa os botoes dos outros relatórios na cor branca, e pinta o selecionado de amarelo
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    ActiveSheet.Shapes.Range(Array("Retângulo 1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 29")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 58")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    ActiveSheet.Shapes.Range(Array("Retângulo 61")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 62")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 76")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    'ActiveSheet.Range("a100").Select
    Application.ScreenUpdating = True

End Sub
Sub bt_relTAEvaz() 'vai para a planilha principal de geração com o tipo de relatorio selecionado

    Application.ScreenUpdating = False
    Plan5.Range("c2").Value = "VAZAMENTO DE GÁS"
    Plan2.Range("c4").Value = "SG004"
    ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    
    'exibir seletor de mês e ano

    Plan2.Range("g6").Value = Year(Now) - 1
    Call addAno
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    
    Plan2.Range("j4").Value = Month(Now)
    If Month(Now) = 1 Then
        Plan2.Range("j6").Value = 12
        Plan2.Range("j8").Value = 11
    Else
        Plan2.Range("j6").Value = Month(Now) - 1
        Plan2.Range("j8").Value = Month(Now) - 2
    End If

    Call addMês
     'deixa os botoes dos outros relatórios na cor branca, e pinta o selecionado de amarelo
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    ActiveSheet.Shapes.Range(Array("Retângulo 1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 29")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 58")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 61")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    ActiveSheet.Shapes.Range(Array("Retângulo 62")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 76")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    'ActiveSheet.Range("a100").Select
    Application.ScreenUpdating = True
    
End Sub
Sub bt_relTAEfalta() 'vai para a planilha principal de geração com o tipo de relatorio selecionado

    Application.ScreenUpdating = False
    Plan5.Range("c2").Value = "FALTA DE GÁS"
    Plan2.Range("c4").Value = "SG005"
    ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    
    'exibir seletor de mês e ano

    Plan2.Range("g6").Value = Year(Now) - 1
    Call addAno
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    
    Plan2.Range("j4").Value = Month(Now)
    If Month(Now) = 1 Then
        Plan2.Range("j6").Value = 12
        Plan2.Range("j8").Value = 11
    Else
        Plan2.Range("j6").Value = Month(Now) - 1
        Plan2.Range("j8").Value = Month(Now) - 2
    End If

    Call addMês
     'deixa os botoes dos outros relatórios na cor branca, e pinta o selecionado de amarelo
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    ActiveSheet.Shapes.Range(Array("Retângulo 1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 29")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 58")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 61")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 62")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    ActiveSheet.Shapes.Range(Array("Retângulo 76")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    'ActiveSheet.Range("a100").Select
    Application.ScreenUpdating = True


End Sub
Sub bt_relReligações() 'vai para a planilha principal de geração com o tipo de relatorio selecionado

    Application.ScreenUpdating = False
    Plan2.Range("c4").Value = "CI002"
    ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    
    'exibir seletor de mês e ano

    Plan2.Range("g6").Value = Year(Now) - 1
    Call addAno
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    
    Plan2.Range("j4").Value = Month(Now)
    If Month(Now) = 1 Then
        Plan2.Range("j6").Value = 12
        Plan2.Range("j8").Value = 11
    Else
        Plan2.Range("j6").Value = Month(Now) - 1
        Plan2.Range("j8").Value = Month(Now) - 2
    End If

    Call addMês
     'deixa os botoes dos outros relatórios na cor branca, e pinta o selecionado de amarelo
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    ActiveSheet.Shapes.Range(Array("Retângulo 1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 29")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 58")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 61")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 62")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Retângulo 76")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    'ActiveSheet.Range("a100").Select
    Application.ScreenUpdating = True

End Sub
Sub bt_gerarRel() 'identifica qual o relatório pela "tag" e gera

'identificar o tipo de relatório
    Select Case Plan2.Range("c4").Value
        Case Is = "CI006":
            Call gerarRelRAF
        Case Is = "QP001":
            Call gerarRelPressao
        Case Is = "SG002":
            Call gerarRelVazamentos
        Case Is = "SG004":
            Call gerarRelTAEvaz
        Case Is = "SG005":
            Call gerarRelTAEfalta
        Case Is = "CI002":
            Call gerarRelReligações
    End Select

End Sub

Sub bt_verRel() 'possibilita a visualização de um relatório já existente
Dim Apagar As String

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="12345"

    'identificar o tipo de relatório
    Select Case Plan2.Range("c4").Value
        Case Is = "CI006":
            Call verRelRAF
        Case Is = "QP001": 
            'fecha o seletor de data
            ActiveSheet.Shapes.Range(Array("seletorData")).Visible = msoFalse
            ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
            ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
            'mostra todas as estações
            ActiveSheet.Shapes.Range(Array("escolherEstação")).Visible = msoTrue
            Plan2.Range("B11:B50").Copy
            Plan4.Activate
            Plan4.Range("a2").Select
            ActiveSheet.Paste
            Plan1.Activate
            'mostra o filtro de estações
            ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoTrue
            ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoTrue
            'apaga texto da caixa
            caixaPesquisa.Value = Empty
            varClick = 0
            'apaga a seleção e coloca as estações na interface
            For i = 1 To 5
                ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange = Plan4.Range("a1").Offset(i, 0)
            Next
    
        Case Is = "SG002": 
            Call verRelVazamentos
        Case Is = "SG004":
            Call verRelTAEvaz
        Case Is = "SG005":
            Call verRelTAEfalta
        Case Is = "CI002":
            Call verRelReligações
    End Select
    
    ActiveSheet.Protect Password:="12345"
    Application.ScreenUpdating = True
    
End Sub

Sub bt_pesquisarEstação() 'referente ao relatorio de pressoes, que possibilita o usuário a pesquisar a estação desejada

Dim Texto As String

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="12345"

    'copiar texto da caixa
    Plan2.Range("G11").Value = caixaPesquisa.Value

    'filtrar estações
    Plan4.Range("a2:a40").ClearContents
    Plan2.Range("B10:B50").AutoFilter
    Plan2.Range("$B$10:$B$50").AutoFilter Field:=1, Criteria1:="=*" & Plan2.Range("g11") & "*", Operator:=xlAnd
    Plan2.Range("$B$11:$B$50").Copy
    Plan4.Activate
    Plan4.Range("a2").Select
    ActiveSheet.Paste
    Plan2.Range("B10:B50").AutoFilter
    
    'mostrar roleta de estações filtradas
    Plan1.Activate
    ActiveSheet.Shapes.Range(Array("selecionarEstação")).Visible = msoTrue
    
    'colocar estações na interface
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange = Plan4.Range("a1").Offset(i, 0)
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next
    
    'colocar (ou não) setinhas
    ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoTrue
    If Plan4.Range("a7") = Empty Then
        ActiveSheet.Shapes.Range(Array("roletaEstações")).Visible = msoFalse
    End If
     
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
        
    Plan1.Activate
    
    ActiveSheet.Protect Password:="12345"
    Application.ScreenUpdating = True

End Sub
Sub bt_verRelPressão()
    Call verRelPressao
End Sub
Sub cliqueEstação1() 'o seletor de estações sempre mostra 5 estações, que se movimentam como uma roleta
Dim i As Integer

    'Deixar todos os nomes brancos
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next

    'Selecionar estação
    ActiveSheet.Shapes.Range(Array("caixa1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
    Plan2.Range("g11") = Plan4.Range("a2").Value
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoTrue
End Sub
Sub cliqueEstação2()
Dim i As Integer

    'Deixar todos os nomes brancos
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next

    'Selecionar estação
    ActiveSheet.Shapes.Range(Array("caixa2")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
    Plan2.Range("g11") = Plan4.Range("a3").Value
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoTrue
    
End Sub
Sub cliqueEstação3()
Dim i As Integer

    'Deixar todos os nomes brancos
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next

    'Selecionar estação
    ActiveSheet.Shapes.Range(Array("caixa3")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
    Plan2.Range("g11") = Plan4.Range("a4").Value
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoTrue
    
End Sub
Sub cliqueEstação4()
Dim i As Integer

    'Deixar todos os nomes brancos
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next

    'Selecionar estação
    ActiveSheet.Shapes.Range(Array("caixa4")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
    Plan2.Range("g11") = Plan4.Range("a5").Value
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoTrue
    
End Sub
Sub cliqueEstação5()
Dim i As Integer

    'Deixar todos os nomes brancos
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next

    'Selecionar estação
    ActiveSheet.Shapes.Range(Array("caixa5")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 192, 0)
    
    Plan2.Range("g11") = Plan4.Range("a6").Value
    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoTrue
    
End Sub
Sub subirEstação() 'faz com que haja a rotatividade das estações mostradas no seletor, corresponde ao botao de baixo

ActiveSheet.Unprotect Password:="12345"

If varClick = 0 Then
    Exit Sub
End If
varClick = varClick - 1
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange = Plan4.Range("a1").Offset(i + varClick, 0)
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next
    
ActiveSheet.Protect Password:="12345"
Range("a100").Select

End Sub
Sub descerEstação() 'faz com que haja a rotatividade das estações mostradas no seletor, corresponde ao botao de cima
Dim varCont As Integer

ActiveSheet.Unprotect Password:="12345"

varCont = 0
For i = 2 To 41
    If Plan4.Range("a" & i).Value <> Empty Then
    varCont = varCont + 1
    End If
Next

If varClick = varCont - 5 Then
    Exit Sub
End If
varClick = varClick + 1
 For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange = Plan4.Range("a1").Offset(i + varClick, 0)
        ActiveSheet.Shapes.Range(Array("caixa" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next
ActiveSheet.Protect Password:="12345"
Range("a100").Select
End Sub

Sub addAno() 'faz com que haja a rotatividade do seletor de ano, sendo o botao de baixo
    
    If Plan2.Range("g6").Value = Year(Now) Then
        Exit Sub
    Else
        Plan2.Range("g6").Value = Plan2.Range("g6").Value + 1
        Plan2.Range("g8").Value = Plan2.Range("g6").Value - 1
        Plan2.Range("g4").Value = Plan2.Range("g6").Value + 1
    End If
    If Plan2.Range("g6").Value = Year(Now) Then
        Plan2.Range("g4").Value = Empty
    End If
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) = DateAdd("m", -1, DateSerial(Year(Now), Month(Now), 1)) Then
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    Else
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    End If
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) >= DateSerial(Year(Now), Month(Now), 1) Then
        If Plan2.Range("c4").Value = "SG002" Or Plan2.Range("c4").Value = "CI006" Or Plan2.Range("c4").Value = "CI002" Or Plan2.Range("c4").Value = "SG004" Or Plan2.Range("c4").Value = "SG005" Then
            ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
        Else
            If Plan2.Range("c4").Value = "QP001" And DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) > DateSerial(Year(Now), Month(Now), 1) Then
                ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
            End If
        End If
    Else
        ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoTrue
    End If
    
    If Plan2.Range("j6").Value = Month(Now) And Plan2.Range("c4").Value = "CI002" Then
        ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    End If
        
    
End Sub

Sub dimAno() 'faz com que haja a rotatividade do seletor de ano, sendo o botao de cima

    If Plan2.Range("g6").Value = Year(Now) - 4 Then
        Exit Sub
    Else
        Plan2.Range("g6").Value = Plan2.Range("g6").Value - 1
        Plan2.Range("g8").Value = Plan2.Range("g6").Value - 1
        Plan2.Range("g4").Value = Plan2.Range("g6").Value + 1
    End If
    If Plan2.Range("g6").Value = Year(Now) - 4 Then
        Plan2.Range("g8").Value = Empty
    End If
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) = DateAdd("m", -1, DateSerial(Year(Now), Month(Now), 1)) Then
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    Else
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    End If
   If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) >= DateSerial(Year(Now), Month(Now), 1) Then
        If Plan2.Range("c4").Value = "SG002" Or Plan2.Range("c4").Value = "CI002" Or Plan2.Range("c4").Value = "CI006" Or Plan2.Range("c4").Value = "SG004" Or Plan2.Range("c4").Value = "SG005" Then
            ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
        Else
            If Plan2.Range("c4").Value = "QP001" And DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) > DateSerial(Year(Now), Month(Now), 1) Then
                ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
            End If
        End If
    Else
        ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoTrue
    End If
    
    If Plan2.Range("j6").Value = Month(Now) And Plan2.Range("c4").Value = "CI002" Then
        ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    End If
    
End Sub
Sub addMês() 'faz com que haja a rotatividade do seletor de mês

    Plan2.Range("j8").Value = Plan2.Range("j6").Value
    Plan2.Range("j6").Value = Plan2.Range("j4").Value
    Plan2.Range("j4").Value = Plan2.Range("j6").Value + 1
    
    If Plan2.Range("j6").Value = 12 Then
        Plan2.Range("j4").Value = 1
    End If
    'só permite que possa ser gerado relatorio do mes anterior ao atual, com exceções
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) = DateAdd("m", -1, DateSerial(Year(Now), Month(Now), 1)) Then
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    Else
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    End If
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) >= DateSerial(Year(Now), Month(Now), 1) Then
        If Plan2.Range("c4").Value = "SG002" Or Plan2.Range("c4").Value = "CI002" Or Plan2.Range("c4").Value = "CI006" Or Plan2.Range("c4").Value = "SG004" Or Plan2.Range("c4").Value = "SG005" Then
            ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
        Else
            If Plan2.Range("c4").Value = "QP001" And DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) > DateSerial(Year(Now), Month(Now), 1) Then
                ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
            End If
        End If
    Else
        ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoTrue
    End If
	
    If Plan2.Range("j6").Value = Month(Now) And Plan2.Range("c4").Value = "CI002" Then
        ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    End If
    
  
    Plan2.Range("k4").Value = UCase(Format(DateSerial(2023, Plan2.Range("j4").Value, 1), "mmmm"))
    Plan2.Range("k6").Value = UCase(Format(DateSerial(2023, Plan2.Range("j6").Value, 1), "mmmm"))
    Plan2.Range("k8").Value = UCase(Format(DateSerial(2023, Plan2.Range("j8").Value, 1), "mmmm"))

End Sub
Sub dimMês() 'faz com que haja a rotatividade do seletor de ano

    Plan2.Range("j4").Value = Plan2.Range("j6").Value
    Plan2.Range("j6").Value = Plan2.Range("j8").Value
    Plan2.Range("j8").Value = Plan2.Range("j6").Value - 1
    
    If Plan2.Range("j6").Value = 1 Then
        Plan2.Range("j8").Value = 12
    End If
    'só permite que possa ser gerado relatorio do mes anterior ao atual, com exceções
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) = DateAdd("m", -1, DateSerial(Year(Now), Month(Now), 1)) Then
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    Else
         ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoFalse
    End If
    If DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) >= DateSerial(Year(Now), Month(Now), 1) Then
        If Plan2.Range("c4").Value = "SG002" Or Plan2.Range("c4").Value = "CI002" Or Plan2.Range("c4").Value = "CI006" Or Plan2.Range("c4").Value = "SG004" Or Plan2.Range("c4").Value = "SG005" Then
            ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
        Else
            If Plan2.Range("c4").Value = "QP001" And DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1) > DateSerial(Year(Now), Month(Now), 1) Then
                ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoFalse
            Else
                ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoTrue
            End If
        End If
    Else
        ActiveSheet.Shapes.Range(Array("verRelatorio")).Visible = msoTrue
    End If
    
    If Plan2.Range("j6").Value = Month(Now) And Plan2.Range("c4").Value = "CI002" Then
        ActiveSheet.Shapes.Range(Array("gerarRelatorio")).Visible = msoTrue
    End If
    
    Plan2.Range("k4").Value = UCase(Format(DateSerial(2023, Plan2.Range("j4").Value, 1), "mmmm"))
    Plan2.Range("k6").Value = UCase(Format(DateSerial(2023, Plan2.Range("j6").Value, 1), "mmmm"))
    Plan2.Range("k8").Value = UCase(Format(DateSerial(2023, Plan2.Range("j8").Value, 1), "mmmm"))
End Sub

Sub gerarRelRAF() 'gera o relatorio de adequação de fornecimento
Dim varRel As String
Dim varPasta As Object
Dim varCaminho As String
Dim varRNG As Range
Dim varComp As Integer


    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
     'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    If Plan2.Range("g13").Value = "SIM" Then
        GoTo Gerar
    End If
    
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx") = "" Then
            GoTo Gerar
        Else
            ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoTrue
            Exit Sub
        End If
    End If
    
Gerar:
    Call gerandoRel
    'selecionar data relatorio anterior 
    varM = Plan2.Range("c6").Value - 1
    Call selectData
    
    'definir o nome de relatório
    'verifica se teve alguma revisão do relatorio
    If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx") = "" Then
        varRel = "ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx"
    Else
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
            & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx") = "" Then
            varRel = "ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx"
        Else
            If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
                & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx") = "" Then
                varRel = "ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx"
            Else
                varRel = "ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx"
            End If
        End If
    End If
    
    'abrir o padrão de relatório baseado no anterior  (existem informações necessárias)
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\" & varRel
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'apagar conteúdos das linhas
    Workbooks(varRel).Activate
    Workbooks(varRel).Sheets("7S").Range("a99:a243").Select
    Selection.EntireRow.Hidden = False
    Workbooks(varRel).Sheets("7S").Range("h23:h47").ClearContents
    Workbooks(varRel).Sheets("7S").Range("h55:h96").ClearContents
    Workbooks(varRel).Sheets("7S").Range("h104:h145").ClearContents
    Workbooks(varRel).Sheets("7S").Range("h153:h193").ClearContents
    Workbooks(varRel).Sheets("7S").Range("h202:h223").ClearContents
    Workbooks("Gerador de Relatórios.xlsm").Activate
    If Plan2.Range("c6").Value - 1 = 0 Then
        Workbooks(varRel).Sheets("7S").Range("i8:l19").ClearContents
    End If
    
    'redefinir conteúdo caixa de texto obs
    Workbooks(varRel).Activate
    ActiveSheet.Shapes.Range(Array("TextBox 11")).TextFrame2.TextRange.Characters.Text = "Observações:" & Chr(13)
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'datas dos rodapés
    Workbooks(varRel).Sheets("7S").Range("m49").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    Workbooks(varRel).Sheets("7S").Range("m98").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    Workbooks(varRel).Sheets("7S").Range("m147").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    Workbooks(varRel).Sheets("7S").Range("m195").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    Workbooks(varRel).Sheets("7S").Range("m242").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    
    'data referencia
     Workbooks(varRel).Sheets("7S").Range("b9").Value = Plan2.Range("c8").Value
     Workbooks(varRel).Sheets("7S").Range("b58").Value = Plan2.Range("c8").Value
     Workbooks(varRel).Sheets("7S").Range("b107").Value = Plan2.Range("c8").Value
     Workbooks(varRel).Sheets("7S").Range("b156").Value = Plan2.Range("c8").Value
     Workbooks(varRel).Sheets("7S").Range("b204").Value = Plan2.Range("c8").Value
     
     'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    If Plan2.Range("c6").Value = 1 Then
        Plan2.Range("c8").Value = Plan2.Range("c8").Value + 1
    End If
     
    'Escolher pasta de dados
    varCaminho = "\\INFOSERVER\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\04 - Registro de Adequação de Fornecimento\" & Plan2.Range("c8").Value & "\" & varMes
    Set varPasta = CreateObject("Scripting.FilesystemObject").GetFolder(varCaminho)
    Set varRNG = Workbooks(varRel).Sheets("7S").Range("h22")
    
    ContArquivos = 0
    For Each arquivo In varPasta.Files
        ContArquivos = ContArquivos + 1
    Next
    Workbooks(varRel).Activate
	'pelo numero de arquivos, ele consegue gerar o relatório com o número necessário de laudas
    For i = 1 To 5
                ActiveSheet.Shapes.Range(Array("algás " & i)).Visible = msoTrue
                ActiveSheet.Shapes.Range(Array("arsal " & i)).Visible = msoTrue
    Next
        If ContArquivos <= 50 Then
            Workbooks(varRel).Sheets("7S").Range("a99:a243").Select
            Selection.EntireRow.Hidden = True
            For i = 3 To 5
                ActiveSheet.Shapes.Range(Array("algás " & i)).Visible = msoFalse
                ActiveSheet.Shapes.Range(Array("arsal " & i)).Visible = msoFalse
            Next
            Workbooks(varRel).Sheets("7S").Range("a49").Value = "pág. 1/2"
            Workbooks(varRel).Sheets("7S").Range("a98").Value = "pág. 2/2"
            ActiveSheet.Shapes.Range(Array("TextBox 11")).Select
            Selection.ShapeRange.IncrementTop -5000
            Selection.ShapeRange.IncrementTop 1214
        ElseIf ContArquivos <= 90 Then
            Workbooks(varRel).Sheets("7S").Range("a148:a243").Select
            Selection.EntireRow.Hidden = True
            For i = 4 To 5
                ActiveSheet.Shapes.Range(Array("algás " & i)).Visible = msoFalse
                ActiveSheet.Shapes.Range(Array("arsal " & i)).Visible = msoFalse
            Next
            Workbooks(varRel).Sheets("7S").Range("a49").Value = "pág. 1/3"
            Workbooks(varRel).Sheets("7S").Range("a98").Value = "pág. 2/3"
            Workbooks(varRel).Sheets("7S").Range("a147").Value = "pág. 3/3"
            ActiveSheet.Shapes.Range(Array("TextBox 11")).Select
            Selection.ShapeRange.IncrementTop -5000
            Selection.ShapeRange.IncrementTop 1969
        ElseIf ContArquivos <= 132 Then
            Workbooks(varRel).Sheets("7S").Range("a196:a243").Select
            Selection.EntireRow.Hidden = True
            ActiveSheet.Shapes.Range(Array("algás 5")).Visible = msoFalse
            ActiveSheet.Shapes.Range(Array("arsal 5")).Visible = msoFalse
            Workbooks(varRel).Sheets("7S").Range("a49").Value = "pág. 1/4"
            Workbooks(varRel).Sheets("7S").Range("a98").Value = "pág. 2/4"
            Workbooks(varRel).Sheets("7S").Range("a147").Value = "pág. 3/4"
            Workbooks(varRel).Sheets("7S").Range("a195").Value = "pág. 4/4"
            ActiveSheet.Shapes.Range(Array("TextBox 11")).Select
            Selection.ShapeRange.IncrementTop -5000
            Selection.ShapeRange.IncrementTop 2706
        Else
            Workbooks(varRel).Sheets("7S").Range("a49").Value = "pág. 1/5"
            Workbooks(varRel).Sheets("7S").Range("a98").Value = "pág. 2/5"
            Workbooks(varRel).Sheets("7S").Range("a147").Value = "pág. 3/5"
            Workbooks(varRel).Sheets("7S").Range("a147").Value = "pág. 4/5"
            Workbooks(varRel).Sheets("7S").Range("a242").Value = "pág. 5/5"
            ActiveSheet.Shapes.Range(Array("TextBox 11")).Select
            Selection.ShapeRange.IncrementTop -5000
            Selection.ShapeRange.IncrementTop 3425
        End If
    'preenche os relatorios com os nomes dos arquivos em uma determinada pasta
    For Each arquivo In varPasta.Files
        If varRNG.Row = 47 Or varRNG.Row = 96 Or varRNG.Row = 145 Or varRNG.Row = 193 Then
                Set varRNG = varRNG.Offset(8, 0)
            Else
                Set varRNG = varRNG.Offset(1, 0)
        End If
        varComp = Len(arquivo.Name)
        varRNG.Value = Left(arquivo.Name, varComp - 4)
        If varRNG.Value = "Thumb" Then
            varRNG.Value = Empty
            Set varRNG = varRNG.Offset(-1, 0)
        End If
    Next

    'Número de atendimentos
    Workbooks(varRel).Sheets("7S").Range("h7").Offset(Plan2.Range("c6").Value, 0) = varRNG.Offset(0, -1).Value
    Workbooks(varRel).Sheets("7S").Range("j7").Offset(Plan2.Range("c6").Value, 0) = 0
    Workbooks(varRel).Sheets("7S").Range("l7").Offset(Plan2.Range("c6").Value, 0) = 0
    
    'Salvar arquivo na pasta
    Workbooks(varRel).SaveAs Filename:= _
    "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") GR.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    'fechar arquivo gerado
    Workbooks("ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") GR.xlsx").Saved = True
    Workbooks("ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") GR.xlsx").Close

    'mensagem finalizado
    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoTrue
    
    Plan2.Range("g13").Value = Empty
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        

End Sub
Sub selectData()

'preenchimento do relatório baseado no anterior
'nome padrão das pastas de acordo com o mes

    Select Case varM
        Case Is = 1:
            varMes = "01 - Janeiro"
        Case Is = 2:
            varMes = "02 - Fevereiro"
        Case Is = 3:
            varMes = "03 - Março"
        Case Is = 4
            varMes = "04 - Abril"
        Case Is = 5
            varMes = "05 - Maio"
        Case Is = 6
            varMes = "06 - Junho"
        Case Is = 7
            varMes = "07 - Julho"
        Case Is = 8
            varMes = "08 - Agosto"
        Case Is = 9
            varMes = "09 - Setembro"
        Case Is = 10
            varMes = "10 - Outubro"
        Case Is = 11
            varMes = "11 - Novembro"
        Case Is = 0
            varMes = "12 - Dezembro"
            Plan2.Range("c8").Value = Plan2.Range("c8").Value - 1
        Case Is = 12
            varMes = "12 - Dezembro"
    End Select
    
End Sub
    
Sub verRelRAF() 'possibilita ver um relatorio existente

Dim varRel As String
Dim varPasta As Object
Dim varCaminho As String
Dim varRNG As Range
Dim varComp As Integer


    Application.ScreenUpdating = False
    'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    
    
    varM = Plan2.Range("c6").Value
    Call selectData
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) > 1 Then
        Call abrirArquivosPDF
        Exit Sub
    ElseIf Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\4. COMERCIAL_" & _
                Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 6C Comercial Individual 006 (" & Plan2.Range("c8").Value & "_" & Left(Plan1.varMes, 2) & ").pdf") = "" Then
            ActiveSheet.Shapes.Range(Array("msgnãoexisteRel")).Visible = msoTrue
        Else
            Call abrirArquivosPDF
            Exit Sub
        End If
    End If
    
    Application.ScreenUpdating = True
        
End Sub

Sub gerarRelPressao() 'gera os relatórios de pressão

Dim varRel As String
Dim varRNG As Range
Dim varNumero As Integer
Dim varRNG2 As Range
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim varCont As Integer

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
   'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    If Plan2.Range("g13").Value = "SIM" Then
        GoTo Gerar
    End If

   'verificar se já existe relatório oficial
   If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\1. QUALIDADE_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value)
   
        varCont = 0
        For Each objFile In objFolder.Files
            varCont = varCont + 1
        Next
        If varCont <= 1 Then
            GoTo Gerar
        Else
            ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoTrue
            GoTo Encerrar
        End If
    End If
   
Gerar:

    Call gerandoRel
    Set varRNG = Plan2.Range("b11")
    varNumero = 1
    'Escolher padrao de relatorio
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\01 - Gerador de Relatórios\01 - Padrões de relatórios\ARSAL Qualidade 001 PRESSÕES.xlsx"
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    varRel = "ARSAL " & varNumero & " Qualidade 001 PRESSÕES - " & varRNG & ".xlsx"
    varRelAnt = "ARSAL Qualidade 001 PRESSÕES.xlsx"   'padrão
    
    Do Until varRNG.Value = Empty
            
        'apagar conteúdos das linhas
        Workbooks(varRelAnt).Sheets("1Q").Range("g8:h38").ClearContents
        'redefinir conteúdo caixa de texto obs
        Workbooks(varRelAnt).Activate
        ActiveSheet.Shapes.Range(Array("CaixaDeTexto 4")).TextFrame2.TextRange.Characters.Text = "Observações:" & Chr(13)
        Workbooks("Gerador de Relatórios.xlsm").Activate
        'redefinir data
        Workbooks(varRelAnt).Sheets("1Q").Range("b16") = DateSerial(Plan2.Range("c8").Value, Plan2.Range("c6").Value, 1)
        'data rodapé
        Workbooks(varRelAnt).Sheets("1Q").Range("l48").Value = DateSerial(Year(Now), Month(Now), Day(Now))
        'preencher relatorio
        Workbooks(varRelAnt).Sheets("1Q").Range("b9").Value = varRNG
        Set varRNG2 = Plan3.Range("a:a").Find(varRNG, , , xlWhole)
            If Not varRNG2 Is Nothing Then
                'indicadores de pressao
                 Workbooks(varRelAnt).Sheets("1Q").Range("b13") = varRNG2.Offset(0, 3).Value
                 Workbooks(varRelAnt).Sheets("1Q").Range("c13") = varRNG2.Offset(0, 4).Value
                 Workbooks(varRelAnt).Sheets("1Q").Range("d13") = varRNG2.Offset(0, 5).Value
                 'preencher pressoes
                 tag = varRNG2.Offset(0, 1).Value
                 modelo = varRNG2.Offset(0, 2).Value
                Call preencher
            End If
            
        'Salvar arquivo na pasta
        
        'verificação de inconsistencia no relatorio
        If Plan2.Range("g15").Value = "erro" Then
            Workbooks(varRelAnt).SaveAs Filename:= _
            "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
            & "\1. QUALIDADE_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\@@@ " & varRel _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Else

            Workbooks(varRelAnt).SaveAs Filename:= _
            "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
            & "\1. QUALIDADE_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\" & varRel _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        End If
        
        Set varRNG = varRNG.Offset(1, 0)
        varNumero = varNumero + 1
        
        varRel = "ARSAL " & varNumero & " Qualidade 001 PRESSÕES - " & varRNG & ".xlsx"
        If Plan2.Range("g15").Value = "erro" Then
            varRelAnt = "@@@ ARSAL " & varNumero - 1 & " Qualidade 001 PRESSÕES - " & varRNG.Offset(-1, 0).Value & ".xlsx"
        Else
            varRelAnt = "ARSAL " & varNumero - 1 & " Qualidade 001 PRESSÕES - " & varRNG.Offset(-1, 0).Value & ".xlsx"
        End If
        
        Plan2.Range("g15").Value = ""
        
        
    Loop
    'fechar ultimo arquivo
    Workbooks(varRelAnt).Saved = True
    Workbooks(varRelAnt).Close
    
    'mensagem finalizado
    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoTrue

Encerrar:
    
    Plan2.Range("g13").Value = Empty
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
Sub fecharMsgRel() 'fecha pop up que informa que o relatório foi gerado

    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse

End Sub
Sub fecharMsgExisteRel() 'fecha pop up que informa que já existe o relatório que eu quero gerar

    ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoFalse

End Sub
Sub fecharMsgNãoExisteRel() 'fecha pop up que informa que não existe relatorio oficial para visualização

     ActiveSheet.Shapes.Range(Array("msgnãoexisteRel")).Visible = msoFalse

End Sub
Sub bt_SIM() 'botão clicado quando aceito gerar umm novo relatorio mesmo ja existindo um

    Plan2.Range("g13").Value = "SIM"
    Call fecharMsgExisteRel
    
     Select Case Plan2.Range("c4").Value
        Case Is = "CI006":
            Call gerarRelRAF
        Case Is = "QP001":
            ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoTrue
            Call gerarRelPressao
        Case Is = "SG002":
            Call gerarRelVazamentos
        Case Is = "SG004":
            Call gerarRelTAEvaz
        Case Is = "SG005":
            Call gerarRelTAEfalta
         Case Is = "CI002":
            Call gerarRelReligações
    End Select
    
    Call fecharMsgExisteRel
   
End Sub
Sub gerandoRel() 'pop up que indica que o relatorio está sendo gerado

ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoTrue

End Sub 
Sub bt_NAO() 'botão clicado quando não aceito gerar umm novo relatorio mesmo ja existindo um
    Call fecharMsgExisteRel
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub verRelPressao() 'posso visualizar as pressoes em um ponto de consumo mesmo sem ter fechado o mes

Dim varEstação As String
Dim varRel As String
Dim varRelPadrão As String
Dim varRNG As Range
Dim varRNG2 As Range

    ActiveSheet.Shapes.Range(Array("verRelatorioPressao")).Visible = msoFalse
    'Deixar todos os nomes brancos
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("caixa" & i)).Select
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next
    
    Application.ScreenUpdating = False
    'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    'compor nome do relatorio
    varEstação = Plan2.Range("g11").Value
    Set varRNG2 = Plan2.Range("b11:b50").Find(varEstação, , , xlWhole)
    varRel = "\ARSAL " & varRNG2.Offset(0, -1) & " Qualidade 001 PRESSÕES - " & varEstação
    
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) > 1 Then
        Call abrirArquivosPDF
        Exit Sub
    ElseIf Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) & " " & _
        Plan2.Range("c8").Value & "\1. QUALIDADE_" & Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & varRel & ".pdf") = "" Then
            GoTo Criar
        Else
            Call abrirArquivosPDF
            Exit Sub
        End If
    End If
    
Criar:
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
       
    'Escolher padrao de relatorio
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\01 - Gerador de Relatórios\01 - Padrões de relatórios\ARSAL Qualidade 001 PRESSÕES.xlsx"
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    varRelPadrão = "ARSAL Qualidade 001 PRESSÕES.xlsx"   'padrão

        'apagar conteúdos das linhas
        Workbooks(varRelPadrão).Sheets("1Q").Range("g8:h38").ClearContents
        'redefinir data
        Workbooks(varRelPadrão).Sheets("1Q").Range("b16") = DateSerial(Plan2.Range("c8").Value, Plan2.Range("c6").Value, 1)
        'data rodapé
        Workbooks(varRelPadrão).Sheets("1Q").Range("l48").Value = DateSerial(Year(Now), Month(Now), Day(Now))
        
        'redefinir conteúdo caixa de texto obs
        Workbooks(varRelPadrão).Activate
        ActiveSheet.Shapes.Range(Array("CaixaDeTexto 4")).TextFrame2.TextRange.Characters.Text = "Observações:" & Chr(13)
        Workbooks("Gerador de Relatórios.xlsm").Activate

        'preencher relatorio
        Set varRNG = Plan3.Range("a:a").Find(varEstação, , , xlWhole)
        Workbooks(varRelPadrão).Sheets("1Q").Range("b9").Value = varEstação
        'indicadores de pressao
        Workbooks(varRelPadrão).Sheets("1Q").Range("b13") = varRNG.Offset(0, 3).Value
        Workbooks(varRelPadrão).Sheets("1Q").Range("c13") = varRNG.Offset(0, 4).Value
        Workbooks(varRelPadrão).Sheets("1Q").Range("d13") = varRNG.Offset(0, 5).Value
        
        'preencher pressoes
        tag = varRNG.Offset(0, 1).Value
        modelo = varRNG.Offset(0, 2).Value
        varRelAnt = varRelPadrão
        Call preencher
            
        'abrir arquivo
    Workbooks(varRelPadrão).Activate
    
    Plan2.Range("g15").Value = ""
    Plan2.Range("g11") = ""
    Application.ScreenUpdating = True
    

End Sub

Sub gerarRelVazamentos() 'gerar o relatorio de indice de vazamentos
Dim varRNGfiltro As Range
Dim varRNG As Range
Dim varRNGsoma As Range

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
     'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    If Plan2.Range("g13").Value = "SIM" Then
        GoTo Gerar
    End If
    
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx") = "" Then
            GoTo Gerar
        Else
            ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoTrue
            Exit Sub
        End If
    End If
    
Gerar:
    Call gerandoRel
    'selecionar data relatorio anterior
    varM = Plan2.Range("c6").Value - 1
    Call selectData
    
	'verifica se teve alguma revisão do relatorio
    If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx") = "" Then
        varRel = "ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx"
    Else
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
            & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx") = "" Then
            varRel = "ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx"
        Else
            If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
                & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx") = "" Then
                varRel = "ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx"
            Else
                varRel = "ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx"
            End If
        End If
    End If
    
    'abrir o padrão de relatório baseado no mes anterior (existem informações necessárias)
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\" & varRel
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'apagar conteúdos das linhas
    Workbooks(varRel).Sheets("2S").Range("h37:h45").ClearContents
    Workbooks(varRel).Sheets("2S").Range("h55:h69").ClearContents

    'ajustar data do filtro
    Plan5.Range("a2").Value = ">=" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6"), 1), 0)
    Plan5.Range("b2").Value = "<" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6") + 1, 1), 0)
    
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    If Plan2.Range("c6").Value = 1 Then
        Plan2.Range("c8").Value = Plan2.Range("c8").Value + 1
    End If
    
    'preencher
    Call Plan5.filtro
    Set varRNGsoma = Workbooks(varRel).Sheets("2S").Range("h7").Offset(Plan2.Range("c6").Value, 0)
    Set varRNG = Workbooks(varRel).Sheets("2S").Range("h37")
    Set varRNGfiltro = Plan5.Range("k7")
    Do Until varRNGfiltro.Value = Empty
        'quantitativo de vazamentos por mês de acordo com o local
        Select Case varRNGfiltro.Offset(0, 9).Value
            Case Is = "CRM":    varRNGsoma.Offset(0, 1).Value = varRNGsoma.Offset(0, 1).Value + 1
            Case Is = "ERPM":   varRNGsoma.Offset(0, 2).Value = varRNGsoma.Offset(0, 2).Value + 1
            Case Is = "ERP":    varRNGsoma.Offset(0, 3).Value = varRNGsoma.Offset(0, 3).Value + 1
            Case Is = "ETC":    varRNGsoma.Offset(0, 4).Value = varRNGsoma.Offset(0, 4).Value + 1
            Case Is = "REDE PEAD":  varRNGsoma.Offset(0, 5).Value = varRNGsoma.Offset(0, 5).Value + 1
            Case Is = "REDE AÇO":   varRNGsoma.Offset(0, 6).Value = varRNGsoma.Offset(0, 6).Value + 1
        End Select
        
        For i = 1 To 6
            If varRNGsoma.Offset(0, i) = Empty Then
                varRNGsoma.Offset(0, i).Value = "0"
                varRNGsoma.Offset(15, i).Value = "0"
            End If
            varRNGsoma.Offset(15, i).Value = varRNGsoma.Offset(14, i).Value + varRNGsoma.Offset(0, i).Value
        Next
        
    
        varRNG.Value = varRNGfiltro.Offset(0, 9).Value & " - " & varRNGfiltro.Value
        If varRNG.Row = 45 Then
            Set varRNG = varRNG.Offset(10, 0)
        Else
            Set varRNG = varRNG.Offset(1, 0)
        End If
        Set varRNGfiltro = varRNGfiltro.Offset(1, 0)
    Loop
    
    'datas dos rodapés
    Workbooks(varRel).Sheets("2S").Range("N48").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    Workbooks(varRel).Sheets("2S").Range("N96").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    
    'ano relatório
    Workbooks(varRel).Sheets("2S").Range("b9").Value = Plan2.Range("c8").Value
    
    'redefinir conteúdo caixa de texto obs
    Workbooks(varRel).Activate
    ActiveSheet.Shapes.Range(Array("CaixaDeTexto 4")).TextFrame2.TextRange.Characters.Text = "Observações:" & Chr(13)
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'Salvar arquivo na pasta
    Workbooks(varRel).SaveAs Filename:= _
    "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") GR.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    Plan2.Range("g13").Value = Empty
    'mensagem finalizado
    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoTrue
    
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub verRelVazamentos() 'possibilita visualizar relatorio existente


Application.ScreenUpdating = False
    'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    
    
    varM = Plan2.Range("c6").Value
    Call selectData
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) > 1 Then
        Call abrirArquivosPDF
        Exit Sub
    ElseIf Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\3. SEGURANÇA_" & _
                Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2S Segurança 002 (" & Plan2.Range("c8").Value & "_" & Left(Plan1.varMes, 2) & ").pdf") = "" Then
            ActiveSheet.Shapes.Range(Array("msgnãoexisteRel")).Visible = msoTrue
        Else
            Call abrirArquivosPDF
            Exit Sub
        End If
    End If
    

Application.ScreenUpdating = True
    
End Sub

Sub gerarRelTAEvaz() 'gera o relatorio de tempo de atendimento de emergencia de vazamento
Dim varRNGfiltro As Range
Dim varRNG As Range
Dim varRNGTAE As Range

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
     'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    If Plan2.Range("g13").Value = "SIM" Then
        GoTo Gerar
    End If
    
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx") = "" Then
            GoTo Gerar
        Else
            ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoTrue
            Exit Sub
        End If
    End If
    
Gerar:
    Call gerandoRel
    'selecionar data relatorio anterior
    varM = Plan2.Range("c6").Value - 1
    Call selectData
    'verifica se teve alguma revisão do relatorio
    If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx") = "" Then
        varRel = "ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx"
    Else
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
            & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx") = "" Then
            varRel = "ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx"
        Else
            If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
                & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx") = "" Then
                varRel = "ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx"
            Else
                varRel = "ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx"
            End If
        End If
    End If
    
    'abrir o padrão de relatório baseado no anterior
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\" & varRel
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'apagar conteúdos das linhas
    Workbooks(varRel).Sheets("4S").Range("h23:l38").ClearContents

    'ajustar data do filtro
    Plan5.Range("a2").Value = ">=" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6"), 1), 0)
    Plan5.Range("b2").Value = "<" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6") + 1, 1), 0)
    
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    If Plan2.Range("c6").Value = 1 Then
        Plan2.Range("c8").Value = Plan2.Range("c8").Value + 1
    End If
    
    'preencher
	'filtra as ocorrencias de acordo com o banco de dados
    Call Plan5.filtro 'importa do banco de dados
    Set varRNG = Workbooks(varRel).Sheets("4S").Range("h23")
    Set varRNGfiltro = Plan5.Range("k7")
    Set varRNGTAE = Workbooks(varRel).Sheets("4S").Range("g7").Offset(Plan2.Range("c6").Value, 0)
    Do Until varRNGfiltro = Empty
        varRNG.Value = varRNGfiltro.Value
        varRNG.Offset(0, 2).Value = varRNGfiltro.Offset(0, -10).Value
        varRNG.Offset(0, 4).Value = Format(varRNGfiltro.Offset(0, -3).Value - varRNGfiltro.Offset(0, -7).Value, "hh:mm")

        Select Case varRNGfiltro.Offset(0, -3).Value - varRNGfiltro.Offset(0, -7).Value
            Case Is <= TimeSerial(0, 30, 0)
                    varRNGTAE.Offset(0, 1).Value = varRNGTAE.Offset(0, 1).Value + 1
            Case Is <= TimeSerial(1, 0, 0)
                    varRNGTAE.Offset(0, 2).Value = varRNGTAE.Offset(0, 2).Value + 1
            Case Is > TimeSerial(1, 0, 0)
                    varRNGTAE.Offset(0, 4).Value = varRNGTAE.Offset(0, 4).Value + 1
        End Select
        For i = 1 To 4
            If varRNGTAE.Offset(0, i).Value = Empty Then
                varRNGTAE.Offset(0, i).Value = "0"
            End If
        Next
        Set varRNGfiltro = varRNGfiltro.Offset(1, 0)
        Set varRNG = varRNG.Offset(1, 0)
    Loop
    
    'data do rodapé
     Workbooks(varRel).Sheets("4S").Range("l48").Value = DateSerial(Year(Now), Month(Now), Day(Now))
     
    'ano referência
     Workbooks(varRel).Sheets("4S").Range("b9").Value = Plan2.Range("c8").Value
     
    'redefinir conteúdo caixa de texto obs
    Workbooks(varRel).Activate
    ActiveSheet.Shapes.Range(Array("CaixaDeTexto 4")).TextFrame2.TextRange.Characters.Text = "Observações:" & Chr(13)
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'Salvar arquivo na pasta
    Workbooks(varRel).SaveAs Filename:= _
    "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") GR.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    Plan2.Range("g13").Value = Empty
    
    'mensagem finalizado
    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoTrue
    
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub verRelTAEvaz() 'possibilita visualizar relatorio existente


Application.ScreenUpdating = False
    'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    
    
    varM = Plan2.Range("c6").Value
    Call selectData
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) > 1 Then
        Call abrirArquivosPDF
        Exit Sub
    ElseIf Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\3. SEGURANÇA_" & _
                Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 4S Segurança 004 (" & Plan2.Range("c8").Value & "_" & Left(Plan1.varMes, 2) & ").pdf") = "" Then
            ActiveSheet.Shapes.Range(Array("msgnãoexisteRel")).Visible = msoTrue
        Else
            Call abrirArquivosPDF
            Exit Sub
        End If
    End If
Application.ScreenUpdating = True

End Sub

Sub gerarRelTAEfalta() 'gera relatorio de tempo de atendimento de emergencia de falta de gas

Dim varRNGfiltro As Range
Dim varRNG As Range
Dim varRNGTAE As Range

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
     'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    If Plan2.Range("g13").Value = "SIM" Then
        GoTo Gerar
    End If
    
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operaçã o\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx") = "" Then
            GoTo Gerar
        Else
            ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoTrue
            Exit Sub
        End If
    End If
    
Gerar:
    Call gerandoRel
    'selecionar data relatorio anterior
    varM = Plan2.Range("c6").Value - 1
    Call selectData
    'verifica se teve alguma revisão do relatorio
    If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx") = "" Then
        varRel = "ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx"
    Else
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
            & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx") = "" Then
            varRel = "ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV01.xlsx"
        Else
            If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
                & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx") = "" Then
                varRel = "ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV02.xlsx"
            Else
                varRel = "ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") REV03.xlsx"
            End If
        End If
    End If
    
    'abrir o padrão de relatório baseado no anterior
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\" & varRel
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'apagar conteúdos das linhas
    Workbooks(varRel).Sheets("5S").Range("h23:l47").ClearContents
    Workbooks(varRel).Sheets("5S").Range("h56:l76").ClearContents
    If Plan2.Range("c6").Value = 1 Then
        Workbooks(varRel).Sheets("5S").Range("h8:k18").ClearContents
    End If
    
    'ajustar data do filtro
    Plan5.Range("a2").Value = ">=" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6"), 1), 0)
    Plan5.Range("b2").Value = "<" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6") + 1, 1), 0)
    
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    If Plan2.Range("c6").Value = 1 Then
        Plan2.Range("c8").Value = Plan2.Range("c8").Value + 1
    End If
    
    'preencher
	'filtra as ocorrencias de acordo com o banco de dados
    Call Plan5.filtro 'importa do banco de dados
    Set varRNG = Workbooks(varRel).Sheets("5S").Range("h23")
    Set varRNGfiltro = Plan5.Range("k7")
    Set varRNGTAE = Workbooks(varRel).Sheets("5S").Range("g7").Offset(Plan2.Range("c6").Value, 0)
    Do Until varRNGfiltro = Empty
        varRNG.Value = varRNGfiltro.Value
        varRNG.Offset(0, 3).Value = varRNGfiltro.Offset(0, -10).Value
        varRNG.Offset(0, 4).Value = Format(varRNGfiltro.Offset(0, -4).Value - varRNGfiltro.Offset(0, -7).Value, "hh:mm")

        
        Select Case varRNGfiltro.Offset(0, -4).Value - varRNGfiltro.Offset(0, -7).Value
            Case Is <= TimeSerial(1, 0, 0)
                    varRNGTAE.Offset(0, 1).Value = varRNGTAE.Offset(0, 1).Value + 1
            Case Is <= TimeSerial(2, 0, 0)
                    varRNGTAE.Offset(0, 2).Value = varRNGTAE.Offset(0, 2).Value + 1
            Case Is <= TimeSerial(3, 0, 0)
                    varRNGTAE.Offset(0, 3).Value = varRNGTAE.Offset(0, 3).Value + 1
            Case Is > TimeSerial(3, 0, 0)
                    varRNGTAE.Offset(0, 4).Value = varRNGTAE.Offset(0, 4).Value + 1
        End Select
        For i = 1 To 5
            If varRNGTAE.Offset(0, i).Value = Empty Then
                varRNGTAE.Offset(0, i).Value = "0"
            End If
        Next
        If varRNG.Row = 47 Then
            Set varRNG = varRNG.Offset(9, 0)
        Else
            Set varRNG = varRNG.Offset(1, 0)
        End If
        Set varRNGfiltro = varRNGfiltro.Offset(1, 0)
    Loop
    
    'data do rodapé
     Workbooks(varRel).Sheets("5S").Range("l49").Value = DateSerial(Year(Now), Month(Now), Day(Now))
     Workbooks(varRel).Sheets("5S").Range("l97").Value = DateSerial(Year(Now), Month(Now), Day(Now))
     
    'ano referência
     Workbooks(varRel).Sheets("5S").Range("b9").Value = Plan2.Range("c8").Value
     
    'redefinir conteúdo caixa de texto obs
    Workbooks(varRel).Activate
    ActiveSheet.Shapes.Range(Array("CaixaDeTexto 6")).TextFrame2.TextRange.Characters.Text = "Observações:" & Chr(13)
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'Salvar arquivo na pasta
    Workbooks(varRel).SaveAs Filename:= _
    "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
    & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ") GR.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    'mensagem finalizado
    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoTrue
    
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Plan2.Range("g13").Value = Empty

End Sub

Sub verRelTAEfalta() 'possibilita visualizar relatorio existente

Application.ScreenUpdating = False
    'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    
    
    varM = Plan2.Range("c6").Value
    Call selectData
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) > 1 Then
        Call abrirArquivosPDF
        Exit Sub
    ElseIf Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\3. SEGURANÇA_" & _
                Left(Plan1.varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 5S Segurança 005 (" & Plan2.Range("c8").Value & "_" & Left(Plan1.varMes, 2) & ").pdf") = "" Then
            ActiveSheet.Shapes.Range(Array("msgnãoexisteRel")).Visible = msoTrue
        Else
            Call abrirArquivosPDF
            Exit Sub
        End If
    End If
    
Application.ScreenUpdating = True

End Sub

Sub gerarRelReligações() 'gera o relatorio de religações de fornecimento

Dim varRNGfiltro As Range
Dim varRNG As Range
Dim varRNGfiltroAUE As Range

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
     'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    'selecionar data relatorio atual
    varM = Plan2.Range("c6").Value
    Call selectData
    
    If Plan2.Range("g13").Value = "SIM" Then
        GoTo Gerar
    End If
    
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2CA Anexo Comercial Individual 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx") = "" Then
            GoTo Gerar
        Else
            ActiveSheet.Shapes.Range(Array("msgExisteRel")).Visible = msoTrue
            Exit Sub
        End If
    End If
    
Gerar:
    
    Call gerandoRel
    'abrir o padrão de relatório
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\01 - Gerador de Relatórios\01 - Padrões de relatórios\ARSAL 2CA_ Religações 002 (aaaa_mm).xlsx"
    Workbooks("Gerador de Relatórios.xlsm").Activate
    varRel = "ARSAL 2CA_ Religações 002 (aaaa_mm)"
    
    'data dos rodapés
    Workbooks(varRel).Sheets("5S").Range("K50").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    Workbooks(varRel).Sheets("5S").Range("K98").Value = DateSerial(Year(Now), Month(Now), Day(Now))
    
    'Apagar conteúdos de linhas
    Workbooks(varRel).Sheets("5S").Range("h8:l48").ClearContents
    Workbooks(varRel).Sheets("5S").Range("h57:l88").ClearContents
     
     
    'ano referência
    Workbooks(varRel).Sheets("5S").Range("b9").Value = DateSerial(Plan2.Range("c8").Value, Plan2.Range("c6").Value, 1)
    Workbooks(varRel).Sheets("5S").Range("b59").Value = DateSerial(Plan2.Range("c8").Value, Plan2.Range("c6").Value, 1)
    'ajustar data dos filtros
    Plan6.Range("a2").Value = ">=" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6"), 1), 0)
    Plan6.Range("b2").Value = "<" & Format(DateSerial(Plan2.Range("c8"), Plan2.Range("c6") + 1, 1), 0)
    Plan6.Range("a3").Value = Plan6.Range("a2").Value
    Plan6.Range("b3").Value = Plan6.Range("b2").Value
    Plan7.Range("a2").Value = Plan6.Range("a2").Value
    Plan7.Range("b2").Value = Plan6.Range("b2").Value
    
    'preencher religações agendadas
    Call Plan6.filtroBanco
    Set varRNG = Workbooks(varRel).Sheets("5S").Range("h8")
    Set varRNGfiltro = Plan6.Range("c7")
    Do Until varRNGfiltro = Empty
        varRNG.Value = varRNGfiltro.Value
        varRNG.Offset(0, 2).Value = varRNGfiltro.Offset(0, 3).Value
        varRNG.Offset(0, 1).Value = DateSerial(Left(varRNGfiltro.Offset(0, 4).Value, 2), Mid(varRNGfiltro.Offset(0, 4).Value, 3, 2), Mid(varRNGfiltro.Offset(0, 4).Value, 5, 2))
        varRNG.Offset(0, 3).Value = varRNGfiltro.Offset(0, 4).Value
        If varRNG.Row = 48 Then
            Set varRNG = varRNG.Offset(9, 0)
        Else
            Set varRNG = varRNG.Offset(1, 0)
        End If
        Set varRNGfiltro = varRNGfiltro.Offset(1, 0)
    Loop
    
    'preencher religações de urgencia e emergencia
    Call Plan7.filtroBancoAUE
    Set varRNGfiltroAUE = Plan7.Range("c7")
    Do Until varRNGfiltroAUE = Empty
        varRNG.Value = varRNGfiltroAUE.Value
        varRNG.Offset(0, 1).Value = varRNGfiltroAUE.Offset(0, 3).Value
        varRNG.Offset(0, 2).Value = varRNGfiltroAUE.Offset(0, 3).Value
        varRNG.Offset(0, 3).Value = varRNGfiltroAUE.Offset(0, 4).Value
        'verificar hora de finalização da religação
        If varRNGfiltroAUE.Offset(0, 6).Value >= TimeSerial(20, 0, 0) Then
            varRNG.Offset(0, 4).Value = "*"
        End If
        
        If varRNG.Row = 48 Then
            Set varRNG = varRNG.Offset(9, 0)
        Else
            Set varRNG = varRNG.Offset(1, 0)
        End If
        Set varRNGfiltroAUE = varRNGfiltroAUE.Offset(1, 0)
    Loop
    
    'Salvar arquivo na pasta
    Workbooks(varRel).SaveAs Filename:= _
    "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2CA Anexo Comercial Individual 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks("Gerador de Relatórios.xlsm").Activate
    
    'mensagem finalizado
    ActiveSheet.Shapes.Range(Array("msgRelPronto")).Visible = msoTrue
    
    ActiveSheet.Shapes.Range(Array("gerandoRelatorio")).Visible = msoFalse
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Plan2.Range("g13").Value = Empty
    
End Sub

Sub verRelReligações() 'possibilita visualizar relatorio existente

Application.ScreenUpdating = False
    'atualizar data
    Plan2.Range("c8").Value = Plan2.Range("g6").Value
    Plan2.Range("c6").Value = Plan2.Range("j6").Value
    Plan2.Range("c6").NumberFormat = "00"
    
    
    varM = Plan2.Range("c6").Value
    Call selectData
    'verificar se já existe relatório oficial
    If Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) > 1 Then
        Call abrirArquivosPDF
        Exit Sub
    ElseIf Abs(DateDiff("m", DateSerial(Plan2.Range("g6").Value, Plan2.Range("j6").Value, 1), DateSerial(Year(Now), Month(Now), 1))) = 1 Then
        If Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) _
        & " " & Plan2.Range("c8").Value & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2CA_ Planilha Anexo Comercial Individual 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").pdf") = "" Or Dir("\\infoserver\Arsal\" & Plan2.Range("c8").Value & "\RELATORIOS MENSAIS " & Left(Plan1.varMes, 2) & "_" & Plan2.Range("c8").Value & "\PDF\1. GEOP_" & Left(Plan1.varMes, 2) _
        & " " & Plan2.Range("c8").Value & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & "\ARSAL 2CA_ Anexo Comercial Individual 002 (" & Plan2.Range("c8").Value & "_" & Left(varMes, 2) & ").pdf") = "" Then
            ActiveSheet.Shapes.Range(Array("msgnãoexisteRel")).Visible = msoTrue
        Else
            Call abrirArquivosPDF
            Exit Sub
        End If
    End If
    
Application.ScreenUpdating = True

End Sub

Sub PDFsSegurança() 'converte em pdf todos os arquivos da pasta segurança
Dim saveLocation As String
Dim varPasta As Object
Dim varData As Date

    Plan2.Range("c6").Value = Month(Now) - 1
    Plan2.Range("c6").NumberFormat = "00"
    varM = Plan2.Range("c6").Value
    Call Plan1.selectData
    
    Application.ScreenUpdating = False
    
    'pega mes anterior
    'verifica se é janeiro
    If Month(Now) = 1 Then
            varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) - 1 & "\" & varMes _
            & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Year(Now) - 1
        Else
             varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) & "\" & varMes _
            & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Year(Now)
    End If
    Set varPasta = CreateObject("Scripting.FilesystemObject").GetFolder(varCaminho)
    
    'percorre cada arquivo e converte em pdf
    For Each arquivo In varPasta.Files
        'saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\22 - ARSAL PDFs\2. SEGURANÇA\" & arquivo.Name & ".pdf"
        saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes _
        & "\3. SEGURANÇA_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & " - PDF\" & Left(arquivo.Name, Len(arquivo.Name) - 5) & ".pdf"
        If arquivo.Name = "Thumbs.db" Then
            Exit Sub
        End If
        Workbooks.Open Filename:=varCaminho & "\" & arquivo.Name
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveLocation
        Workbooks(arquivo.Name).Saved = True
        Workbooks(arquivo.Name).Close
    Next
    
    Application.ScreenUpdating = True
    
End Sub

Sub PDFsComercial() 'converte em pdf todos os arquivos da pasta comercial
Dim saveLocation As String
Dim varPasta As Object
Dim varData As Date

    Plan2.Range("c6").Value = Month(Now) - 1
    Plan2.Range("c6").NumberFormat = "00"
    varM = Plan2.Range("c6").Value
    Call Plan1.selectData
    
    Application.ScreenUpdating = False
    
    'pega mes anterior
    'verifica se é janeiro
    If Month(Now) = 1 Then
            varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) - 1 & "\" & varMes _
            & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Year(Now) - 1
        Else
             varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) & "\" & varMes _
            & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Year(Now)
            
    End If
    Set varPasta = CreateObject("Scripting.FilesystemObject").GetFolder(varCaminho)
    
    'percorre cada arquivo e converte em pdf
    For Each arquivo In varPasta.Files
        'saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\22 - ARSAL PDFs\3. COMERCIAL\" & arquivo.Name & ".pdf"
        saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes & "\4. COMERCIAL_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & " - PDF\" & Left(arquivo.Name, Len(arquivo.Name) - 5) & ".pdf"
        If arquivo.Name = "Thumbs.db" Then
            Exit Sub
        End If
        Workbooks.Open Filename:=varCaminho & "\" & arquivo.Name
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveLocation
        Workbooks(arquivo.Name).Saved = True
        Workbooks(arquivo.Name).Close
    Next
    
    Application.ScreenUpdating = True
    
End Sub

Sub PDFsQualidade() 'converte em pdf todos os arquivos da pasta qualidade
Dim saveLocation As String
Dim varPasta As Object
Dim varData As Date

    Plan2.Range("c6").Value = Month(Now) - 1
    Plan2.Range("c6").NumberFormat = "00"
    varM = Plan2.Range("c6").Value
    Call Plan1.selectData
    
    Application.ScreenUpdating = False
    
    'pega mes anterior
    'verifica se é janeiro
    If Month(Now) = 1 Then
            varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) - 1 & "\" & varMes _
            & "\1. QUALIDADE_" & Left(varMes, 2) & " " & Year(Now) - 1
        Else
             varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) & "\" & varMes _
            & "\1. QUALIDADE_" & Left(varMes, 2) & " " & Year(Now)
            
    End If
    Set varPasta = CreateObject("Scripting.FilesystemObject").GetFolder(varCaminho)
        
    'percorre cada arquivo e converte em pdf
    For Each arquivo In varPasta.Files
        'saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\22 - ARSAL PDFs\1. QUALIDADE\" & arquivo.Name & ".pdf"
        saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes & "\1. QUALIDADE_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & " - PDF\" & Left(arquivo.Name, Len(arquivo.Name) - 5) & ".pdf"
        If arquivo.Name = "Thumbs.db" Then
            Exit Sub
        End If
        Workbooks.Open Filename:=varCaminho & "\" & arquivo.Name
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveLocation
        Workbooks(arquivo.Name).Saved = True
        Workbooks(arquivo.Name).Close
    Next
    
    Application.ScreenUpdating = True
    
End Sub

Sub PDFsLaudos() 'converte em pdf todos os arquivos da pasta laudos
Dim saveLocation As String
Dim varPasta As Object
Dim varData As Date

    Plan2.Range("c6").Value = Month(Now) - 1
    Plan2.Range("c6").NumberFormat = "00"
    varM = Plan2.Range("c6").Value
    Call Plan1.selectData
    
    Application.ScreenUpdating = False
    
    'pega mes anterior
    'verifica se é janeiro
    If Month(Now) = 1 Then
            varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) - 1 & "\" & varMes _
            & "\5. LAUDOS_" & Left(varMes, 2) & " " & Year(Now) - 1
        Else
             varCaminho = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Year(Now) & "\" & varMes _
            & "\5. LAUDOS_" & Left(varMes, 2) & " " & Year(Now)
    End If
    Set varPasta = CreateObject("Scripting.FilesystemObject").GetFolder(varCaminho)
    
    'percorre cada arquivo e converte em pdf
        For Each arquivo In varPasta.Files
            'saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\22 - ARSAL PDFs\4. LAUDOS\" & arquivo.Name & ".pdf"
            saveLocation = "\\infoserver\Nod - Nucleo de Op e Dist\COP - Coordenadoria de Operação\01 - ARSAL\01 - Relatórios Mensais\" & Plan2.Range("c8").Value & "\" & varMes & "\5. LAUDOS_" & Left(varMes, 2) & " " & Plan2.Range("c8").Value & " - PDF\" & Left(arquivo.Name, Len(arquivo.Name) - 5) & ".pdf"
            If arquivo.Name = "Thumbs.db" Then
                Exit Sub
            End If
            If Left(arquivo.Name, 1) = "A" Or Left(arquivo.Name, 1) = "C" Then
                Workbooks.Open Filename:=varCaminho & "\" & arquivo.Name
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveLocation
                Workbooks(arquivo.Name).Saved = True
                Workbooks(arquivo.Name).Close
            End If
        Next
    Application.ScreenUpdating = True
End Sub

Sub AllPFDs() 'converte em pdf todos os arquivos de todas as 4 pastas de uma vez

    Call PDFsQualidade
    Call PDFsLaudos
    Call PDFsComercial
    Call PDFsSegurança

End Sub

Sub MenuPDF() 'mostra a tela de conversão de pdf para o usuario

Plan10.Activate

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'código para capturar as médias das pressoes diárias disponíveis no banco de dados
'para preencher os relatórios de pressoes diárias
Private Function CapturaMediaPressao(ByVal d As Date, ByVal tag As String)
     Dim rs As New ADODB.Recordset
     
     On Error GoTo Erro
     
     data = "TO_DATE('" & Format(DateAdd("d", 1, d), "dd/MM/yyyy 00:00:00") & "', 'dd/mm/yyyy hh24:mi:ss')"
     
     If Not DBIsConnected Then DBConnect
     
     QueryString = "SELECT * FROM HISTORICOS WHERE HSTTAG = '" & tag & "' AND HSTHORDIA = 'D' AND HSTDAT = " & data
     rs.CursorLocation = adUseServer
     rs.Open QueryString, Cnt
     
     PressaoMedia = -1
     While Not rs.EOF
       PressaoMedia = rs!HSTPRE
       rs.MoveNext
     Wend
     rs.Close
     
     CapturaMediaPressao = Round(PressaoMedia, 2)
     
     On Error GoTo 0
     Exit Function
    
Erro:
    MsgBox ("Erro ao tentar capturar a pressão média." & Err.Description)
     
End Function
Private Function CalculaMediaPressao(ByVal d As Date, ByVal tag As String)
    
    Dim rs As New ADODB.Recordset
    Dim soma, cont As Integer
    
    On Error GoTo Erro
    
    If Not DBIsConnected Then DBConnect
    
    DataInicial = "TO_DATE('" & Format(d, "dd/MM/yyyy 00:00:00") & "', 'dd/mm/yyyy hh24:mi:ss')"
    DataFinal = "TO_DATE('" & Format(DateAdd("d", 1, d), "dd/MM/yyyy 00:00:00") & "', 'dd/mm/yyyy hh24:mi:ss')"
     
    QueryString = "SELECT * FROM HISTORICOS WHERE HSTTAG = '" & tag & "' AND HSTHORDIA = 'H' AND HSTDAT >= " & DataInicial & " AND HSTDAT <= " & DataFinal & " ORDER BY HSTDAT"
    rs.CursorLocation = adUseServer
    rs.Open QueryString, Cnt
    
    soma = 0
    cont = 0
    While Not rs.EOF
      soma = soma + Round(rs!HSTPRE, 2)
      cont = cont + 1
      rs.MoveNext
    Wend
    rs.Close
    
    If cont = 0 Then
        CalculaMediaPressao = -1
    Else
       CalculaMediaPressao = soma / cont
    End If
    
    On Error GoTo 0
    Exit Function
    
Erro:
    MsgBox ("Erro ao tentar calcular a média." & Err.Description)
End Function

Private Sub CapturaInfo(ByVal estacao As String, ByRef tag As String, ByRef modelo As String, ByRef pressao As Double)
    
    On Error GoTo Erro
    Dim i As Integer
      
    For i = 2 To 200
        If Planilha2.Cells(i, 3).Value = estacao Then
            tag = Planilha2.Cells(i, 4)
            modelo = Planilha2.Cells(i, 5)
            pressao = Planilha2.Cells(i, 6)
            Exit For
        End If
    Next
          
    On Error GoTo 0
    Exit Sub
    
Erro:
    MsgBox ("Erro ao tentar capturar a tag." & Err.Description)
    
End Sub

Sub preencher()
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim pressao As Double
    Dim data As Date
   
    
    On Error GoTo Erro
    
    If Not DBIsConnected Then DBConnect
            
    data = DateSerial(Plan2.Range("c8").Value, Plan2.Range("c6").Value, 1)
    
    
    i = 8
    While Month(data) = Plan2.Range("c6").Value
        media = " "
        If Plan1.modelo = "XARTU" Then
            media = CapturaMediaPressao(data, Plan1.tag)
        ElseIf Plan1.modelo = "FLOBOSS" Then
            media = CalculaMediaPressao(data, Plan1.tag)
        ElseIf Plan1.modelo = "ELCOR" Then
            media = CapturaMediaPressao(data, Plan1.tag)
        End If
        If media < 0 Then
            media = Empty
        End If
       
        Workbooks(Plan1.varRelAnt).Sheets("1Q").Cells(i, 7) = data
        Workbooks(Plan1.varRelAnt).Sheets("1Q").Cells(i, 8) = media
        
        If media < Workbooks(Plan1.varRelAnt).Sheets("1Q").Range("b13").Value Or media > Workbooks(Plan1.varRelAnt).Sheets("1Q").Range("d13").Value Then
            Plan2.Range("g15").Value = "erro"
        End If
        
        
        data = DateAdd("d", 1, data)
        i = i + 1
        
    Wend
     
    On Error GoTo 0
    Exit Sub
    
Erro:
    MsgBox ("Erro ao tentar capturar dados." & Err.Description)

    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub filtroBanco()'acessa banco de dados e filtra

'filtro
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\Z_Programs\Agenda de Atendimentos\01 - Banco de Dados\Banco.xlsx", ReadOnly:=True
    Workbooks("Gerador de Relatórios.xlsm").Activate
    Workbooks("Banco.xlsx").Sheets("plan1").Range("TabelaBanco").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
            "A1:d3"), CopyToRange:=Range("A6:u6"), Unique:=False
    
    Workbooks("Banco.xlsx").Saved = True
    Workbooks("Banco.xlsx").Close

End Sub
Sub filtroBancoAUE() 'acessa banco de dados e filtra

'filtro
    Workbooks.Open Filename:="\\infoserver\Nod - Nucleo de Op e Dist\Z_Programs\Agenda de Atendimentos\01 - Banco de Dados\BancoAUE.xlsx", ReadOnly:=True
    Workbooks("Gerador de Relatórios.xlsm").Activate
    Workbooks("BancoAUE.xlsx").Sheets("plan1").Range("TabelaBanco").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("A1:d2"), CopyToRange:=Range("A6:v6"), Unique:=False
    
    Workbooks("BancoAUE.xlsx").Saved = True
    Workbooks("BancoAUE.xlsx").Close

End Sub