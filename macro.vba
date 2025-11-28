Sub GerarFontes()
    Dim nome As String
    Dim linhas() As String
    Dim fontes(1 To 14) As String
    Dim i As Integer, j As Integer
    Dim xPos As Double, yPos As Double
    Dim paginaLargura As Double, paginaAltura As Double
    Dim s As Shape, sEnum As Shape
    Dim margemVertical As Double, margemHorizontal As Double
    Dim margemEntreLinhas As Double, margemEntreNumeroETexto As Double
    Dim ajusteVerticalTexto As Double
    Dim tamanhoFonte As Double
    Dim maxLarguraColuna As Double
    Dim alturaLinha As Double
    Dim larguraNumero As Double, larguraTexto As Double
    Dim larguraTotal As Double
    Dim alinhamentoEscolhido As String
    Dim xInicial As Double
    
    ' === Entrada do nome ===
    nome = InputBox("Digite o nome que será usado (use | para quebra de linha):", "Entrada de Nome")
    If nome = "" Then
        MsgBox "Nenhum nome foi inserido. Macro cancelado.", vbExclamation
        Exit Sub
    End If
    
    ' === Escolha de alinhamento ===
    alinhamentoEscolhido = InputBox("Escolha o alinhamento:" & vbCrLf & _
        "1 = Esquerda" & vbCrLf & "2 = Centralizado" & vbCrLf & "3 = Direita", _
        "Alinhamento do Texto", "1")
    
    ' Divide em linhas
    linhas = Split(nome, "|")
    
    ' === Lista de fontes ===
    fontes(1) = "Arial": fontes(2) = "Ananda": fontes(3) = "Birds of Paradise"
    fontes(4) = "Love": fontes(5) = "joseph sophia": fontes(6) = "Bella Donna"
    fontes(7) = "Avance": fontes(8) = "Best Valentina": fontes(9) = "Autography"
    fontes(10) = "Bernadette": fontes(11) = "Pacifico": fontes(12) = "Fiolex Girls"
    fontes(13) = "myloves": fontes(14) = "Amarillo"
    
    ' === Configurações de página ===
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    ActivePage.GetSize paginaLargura, paginaAltura
    
    ' === Parâmetros ===
    margemVertical = 12            ' entre blocos
    margemEntreLinhas = 6          ' entre linhas do mesmo bloco
    margemHorizontal = 12
    margemEntreNumeroETexto = 8
    ajusteVerticalTexto = 5
    xPos = 12
    yPos = paginaAltura - 30
    maxLarguraColuna = 0
    
    ' === Otimização ===
    Application.Optimization = True
    Application.EventsEnabled = False
    ActiveDocument.BeginCommandGroup "Gerar Fontes"
    
    ' === Loop das fontes ===
    For i = 1 To 14
        ' Define tamanho base
        If LCase(fontes(i)) = "amarillo" Then
            tamanhoFonte = 32
        Else
            tamanhoFonte = 64
        End If
        
        ' Cria o número à esquerda do primeiro texto
        Set sEnum = ActiveLayer.CreateArtisticText(0, 0, CStr(i) & ".")
        sEnum.Text.Story.Font = "Arial"
        sEnum.Text.Story.Size = 32
        larguraNumero = sEnum.SizeWidth
        
        ' Loop por cada linha do texto
        Dim alturaTotalTexto As Double
        alturaTotalTexto = 0
        
        For j = LBound(linhas) To UBound(linhas)
            Set s = ActiveLayer.CreateArtisticText(0, 0, linhas(j))
            s.Text.Story.Font = fontes(i)
            s.Text.Story.Size = tamanhoFonte
            alturaLinha = s.SizeHeight
            larguraTexto = s.SizeWidth
            
            ' Calcula largura total do conjunto (número + margem + linha)
            larguraTotal = larguraNumero + margemEntreNumeroETexto + larguraTexto
            
            ' Calcula X inicial do conjunto conforme alinhamento
            Select Case alinhamentoEscolhido
                Case "2" ' Centralizado
                    xInicial = xPos - larguraTotal / 2
                Case "3" ' Direita
                    xInicial = xPos - larguraTotal
                Case Else ' Esquerda
                    xInicial = xPos
            End Select
            
            ' Posiciona o número somente na primeira linha
            If j = LBound(linhas) Then
                sEnum.SetPosition xInicial, yPos
            End If
            
            ' Posiciona a linha
            s.SetPosition xInicial + larguraNumero + margemEntreNumeroETexto, _
                yPos - alturaTotalTexto - ajusteVerticalTexto
            
            ' Acumula altura para a próxima linha incluindo margem entre linhas
            alturaTotalTexto = alturaTotalTexto + alturaLinha + margemEntreLinhas
        Next j
        
        ' Atualiza max largura da coluna
        If larguraTotal > maxLarguraColuna Then maxLarguraColuna = larguraTotal
        
        ' Próxima posição vertical (após todas as linhas + margem entre blocos)
        yPos = yPos - (alturaTotalTexto - margemEntreLinhas + margemVertical)
        
        ' Após 7 textos, muda para segunda coluna
        If i = 7 Then
            xPos = xPos + maxLarguraColuna + margemHorizontal
            yPos = paginaAltura - 30
            maxLarguraColuna = 0
        End If
    Next i
    
    ' === Finaliza ===
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.EventsEnabled = True
    ActiveWindow.Refresh
    
    MsgBox "Foram criadas " & UBound(fontes) & " versões com fontes diferentes, com quebra de linha separada em objetos e margens configuráveis.", vbInformation
End Sub

