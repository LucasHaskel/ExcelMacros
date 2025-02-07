Sub EmMaiusculas()
    'Preciso armazenar o "RANGE" dentro de uma variável
    Dim celulas As Range
    
    'Para evitar erros, preciso verificar se quando a macro foi selecionada, o usuário selecionou antes alguma célula
    If Selection Is Nothing Then Exit Sub
    
    'Agora um comando para que a macro acesse todas as células armazenadas na variável e coloque em letras maiúsculas
    For Each celulas In Selection
    
    'Agora preciso espeficiar que é para aplicar a ação apenas para células que não estejam vazias:
        If Not IsEmpty(celulas.Value) Then
            celulas.Value = UCase(celulas.Value)
        End If
        
    'Agora, um comando para que continue a ação para cada célula dentro da seleção
    Next celulas
End Sub