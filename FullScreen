
Sub FullScreen()
    'Oculta todas as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de fórmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = False
    
    'Alterar o nome do Excel
    'Application.Caption = "Label Name"
    
    With ActiveWindow
        'Ocultar barra horizontal
        .DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        .DisplayVerticalScrollBar = False
        
        'Ocultar guias das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os títulos de linha e coluna
        .DisplayHeadings = False
        
        'Oculta valores zero na planilha
        .DisplayZeros = False
        
        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With
End Sub       
