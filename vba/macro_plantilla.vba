Option Explicit
Sub Crear_plantilla_JT()
    Dim CarpetaPrincipal As String
    Dim CarpetaSecundaria As String
    Dim ArchAct As String
    Dim ObjetoFSO As Object
    Dim ObjetoCarpeta As Object
    Dim ObjetoCarpeta2 As Object
    Dim ObjetoSubcarpeta As Object
    Dim ObjetoSubcarpeta2 As Object
    Dim Serie As String
    Dim wbDestino As Workbook

    
    ' Aqui se ingresa el año y mes con formato YYMM para poder correr incluso messes anteriores.
    Serie = InputBox("Ingrese el número de consecutivo.")
    
    If Serie = Empty Then Exit Sub

    ' Verificación Asignacion, CarpetaPrincipal = "C:\Users\jpineda\Desktop\CUBO\IMPLEMENTAR\1.MACRO\TC JEFES DE TRAZA"
    
    CarpetaPrincipal = "C:\Users\jpineda\Desktop\CUBO\PROYECTOS\Laboratorio\Laboratorio"
    
    Set ObjetoFSO = CreateObject("Scripting.FileSystemObject")
    
    Set ObjetoCarpeta = ObjetoFSO.GetFolder(CarpetaPrincipal)
    ' crea obj tipo file system (sistema de archivos)del SO
    
    For Each ObjetoSubcarpeta In ObjetoCarpeta.SubFolders
    ' Recorre las Carpetas hijas del obj carpeta
    
        CarpetaSecundaria = CarpetaPrincipal & "\" & ObjetoSubcarpeta.Name
        ' Armado de la ruta de la subcarpeta
        Debug.Print CarpetaSecundaria
        
                
        ArchAct = CarpetaSecundaria & "\" & Serie & " - TC Comercial (Jefe de Traza) v2.0.xlsx"
        'Arma laruta archivo nuevo
        Application.DisplayAlerts = False
        'Desactivan alertas para no leer los datos
        
        Dim wb As Workbook
        ' Se define la variable
        Set wb = Workbooks.Add("C:\Users\jpineda\Desktop\CUBO\PROYECTOS\Laboratorio\2403 - TC Comercial (Jefe de Traza) v2.0.xltm")
        ' Se establece valor variable, creando excel desde un tamplate
        
        Debug.Print ArchAct
        ' Se imp el archiv junto a la ruta
        
        Application.Wait (Now + TimeValue("0:00:10"))
        
        wb.Sheets("Resumen").Range("B7").Value = ObjetoSubcarpeta.Name
        
        With wb
            wb.SaveAs ArchAct
            wb.Close
        End With
        Application.DisplayAlerts = True
        ' Habilita alertas
    
    Next
    ' Continua con el for
    MsgBox ("Se Finalizo la creación de las plantillas")
    ' finaizado se muestra mensaje
    
 

End Sub
