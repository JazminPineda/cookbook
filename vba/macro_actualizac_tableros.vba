
' Este ejemplo Actualización masiva desde un archivo Excel principal a archivos secundarias, 
' visible únicamente para la persona responsable.  
Option Explicit
Sub Ejecutar_Carpetas()
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

    ' CarpetaPrincipal = "C:\Users\jpineda\Desktop\CUBO\IMPLEMENTAR\1.MACRO\TC JEFES DE TRAZA"
    
    CarpetaPrincipal = "C:\Users\jpineda\Dropbox\COMERCIAL\TC COMERCIALES (NR)\TC JEFES DE TRAZA"
    
    Set ObjetoFSO = CreateObject("Scripting.FileSystemObject")
    
    Set ObjetoCarpeta = ObjetoFSO.GetFolder(CarpetaPrincipal)
    
    For Each ObjetoSubcarpeta In ObjetoCarpeta.SubFolders
    
        CarpetaSecundaria = CarpetaPrincipal & "\" & ObjetoSubcarpeta.Name
    
        Debug.Print CarpetaSecundaria
        
        Set ObjetoCarpeta2 = ObjetoFSO.GetFolder(CarpetaSecundaria)
        
        For Each ObjetoSubcarpeta2 In ObjetoCarpeta2.Files
            
            If Left(ObjetoSubcarpeta2.Name, 4) = Serie Then
            
                ArchAct = CarpetaSecundaria & "\" & ObjetoSubcarpeta2.Name
                ' Se abre el archivo que se quiere actualizar,
                ' para que se ejecute la macro al abrirse
                Set wbDestino = Workbooks.Open(ArchAct)
                
                ' Se espera 3 minutos para dar tiempo a la actualización
                ' Application.Wait (Now + TimeValue("0:03:30"))
                ' wbDestino.RefreshAll No se utilizó esta opción porque se tenía que actualizar la fecha
                
                ' Para actualizar la fecha y las conexiónes se llama a m_ActAll del otro workbook
                Run "'" & wbDestino.Name & "'!" & "m_ActAll"
                Workbooks(wbDestino.Name).Close savechanges:=True
                
                Debug.Print ObjetoSubcarpeta2.Name
            
            End If
            
        Next
    
    Next
    MsgBox ("Se Finalizo actualización de las plantillas")
