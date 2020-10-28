Public conexion As ADODB.Connection
Public rsConciliacion As ADODB.Recordset
Public CadenaConexion As String
Option Explicit
Sub ConectarBase()
Dim miBase As String

miBase = "D:\PRUEBAS INFORMES\BASEtRABAJO.accdb"
CadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0; " & "data source=" & miBase & ";"
If Len(Dir(miBase)) = 0 Then
    MsgBox "La base que intenta conectar no se encuentra disponible", vbCritical
    Exit Sub
End If
Set conexion = New ADODB.Connection
    If conexion.State = 1 Then
        conexion.Close
    End If
        conexion.Open (CadenaConexion)
End Sub
Sub abrirConciliacion()
    Set rsConciliacion = New ADODB.Recordset
    If rsConciliacion.State = 1 Then
        rsConciliacion.Close
    Else
        rsConciliacion.Open "Conciliacion", conexion, adOpenKeyset, adLockOptimistic, adCmdTable
    End If
End Sub
Sub consultarmaxPagos()
Dim Sql As String
Dim RS As ADODB.Recordset
    Call ConectarBase
     Set RS = New ADODB.Recordset
    Sql = "SELECT MAX(Conciliacion.[Fecha de Transmisión]) FROM Conciliacion"
    RS.Open Sql, conexion
    Sheets("Conciliacion").Range("C2").ClearContents
    Sheets("Conciliacion").Range("C2").CopyFromRecordset RS
    RS.Close
    Set RS = Nothing
    conexion.Close
    Set conexion = Nothing
    MsgBox "La fecha maxima es: " & Sheets("Conciliacion").Range("C2"), vbInformation, "Fecha Maxima"
End Sub
Sub cargarPagos()
Dim i, largConci As Long
    
     Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call ConectarBase
    Call abrirConciliacion
    largConci = Sheets("Conciliacion").Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 5 To largConci + 1
        With rsConciliacion
        .AddNew
            .Fields("Proceso") = Cells(i, "A")
            .Fields("Fecha de Transmisión") = Cells(i, "B")
            .Fields("Fecha de Aplicación") = Cells(i, "C")
            .Fields("Convenio") = Cells(i, "D")
            .Fields("ALIAS") = Cells(i, "E")
            .Fields("Consecutivo") = Cells(i, "F")
            .Fields("Cédula") = Cells(i, "G")
            .Fields("Nombre") = Cells(i, "H")
            .Fields("Obligación") = Cells(i, "I")
            .Fields("Valor Transmitido") = Cells(i, "J")
            .Fields("Valor Unificado aplicado") = Cells(i, "K")
            .Fields("Valor Sobrante Actual") = Cells(i, "L")
            .Fields("Valor Sobrante Anterior") = Cells(i, "M")
            .Fields("Valor Devolución Créditos") = Cells(i, "N")
            .Fields("Estado Obligación") = Cells(i, "O")
            .Fields("Sobrante actual - sobrante anterior") = Cells(i, "P")
            .Fields("Transmitido menos Aplicado") = Cells(i, "Q")
            .Fields("Estado de conciliación") = Cells(i, "R")
            .Fields("Saldo Vencido") = Cells(i, "S")
            .Fields("OBSERVACION") = Cells(i, "T")
            .Fields("Cedula con Doble Credito") = Cells(i, "U")
            .Fields("validacion") = Cells(i, "V")
           
        End With
    Next i
'    rsConciliacion.Close
'    Set rsConciliacion = Nothing
'    conexion.Close
'    Set conexion = Nothing
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
