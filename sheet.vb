Private Sub CommandButton1_Click()
    
    Call cargarPagos
    Sheets("Conciliacion").Cells(1, 3) = Format(Now(), "dd/mm/yyyy h:N:S")
    Sheets("Conciliacion").Range("A5:W1000000").ClearContents
    MsgBox "PAGOS CARGADOS CORRECTAMENTE", vbInformation, "Confirmacion"
End Sub

Private Sub CommandButton2_Click()
    Call consultarmaxPagos
End Sub
