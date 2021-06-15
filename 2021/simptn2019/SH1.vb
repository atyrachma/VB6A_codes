Private Sub PRODI_Change()
  Sheet1.KODEPRODI.Text = ""
  Sheet1.PREDIKSI.Text = ""
  Sheet1.MINIMAL.Text = ""

  Sheet1.PREDIKSI.BackColor = &HFFFFFF
  Sheet1.PREDIKSI.ForeColor = &H0&
End Sub

Private Sub PTN_Change()
    Sheet1.KODEPRODI.Text = ""
    Sheet1.PREDIKSI.Text = ""
    Sheet1.MINIMAL.Text = ""

    Sheet1.PREDIKSI.BackColor = &HFFFFFF
    Sheet1.PREDIKSI.ForeColor = &H0&
    Sheet1.PRODI.Clear
    mim = Sheet4.Application.WorksheetFunction.Match(Sheet1.PTN, Sheet4.Range("B:B"), 0) - 1 'NOMOR PRODI DI PTN
    helo = Sheet9.Application.WorksheetFunction.Match(mim, Sheet9.Range("B:B"), 0) 'BARIS PRODI PERTAMA DI BAMAN
    Sheet1.PRODI.Text = "Pilih PRODI"
    i = 0
    Do While i < Sheet4.Cells(mim + 1, 5)
        With Sheet1.PRODI
            .AddItem Sheet9.Cells(helo + i, 4)
        End With
        i = i + 1
    Loop
    
End Sub
