Private Sub Workbook_Open()
  With Sheet1.PTN
      For baris = 2 To 86
          .AddItem Sheet4.Cells(baris, 2)
      Next
  End With
  Sheet1.PTN.Text = "Pilih PTN"

  With Sheet1.SAINSOS
      .AddItem "SAINTEK/CAMPURAN"
      .AddItem "SOSHUM/CAMPURAN"
  End With
  Sheet1.SAINSOS.Text = "Kelompok Tes"
  Sheet1.AVGUTBK.Text = "Skor"
  Sheet1.KODEPRODI.Text = ""
  Sheet1.PREDIKSI.Text = ""
  Sheet1.MINIMAL.Text = ""

  Sheet1.PREDIKSI.BackColor = &HFFFFFF
  Sheet1.PREDIKSI.ForeColor = &H0&

  Sheet1.PTN.Select
End Sub
