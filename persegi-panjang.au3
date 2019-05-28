#include <Constants.au3>

$applName = "Menghitung luas dan keliling bangun persegi-panjang"
$panjang = InputBox($applName, "Masukan panjang (cm)",6)
$lebar = InputBox($applName, "Masukan lebar (cm)",5)
$luas = $panjang * $lebar
$keliling = 2 * ($panjang+$lebar)

$textPanjang = "Panjang " & $panjang & " cm. " & CHR(13)
$textLebar = "Lebar " & $lebar & " cm." & chr(13) & chr(13)

$textKeliling = "Maka, keliingnya adalah " & $keliling & " cm." & Chr(13)
$textLuas = "Dan luasnya adalah " & $luas & " cm2."

$textHasil = $textPanjang & $textLebar & $textKeliling & $textLuas

MsgBox($MB_SYSTEMMODAL, $applName, $textHasil)

Local $oMyExcel = ObjCreate("Excel.Application") ; Membuat object Excel
If @error Then
	MsgBox($MB_SYSTEMMODAL, $applName, "Ada kesalahan membuat Object Excel, Kode errornya " & @error)
	Exit
EndIf
If Not IsObj($oMyExcel) Then
	MsgBox($MB_SYSTEMMODAL, $applName, "Gagal membuat object Excel! Mungkin di komputermu tidak ada Excel")
	Exit
EndIf

$oMyExcel.Visible = 1
$oMyExcel.workbooks.add

With $oMyExcel.activesheet
.cells(1,1).value = $applName
.cells(3,1).value = $textPanjang
.cells(4,1).value = $textLebar
.cells(5,1).value = $textKeliling
.cells(6,1).value = $textLuas
EndWith

$oMyExcel.activeworkbook.saved = 1 ; Jangan tampilkan pertanyaan (yes/no) dari Excel
