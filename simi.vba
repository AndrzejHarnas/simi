Sub similarity()
Sheets("Dane").Select

lv = Range("B2").Value
ld = Range("D2").Value

For i = 1 To ld
 Range("F3").Offset("" & i & "", 0).Value = Range("D3").Offset("" & i & "", 0).Value
 Range("F3").Offset("" & i & "", 0).Select
 pelneObra
Next i

For lmw = 0 To ld - 1
 For i = 1 To lv
 
  Range("F3").Offset(0, "" & i & "").Value = Range("B3").Offset("" & i & "", 0).Value
  Range("F3").Offset(0, "" & i & "").Select
  pelneObra
  lnv = Len(Range("C3").Offset("" & i & "", 0).Value)
  dimens = Range("C3").Offset("" & i & "", 0).Value
  mainDimension = Range("D3").Offset("" & lmw + 1 & "", 0).Value
  
  znak = 1
  mvalue = 1
  flag = 0
 
  For k = 1 To lnv
    x = Mid(dimens, k, 1)
   xn = Mid(dimens, k + 1, 1)
   xn2 = Mid(dimens, k + 2, 1)
 
   If x = "/" Then
   znak = -1
   End If
 
   If x = mainDimension Then
    If xn = "^" Then
    mvalue = xn2
    End If
    Range("F3").Offset("" & lmw + 1 & "", "" & i & "").Value = mvalue * znak
    Range("F3").Offset("" & lmw + 1 & "", "" & i & "").Select
    pelneObra
    flag = 1
   End If
   
   If flag = 0 Then
   
   Range("F3").Offset("" & lmw + 1 & "", "" & i & "").Value = 0
   Range("F3").Offset("" & lmw + 1 & "", "" & i & "").Select
   pelneObra
   
   End If
    
  Next k
 Next i
Next lmw

prodlos
uklady

End Sub
Sub pelneObra()
'
' Makro1 Makro
'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub prodlos()
Worksheets("Dane").Activate

Dim vmax As Integer

vmax = Range("B2").Value
dimmax = Range("D2").Value

ReDim onlyvar(vmax)
ReDim dimens(vmax)
ReDim newvar(vmax - 1)
ReDim dimzest(dimmax)
ReDim mainDim(dimmax)

For i = 1 To vmax
onlyvar(i) = Range("B3").Offset("" & i & "", 0).Value
dimens(i) = Range("C3").Offset("" & i & "", 0).Value
Next i

For i = 1 To dimmax
mainDim(i) = Range("D3").Offset("" & i & "", 0).Value
Next i




'obliczenie ilosci kombinacji
zs = Silnia(vmax)
x = Ikombezpow(zs, vmax, dimmax)

Range("A1").Value = "ilość kombinacji="
Range("B1").Value = x

'wypisanie wszystkich kombinacji
v = GLKBP(vmax, dimmax, x, onlyvar)

Worksheets("Produkty").Activate

'sprawdzanie wszystkich kombinacji
For i = 0 To x - 1

    For j = 1 To dimmax
       x = Range("C1").Offset(i, j).Value
       
       For y = 1 To vmax
       
        If onlyvar(y) = x Then
       
        xd = dimens(y)
        dimzest(j) = xd
        xdd = xdd + xd
        
  ' sprawdzenie czy wymiary w kombinacji sie nie powtarzają
            For d = 1 To dimmax
                
                For dd = 1 To dimmax
                     
                  If dimzest(dd) = dimzest(d) And d <> dd Then
                    Range("C1").Offset(i, j + dimmax + 1).Value = 1
                  End If
                
                Next dd
            
            Next d
       
        End If
       
       Next y
       
       
       Range("C1").Offset(i, j + dimmax).Value = xd

    Next j
    
    liczbDims = 0
    flag = 0
    For k = 1 To dimmax
        
      
        
        For j = 1 To Len(xdd)
    
        wymiar = Mid(xdd, j, 1)
         
        If wymiar = mainDim(k) Then
        
        flag = 1
        
        End If
        
        Next j
        
        liczbDims = liczbDims + flag
        flag = 0
        
        
    Next k
    
     If liczbDims < 3 Then
     
     Range("C1").Offset(i, dimmax + dimmax + 2).Value = 1
     
     End If
     
     xdd = ""
     
     
     W1 = Range("C1").Offset(i, dimmax + dimmax + 1).Value
     W2 = Range("C1").Offset(i, dimmax + dimmax + 2).Value
     
    
     If W1 <> 1 And W2 <> 1 Then
     
     

     

     
     Range("C1").Offset(liczkomb, dimmax + dimmax + 3).Value = liczkomb + 1
     Range("C1").Offset(liczkomb, dimmax + dimmax + 4).Value = Range("C1").Offset(i, 0).Value
     
     liczkomb = liczkomb + 1

     End If
    
Next i


Worksheets("ukladyrownan").Activate
Range("A1").Value = liczkomb

For i = 1 To liczkomb

Worksheets("Produkty").Activate

x = Range("M1").Offset(i - 1, 0).Value

Worksheets("ukladyrownan").Activate

Range("B1").Offset(i, 0).Value = "Qzest" & i
Range("C1").Offset(i, 0).Value = x



Next i




End Sub
Sub uklady()
Worksheets("Dane").Activate

lvar = Range("B2").Value
ldim = Range("D2").Value

ReDim Tabvar(lvar)
ReDim Tabdim(ldim)
ReDim TabNvar(lvar - ldim)

For i = 1 To ldim

Tabdim(i) = Range("F3").Offset(i, 0).Value

Next i

For i = 1 To lvar

Tabvar(i) = Range("F3").Offset(0, i).Value

Next i

ReDim TabZest(ldim + 1)
ReDim CurentZest(ldim + 1)
ReDim TabDimens(ldim)
ReDim Matrix(ldim + 1, ldim)
ReDim Wspolczynniki(ldim + 1)

Worksheets("ukladyrownan").Activate
Lzestaw = Range("A1").Value

    For i = 1 To Lzestaw
'    For i = 74 To 74
     x = 1
     zestaw = Range("C1").Offset(i, 0).Value
        
        For j = 1 To ldim
        TabZest(j) = Mid(zestaw, j, 1)
        Next j
            
            For k = 1 To lvar
                
                flag = 0
                For d = 1 To ldim
                    
                    If Tabvar(k) = TabZest(d) Then
                    flag = 1
                    End If
            
                Next d
                
                If flag = 0 Then
                TabNvar(x) = Tabvar(k)
                x = x + 1
                End If
                
            Next k
            
    For k = 1 To lvar - ldim
        
        Range("C1").Offset(k, 1).Value = "Pi_" & k
        Range("C1").Offset(k, 2).Value = zestaw + TabNvar(k)
        CurZest = Range("C1").Offset(k, 2).Value
        
            For j = 1 To ldim + 1
            CurentZest(j) = Mid(CurZest, j, 1)
            Next j
        
        
       
        
        Worksheets("Dane").Activate
        
        For j = 1 To ldim + 1
            
            For z = 1 To lvar
            curvar = Range("F3").Offset(0, z)
            
                If curvar = CurentZest(j) Then
            
                    For m = 1 To ldim
            
                    Matrix(j, m) = Range("F3").Offset(m, z).Value
            
                    Next m
            
                End If
        
            Next z
            
        Next j
        
        
        Worksheets("ukladyrownan").Activate
        
    
        
   
        For j = 1 To ldim + 1
            Range("F1").Offset(1, j + 1).Value = CurentZest(j)
            For m = 1 To ldim
            Range("F1").Offset(m + 1, 1).Value = Tabdim(m)
            Range("F1").Offset(m + 1, j + 1).Value = Matrix(j, m)
            Next m
        Next j
        
      
      Range("H1").Offset(1, ldim + 1).Value = "-" & Range("H1").Offset(1, ldim).Value
      
      For j = 1 To ldim
      Range("H1").Offset(1 + j, ldim + 1).Value = -Range("H1").Offset(1 + j, ldim).Value
      Next j
        
      Range("G1").Offset(ldim + 3, 0).Select
      ActiveCell.Value = "wyznacznik="
      Range("H1").Offset(ldim + 3, 0).Select
      ActiveCell.FormulaR1C1 = "=MDETERM(R[-" & ldim + 1 & "]C:R[-" & ldim - 1 & "]C[" & ldim - 1 & "])"
      wyzWart = Range("H1").Offset(ldim + 3, 0).Value
      Range("H1").Offset(ldim + 6, 0).Select
      ActiveCell.Value = "macierz odwrotna"
      Range("H1").Offset(ldim + 6, 4).Select
      ActiveCell.Value = "wartości współczynników"
      Range(Cells(8 + ldim, 8), Cells(7 + ldim + ldim, 7 + ldim)).Select
      Selection.FormulaArray = "=MINVERSE(R[-" & ldim + 5 & "]C:R[-" & ldim + 3 & "]C[" & ldim - 1 & "])"
      Range(Cells(8 + ldim, 8 + ldim + 1), Cells(7 + ldim + ldim, 7 + ldim + 2)).Select
      Selection.FormulaArray = "=MMULT(RC[-" & ldim + 1 & "]:R[" & ldim - 1 & "]C[-" & ldim - 1 & "],R[-" & ldim + 5 & "]C[0]:R[-6]C[0])"
      Range("H1").Offset(ldim + 6, 4 + ldim + 2).Select
      
       For j = 1 To ldim
       
       Wspolczynniki(j) = Range(Cells(7 + ldim + j, 8 + ldim + 1), Cells(7 + ldim + j, 8 + ldim + 1)).Value
      
       Next j
      
       Wspolczynniki(ldim + 1) = 1
      
      

      
        Pi = Range("D1").Offset(0 + k, 0).Value
        ZestP = Range("D1").Offset(0 + k, 1).Value
        
        If wyzWart <> 0 Then
        Worksheets("LiczbyKryterialne").Activate
        Range("A" & i + c).Value = "Qzest" & i
        Range("B" & i + c).Value = Pi
        Range("C" & i + c).Value = ZestP
        
      
            For v = 1 To ldim
            
            x = Mid(ZestP, v, 1)
            Range("C" & i + c).Offset(0, v).Value = Wspolczynniki(v)
            
                If Wspolczynniki(v) < 0 Then
            
                wyk = -Wspolczynniki(v)
            
                    If wyk <> 1 Then
                    mianownik = mianownik & "(" & x & "^" & wyk & ")"
                    Else
                    mianownik = mianownik & "" & x
                    End If
            
                End If
                
                If Wspolczynniki(v) > 0 Then
                
                    wyk = Wspolczynniki(v)
            
                    If wyk <> 1 Then
                    licznik = licznik & "(" & x & "^" & wyk & ")"
                    Else
                    licznik = licznik & "" & x
                    End If
                
                End If
                
                
            
            Next v
            
            
            Range("C" & i + c).Offset(0, v).Value = Wspolczynniki(v)
            
                        x = Mid(ZestP, v, 1)
            Range("C" & i + c).Offset(0, v).Value = Wspolczynniki(v)
            
                If Wspolczynniki(v) < 0 Then
            
                wyk = -Wspolczynniki(v)
            
                    If wyk <> 1 Then
                    mianownik = mianownik & "(" & x & "^" & wyk & ")"
                    Else
                    mianownik = mianownik & "" & x
                    End If
            
                End If
                
                If Wspolczynniki(v) > 0 Then
                
                    wyk = Wspolczynniki(v)
            
                    If wyk <> 1 Then
                    licznik = licznik & "(" & x & "^" & wyk & ")"
                    Else
                    licznik = licznik & "" & x
                    End If
                
                End If
            
            
            
            Range("C" & i + c).Offset(0, v + 1).Value = licznik & "/" & mianownik
            
            c = c + 1
            mianownik = ""
            licznik = ""
      End If
      Worksheets("ukladyrownan").Activate
        
  
      
      

     Next k
Next i





End Sub

Sub genkomb()

n = 11
m = 3

z = Silnia(n)
x = Ikombezpow(z, n, m)

MsgBox "ilość kombinacji: " & x

v = GLKBP(n, m, x)



End Sub
'Funkcja obliczająca silnie
Function Silnia(wielkosczbioru)

If wielkosczbioru = 1 Or wielkosczbioru = 0 Then

Silnia = 1

Else

Silnia = wielkosczbioru * Silnia(wielkosczbioru - 1)

End If

End Function

'ilosc kombinacji bez powtórzen

Function Ikombezpow(silnian, wielkosczbioru, wielkoscpodzbioru)

'wzór: n! /(n-m)!*m!
' n=wielkosc zbioru m = wielkosc pozdzbioru

mianownik = Silnia(wielkosczbioru - wielkoscpodzbioru) * Silnia(wielkoscpodzbioru)
licznik = silnian

Ikombezpow = licznik / mianownik

End Function

'generowanie k losowych kombinacji bez poqwtórzeń

Function GLKBP(wielkosczbioru, wielkoscpodzbioru, k, vartab)

Dim tmp(20, 1000) 'pamieć już wylosowanych kombinacji
Dim podzbior(20) 'podzbior przechowujący wygenerowaną losowo kombinacje bez powtorzeń
Dim i, j, z, x, e, q As Integer 'liczniki pętli
Dim zmp As Integer 'wylosowana pojedyncza liczba w zakresie od 1 do wielkosczbioru
Dim powtorka As Integer 'zmienna odpowiedzialna za losowanie podzbioru, tak długo az to będzie właściwa z punktu widzenia matematycznego kombinacja wielkoscpodzbioru elementowa bez powtórzeń ze zbioru wielkosczbioru elementowego
Dim powtorka2 As Integer 'powtarzanie generowania kombinacji bez powtórzeń, tak długo aż bedzie się powtarzać. Celem zbioru jest uzyskanie nie wygenerowanej wcześniej kombinacji bez powtórzeń
Dim ilosc As Integer 'zmienna odpowiedzialna za wylosowanie tylko wielkoscpodzbioru elementów do podzbioru podzbior
Dim przecinek As Integer 'czy wypisywac przecinek
Dim wynik As String


'zerowanie pamieci zapamietanych kombinacji

For i = 1 To 1000

    For z = 1 To wielkoscpodzbioru
    
    tmp(z - 1, i) = 0
    
    Next z

Next i

'zerowanie podzbioru

For i = 1 To wielkoscpodzbioru

podzbior(i - 1) = 0

Next i

q = 1

Worksheets("Produkty").Cells.ClearContents

For j = 1 To k

    przecinek = 0
    ilosc = 0
    
    'losowanie kombinacji wielkoscpodzbioru elementowej bez powtórzeń ze zbioru wielkosczbioru elementowego
    powtorka2 = 0
        
    Do While powtorka2 = 0
        
            'wylosowanie prawidłowej z punktu widzenia matematyki kombinacji bez powtórzeń
            Do While ilosc < wielkoscpodzbioru
            
                powtorka = 0
                Randomize
                zmp = 1 + Round(Rnd * (wielkosczbioru - 1), 1)
                
                For z = 1 To wielkoscpodzbioru
                
                    If podzbior(z - 1) = zmp Then
                    
                    powtorka = 1
                    
                    Exit For
                    
                    End If
                
                Next z
                
                If powtorka = 0 Then
                
                    ilosc = ilosc + 1
                    podzbior(ilosc - 1) = zmp
                
                End If
        
            Loop
                
        'sprawdzenie czy ta kombinacja bez powtorzen była już wczesniej wylosowana i jak tak to powtórka losowania
        powtorka2 = 1
        
        For i = 1 To k
        
            If tmp(0, i - 1) = 0 Then
            
            Exit Do
            
            End If
            
            For z = 1 To wielkoscpodzbioru
            
                For x = 1 To wielkoscpodzbioru
                
                    If tmp(z - 1, i - 1) = podzbior(x - 1) Then
                        powtorka2 = 0
                        Exit For
                    End If
                    
                    powtorka2 = 1
                
                Next x
                
                If powtorka2 = 1 Then
                
                Exit For
                
                End If
            
            Next z
         
            If powtorka2 = 0 Then
            
                For e = 1 To wielkoscpodzbioru
                
                    podzbior(e - 1) = 0
                
                Next e
                ilosc = 0
                
                Exit For
            End If
            
           
         
         Next i
        
        
   Loop
        
   'wypisywanie i zapamietanie kombinacji bez powtorzen
   For i = 1 To wielkoscpodzbioru
        
        wynik = podzbior(i - 1)
        tmp(i - 1, j - 1) = podzbior(i - 1)
        przecinek = przecinek + 1
        Worksheets("Produkty").Activate
        
        
        wynikZ = vartab(wynik)
        Range("C" & q).Offset(0, i).Value = wynikZ

        wynik2 = "" & wynik2 & "" & wynik & ""
        wynikZ2 = "" & wynikZ2 & "" & wynikZ & ""
        
        If przecinek = wielkoscpodzbioru Then
            Range("A" & q).Value = q
            Range("B" & q).Value = wynik2
            Range("C" & q).Value = wynikZ2
            
            q = q + 1
            wynik2 = ""
            wynikZ2 = ""
        
        End If
        
   Next i
   
   'ponowne zerowanie podzbioru i przygotowanie do następnego generowania
   
   For i = 1 To wielkoscpodzbioru
   
   podzbior(i - 1) = 0
   
   Next i
   
Next j

End Function

