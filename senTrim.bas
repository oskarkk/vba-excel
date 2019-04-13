' zmienne troche globalne do daty poczatku arkusza,
' liczby odrzucanych wartosci ekstremalnych
' oraz rodzaju liczonej średniej (wolne/szkolne)
Private workbookStartDate As Double
Private trim As Integer
Private rodzajAvg As String




Sub senTrimmed()

' "ladujemy" modul stalych numerow potrzebnych kolumn z "Day"
klmn

' generowanie "raportów" w "MonthSen"
' tj. liczone obcięte średnie dla n miesięcy i wstawiane w wybrane kolumny
senTrimmedSummary 400, 15, "all", 3
senTrimmedSummary 400, 18, "L", 2
senTrimmedSummary 400, 21, "W", 2

End Sub




Sub senTrimmedSummary(daysBack, column, toRodzajAvg, toTrim)

' zeby bylo mniej argumentow nastepnych funkcji to ladujemy rodzaj sredniej
' i liczbe obcinanych wartosci do zmiennych troche globalnych
trim = toTrim
rodzajAvg = toRodzajAvg


' wylaczamy update ekranu by nie mulilo, choc akurat tutaj to chyba nic nie zmienia
Application.ScreenUpdating = False


' pobieramy datę początku arkusza z komórki D3
workbookStartDate = Worksheets(1).Cells(3, 4).Value


' znajdujemy ostatni z wierszy z miesiącami ("rowEnd") z jednej z kolumn "MonthSen"
' żeby wiedzieć gdzie jesteśmy czasowo
rowEnd = FindLastCellRow(5, 4, 1)
' pierwszy z wierszy z miesiącami ("rowStart"), czyli od ktorego miesiaca (do teraz) liczymy
rowStart = rowEnd - daysBack


' jesli liczba miesecy w tyl jest tak duza ze osiagamy gorna granice arkusza
' to zapobiegamy robieniu czegos tam gdzie juz nie ma danych
' co znacza ze jesli chce sie przeliczyc wszystkie miesiace
' to mozna podac jakas duza liczbe "daysBack"
' zamiast kminic ile jest miesiecy w arkuszu
If rowStart < 3 Then
    rowStart = 3
End If


' dla kazdego z miesiecy dla ktorych mamy liczyc srednia pobieramy daty graniczne miesiaca
' pobrane daty dajemy funkcji liczacej srednia i wyniki zapisujemy w trzech sąsiednich kolumnach "MonthSen"
For row = rowStart To rowEnd

    dateStart = Worksheets(4).Cells(row, 2).Value
    dateEnd = Worksheets(4).Cells(row + 1, 2).Value - 1
    
    srednie = trimmedAvgs(dateStart, dateEnd)
    
    Worksheets(4).Cells(row, column).Value = srednie(0)
    Worksheets(4).Cells(row, column + 1).Value = srednie(1)
    Worksheets(4).Cells(row, column + 2).Value = srednie(2)

Next row


' wlaczamy update ekranu
Application.ScreenUpdating = True

' pykniecie na koniec
Beep
    
End Sub




Function trimmedAvgs(dateStart, dateEnd) As Variant

' z pomoca daty poczatku arkusza liczymy numery wierszy ktore nas obchodza
' bez mozolnego porownywania daty kazdego wiersza z granicami miesiecy
rowOfFirstDay = 3 + dateStart - workbookStartDate
rowOfLastDay = 3 + dateEnd - workbookStartDate


' deklarujemy tablice do ktorych zaladujemy godziny
' zwrocmy uwage ze te tablice zaczynaja sie od 1
Dim wartosciZ() As Variant
ReDim Preserve wartosciZ(1 To 1)

Dim wartosciW() As Variant
ReDim Preserve wartosciW(1 To 1)

Dim wartosciS() As Variant
ReDim Preserve wartosciS(1 To 1)


' dla kazdego dnia miesiaca sprawdzamy czy to dzien wolny czy szkolny
' a potem uruchamiamy procedury biorace odpowiednie godziny z arkusza,
' dajac im te tablice, wszysktie dni danego miesiaca po kolei
' oraz stale numerow kolumn
For row = rowOfFirstDay To rowOfLastDay
    
    rodzajDnia = Worksheets(1).Cells(row, 5)
    
    getTimes wartosciZ, rodzajDnia, row, klmnSenZ
    getTimes wartosciW, rodzajDnia, row, klmnSenW
    getTimes wartosciS, rodzajDnia, row, klmnSenS
        
Next row

' zwracamy tablice zawierajaca wyniki funkcji obcinajacej tablice i liczacej srednia
trimmedAvgs = Array(trimDataAndAvg(wartosciZ), trimDataAndAvg(wartosciW), trimDataAndAvg(wartosciS))


End Function




Sub getTimes(wartosci As Variant, rodzajDnia, row, column)

' robimy sobie skrot do biezacej komorki
xCell = Worksheets(1).Cells(row, column)


' zaleznie od rodzaju sredniej wybieramy te komorki ktore sa z dni ktore nas interesuja i nie sa puste
' i jesli biezaca komorka pasuje, to wrzucamy godzine na ostatnie miejsce talbicy
' a potem powiekszamy tablice by zrobic miejsce na nastepna godzine
Select Case rodzajAvg
Case "all"
    If TypeName(xCell) = "Double" Then
        wartosci(UBound(wartosci)) = xCell
        ReDim Preserve wartosci(1 To UBound(wartosci) + 1)
    End If
Case "L"
    If (TypeName(xCell) = "Double") And (rodzajDnia = "L") Then
        wartosci(UBound(wartosci)) = xCell
        ReDim Preserve wartosci(1 To UBound(wartosci) + 1)
    End If
Case "W"
    If (TypeName(xCell) = "Double") And (rodzajDnia = "W") Then
        wartosci(UBound(wartosci)) = xCell
        ReDim Preserve wartosci(1 To UBound(wartosci) + 1)
    End If
End Select


End Sub




Function trimDataAndAvg(wartosci)


' jezeli w tablicy mamy tyle godzin że cokolwiek zostanie po obcieciu
If UBound(wartosci) > (2 * trim - 1) Then
    
    ' to zmniejszamy tablice, bo w ostatniej iteracji petli zrobilismy w tablicy
    ' niepotrzebne miejsce dla nastepnej godziny
    ReDim Preserve wartosci(1 To UBound(wartosci) - 1)

    ' sortujemy tablice kodem ze Stack Overflow xD
    QuickSort wartosci, LBound(wartosci), UBound(wartosci)
    
    ' robimy nowa tablice z miejscami na godziny ktore zostana po obcieciu
    Dim wartosciTrimmed() As Variant
    ReDim Preserve wartosciTrimmed(1 To UBound(wartosci) - trim * 2)
    
    ' i na kazde miejsce tej nowej tablicy wrzucamy po kolei godziny z posortowanej tablicy
    ' zaczynajac od "i + trim", czyli pomijajac wybrana liczbe najnizszych wartosci
    ' i konczac gdy konczy sie miejsce w tej nowej tablicy
    ' czyli pomijajac tez te sama wybrana liczbe ostatnich godzin w posortowanej tablicy
    For i = 1 To UBound(wartosciTrimmed)
        wartosciTrimmed(i) = wartosci(i + trim)
    Next i
    
    ' deklarujemy zmienna sumy z ktorej bedziemy liczyc srednia
    suma = 0
    
    ' i kazdy element tablicy z ucietymi srednimi dodajemy do tej sumy
    For Each x In wartosciTrimmed
        suma = suma + x
    Next
    
    ' liczymy i zwracamy srednia
    ' otrzymujemy ją dzielac sumę godzin przez liczbę wartości w uciętej tablicy
    trimDataAndAvg = suma / UBound(wartosciTrimmed)

' a jeżeli godzin pobranych z arkusza jest za mało, to zwracamy nic, nie ma sredniej, mowi sie trudno
Else

    trimDataAndAvg = ""

End If


End Function



' funkcja sprawdza do ktorego wiersza sa dane w wybranej kolumnie "column" w arkuszu "Worksheet"
' można wybrać wiersz poczatkowy "startRow", by ominąć puste wiersze
' np. nagłówkowe albo po prostu jakieś zjebane puste miejsca
Function FindLastCellRow(column, Worksheet, Optional startRow = 1) As Double

' ustalamy ze sprawdzanie zaczynamy od tego wybranego wiersza startowego
' TODO: sprawdzic czy bez podania argumentu "startRow" to wszystko sie nie sypie
' a jak sypie to naprawic
row = startRow

' ladujemy sobie pierwsza rozpatrywana komorke
Dim CurrentCell As Range
Set CurrentCell = Worksheets(Worksheet).Cells(row, column)

' i sprawdzamy czy nie jest pusta
While Not (TypeName(CurrentCell.Value) = "Empty")
    ' jesli ma zawartosc to ladujemy komorke z nastepnego wiersza
    row = row + 1
    Set CurrentCell = Worksheets(Worksheet).Cells(row, column)
    ' i petla dopoki nie znajdziemy czegoś pustego
Wend

' zwracamy numer ostatniego wiersza z niepusta komorka w wybranej kolumnie
FindLastCellRow = row - 1

End Function

