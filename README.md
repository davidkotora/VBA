# VBA - Základné funkcie automatizácie (Vytvorenie hlavičky tabuľky + formátovanie)
Označenie danej oblasti
Range("").Select
Napísanie hodnoty (Value) do danej oblasti
Range("").Value
Označenie pracovného hárku
Worksheets("").Select
Označenie pracovného hárku + oblasti
Worksheets("").Range("").Select
Zapísanie hodnoty do pracovného hárku + oblasti
Worksheets("").Range("").Value
             Formátovanie textu =
Rem format - font.color pre hlavicku
Range("A1:G1").Font.Color = RGB(40, 180, 50)
Rem format - Tučné písmo
Range("A1:G1").Font.Bold = True
Rem format - veľkosť písma
Range("A1:G1").Font.Size = 18
Rem format - Podfarbenie písma
Range("A1:G1").Interior.Color = RGB(0, 0, 0)
Rem pridá orámovanie podľa zadania
Range("A1:G20").Borders.LineStyle = True
Range("A21:G27").Interior.Color = RGB(245, 245, 220)
