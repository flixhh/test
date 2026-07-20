Attribute VB_Name = "ImportCsvMacro"
'==============================================================================
' Modul:   ImportCsvMacro
' Zweck:   Liest CSV-Dateien aus dem "aktuellsten" Timestamp-Ordner und kopiert
'          deren Inhalte in vorhandene Tabellenblaetter der Arbeitsmappe.
'
' Ablauf:
'   1. Im Basisordner (BASE_FOLDER) werden alle Unterordner betrachtet, deren
'      Name einem Timestamp-Format entspricht. Der "aktuellste" (groesste) Name
'      wird ausgewaehlt.
'   2. Fuer jede CSV-Datei in diesem Ordner wird anhand des Dateinamens (ohne
'      ".csv") das passende Tabellenblatt bestimmt.
'   3. Der bisherige Inhalt des Tabellenblatts wird geloescht.
'   4. Die Inhalte der CSV werden eingefuegt.
'   5. Schleife ueber alle CSV-Dateien.
'   6. Im Tabellenblatt "Log" wird eine Zeile mit Ausfuehrungszeitpunkt und
'      dem Timestamp des Ordners angehaengt.
'
' Konfiguration: Siehe Konstanten im Abschnitt "Einstellungen".
'
' Empfohlenes Timestamp-Format fuer die Ordnernamen (inkl. Uhrzeit):
'   YYYY-MM-DD_HH-MM-SS     z. B. 2026-06-30_14-05-30
'   In VBA erzeugbar mit:   Format(Now, "yyyy-mm-dd_hh-nn-ss")
'   (Hinweis: in VBA ist "nn" = Minuten, "mm" = Monat.)
' Wichtig: Stellen muessen von gross nach klein angeordnet sein (Jahr -> Sekunde),
' damit der alphabetische Vergleich = chronologische Reihenfolge ist und der
' "aktuellste" Ordner korrekt erkannt wird. Ebenfalls moeglich:
'   YYYYMMDD_HHMMSS         z. B. 20260630_140530   (kompakt)
'   YYYY-MM-DD_HH-MM-SS-fff z. B. 2026-06-30_14-05-30-123 (mit Millisekunden)
' Nicht verwenden: DD.MM.YYYY oder MM-DD-YYYY -> sortieren nicht chronologisch.
'==============================================================================
Option Explicit

'------------------------------- Einstellungen --------------------------------
' Basisordner, der die Timestamp-Unterordner enthaelt.
' Tipp: Leer lassen ("") um den Ordner der aktuellen Arbeitsmappe zu verwenden.
Private Const BASE_FOLDER As String = ""

' Trennzeichen der CSV-Dateien. In deutschen Excel-Umgebungen meist ";".
' Auf "" setzen, um das Trennzeichen automatisch zu erkennen (; , Tab).
Private Const CSV_DELIMITER As String = ";"

' Name des Log-Tabellenblatts.
Private Const LOG_SHEET_NAME As String = "Log"

' Wenn True: fehlt ein passendes Tabellenblatt, wird die Datei uebersprungen.
' Wenn False: es wird ein neues Tabellenblatt mit dem Namen angelegt.
Private Const SKIP_IF_SHEET_MISSING As Boolean = True
'------------------------------------------------------------------------------


'==============================================================================
' Haupt-Makro: per Button/Alt+F8 aufrufen.
'==============================================================================
Public Sub ImportiereCsvDateien()
    Dim wb As Workbook
    Dim fso As Object
    Dim baseFolderPath As String
    Dim latestFolder As Object
    Dim latestTimestamp As String
    Dim file As Object
    Dim sheetName As String
    Dim ws As Worksheet
    Dim importedCount As Long
    Dim skippedCount As Long
    Dim startTime As Date
    Dim details As String
    Dim reason As String
    Dim hint As String

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    startTime = Now

    ' --- Basisordner bestimmen --------------------------------------------
    baseFolderPath = BASE_FOLDER
    If Len(Trim$(baseFolderPath)) = 0 Then baseFolderPath = wb.Path
    If Len(Trim$(baseFolderPath)) = 0 Then
        MsgBox "Kein Basisordner gesetzt und die Arbeitsmappe wurde noch nicht gespeichert." & vbCrLf & _
               "Bitte BASE_FOLDER im Modul setzen.", vbExclamation, "CSV-Import"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseFolderPath) Then
        MsgBox "Basisordner nicht gefunden:" & vbCrLf & baseFolderPath, vbCritical, "CSV-Import"
        Exit Sub
    End If

    ' --- Aktuellsten Timestamp-Ordner suchen ------------------------------
    Set latestFolder = GetLatestTimestampFolder(fso, baseFolderPath)
    If latestFolder Is Nothing Then
        MsgBox "Im Basisordner wurde kein Timestamp-Unterordner gefunden:" & vbCrLf & baseFolderPath, _
               vbExclamation, "CSV-Import"
        Exit Sub
    End If
    latestTimestamp = latestFolder.Name

    ' --- Performance: Bildschirm/Berechnung anhalten ----------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Schleife ueber alle CSV-Dateien ----------------------------------
    importedCount = 0
    skippedCount = 0
    details = ""
    For Each file In latestFolder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "csv" Then
            sheetName = fso.GetBaseName(file.Name)   ' Dateiname ohne ".csv"

            If SheetExists(wb, sheetName) Then
                ' Blatt existiert -> leeren und importieren.
                Set ws = wb.Worksheets(sheetName)
                reason = ""
                If TryImport(file.Path, ws, reason) Then
                    importedCount = importedCount + 1
                Else
                    skippedCount = skippedCount + 1
                    details = details & "- " & file.Name & " -> Blatt '" & sheetName & _
                              "': " & reason & vbLf
                End If

            ElseIf Not SKIP_IF_SHEET_MISSING Then
                ' Blatt fehlt, soll aber angelegt werden.
                Set ws = GetOrCreateSheet(wb, sheetName)
                reason = ""
                If Not ws Is Nothing Then
                    If TryImport(file.Path, ws, reason) Then
                        importedCount = importedCount + 1
                    Else
                        skippedCount = skippedCount + 1
                        details = details & "- " & file.Name & " -> neues Blatt '" & _
                                  sheetName & "': " & reason & vbLf
                    End If
                Else
                    skippedCount = skippedCount + 1
                    details = details & "- " & file.Name & " -> Blatt '" & sheetName & _
                              "' konnte nicht angelegt werden" & vbLf
                End If

            Else
                ' Blatt fehlt und wird uebersprungen -> Grund + Hinweis protokollieren.
                skippedCount = skippedCount + 1
                hint = FindSheetHint(wb, sheetName)
                details = details & "- " & file.Name & " -> kein Tabellenblatt '" & _
                          sheetName & "' vorhanden (uebersprungen)"
                If Len(hint) > 0 Then
                    details = details & " | Hinweis: aehnliches Blatt '" & hint & _
                              "' gefunden - Name/Leerzeichen/Gross-Kleinschreibung pruefen"
                End If
                details = details & vbLf
            End If
        End If
    Next file

    ' --- Log-Zeile schreiben ----------------------------------------------
    WriteLog wb, startTime, latestTimestamp, importedCount, skippedCount, details

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "CSV-Import abgeschlossen." & vbCrLf & vbCrLf & _
           "Ordner (Timestamp): " & latestTimestamp & vbCrLf & _
           "Importiert: " & importedCount & " Datei(en)" & vbCrLf & _
           "Uebersprungen: " & skippedCount & " Datei(en)" & _
           IIf(Len(details) > 0, vbCrLf & vbCrLf & "Details:" & vbCrLf & details, ""), _
           vbInformation, "CSV-Import"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Fehler beim CSV-Import:" & vbCrLf & _
           "Nr. " & Err.Number & " - " & Err.Description, vbCritical, "CSV-Import"
End Sub


'==============================================================================
' Sucht im Basisordner den "aktuellsten" Timestamp-Unterordner.
' "Aktuellster" = groesster Ordnername gemaess Textvergleich. Bei sortierbaren
' Timestamp-Formaten (z. B. YYYYMMDD_HHMMSS oder YYYY-MM-DD_HH-MM-SS) entspricht
' das dem neuesten Zeitpunkt.
'==============================================================================
Private Function GetLatestTimestampFolder(ByVal fso As Object, ByVal baseFolderPath As String) As Object
    Dim subFolder As Object
    Dim best As Object

    Set best = Nothing
    For Each subFolder In fso.GetFolder(baseFolderPath).SubFolders
        If LooksLikeTimestamp(subFolder.Name) Then
            If best Is Nothing Then
                Set best = subFolder
            ElseIf StrComp(subFolder.Name, best.Name, vbTextCompare) > 0 Then
                Set best = subFolder
            End If
        End If
    Next subFolder

    Set GetLatestTimestampFolder = best
End Function


'==============================================================================
' Prueft heuristisch, ob ein Ordnername ein Timestamp ist:
' beginnt mit einer Ziffer (Jahr), enthaelt nur Ziffern und gaengige
' Trennzeichen (- _ . Leerzeichen :) und hat insgesamt mindestens 8 Ziffern
' (Datum YYYYMMDD). Trennzeichen sind erlaubt, daher werden sowohl
' "20260701_143106" als auch "2026-07-01_14-31-06" erkannt.
'==============================================================================
Private Function LooksLikeTimestamp(ByVal name As String) As Boolean
    Dim i As Long, ch As String, digitCount As Long

    If Len(name) = 0 Then Exit Function

    ' Muss mit einer Ziffer beginnen (Jahr).
    ch = Left$(name, 1)
    If ch < "0" Or ch > "9" Then Exit Function

    digitCount = 0
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        If ch >= "0" And ch <= "9" Then
            digitCount = digitCount + 1
        ElseIf InStr("-_. :", ch) > 0 Then
            ' erlaubtes Trennzeichen -> ok
        Else
            LooksLikeTimestamp = False
            Exit Function
        End If
    Next i

    ' Insgesamt genug Ziffern fuer mindestens ein Datum (YYYYMMDD = 8).
    LooksLikeTimestamp = (digitCount >= 8)
End Function


'==============================================================================
' Liefert das Tabellenblatt mit dem gewuenschten Namen. Existiert es nicht,
' wird es entweder uebersprungen (Nothing) oder neu angelegt.
'==============================================================================
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim safeName As String

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing And Not SKIP_IF_SHEET_MISSING Then
        safeName = SanitizeSheetName(sheetName)
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = safeName
    End If

    Set GetOrCreateSheet = ws
End Function


'==============================================================================
' Prueft, ob ein Tabellenblatt mit exakt diesem Namen existiert.
'==============================================================================
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function


'==============================================================================
' Versucht, die CSV in das Blatt zu importieren. Bei Erfolg True; bei einem
' Fehler False, wobei "reason" die Fehlerbeschreibung enthaelt. So wird ein
' einzelner Fehler protokolliert, ohne das gesamte Makro abzubrechen.
'==============================================================================
Private Function TryImport(ByVal filePath As String, ByVal ws As Worksheet, _
                           ByRef reason As String) As Boolean
    On Error GoTo Failed
    ws.Cells.Clear                       ' bisherigen Inhalt loeschen
    ImportCsvIntoSheet filePath, ws      ' CSV-Inhalt einfuegen
    TryImport = True
    Exit Function
Failed:
    reason = "Fehler beim Import (Nr. " & Err.Number & " - " & Err.Description & ")"
    TryImport = False
End Function


'==============================================================================
' Sucht ein Blatt, dessen Name dem gesuchten Namen "aehnlich" ist (gleich nach
' Trimmen und ohne Beachtung der Gross-/Kleinschreibung). Dient als Hinweis bei
' typischen Ursachen fuer uebersprungene Dateien (Leerzeichen, Schreibweise).
' Liefert den tatsaechlichen Blattnamen oder "" wenn keiner passt.
'==============================================================================
Private Function FindSheetHint(ByVal wb As Workbook, ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim target As String
    target = LCase$(Trim$(sheetName))
    For Each ws In wb.Worksheets
        If LCase$(Trim$(ws.Name)) = target Then
            FindSheetHint = ws.Name
            Exit Function
        End If
    Next ws
    FindSheetHint = ""
End Function


'==============================================================================
' Entfernt fuer Tabellenblattnamen unzulaessige Zeichen und kuerzt auf 31 Zeichen.
'==============================================================================
Private Function SanitizeSheetName(ByVal name As String) As String
    Dim invalid As Variant, ch As Variant, result As String
    result = name
    invalid = Array("\", "/", "?", "*", "[", "]", ":")
    For Each ch In invalid
        result = Replace(result, CStr(ch), "_")
    Next ch
    If Len(result) = 0 Then result = "Sheet"
    If Len(result) > 31 Then result = Left$(result, 31)
    SanitizeSheetName = result
End Function


'==============================================================================
' Liest eine CSV-Datei und schreibt deren Inhalt ab Zelle A1 in das Tabellenblatt.
' Unterstuetzt in Anfuehrungszeichen eingeschlossene Felder mit Trennzeichen,
' Zeilenumbruechen und doppelten Anfuehrungszeichen ("").
'==============================================================================
Private Sub ImportCsvIntoSheet(ByVal filePath As String, ByVal ws As Worksheet)
    Dim content As String
    Dim delim As String
    Dim rows As Collection
    Dim rowFields As Variant
    Dim r As Long, c As Long
    Dim maxCols As Long
    Dim outArr() As Variant
    Dim line As Variant

    content = ReadTextFile(filePath)
    If Len(content) = 0 Then Exit Sub

    delim = CSV_DELIMITER
    If Len(delim) = 0 Then delim = DetectDelimiter(content)

    Set rows = ParseCsv(content, delim)
    If rows.Count = 0 Then Exit Sub

    ' Maximale Spaltenzahl ermitteln.
    maxCols = 0
    For Each line In rows
        If UBound(line) + 1 > maxCols Then maxCols = UBound(line) + 1
    Next line
    If maxCols = 0 Then Exit Sub

    ' Ausgabe-Array fuellen (in einem Rutsch -> schnell).
    ReDim outArr(1 To rows.Count, 1 To maxCols)
    For r = 1 To rows.Count
        rowFields = rows(r)
        For c = 0 To UBound(rowFields)
            outArr(r, c + 1) = rowFields(c)
        Next c
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(rows.Count, maxCols)).Value = outArr
End Sub


'==============================================================================
' Liest eine Textdatei (CSV) ein. Versucht UTF-8 zu erkennen, faellt sonst auf
' ANSI zurueck.
'==============================================================================
Private Function ReadTextFile(ByVal filePath As String) As String
    Dim stream As Object
    Dim text As String

    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText(-1)
    stream.Close

    ' BOM entfernen, falls vorhanden.
    If Len(text) > 0 Then
        If AscW(Left$(text, 1)) = &HFEFF Then text = Mid$(text, 2)
    End If
    ReadTextFile = text
    Exit Function

Fallback:
    ' Klassisches Einlesen als ANSI.
    Dim fnum As Integer
    Dim raw As String
    fnum = FreeFile
    Open filePath For Input As #fnum
    raw = Input$(LOF(fnum), fnum)
    Close #fnum
    ReadTextFile = raw
End Function


'==============================================================================
' Erkennt das wahrscheinlichste Trennzeichen aus der ersten Datenzeile.
'==============================================================================
Private Function DetectDelimiter(ByVal content As String) As String
    Dim firstLine As String, p As Long
    p = InStr(content, vbLf)
    If p > 0 Then firstLine = Left$(content, p - 1) Else firstLine = content
    firstLine = Replace(firstLine, vbCr, "")

    Dim semi As Long, comma As Long, tab As Long
    semi = CountChar(firstLine, ";")
    comma = CountChar(firstLine, ",")
    tab = CountChar(firstLine, vbTab)

    If semi >= comma And semi >= tab And semi > 0 Then
        DetectDelimiter = ";"
    ElseIf tab >= comma And tab > 0 Then
        DetectDelimiter = vbTab
    ElseIf comma > 0 Then
        DetectDelimiter = ","
    Else
        DetectDelimiter = ";"   ' Standard
    End If
End Function


Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    Dim pos As Long, n As Long
    pos = InStr(1, s, ch)
    Do While pos > 0
        n = n + 1
        pos = InStr(pos + 1, s, ch)
    Loop
    CountChar = n
End Function


'==============================================================================
' Parst CSV-Inhalt zu einer Collection von Zeilen. Jede Zeile ist ein
' 0-basiertes Array von Feldern. Beruecksichtigt Anfuehrungszeichen.
'==============================================================================
Private Function ParseCsv(ByVal content As String, ByVal delim As String) As Collection
    Dim result As New Collection
    Dim fields As Collection
    Dim field As String
    Dim i As Long, n As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim delimCh As String

    delimCh = delim
    n = Len(content)
    Set fields = New Collection
    field = ""
    inQuotes = False

    i = 1
    Do While i <= n
        ch = Mid$(content, i, 1)

        If inQuotes Then
            If ch = """" Then
                If i < n And Mid$(content, i + 1, 1) = """" Then
                    field = field & """"      ' doppeltes Anfuehrungszeichen -> "
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                field = field & ch
            End If
        Else
            If ch = """" Then
                inQuotes = True
            ElseIf ch = delimCh Then
                fields.Add field
                field = ""
            ElseIf ch = vbCr Then
                ' ignorieren; Zeilenende ueber vbLf behandeln
            ElseIf ch = vbLf Then
                fields.Add field
                result.Add CollectionToArray(fields)
                Set fields = New Collection
                field = ""
            Else
                field = field & ch
            End If
        End If

        i = i + 1
    Loop

    ' Letzte Zeile (falls Datei nicht mit Zeilenumbruch endet).
    If Len(field) > 0 Or fields.Count > 0 Then
        fields.Add field
        result.Add CollectionToArray(fields)
    End If

    Set ParseCsv = result
End Function


'==============================================================================
' Wandelt eine Collection von Feldern in ein 0-basiertes Array um.
'==============================================================================
Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim arr() As String
    Dim i As Long
    If c.Count = 0 Then
        ReDim arr(0 To 0)
        arr(0) = ""
    Else
        ReDim arr(0 To c.Count - 1)
        For i = 1 To c.Count
            arr(i - 1) = c(i)
        Next i
    End If
    CollectionToArray = arr
End Function


'==============================================================================
' Haengt im Log-Tabellenblatt eine Zeile an:
'   Ausfuehrungszeitpunkt | Timestamp des Ordners | importiert | uebersprungen
' Existiert das Log-Blatt nicht, wird es mit Kopfzeile angelegt.
'==============================================================================
Private Sub WriteLog(ByVal wb As Workbook, ByVal execTime As Date, _
                     ByVal folderTimestamp As String, _
                     ByVal importedCount As Long, ByVal skippedCount As Long, _
                     ByVal details As String)
    Dim logWs As Worksheet
    Dim nextRow As Long
    Dim detailText As String

    On Error Resume Next
    Set logWs = wb.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If logWs Is Nothing Then
        Set logWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        logWs.Name = LOG_SHEET_NAME
        logWs.Range("A1").Value = "Ausfuehrungszeitpunkt"
        logWs.Range("B1").Value = "Ordner-Timestamp"
        logWs.Range("C1").Value = "Importierte Dateien"
        logWs.Range("D1").Value = "Uebersprungene Dateien"
        logWs.Range("E1").Value = "Details (uebersprungen / Fehler)"
        logWs.Range("A1:E1").Font.Bold = True
    End If

    nextRow = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ' Zeilenumbrueche fuer die Zelle vereinheitlichen und Endezeichen entfernen.
    detailText = details
    Do While Len(detailText) > 0 And Right$(detailText, 1) = vbLf
        detailText = Left$(detailText, Len(detailText) - 1)
    Loop
    If Len(detailText) = 0 Then detailText = "OK - alle Dateien importiert"

    logWs.Cells(nextRow, "A").Value = execTime
    logWs.Cells(nextRow, "A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    logWs.Cells(nextRow, "B").Value = folderTimestamp
    logWs.Cells(nextRow, "C").Value = importedCount
    logWs.Cells(nextRow, "D").Value = skippedCount
    logWs.Cells(nextRow, "E").Value = detailText
    logWs.Cells(nextRow, "E").WrapText = True
End Sub
