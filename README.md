# Ordini Bloccati Macro Excel VBA
Questo progetto descrive l'automazione della reportistica degli ordini bloccati nell'ambito delle attività di recupero credito di un'azienda del settore farmaceutico

``` vb

Private Sub Orders_Blocked()
''Private prima di Sub per non fare apparire la macro nell'elenco delle macro ( accessibile con F8)
'' Macro VBA per Elaborazione ordini bloccati per assenza pagamenti
'' Prima fase: preparare il foglio all'inserimento dei commenti
''la struttura dei fogli deve rimanere la stessa

''All'avvio della macro il file OB_C.xlsx dovra essere aperto!!!

Dim infos As Variant  '' Dichiarazione della variabile infos a cui verrà assegnato la scelta dell'utente
                        ''se eseguire o no la Macro: Utile nel caso viene premuto inavertitamente
''
    ''Msgbox per fornire infos all'utente di cio' che succedera' una volta che avra premuto su ok
    infos = MsgBox("Elaborazione file Ordini Bloccati..." & vbNewLine & vbNewLine & _
    "Salvare Prima!!! ==>> CTRL-Z non funziona!" & vbNewLine & vbNewLine & "Accertarsi salvataggio File necessari nel seguente percorso." _
    & vbCrLf & "T:\CONTABILITA'\RECUPERO CREDITI\macraris_kl\MacrAris\Orders_Blocked_Macro " & vbCrLf & vbCrLf _
    & "Pregasi NON ALTERARE il nome dei File" _
    & vbNewLine & vbNewLine & "" & vbNewLine & "Qui per sbaglio -->  Click 'NO'", _
                    vbYesNo + vbInformation + vbDefaultButton2, "Macr@ris Ordini Bloccati")
                    
If infos = vbNo Then ''Se l'utente avrà cliccato su NO allora l'istruzione successiva verra eseguita quindi verrà interrota
                        ''la macro. l'esecuzione successiva sara' quindi Exit Sub equivale Esci dalla routine
    Exit Sub '' Non esegue la macro in quanto l'utente ha lanciato la macro per errore
    End If '' Fine esecuzione Macro
'
''
''@@@ ESECUZIONE PROCEDURA CREAZIONE CARTELLE ANNO SEGUENTE
            creazioneCartelleYr1_ordersblock     '' Macro nidificata in un'altra: richiama un'altra macro ( codici vba in indice)
                                                    ''Utile a fine Anno Dicembre: crea automaticamente la cartella relativa
                                                    ''al nuovo anno di riferimento in cui salvare i file elaborati
''
''####

    Dim wOb_C As Workbook  '' Dichiara una variabile di tipo cartella
         Set wOb_C = ActiveWorkbook  ''Assegna ed inizializza la cartella attiva alla variabile wOb_C
         

If wOb_C.Name <> "OB_C.xlsx" Then   ''controlla che il file attivo sia effettivamente quello di interesse e se negativo
                                        '' visualizza un messaggio informativo e poi interruzione. E' utile includere la
                                        '' la gestione di errori prevedibili.
    MsgBox "File Excel NON Corretto!" & _
                                            vbCrLf & vbCrLf & "verifica che file attivo sia OB_C.xlsx" & _
                                            vbCrLf & vbCrLf & "Interuzione Macro senza Alcuna Conseguenza!", _
                                            vbCritical, "Macr@ris messaggio di errore"
    Exit Sub ''Fine esecuzione Macro
End If

''Macro tracking time  '' Dichiarazione variabili per rilevare la durata di elaborazione della Macro
Dim triggerChrono As Date, endtriggerChrono As Date, Interval As Date, strOutput As String
''Web source data
''http://msdn.microsoft.com/en-us/library/office/ff197413(v=office.15).aspx

triggerChrono = Now  '' alla Variabile triggerChrono viene assegnata l'ora al momento in cui viene eseguita l'attuale riga
''

''Messaggio all'utente dell'esecuzione in corso della Macro nella barra di stato
Application.StatusBar = "Elaborazione Dati in Corso... Un PO' di Relax!"

Application.ScreenUpdating = False '' Movimenti dello schermo possono rallentare l'esecuzione delle macro
                                    ''Il valore False disattiva i movimenti dello schermo

On Error GoTo ErrorHandler '' Ad ogni errore l'esecuzione del codice verrà rinviato a ErrorHandler dove ci sono delle istruzioni sul
                                '' su cosa e' accaduto e suggerimenti su cosa fare

Columns("A:A").EntireColumn.AutoFit ''auto adattamento larghezza colonna A
[b1].Value = "Valore Ordini"    '' nuovo valore della cella B1
    Range("B1").Font.Bold = True ''Applica il grassetto alla cella B1
        Columns("B:B").EntireColumn.AutoFit '' auto adattamento larghezza colonna B
    
    [H1].Value = "Rag. Sociale"  '' inserisce il testo tra virgolette nella cella H1
    [T1].Value = "Rif. Ord. Cliente"

''Applica larghezza fissa alle colonne indicate
    Columns("C:C").ColumnWidth = 1.5
    Columns("D:D").ColumnWidth = 4.33
    Columns("E:E").ColumnWidth = 2.17
    Columns("F:F").ColumnWidth = 2.5
    
    [G:G].Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove  ''inserisce una colonna in posizione G
       Columns("G:G").ColumnWidth = 55
       
         With Range("G1")  '' con la struttura "With...End With" attribuisce una serie di caratteristiche alla cella G1
              .Value = "Azioni / Commenti"
              .Font.Bold = True
              .Font.Name = "Arial"
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlCenter
              .Font.Size = 12
          End With
    
    Columns("K:L").NumberFormat = "dd-mmm-yy"  '' applica il formato date indicato alle colonne K e L
    
[U:U].Cut '' Rimuove la colonna U
[N:N].Insert Shift:=xlToRight '' inserisce la colonna U prima della colonna N spostandola verso destra
       
Columns("O:Z").Delete Shift:=xlToLeft '' Elimina le colonne O a Z spostando le altre verso sinistra
    
    
    [H:N].EntireColumn.AutoFit ''Applica adatta colonna alle seguenti colonne
    
    Dim shOb_C As Worksheet '' Dichiarazione di una variabile come oggetto foglio

    Set shOb_C = wOb_C.Sheets(1) '' Assegna il foglio 1 alla variabile foglio dichiarata alla riga precedente
    
    shOb_C.Name = "Ordini Bloccati-clienti arancio" '' Rinomina il foglio rappresentato da shOb_C
      
'       '@Adding Finservice Sheet
  '@ Ciclo For ... Each per eseguire un stesso gruppo di commando su tutti i fogli nella cartella di lavoro
        Dim CrSh As Worksheet '' Dichiarazione di una variabile come oggetto Foglio
    Application.DisplayAlerts = False   '' Disattiva tutti messaggi d'avvertimento quando si elimina un oggetto
        
        For Each CrSh In Worksheets '' ciclo per esecuzione su tutti i fogli
            If CrSh.Name <> shOb_C.Name Then ''Condizione IF da controllare per ogni ciclo
            CrSh.Delete ''Eliminazione foglio se nome non rispecchia shOb_C.Name  = "Ordini Bloccati-clienti arancio"
            End If '' fine esecuzione condizione
        Next  '' Riporta il codice di esecuzione a For per esecuzione gruppo di codici IF...End IF sul foglio successivo
        
    Application.DisplayAlerts = True '' Riattiva i messaggi di avvertimento a seguito eliminazione oggetto
    
    ''Matrice Aris Cerca verticale su file compensi collectors
            
 Dim Cy As String
    Cy = Format(Now, "YYYY")
    
 Dim wCompensiColl As Workbook   ''Dichiarazione di una variabile oggetto Cartella
        ''apertura di un foglio di lavoro e attribuzione di quel foglio alla variabile tipo foglio appena creata
        Set wCompensiColl = Workbooks.Open(Filename:= _
                "T:\CONTABILITA'\RECUPERO CREDITI\Credit_Collectors\Contratti_Credit_Collectors\" & _
                         Cy & "\tabella_compensi_collectors.xlsx")

Dim ShCompensiC As Worksheet '' Dichiarazione di una variabile di tipo foglio
   Set ShCompensiC = wCompensiColl.Sheets("BBMI_Priv_Collectors") '' Assegnazione del foglio BBMI_PRIV_COLLECTORS alla variabile di tipo foglio

shOb_C.Activate '' attiva il foglio denominato "Ordini Bloccati-clienti arancio" tramite la variabile shOb_C

''In questo blocco si difinisce la matrice virtuale in cui copiare i dati per eseguire  la ricerca verticale
''la creazione di una matrice virtuale rende piu' veloce l'elaborazione
Dim vAIndexCliente As Variant, vaNameColl As Variant, avlookup As Variant, avResult() As Variant

With ShCompensiC ''Uso di With....end come scorciatioa per assegnare i valori.... alla matrice nominata VaNameColl
'    vaNameColl = .Range(.Cells(Rows.Count, "A").End(xlUp), "I2")
    vaNameColl = .Range(.Cells(.Range("A1").SpecialCells(xlCellTypeLastCell).Row, "A"), "I2")
End With

Application.DisplayAlerts = False
    wCompensiColl.Close ''Chiusura della cartella
Application.DisplayAlerts = True

'' avlookup = Range(Cells(Rows.Count, "H").End(xlUp), "H2")''Questa alternativa eliminata perche rendeva il file molto pesante e voluminoso
    avlookup = Range(Cells(Range("H1").SpecialCells(xlCellTypeLastCell).Row - 1, "H"), "H2") ' Selezione dinamica dell'intervallo di dati H2 e
                                                                                            ''l'ultima cella dell'intervallo che contiene i dati.
                                                                                            ''notare l'uso "Vai a formato speciale Ultima cella.
                                                                                            ''identifica la riga e meno 1 per avere il numero di riga che contiene
                                                                                            ''l'ultimo dato
                                                                                            ''Assegna poi l'intervallo di dati selezionati alla matrice "avLookup"
 
    ReDim avResult(1 To UBound(avlookup, 1), 1 To 1) '' Con "ReDim" crea un intervallo di dati in matrice di n righe, 1 colonna della stessa
                                                        '' dimensione della matrice "AvLookup"
      

For i = 1 To UBound(avlookup, 1) ''ciclo di ripetizione con limite di esecuzione n esima riga della matrice avlookup

    On Error Resume Next   ''ignorare eventuali errori generati nell'esecuzione della macro specie quando il risultato della
                        ''formula CERCA.VERT restituisce un #N/D
    
    avResult(i, 1) = WorksheetFunction.VLookup(avlookup(i, 1) * 1, vaNameColl, 9, 0)  '' CERCA.VERT del dato in riga i nell'intervallo di dati VanameColl
        If Err.Number = 1004 Then avResult(i, 1) = "Mombrini"  '' se errore di tipo 1004 allora risultato CERCA.VERT = #N/D quindi sostituisci col Nome

Next i

On Error GoTo ErrorHandler '' Ripristina la gestione degli errori generici definita per l'insieme della Macro
    [O2].Resize(UBound(avlookup, 1), 1).Value = avResult  ''copia i risultati della ricerca verticale nell'intervallo limite inferiore cella O2
                                                        '' e limite superiore n righe della matrice Avlookup
    
With Range("N1")
    .Copy [O1]
        With .Offset(0, 1)
            .ClearContents
            .Value = "Collectors"
        End With
    
End With
  
       
    ''Arrayaris fine cerca vert su compensi collectors
    
 '' cerca vert on previous file orders block
 
     ''Arrayaris Cerca verticale file bloccati settimana precedente
    
 Dim wOb_P As Workbook ''Defizione di variabile oggetto di tipo cartella
   Set wOb_P = Workbooks.Open(Filename:="T:\CONTABILITA'\RECUPERO CREDITI\macraris_kl\MacrAris\Orders_Blocked_Macro\OB_P.xlsx") ''apre un cartella specifica di nome.... e
''assegna il tutto alla variabile inizializzata cartella "wOb_P

Dim shOb_P As Worksheet ''Defizione di variabile oggetto di tipo foglio
   Set shOb_P = wOb_P.Sheets("Ordini Bloccati-clienti arancio") ''Assegna il foglio denominato .... alla variabile "shOb_P

Columns("H:H").Cut '' taglia la colonna
    Columns("G:G").Insert Shift:=xlToRight '' inserisce la colonna h tagliata e sposta la colonna G verso destra

With shOb_P  ''
.Range(.Cells(.Range("G1").SpecialCells(xlCellTypeLastCell).Row, "G"), "G2").Select ''Seleziona l'intervallo di dati utile nella colonna G2
End With
         Selection.TextToColumns Destination:=Range("G2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 2), TrailingMinusNumbers:=True                             ' applica la funzionalita testo in colonne per trasformare i dati in formato testo
        
shOb_C.Activate ''Attiva il foglio assegnato alla variabile shOb_C

wOb_P.Sheets(Array("Finservice_Affidati", "clienti a contenz o pro concors", _
            "privati payer")).Copy After:=shOb_C                    ' Copia fogli dalla cartella (file) settimana precedente dentro la corrente cartella
 
 shOb_C.Activate   ''attiva il foglio ; vedi il foglio a cui e' stato assegnato la variabile

With shOb_P
    vaNameColl = .Range(.Cells(.Range("G1").SpecialCells(xlCellTypeLastCell).Row, "G"), "H2") ''Nel file settimana precedente assegna l'intervallo di dati
                    '' tra nelle colonne G a H alla matrice vaNameColl
End With


Application.DisplayAlerts = False   ''disattiva le notitiche di Excel
    wOb_P.Close     ''chiude la cartella di lavoro settimana precedente senza salvare
Application.DisplayAlerts = True    ''Riattiva le notifiche di Excel

'
' avlookup = Range(Cells(Rows.Count, "H").End(xlUp), "H2")
'    ReDim avResult(1 To UBound(avlookup, 1), 1 To 1)
      
For i = 1 To UBound(avlookup, 1)
    On Error Resume Next ''Ignora eventuali errori
    
    avResult(i, 1) = WorksheetFunction.VLookup(avlookup(i, 1), vaNameColl, 2, 0)  ''Cerca verticale sulla matrice di dati per copiare i commenti presenti
        If Err.Number = 1004 Then avResult(i, 1) = CVErr(xlErrNA)                   ''nel file della settimana precedente. Notare prima di chiudere il file
                                                                                    '' Tali valori sono stati assegnati alla matrice vaNameColl. Ove possibile
                                                                                    ''preferire operazioni sulle matrici al posto delle operazioni sulle celle di Excel
                                                                                    ''in quanto la velocita di elaborazione e' di 10+
                                                                                    ''Quando riscontra un errore di tipo 1004 con CVErr(xlErrNA)  assegna il valore #N/D

Next i

On Error GoTo ErrorHandler  ''Ripristina la gestione di errori generici
    [G2].Resize(UBound(avlookup, 1), 1).Value = avResult   ''Restituisce i risultati della ricerca verticale dalla cella G2 in poi.
                                                            '' I risultati sono presi dalla matrice avResult
       
    ''### Arrayaris fine cerca vert su file bloccati settimana precedente

'###FREEZE Panes    ''blocca riquadri da posizione cella J2
    ActiveWindow.ScrollColumn = 1
    Range("J2").Select
    ActiveWindow.FreezePanes = True

'Sub LoopTroughOrdersBlocked()
    ''@Per il buon funzionamento di questa macro, la colonna numero 8 deve essere ordinata in modo crescente.
    ''La macro esamina e seleziona i duplicati nella colonna 8 e quindi nella colonna corrispondente alla selezione fa un UNISCI CELLE
    ''quindi lo scopo e' unire e centrare tutte le celle in corrispondenza di + posizioni dello stesso cliente nella colonna "Azioni/Commenti"

Cells(2, 8).Activate  ''Seleziona la cella riga 2 colonna 8

Dim x As Integer, y As Integer ''Definizione di variabile di traccia riga

x = ActiveCell.Row   ''L'attuale riga viene assegnata a x
y = x + 1              '' la riga x+1 viene assegnata a y

Do While Cells(x, 8).Value <> ""  ''Loop A esegue il blocco di codici ripetutamente finche la cella non sara vuota a quel punto
                                    ''si fermera' il Do While
   
    If Cells(x, 8).Value <> Cells(y, 8).Value Then
        Cells(y, 8).Select
    Else
Do While Cells(x, 8).Value = Cells(y, 8).Value  ''Loop B
 
        If Cells(x, 8).Value = Cells(y, 8).Value Then
            Range(Selection, Selection.Offset(1, 0)).Select
        End If
        x = x + 1
        y = x + 1
        Loop ''Fine Loop B
          Selection.Offset(0, -1).Select
Application.Run "PERSONAL.XLSB!Merge_Cells"
Selection.Offset(1, 1).Select
End If

 x = x + 1
        y = x + 1
Loop ''Fine Loop A


'#----------------
    Cells.Select '' Seleziona tutte le celle
    Selection.RowHeight = 15 '' alla selezione tutte celle attribuisce altezza righe 15

'@# Questa sezione  somma il valore di tutti gli ordini inserendo una formula

Dim PrimaCella, lastsumCella As String ''Dichiarazione di due variabili di tipo stringa

PrimaCella = "B2"  '' assegnazione della stringa B2 alla variabile prima cella; servirà per indirizzo di cella nella formula

lastsumCella = Range("B2").End(xlDown).Offset(-1, 0).Address(rowrelative, columnrelative) ''con la funzione Offset e ADDRESS si rileva l'indirizzo
                                                                                         '' di dell'ultima cella contenente un valore (agisce in modo
                                                                                         '' dinamico)


Range("B2").End(xlDown).Value = "=sum(" & PrimaCella & ":" & lastsumCella & ")"  '' Inserisce la formula somma nell'ultima cella di valore
      
    Range(Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row - 1, "B"), "B2"). _
                                                                NumberFormat = "#,##0.00_);(#,##0.00)" '' applica formato migliaia a intervallo dati colonna B

'Selection.NumberFormat = "€ #,##0"
    Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, "B").NumberFormat = "€ #,##0" '' Applica formato numero con Euro

    
Cells(1, 1).Select '' Seleziona la cella A1
Dim stAttachment As String '' Dichiarazione Variabile stringa
Dim StPath As String
         StPath = "T:\CONTABILITA'\RECUPERO CREDITI\file ordini bloccati\BBMI_" & _
                            Cy & "_Ordini_Bloccati\"   ''Assegnato il percorso ove salvare il file a Stpath

                Dim StFileName As String ''Dichiarazione della variabile per il nome file
                    StFileName = "Ordini_Bloccati_" & Format(Date, "DD-MM-YYYY") ''assegna nome file con stringa + funzione FORMAT per la data
                    stAttachment = StPath & StFileName & ".xlsx" ''percorso completo di salvataggio attribuito a StAttachment
                  
                  Application.DisplayAlerts = False  '' disattiva tutti messaggi di avvertimento
                        With ActiveWorkbook  ''con la struttura with salva il file nel percorso definito prima
                            .SaveAs stAttachment ', FileFormat:=xlOpenXMLWorkbook
                       End With
                Application.DisplayAlerts = True  '' Attiva il messaggio di avvertimento
 
 Application.ScreenUpdating = True  ''Riattia il flash screen di Excel
 
 ''...preso dalla mia Macro tracking time
  endtriggerChrono = Now   ''Rileva l'orario di fine Esecuzione

Interval = endtriggerChrono - triggerChrono   ''Calcola la durata dell'esecuzione

 ' Formato della durata in minuti e secondi
   
     strOutput = Int(CSng(Interval * 24 * 60)) & " Minutes :" & Format(Interval, "ss") _
        & " Seconds"
    ''Debug.Print strOutput

'' Messagio finale di fine elaborazione e durata da variabile strOutput
MsgBox " Durata Elaborazione Bloccati" & vbCrLf & vbCrLf _
            & strOutput, vbOKOnly + vbInformation, "Macr@ris Tracking Time"
 
 Application.StatusBar = ""  '' re inizializza la barra di stato
 Exit Sub ''Fine macro se errori non riscontrati
 
ErrorHandler:  '' Gestione di errore con rilevamento tipo di errore e descrizione
MsgBox "Interruzione Macro Causa Errore in Orders_Blocked" & vbNewLine & "Contattare Macr@ris" & vbNewLine & _
    vbCrLf & "Error number:  # " & Err.Number & vbNewLine & _
      "Description:==> " & Err.Description, vbCritical, "Macr@ris \Error Macro"
 
Application.ScreenUpdating = True  ''riattiva i movimenti dello schermo
Application.DisplayAlerts = True  '' riattiva tutti i messagi di avvertimento
Application.StatusBar = ""        '' riattiva le impostazioni predefinite della barra di stato

End Sub

```
