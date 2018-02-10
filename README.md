## Aggregare cartelle Excel in un unico file

Questa macro apre tutti i file di una cartella specifica e esegue una
 determinata operazione.Utile nel caso tu abbia tanti file excel
 da elaborare. In questo caso aggrega tutti i dati di determinati
 fogli di diverse cartelle Excel in un solo file Excel
 Viene fatto  uso di Cicli `Do while...loop`,
 `For each...Next` Inoltre si fa uso di condizione `IF...End IF`, struttura
 `With ...End With`.Uso di variabili tipo `Worksheet` e `Workbook`. 
 Invito a leggere i `codici VBA` per scoprire di più.
 
 ***
 > Questa Macro puo permettere di risparmiare tempo nella funzione di **Cash Allocation** quando 
 > ricevi da un Ente come Finlombarda pagamenti frazionati in tanti file Excel. Per elaborare questi
 > file Excel potresti aver bisogno di raggrupparli in un unico file: Operazione che puo richiedere
 > tanto tempo se ce ne sono parecchi. 
 > _Con una Macro Excel, l'operazione si risolve in qualche secondo_
 [A questo link un video esempio](https://youtu.be/t6ifrAJB6Xk)
 ***
 
``` vb
Sub UnisciFogliCartelle_UnicaCartella()
''' Questa Macro viene fornita solo a titolo esemplicativo.
'''\richiesta di esecuzione a l'utente
  ''Msgbox per fornire infos all'utente di cio' che succedera' una volta che avra premuto su ok
Dim infos As Variant
    infos = MsgBox("Ciao!" & vbNewLine _
    & "Macro per Unire le Cartelle" & vbNewLine _
    & "Salvare Prima i Dati!!! ==>> CTRL-Z non funziona!" & vbNewLine _
     & vbNewLine & vbNewLine & "" & "Qui per sbaglio -->  Clicca su 'NO'", _
        vbYesNo + vbInformation + vbDefaultButton2, "Macraris unire Cartelle")
  If infos = vbNo Then
  Exit Sub
  End If
  
''istruzione per la gestione degli errori
''La gestione errore permette di dare all'utente
''messaggi personalizzati su errori riscontri e eventualmente
''fare eseguire determinate azioni per errori conosciuti
    
    On Error GoTo erroreGestione
  
  
 Application.StatusBar = "Macraris....Goditi un Caffe mentre lavoro per te..."

'''''''''''''''''''''''''''''''''''''''''''''''''
''@INIZIO Crea la cartella Excel Master in cui incollare dati da vari fogli
'''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False

Dim stFile As String '' Inizializzazione della variabile stFile al quale assegnare
                            ''il nome della cartella Master + percorso di salvataggio
Dim stPathDest As String ''' Dichiarazione del percorso di salvataggio
    
    '' salva sul desktop dell'utente
    stPathDest = "C:\Users\" & Environ("UserName") & "\Desktop\"
    stFile = stPathDest & "Master_File" & Format(Now, "dd-mmm-yy hh-mm-ss") & ".xlsx"
  
 Dim wbNuovaCart As Workbook '' dichiarazione di variabile di tipo cartella
 
    Set wbNuovaCart = Workbooks.Add '' Creazione i una cartella vuota e conseguente
                                    ''assegnazione alla variabile di tipo cartella
                                    '' wbNuovaCart
    
 '' Uso struttura with..end per attribuire proprietà titolo e soggetto
 '' alla cartella wbNuovaCart + salvataggio assegnando nome in stFile
    With wbNuovaCart
        .Title = "macraris_file_unico"
        .Subject = "accorpa_file"
        .SaveAs stFile
    End With

'' Cambia nome foglio attivo
ActiveSheet.Name = "File_Accorpato"

'' Definisce variabile di tipo Foglio
Dim shNuovaCart As Worksheet

'' Assegna il foglio attivo alla variabile shNuovaCart
'' cosi potra essere usato comodamente la variabile invece
'' di riferirsi al foglio con ad esempio
'' sheets("nomefoglio").
    Set shNuovaCart = wbNuovaCart.ActiveSheet
       '''
'''''''''''''''''''''''''''''''''''''''''''''''''
''# FINE Crea la cartella Excel Master in cui incollare dati da vari fogli
'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
'''''''''''''''''''''''''''''''''''''''''''''''''
'' @INZIO Crea ciclo di ripetizione While che permettera
'' di estrarre un file alla volta dalla cartella specificata
'' eseguire determinate azioni poi chiudere il file aperto.
'' Il ciclo viene ripetuto fino a che tutti file della cartella
'' saranno stati elaborati.
'''''''''''''''''''''''''''''''''''''''''''''''''

'' Definizione delle variabili
'' vMyfile sara' la variabile per i file del ciclo (loop)
'' strDirFile variabile tipo stringa e' il percorso dei file da analizzare
'' strTipoFile tipo stringa e' l'estensione del file

Dim vMyfile As Variant, strDirFile As String, strTipoFile As String
    '' Assegnazione del percorso alla variabile.
    '' Cambiare questo percorso per indicare la cartella dove trovare
    '' i tuoi file
    strDirFile = "C:\Users\Public\Temp\Macraris_Gare\FATTURATI USATI"
    
    '' in caso il percorso non abbia \ alla fine lo aggiunge
      If Right(strDirFile, 1) <> "\" Then strDirFile = strDirFile & "\"
    
    strTipoFile = "*.xls"
   ''Alla variabile vMyfile assegnazione percorso + nome file
   '' NB. *. in strTipoFile significa tutti file che terminano con
   ''.xls
   vMyfile = Dir(strDirFile & strTipoFile)
    
'' Dichiarazione variabile di tipo cartella a cui assegnare le varie cartelle
'' o file che che verranno aperte

Dim wbCartAttiva As Workbook

'' Inizio del Ciclo
'' il ciclo verra' ripetuto finche non ci saranno + file da analizzare
'' Se ci sono 50 file il Do While viene ripetuto 50 volte
Do While (vMyfile <> "")
    ''Nella finestra immediata scrive il nome del file considerato
        Debug.Print vMyfile
    '' Apre la cartella Excel
        Workbooks.Open strDirFile & vMyfile

'''Assegna il file alla variabile
Set wbCartAttiva = ActiveWorkbook

'' Dichiarazione di una Variabile di tipo Foglio +
'' 2 variabili di tipo numeri Long
Dim shAtt As Worksheet
    Dim iRiga As Long, iCol As Long

'' Ciclo ripetuto sui fogli della cartella
 For Each shAtt In Worksheets
 
 '' Se il foglio inizia con tot e termina con anno allora
 If shAtt.Name Like "tot*anno" Then

'' Struttura With ...End
 With shAtt
 '' con Specialcells....Row individua il numero riga ultimo dato
 '' del foglio
iRiga = .Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row
'' il seguente codice commentato non funzionerebbe in quanto
'' in caso ultimacella fosse l'ultima colonna di Excel cercare di
'' spostarsi oltre l'ultima cartella darebbe errori
    ''iCol = .Cells(1, 1).SpecialCells(xlCellTypeLastCell).Offset(0, 1).Column
End With
''' Non dinamica percio da adattare al tipo di lavoro

iCol = 12
'' A intervallo riga ultima e colonna 12
'' quindi tutte le righe di dati indicare il nome del file
 Range(shAtt.Cells(iRiga, iCol), shAtt.Cells(2, iCol)) = wbCartAttiva.Name

 '' Attento nel copiare in una sola operazione se rimuovi il shNuovaCart dopo cella allora
 '' prendera' il riferimento dell'ultima cella dalla cartella di partenza e non
 '' è quello la tua intenzione
 shAtt.Cells(1, 1).CurrentRegion.Copy _
    shNuovaCart.Cells(shNuovaCart.[A1].SpecialCells(xlCellTypeLastCell).Row + 1, "A")

End If
 Next
 
'' Disattiva messaggi di avvertimento cosi non salva le modifiche
 Application.DisplayAlerts = False
    ''chiude la cartella Excel aperta all'inizio del ciclo "loop"
    wbCartAttiva.Close
Application.DisplayAlerts = True
   
   '' Assegna il prossimo file della cartella alla variabile vMyfile
        vMyfile = Dir()
'' Riprende il loop e cosi via...
Loop

MsgBox "Ho Finito!" & vbNewLine & "Sono stato veloce Vero!?" & vbNewLine _
        & vbNewLine & "Ci Vediamo presto!" & vbTab & "Ciao!", vbInformation, _
        "Macraris Formazione e Consulenza"
 Application.ScreenUpdating = True
 Application.StatusBar = ""

Exit Sub

erroreGestione: '' Gestione di errore con rilevamento tipo di errore e descrizione
MsgBox "Interruzione Macro Causa Errore in UnisciFogliCartelle_UnicaCartella" & vbNewLine _
	& "Contattare Macraris" & vbNewLine & vbCrLf _
    & "Numero Errore:  # " & Err.Number & vbNewLine _
    & "Descizione Errore :==> " & Err.Description, _
	vbCritical, "Macraris \Error Macro"

 Application.ScreenUpdating = True
 Application.StatusBar = ""
 
''Riferimenti Web:
''https://wellsr.com/vba/2016/excel/vba-loop-through-files-in-folder/
''https://msdn.microsoft.com/en-us/library/cc793964.aspx
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
```
***
