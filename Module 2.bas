'Concatène les entrées de tableaux sur les onglets dont le nom répond à une condition, dans divers fichiers dits "Sources" vers ce présent fichier dit "Cible"
'Validé le 19/04/2023
'Fonctionne pour un nombre variable de fichiers sources, et un nombre variables d'onglets sources de même format dans chaque fichier
'Prérequis :
'- utiliser une trame harmonieuse entre les fichiers sources et le fichier cible
'- déposer les fichiers sources dans un même dossier au même niveau
'Flexibilité :
'- il n'est pas nécessaire que ce fichier soit dans le même dossier que les sources
'- d'autres fichiers peuvent coexister dans le fichier sources
'Paramètres du code :
'- le chemin du dossier contenant les fichiers sources
'- les dimensions de la trame du tableau
'- la condition sur le nom des onglets dont le contenu est collecté
'
' ! le code ouvre et ferme l'ensemble des fichiers sources -> VEILLER A ENREGISTRER VOTRE TRAVAIL SUR L'ENSEMBLE DES FICHIERS SOURCES AVANT DE L'UTILISER

Sub ImportDataFromMultipleFiles()

    Dim strPath As String 'path of the folder containing the Excel files
    Dim strFile As String 'filename of the Excel file
    Dim wbSource As Workbook 'source workbook
    Dim wbTarget As Workbook 'target workbook
    Dim wsSource As Worksheet 'source worksheet
    Dim wsTarget As Worksheet 'target worksheet
    Dim targetLastRow As Long 'last row in the target worksheet
    Dim rCopyRange As Range 'range to copy from the source worksheet
       
    'JB debug variables
    Dim debugJB As Integer
    debugJB = 1
    
    'JB mission constant parameters
    Dim premierRangLibre As Integer
    premierRangLibre = 3 ''PARAMETRER ICI : selon la hauteur de l'entête
    
    
    'JB mission variables
    Dim wsIterator As Integer 'to follow iterations through source files
    wsIterator = 2 'initialized so as to start the source files log at A2

    Dim testWs 'to identify worksheets of interest for importing
    Dim sourceLastRow As Long
    
       
    'Set the path of the folder containing the Excel files
    strPath = "Y:\L3-Data\Job STB NUWARDrevD\2-Remontages\Remontage 20230216\Sources\" 'PARAMETRER ICI
    
    'Set the target workbook and worksheet
    Set wbTarget = ThisWorkbook 'or set to a specific workbook
    Set wsTarget = wbTarget.Worksheets("Compilation")
    
    wbTarget.Worksheets("Debug").Cells(1, 1).Value = debugJB 'DEBUG
    debugJB = debugJB + 1 ' DEBUG
    
    'Loop through all Excel files in the folder
    strFile = Dir(strPath & "*.xlsx")
    Do While strFile <> ""
        
        'Open the source workbook
        Set wbSource = Workbooks.Open(strPath & strFile)
        
       
        wbTarget.Worksheets("Debug").Cells(1, 1).Value = debugJB 'DEBUG
        debugJB = debugJB + 1 'DEBUG
        
        'Loop through all worksheets in the source workbook
        For Each wsSource In wbSource.Worksheets
            
            'JB Le  contenu d'un ws est importé si le nom de la ws contient "Résultat exigences"
            testWs = InStr(1, wsSource.Name, "Résultat exigences") 'PARAMETRER ICI
            If testWs <> 0 Then
                        
                'JB : Note la source dans l'onglet SOURCES du fichier cible
                wbTarget.Worksheets("Sources").Cells(wsIterator, 2).Value = wbSource.Name
                wbTarget.Worksheets("Sources").Cells(wsIterator, 3).Value = wsSource.Name
                
                                                     
                'Find the last row in the target worksheet
                targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row 'JB : la colonne B contenant les identifiants est utilisée pour juger de la fin du tableau car elle est en théorie 100% remplie
                wbTarget.Worksheets("Sources").Cells(wsIterator, 5).Value = targetLastRow + 1 'JB : note sur l'onglet "Sources" du fichier cible le dernier rang considéré occupé dans le fichier cible
                
                'JB : Find the last row in the source worksheet
                sourceLastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row 'JB : la colonne B contenant les identifiants est utilisée pour juger de la fin du tableau car elle est en théorie 100% remplie
                wbTarget.Worksheets("Sources").Cells(wsIterator, 4).Value = sourceLastRow 'JB : note sur l'onglet "Sources" du fichier cible le dernier rang considéré occupé dans le fichier source
                
                'Set the range to copy from the source worksheet
                wsSource.Activate
                Set rCopyRange = wsSource.Range(Cells(premierRangLibre, 1), Cells(sourceLastRow, 23)) 'PARAMETRER ICI : selon la largeur de l'entête
                
                'Copy the data to the target worksheet
                rCopyRange.Copy wsTarget.Cells(targetLastRow + 1, "A")
                
                wsIterator = wsIterator + 1
                
            End If
               
        Next wsSource
        
        
        'Close the source workbook
        wbSource.Close False
        
        
        'Get the next Excel file in the folder
        strFile = Dir
        
    Loop
    
        wbTarget.Worksheets("Debug").Cells(2, 1).Value = "finito" 'DEBUG
       
End Sub

