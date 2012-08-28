Attribute VB_Name = "Module1"
'Provisions

'Déclaration des variables publiques

'Feuilles
Public RSH As Worksheet         'Feuille récap
Public CSH As Worksheet         'Feuille composition
Public VSH As Worksheet         'Feuille cours (valeurs)
Public DSH As Worksheet         'Feuille dictionnaire de codes

Public c As Collection                                  'Collection portefeuilles
Public Const PORTEFEUILLES As String = "TRANS PART PLACT"     'Portefeuilles

'Index
'Feuille CSH - composition
Public Const ICL As Integer = 4     'Ligne de début du tableau
Public Const ICT As Integer = 2     'Colonne titre
Public Const ICC As Integer = 3     'Colonne code
Public Const ICP As Integer = 4     'Colonne portefeuille
Public Const ICN As Integer = 5     'Colonne nombre de titres détenus
Public Const ICA As Integer = 6     'Colonne valeur d'acquisition
Public Const ICV As Integer = 7     'Colonne cours d'acquisition
Public Const ICS As Integer = 8     'Colonne stock de provision
'Feuille VSH - Cours
Public Const IVL As Integer = 4     'Ligne de début du tableau
Public Const IVC As Integer = 2     'Colonne de début du tableau (titre)
Public Const IVCA As Integer = 3    'Colonne du cours actuel
Public Const IVCD As Integer = 4    'Colonne dernier cours de cloture (j-1)
'Feuille DSH - Dictionnaire
Public Const IDL As Integer = 4     'Ligne de début du tableau
Public Const IDC As Integer = 2     'Colonne de début du tableau (intitulé)
Public Const IDCC As Integer = 3    'Colonne code

Sub commencer()

optimisationDébut
variables
réinitialiser

''''''''''''''
'''' Code ''''
''''''''''''''

'Collection portefeuilles
Set c = New Collection
'c.Add New Collection, "TRANS"
'c.Add New Collection, "PART"
'c.Add New Collection, "PLACT"
For Each i In Split(PORTEFEUILLES)
    c.Add New Collection, i
Next i

'Compile la liste des actions
With CSH
'i = 5
i = ICL
Do Until IsEmpty(.Cells(i, ICT))
    c.Item(.Cells(i, ICP)).Add New Collection, .Cells(i, ICC)
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add .Cells(i, ICC), "code"
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add .Cells(i, ICT), "nom"
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add .Cells(i, ICN), "nb"
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add .Cells(i, ICA), "valeur acquisition"
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add .Cells(i, ICV), "cours acquisition"
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add 0, "cours valorisation"
    c.Item(.Cells(i, ICP)).Item(.Cells(i, ICC)).Add .Cells(i, ICS), "provision"
    i = i + 1
Loop
End With


'''''''''''''''''''''''''''''''''''''''''''''''
'''' Met à jour les cours '''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'extract_bourse_casablanca

'''''''''''''''''''''''''''''''''''''''''''''''
'''' Cherche les codes correspondants '''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
With DSH
For Each i In Split(PORTEFEUILLES)
    For Each j In c.Item(i)
        'x = 3
        x = IDL
        Do While Not IsEmpty(.Cells(x, IDC))
            If j.Item("code") = .Cells(x, IDCC) Then
                j.Remove "code"
                j.Add .Cells(x, IDC), "code"
            End If
            x = x + 1
        Loop
    Next j
Next i
End With

'''''''''''''''''''''''''''''''''''''''''''''''
'''' Cherche les cours correspondants '''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
' Quel cours utiliser "Dernier cours" ou "Cours de référence"?
If Worksheets("Cours").OptionButtonCoursActuel.Value = True Then colonneCours = 3
If Worksheets("Cours").OptionButtonCoursCloture.Value = True Then colonneCours = 4

With VSH
For Each i In Split(PORTEFEUILLES)
    For Each j In c.Item(i)
        'x = 3
        x = IVL
        Do Until IsEmpty(.Cells(x, IVC))
            If j.Item("code") = .Cells(x, IVC) Then
                j.Remove "cours valorisation"
                'j.Add .Cells(x, 3), "cours valorisation"
                j.Add .Cells(x, colonneCours), "cours valorisation"
            End If
            x = x + 1
        Loop
    Next j
Next i
End With

'''''''''''''''''''''''''''''''''''''''''''''''
'''' Dessine le tableau récapitulatif '''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
With RSH
x = 5
For Each i In Split(PORTEFEUILLES)

    .Cells(x, 2) = i        'Nom du portefeuille (titre)
    .Cells(x, 2).Font.Bold = True
    titres = "Titre,Nb titres,Valeur d'acquisition,Cours d'acquisition,Cours valorisation,Valeur marché,+/- value latente,VM/VC,Provision fin,Dotation,Reprise"
    'index tableau
    irt = 2         'Colonne titre
    irn = 3         'Colonne nombre de titres
    irva = 4        'Colonne valeur d'acquisition
    irca = 5        'Colonne cours d'acquisition
    ircv = 6        'Colonne cours valorisation
    irvm = 7        'Colonne valeur marché
    irpv = 8        'Colonne plus/moins value latente
    irvc = 9        'Colonne valeur marché / valeur comptable
    irpr = 10       'Colonne provision
    irdt = 11       'Colonne dotation
    irre = 12       'Colonne reprise
        
    k = 2
    For Each j In Split(titres, ",")
        .Cells(x + 1, k) = j
        .Cells(x + 1, k).Font.Bold = True
        .Cells(x + 1, k).HorizontalAlignment = xlCenter
        k = k + 1
    Next j
    x = x + 2
    total_debut = x
    For Each j In c.Item(i)
        .Cells(x, irt) = j.Item("nom")
        .Cells(x, irn) = j.Item("nb")
        .Cells(x, irva) = j.Item("valeur acquisition")
        .Cells(x, irca) = j.Item("cours acquisition")
        .Cells(x, ircv) = j.Item("cours valorisation")
        .Cells(x, irvm) = j.Item("cours valorisation") * j.Item("nb")
        pmv = (j.Item("cours valorisation") * j.Item("nb")) - j.Item("valeur acquisition")      'Plus/moins value
        .Cells(x, irpv) = pmv
        .Cells(x, irvc) = (j.Item("cours valorisation") * j.Item("nb")) / j.Item("valeur acquisition")
        
        'Provision
        'p = (j.Item("cours valorisation") * j.Item("nb")) - j.Item("valeur acquisition") - j.Item("provision")
        dotation = 0
        reprise = 0
        provision1 = 0
        If pmv < 0 Then provision1 = -pmv
            diffprov = provision1 - j.Item("provision")
            If diffprov > 0 Then
                dotation = diffprov
                reprise = 0
            End If
            If diffprov < 0 Then
                dotation = 0
                reprise = diffprov
            End If
        
        .Cells(x, irpr) = provision1
        .Cells(x, irdt) = dotation
        .Cells(x, irre) = -reprise
        
        .Cells(x, irn).NumberFormat = "#,##0"
        .Cells(x, irva).NumberFormat = "#,##0.00"
        .Cells(x, irca).NumberFormat = "#,##0.00"
        .Cells(x, ircv).NumberFormat = "#,##0.00"
        .Cells(x, irvm).NumberFormat = "#,##0.00"
        .Cells(x, irpv).NumberFormat = "#,##0.00"
        .Cells(x, irvc).NumberFormat = "#,##0.00"
        .Cells(x, irpr).NumberFormat = "#,##0.00"
        .Cells(x, irdt).NumberFormat = "#,##0.00"
        .Cells(x, irre).NumberFormat = "#,##0.00"
        x = x + 1
    Next j
    
    colorer_tableau RSH, Int(total_debut), Int(irt), Int(irre)
    
    'Ajout des totaux
    .Cells(x, 2) = "Total"
    .Cells(x, 2).Font.Bold = True
    'For j = 3 To irre
    For Each j In Array(irn, irva, irvm, irpv, irpr, irdt, irre)
        .Cells(x, j).Formula = "=sum(" + .Cells(total_debut, j).Address + ":" + .Cells(x - 1, j).Address + ")"
        .Cells(x, j).Font.Bold = True
    Next j
    .Range(.Cells(x, 2), .Cells(x, irre)).Interior.Color = RGB(64, 64, 64)
    .Range(.Cells(x, 2), .Cells(x, irre)).Font.Color = vbWhite


    x = x + 2
Next i
End With



optimisationFin

End Sub
Function colorer_tableau(sh As Worksheet, ligne As Integer, min As Integer, max As Integer)

Dim ligneGrandTitre As Integer
Dim ligneSousTitre As Integer

ligneGrandTitre = ligne - 2
ligneSousTitre = ligne - 1

With sh
    'Colore le grand titre
    With sh.Cells(ligneGrandTitre, min)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = vbWhite
    End With
    .Range(.Cells(ligneGrandTitre, min), .Cells(ligneGrandTitre, max)).Merge

    'Colore les sous-titres
    With sh.Range(.Cells(ligneSousTitre, min), .Cells(ligneSousTitre, max))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(83, 142, 213)
        .Font.Color = vbWhite
    End With
    'Colore les lignes
    k = 0
    Do Until IsEmpty(.Cells(ligne + k, min))
        If k Mod 2 = 0 Then .Range(.Cells(ligne + k, min), .Cells(ligne + k, max)).Interior.Color = RGB(216, 216, 216) 'Gris clair
        k = k + 1
    Loop
End With


End Function

Sub optimisationDébut()

Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

End Sub


Sub optimisationFin()

Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Sub variables()

Set RSH = Worksheets("Récap")
Set CSH = Worksheets("Composition")
Set VSH = Worksheets("Cours")
Set DSH = Worksheets("Dictionnaire codes")

'Index
'Feuille csh - composition
'ICT = 2     'Colonne titre
'ICC = 3     'Colonne code
'ICP = 4     'Colonne portefeuille
'ICN = 5     'Colonne nombre de titres détenus
'ICA = 6     'Colonne valeur d'acquisition
'ICV = 7     'Colonne cours d'acquisition
'ICS = 8     'Colonne stock de provision
'
'PORTEFEUILLES = "PART PLACT"

End Sub

Sub réinitialiser()

RSH.Range("A1:L1000").Clear
RSH.Range("A1:L1000").ClearFormats

End Sub

Sub extract_bourse_casablanca()
'Bourse de Casablanca
'http://www.casablanca-bourse.com/bourseweb/Cours-Valeurs.aspx?Cat=24&IdLink=300

    variables

    Set re = CreateObject("vbscript.regexp")
    With re
        .MultiLine = True
        .Global = True
        .IgnoreCase = False
    End With
    
    stri = wget("http://www.casablanca-bourse.com/bourseweb/Cours-Valeurs.aspx?Cat=24&IdLink=300")
    'Titre
    re.pattern = "id=""CoursValeurs1_Actionl1_ListActionSecteur_ctl[0-9]{2}_TableAction1_RptrAction_ctl[0-9]{2}_Label11?"">.*?</span>"
    Set titres = re.Execute(stri)
    'Dernier cours (différé 15 minutes)
    re.pattern = "id=""CoursValeurs1_Actionl1_ListActionSecteur_ctl[0-9]{2}_TableAction1_RptrAction_ctl[0-9]{2}_Label31?"">.*?</span>"
    Set cours = re.Execute(stri)
    'Cours de référence (cours de cloture j-1)
    re.pattern = "id=""CoursValeurs1_Actionl1_ListActionSecteur_ctl[0-9]{2}_TableAction1_RptrAction_ctl[0-9]{2}_Label41?"">.*?</span>"
    Set coursCloture = re.Execute(stri)
    
    
    re.pattern = "id=""CoursValeurs1_DateSeance1_LBDateSeance"">.*?</span>"
    miseajour = re.Execute(stri)(0)
    re.pattern = "[0-9]{2}/[0-9]{2}/[0-9]{4}"
    miseajour = re.Execute(miseajour)(0)
    
    
    '''
    'Set sh = Worksheets("Cours")
    'With sh
    With VSH
        
    'Effacer
    'j = 3
    j = IVL
    Do Until IsEmpty(.Cells(j, IVC))
        With VSH.Range(.Cells(j, IVC), .Cells(j, IVCD))
            .ClearContents
            .ClearFormats
            .Interior.Color = xlNone
        End With
'        .Cells(j, 2) = ""
'        .Cells(j, 3) = ""
        j = j + 1
    Loop
    '.Cells(3, 5) = ""
    '.Cells(4, 5) = ""
    
    'Boucle titres
    'j = 3
    j = IVL
    For Each i In titres
        re.pattern = "[A-Z][A-Z\s\(\)\.]+"
        re.pattern = "[A-Z][A-Z0-9\s\(\)\.]+"
        re.pattern = "[A-Z][A-Z0-9\s\.\-]+"
        .Cells(j, IVC) = re.Execute(i)(0)
        j = j + 1
    Next i
    
    'Boucle cours
    'j = 3
    j = IVL
    For Each i In cours
        re.pattern = "[0-9\s]+,[0-9]{2}"
        re.pattern = "[0-9]? ?[0-9]+,[0-9]{2}"
        .Cells(j, IVCA) = CDbl(re.Execute(i)(0))
        .Cells(j, IVCA).NumberFormat = "#,##0.00"
        j = j + 1
    Next i
    
    'Boucle cours cloture
    'j = 3
    j = IVL
    For Each i In coursCloture
        re.pattern = "[0-9\s]+,[0-9]{2}"
        re.pattern = "[0-9]? ?[0-9]+,[0-9]{2}"
        .Cells(j, IVCD) = CDbl(re.Execute(i)(0))
        .Cells(j, IVCD).NumberFormat = "#,##0.00"
        j = j + 1
    Next i
    
    colorer_tableau VSH, IVL, IVC, IVCD

    .Cells(3, 6) = miseajour
    
    End With

End Sub

Function wget(url As String)
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'url = "http://www.mibian.net/serve?eurusd"
    objHttp.Open "GET", url, False
    'objHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5"
    objHttp.send ""
    wget = objHttp.ResponseText
End Function

Function regex(str As String, pattern As String) As Collection
    Dim re As Object
    Dim results As Object
    
    Set re = CreateObject("vbscript.regexp")
    'Set re = New regexp
    With re
        .MultiLine = True
        .Global = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    Set results = re.Execute(str)
    regex = results

End Function

Sub extract_bmce()
'BMCE
    Set re = CreateObject("vbscript.regexp")
    With re
        .MultiLine = True
        .Global = True
        .IgnoreCase = False
        '.pattern = "<span>[A-Z]*?.*?[0-9]*,[0-9]{2}</td> </tr>"
        .pattern = "[A-Z\s\(\)\.]*?</span>.*?[0-9]*,[0-9]{2}</td> </tr>"
    End With
    stri = wget("http://www.bmcek.co.ma/front.aspx?sectionId=278")
    Set results = re.Execute(stri)
    
    Set sh = Worksheets("Cours")
    With sh
    
    'Effacer
    j = 3
    Do While Not IsEmpty(.Cells(j, 2))
        .Cells(j, 2) = ""
        .Cells(j, 3) = ""
        j = j + 1
    Loop
    .Cells(3, 5) = ""
    .Cells(4, 5) = ""
    
    j = 3
    For Each i In results
        're.pattern = "[A-Z\s]*"
        re.pattern = "[A-Z\s\(\)\.]*"
        .Cells(j, 2) = re.Execute(i)(0)
        re.pattern = "[0-9]*,[0-9]{2}</td> </tr>"
        r = re.Execute(i)(0)
        re.pattern = "[0-9]*,[0-9]{2}"
        .Cells(j, 3) = CInt(re.Execute(r)(0))
        .Cells(j, 3).NumberFormat = "#,##0.00"
        '.Cells(j, 5) = i
        j = j + 1
    Next i
    
    re.pattern = "Date de mise à jour .*?</span>"
    results = re.Execute(stri)(0)
    re.pattern = "[0-9]{2}/[0-9]{2}/[0-9]{4}"
    .Cells(3, 5) = re.Execute(results)(0)
    re.pattern = "[0-9]{2}:[0-9]{2}:[0-9]{2}"
    .Cells(4, 5) = re.Execute(results)(0)
    
    End With

End Sub
