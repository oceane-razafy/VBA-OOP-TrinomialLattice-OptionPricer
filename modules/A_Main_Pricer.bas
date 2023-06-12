Attribute VB_Name = "A_Main_Pricer"
Option Explicit
Sub Main_Pricer()
    'Proc�dure permettant d'afficher les prix dans la feuille "Pricer", et si demand�, les graphiques _
    "Underlying Value" & "Option Price")
    
    'Feuilles et fichier excel
    Dim wb As Workbook, wsPricer As Worksheet, wsUnd As Worksheet, wsOpt As Worksheet
    
    'Instances des classes d'input
    Dim mk As New Market, opt As New opt, n As New node, t As New tree
    
    'Variables pour calculer le temps d'ex�cution de la macro
    Dim start_t As Double, elapsed As Double

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'I.INITIALISATION DES VARIABLES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Le temps de d�part
    Let start_t = Timer()
    
    'Affectation du fichier et des feuilles excel
    Set wb = ThisWorkbook
    Set wsPricer = wb.Worksheets("Pricer")
    Set wsUnd = wb.Worksheets("Graph_Under")
    Set wsOpt = wb.Worksheets("Graph_Option")
    
    'Affectation de valeurs aux attributs des instances
    Call mk.FillMarket(Range("InterestRate"), Range("Volatility"), Range("Dividend"), _
    Range("StartPrice"), Range("DF"), Range("Start_date"), Range("Div_date"))
    Call opt.FillOption(Range("Strike"), Range("Maturity"), Range("Time"), Range("IsAmerican"), Range("IsCall"))
    
    t.nbSteps = wsPricer.Range("NbSteps")
    t.Delta_t = opt.time / t.nbSteps
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'II. PARTIE 1 : PRICING
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Calcul alpha de l'arbre
    Call t.compute_alpha(mk, opt)
     
    'Construction de l'arbre
    Call t.TreeBuild(opt, mk)
    
    'Prendre d'abord la racine de l'arbre cr��
    Set n = t.root
    
    'Calcul des prix de l'option
    Call t.Pricer(n, opt, mk)
    
    'Mettre le prix donn� par l'arbre dans la cellule excel
    Range("Tree_price").Value = t.root.Value
    
    'S'occuper du cas avec sans dividende pour le prix BS
    If mk.Dividend <> 0 Then
        Range("BS_price").Value = "Dividend <> 0"
    Else
        Range("BS_price").Value = Price_BS(opt, mk)
    End If
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%
    'III. PARTIE 2 : GRAPHIQUES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Effacer les graphiques pr�c�dents
    wsUnd.Range("place_graph_und").ClearContents
    wsOpt.Range("place_graph_opt").ClearContents
    
    'Si demand�, afficher les graphiques
    If wsPricer.Range("DisplayOrNot").Value = 1 Then
        Call display_graph(wb, t)
    End If
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%
    'IV. TEMPS D'EXECUTION
    '%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Afficher le temps d'ex�cution de la macro
    Let elapsed = Timer() - start_t
    wsPricer.Range("execution_time") = elapsed
End Sub

