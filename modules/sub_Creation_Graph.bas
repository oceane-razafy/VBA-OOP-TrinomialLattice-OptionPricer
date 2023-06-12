Attribute VB_Name = "sub_Creation_Graph"
Option Explicit
Sub display_graph(ByRef wb As Workbook, ByVal t As tree)
    'Proc�dure permettant d'afficher les graphiques (Underlying Value & Option Price) _
    qui prend comme input le fichier excel et l'arbre � cr�er

    'Feuilles de graphiques
    'Variables "noeuds" permettant de cr�er le graphique
     Dim n As node, new_mid As node, node_up As node, node_down As node
    
    'Variable pour remplir cases excel lors de la construction du graphique, Points de d�part des graphiques
    Dim position As Long, start_point_und As Range, start_point_opt As Range
    
    'Le pas de l'arbre o� on se situe
    Dim step As Long
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'I. INITIALISATION DES VARIABLES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Call init_display(wb, t, start_point_und, start_point_opt)
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'II. AFFICHAGE DES GRAPHIQUES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Mettre les valeurs et prix des racines (step 0)
    Set n = t.root
    start_point_und.Value = n.underlying
    start_point_opt.Value = n.Value
    
    'se positionner sur le tronc d'abord
    position = 0
    
    'Boucle sur les pas de l'arbre
    For step = 1 To t.nbSteps
        'positionne sur le bon mid (au d�but : n = racine, new_mid = future_mid de la racine)
        Set new_mid = n.future_mid
        start_point_und.Offset(0, step).Value = new_mid.underlying
        start_point_opt.Offset(0, step).Value = new_mid.Value
        
        'Nodes au dessus et en dessous
        Set node_up = new_mid.up
        Set node_down = new_mid.down
        
        '__________a) Remplir la partie haute de la colonne_______________
        Do
            'Changer la position pour monter et descendre sur les colonnes
            position = position + 1
            
            'Remplir les case Excel avec :
            '--> Valeur du sous-jacent
            start_point_und.Offset(-position, step).Value = node_up.underlying
            '-->Prix de l'option
            start_point_opt.Offset(-position, step).Value = node_up.Value
        
            'Passer au node_up et node_down suivant
            Set node_up = node_up.up
    
        'Boucler jusque tout en haut de la colonne
        Loop Until node_up Is Nothing
        
        'Se reposition sur le tronc
        Let position = 0
        
       '__________a) Remplir la partie basse de la colonne_______________
        
        'Faire la m�me en chose dans la partie basse de la colonne
        Do
            position = position + 1
            start_point_und.Offset(position, step).Value = node_down.underlying
            start_point_opt.Offset(position, step).Value = node_down.Value
            Set node_down = node_down.down
        Loop Until node_down Is Nothing
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        'c) Passer au noeud tronc mid suivant pour le pas d'apr�s
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        position = 0
        Set n = new_mid
    Next step
End Sub
Sub init_display(ByVal wb As Workbook, ByVal t As tree, ByRef start_point_und As Range, ByRef start_point_opt As Range)
    'Proc�dure permettant d'initialiser les variables pour l'affichage des graphiques
    'notamment les points de d�part

    'Feuilles graphiques
    Dim wsUnd As Worksheet, wsOpt As Worksheet
    
    'le nombre de noeuds au dessus du mid de la derni�re colonne
    Dim nb As Long
    
    'Affectation des feuilles dans des variables
    Set wsUnd = wb.Worksheets("Graph_Under")
    Set wsOpt = wb.Worksheets("Graph_Option")
    
    'Prendre le nombre de noeuds au dessus du mid de la derni�re colonne
    nb = t.LastColSizeUp(t.root)

    'Point de d�part des arbres
    Set start_point_und = wsUnd.Range("starting_point_under")
    Set start_point_und = start_point_und.Offset(Round(nb + 1, 0), 0)
    Set start_point_opt = wsOpt.Range("starting_point_option")
    Set start_point_opt = start_point_opt.Offset(Round(nb + 1, 0), 0)
End Sub
Sub treevsbs1()
Attribute treevsbs1.VB_ProcData.VB_Invoke_Func = " \n14"
'Creation du graphique Gap x NbSteps as a function of Number of Steps"
'OPTION SOUHAITE : European call Spot @100 Strike @100 Expiry@1Y Vol@20% IR@2%, with 500 steps
'param�tres de l'option remplies dans param_option

    'Param�tres choisies :
    Const nbSteps1 As Double = 500
    
    'Le pas de l'arbre o� l'on se situe
    'La feuille de donn�es
    'Les rang�es de donn�es
    Dim step As Long, wsTreeBS1 As Worksheet, rg_nbsteps As Range, rg_treeprice As Range, rg_bsprice As Range, rg_gap As Range
    
    'Les instances des classes d'inputs
    'Variables pour calculer le temps d'ex�cution de la macro
    Dim mk As New Market, opt As New opt, t As New tree, start_t As Double, elapsed As Double
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'I.INITIALISATION DES VARIABLES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Le temps de d�part
    Let start_t = Timer()
    
    'Affectation de la feuille
    Set wsTreeBS1 = ThisWorkbook.Worksheets("Tree vs BS (1)")
    
    'Affectation des inputs
    Call param_option(mk, opt)
    
    'Arr�ter la comparaison des prix si le dividende est non nul
    If mk.Dividend <> 0 Then
        MsgBox ("Dividend <> 0, no comparison possible")
        Exit Sub
    End If
    
    'Affectation des rang�es de donnn�es
    Set rg_nbsteps = wsTreeBS1.Range("range_nbsteps1")
    Set rg_treeprice = wsTreeBS1.Range("range_treeprice1")
    Set rg_bsprice = wsTreeBS1.Range("range_bsprice1")
    Set rg_gap = wsTreeBS1.Range("range_gap1")

    
    '%%%%%%%%%%%%%%%%%%%%%%%%
    'II. REMPLISSAGE DES PRIX
    '%%%%%%%%%%%%%%%%%%%%%%%%
    '_________a) Prix Black & Scholes_____________
    rg_bsprice(1, 1).Value = Price_BS(opt, mk)
    rg_bsprice(1, 1).AutoFill Destination:=rg_bsprice, Type:=xlFillValues
    
    '_________b) Prix de l'arbre___________________
    'Boucle sur les pas de l'arbre
    For step = 1 To nbSteps1
        'Initalisation
        t.nbSteps = rg_nbsteps(step, 1)
        
        'Pricing
        Call proc_price_treevsbs1(t, mk, opt)
        
        'Mettre le prix de l'option dans la cellule excel correspondante
        rg_treeprice(step, 1).Value = t.root.Value
        
        'c)_________________GAP_________________
        rg_gap(step, 1).Value = (3 * mk.StartPrice / (8 * Sqr(2 * WorksheetFunction.Pi))) * _
        (((mk.Volatility) ^ 2 * t.Delta_t) / (Sqr(Exp((mk.Volatility) ^ 2 * opt.time) - 1)))
    
        'Lib�rer la m�moire
        Call t.FreeTree(t)
    Next step
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%
    'III. TEMPS D'EXECUTION
    '%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Afficher le temps d'ex�cution de la macro
    Let elapsed = Timer() - start_t
    wsTreeBS1.Range("execution_time2") = elapsed
End Sub
Sub param_option(ByRef mk As Market, ByRef opt As opt)
    'Les constantes permettant de construire l'OPTION SOUHAITE
    'OPTION SOUHAITE : European call Spot @100 Strike @100 Expiry@1Y Vol@20% IR@2%, with 500 steps
    Const start_p As Double = 100, strike As Double = 100, time As Double = 1, vol As Double = 0.2
    Const IR As Double = 0.02, div As Double = 0, isAmerican As Boolean = False, isCall As Boolean = True
    
    'Affectation des inputs
    mk.InterestRate = IR
    mk.Volatility = vol
    mk.Dividend = div
    mk.StartPrice = start_p
    
    opt.strike = strike
    opt.time = time
    opt.isAmerican = isAmerican
    opt.isCall = isCall
End Sub
Sub proc_price_treevsbs1(ByRef t As tree, ByVal mk As Market, ByVal opt As opt)
    'Proc�dure permettant de faire le pricing pour chaque arbre

    'Initialisation des inputs
    t.Delta_t = opt.time / t.nbSteps
    mk.DF = Exp(-mk.InterestRate * (opt.time / t.nbSteps))
    
    'Calcul alpha de l'arbre
    Call t.compute_alpha(mk, opt)
    
    'Construction de l'arbre
    Call t.TreeBuild(opt, mk)
        
    'Calcul du prix de l'option
    Call t.Pricer(t.root, opt, mk)
End Sub
Sub treevsbs2()
'Creation du graphique "Tree and Black-Scholes prices as functions of the strike"
'Avec les inputs de la feuille Pricer
    
    'Le num�ro de la simulation (simulation 1 pour valeur strike 1, etc.)
    Dim simul_strike As Long
    
    'Le fichier et les feuilles excel
    Dim wb As Workbook, wsPricer As Worksheet, wsTreeBS2 As Worksheet
    
    'Les rang�es de donn�es
    Dim rg_strike As Range, rg_treeprice2 As Range, rg_bsprice2 As Range
    
    'Les instances des classes d'input
    Dim mk As New Market, opt As New opt, t As New tree

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'I.INITIALISATION DES VARIABLES
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Affectation du fichier et des feuilles excel
    Set wb = ThisWorkbook
    Set wsPricer = wb.Worksheets("Pricer")
    Set wsTreeBS2 = wb.Worksheets("Tree vs BS (2)")
    
    'Affectation de valeurs pour les attributs des classes d'input
    Call opt.FillOption(Range("Strike"), Range("Maturity"), Range("Time"), Range("IsAmerican"), Range("IsCall"))
    Call mk.FillMarket(Range("InterestRate"), Range("Volatility"), Range("Dividend"), _
    Range("StartPrice"), Range("DF"), Range("Start_date"), Range("Div_date"))
    
    'Arr�ter la comparaison si le dividende est non nul
    If mk.Dividend <> 0 Then
        MsgBox ("Dividend <> 0, no comparison possible")
        Exit Sub
    End If
    
    'Affectation des plages de donn�es
    Set rg_strike = wsTreeBS2.Range("range_strike2")
    Set rg_treeprice2 = wsTreeBS2.Range("range_treeprice2")
    Set rg_bsprice2 = wsTreeBS2.Range("range_bsprice2")
    
    '%%%%%%%%%%%%%%%%%%%%%%%%
    'II. REMPLISSAGE DES PRIX
    '%%%%%%%%%%%%%%%%%%%%%%%%
    'Boucle pour le prix donn� par l'arbre dans chaque simulation
    For simul_strike = 1 To rg_strike.Count
    
        'R�initialiser apr�s la lib�ration de la m�moire
        t.nbSteps = wsPricer.Range("NbSteps")
        t.Delta_t = opt.time / t.nbSteps
        
        'Utiliser le bon strike
        Call opt.FillOption(rg_strike(simul_strike, 1), Range("Maturity"), Range("Time"), Range("IsAmerican"), Range("IsCall"))
        
        'Prix BS
        rg_bsprice2(simul_strike, 1).Value = Price_BS(opt, mk)
        
        'Calcul alpha de l'arbre
        Call t.compute_alpha(mk, opt)
     
        'Construction de l'arbre
        Call t.TreeBuild(opt, mk)
            
        'Calcul du prix de l'option
        Call t.Pricer(t.root, opt, mk)
        
        'Mettre le prix de l'option dans la cellule excel
        rg_treeprice2(simul_strike, 1).Value = t.root.Value
    
        'Lib�rer la m�moire
        Call t.FreeTree(t)
    'Passer � la simulation strike suivant
    Next simul_strike
End Sub



