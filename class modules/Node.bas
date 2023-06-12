VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Sous-jacent et arbre
Public underlying As Double
Public tree As tree

'Le parent mid sur le tronc de l'arbre
Public parent_ As node

'Les noeuds fils
Public future_mid As node
Public future_up As node
Public future_down As node

'Les noeuds voisins
Public up As node
Public down As node

'Les probabilités
Public pup As Double
Public pmid As Double
Public pdown As Double

'Attributs pour l'option:
'prix de l'option inexistant ou non sur le noeud
Public IsValueAvailable As Boolean
'prix de l'option
Public Value As Double

Const sum_prob As Integer = 1
Const Min_underlying As Integer = 0
Function fwd(ByVal mk As Market, ByVal opt As opt, ByVal step As Long, ByVal tree As tree) As Double
'Fonction permettant de calculer le prix forward qui prend en inputs, les charactériques de l'option et du marché, le pas de l'abre, et l'arbre

    'Variables pour connaître la position où le dividende sera pris en compte dans _
    le calcul du forward
    Dim propor_date As Double, propor_steps As Double
    Dim step_div As Long 'Le pas où on prend le dividend en compte
    
    'Si dividende n'est pas nul
    If mk.Dividend <> 0 Then
        'Chercher la position où on doit prendre en compte le dividende
        Let propor_date = (mk.Div_date - mk.start_date) / (opt.time * 365)
        Let step_div = WorksheetFunction.RoundDown(propor_date * tree.nbSteps, 0)
        
        'Si ce n'est pas la bonne position, ne pas prendre en compte
        If step <> step_div Then
            fwd = Me.underlying * 1 / mk.DF
        'Si c'est la bonne position
        Else
            Let fwd = Me.underlying * 1 / mk.DF - mk.Dividend
        End If
    'Si le dividend est nul
    Else
        Let fwd = Me.underlying * 1 / mk.DF
    End If
End Function
Sub creation_links(ByVal mk As Market, ByVal opt As opt, ByVal tree As tree, _
    ByVal mid_trunk As Boolean, ByVal up_position As Boolean, ByVal step As Long)
    'Fonction permettant de créer les liens de parenté, et voisinage des noeuds, de calculer les probabilités correspondantes _
    qui prend comme inputs les caractéristiques du marché, de l'option, l'arbre, sa position par rapport au tronc(mid_trunk ou up ou down), le pas de l'abre

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Cas 1 : Le noeud est un mid sur le tronc
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If mid_trunk = True Then
        Call Me.links_mid(mk, opt, step, tree)
        
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Cas 2 : Le noeud est au dessus d'un noeud sur le tronc
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    ElseIf up_position = True Then
        'Associer avec le noeud mid optimal
        Set Me.future_mid = Me.best_futur_mid(mk, opt, tree, up_position, step)
        
        'Associer avec le node_up correspondant
        'si le node_up n'existe pas, le créer, et créer ses liens
        If Me.future_mid.up Is Nothing Then
            Set Me.future_up = make_node(Me.future_mid.underlying * tree.Alpha, tree)
            Set Me.future_mid.up = Me.future_up
            Set Me.future_up.down = Me.future_mid
        
        'Sinon, lier ce node_up
        Else
            Set Me.future_up = Me.future_mid.up
        End If
        
        'Associer avec le node_down correspondant
        Set Me.future_down = Me.future_mid.down

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Cas 3 : Le noeud est en dessous d'un noeud sur le tronc
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Else
        'Associer avec le node mid optimal
        Set Me.future_mid = Me.best_futur_mid(mk, opt, tree, up_position, step)
        
        'Associer avec le node down correspondant, s'il n'existe pas, le créer
        If Me.future_mid.down Is Nothing Then
            Set Me.future_down = make_node(Me.future_mid.underlying / tree.Alpha, tree)
            
            'création des liens haut et bas
            Set Me.future_mid.down = Me.future_down
            Set Me.future_down.up = Me.future_mid
        Else
        'si le noeud existe, le lier
            Set Me.future_down = Me.future_mid.down
        End If
        
        'Associer avec le node_up correspondant
        Set Me.future_up = Me.future_mid.up
    End If
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    ' FINAL : Calculer les probabilités
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Call Me.compute_probabilities(opt, mk, tree, step) 'car c'est lors de la création des "liens" que l'on associe les probas
End Sub
Sub compute_probabilities(ByVal opt As opt, ByVal mark As Market, ByVal tree As tree, ByVal step As Long)
    'Procédure calculant les probabilités up, down et mid _
    qui prend comme inputs les caractéristiques de l'option, du marché, l'arbre, et le pas de l'arbre
    
    Dim var As Double 'Variance
    Dim expect As Double 'Moyenne
    
    'Calcul de la variance et de la moyenne
    var = (Me.underlying) ^ (2) * Exp(2 * mark.InterestRate * tree.Delta_t) _
            * (Exp(mark.Volatility ^ 2 * tree.Delta_t) - 1)

    expect = Me.fwd(mark, opt, step, tree)

    'Calcul des probabilités
    Me.pdown = ((Me.future_mid.underlying ^ (-2)) * (var + expect ^ 2) - 1 - (tree.Alpha + 1) _
    * ((Me.future_mid.underlying) ^ (-1) * expect - 1)) / ((1 - tree.Alpha) * (tree.Alpha ^ (-2) - 1))
    
    Me.pup = (((1 / Me.future_mid.underlying) * expect - 1) - (1 / tree.Alpha - 1) * Me.pdown) / (tree.Alpha - 1)

    Me.pmid = 1 - Me.pdown - Me.pup
End Sub
Sub links_mid(ByVal mk As Market, ByVal opt As opt, ByVal step As Double, ByVal tree As tree)
    'Procédure permettant de créer les noeuds/liens si le noeud est un noeud mid du tronc
    'qui prend comme inputs les caractéristiques du marché, de l'option, le pas de l'arbre, et l'arbre
    
    'Création des noeuds et liens enfants
    Set Me.future_mid = make_node(Me.fwd(mk, opt, step, tree), tree)
    Set Me.future_up = make_node(Me.future_mid.underlying * tree.Alpha, tree)
    Set Me.future_down = make_node(Me.future_mid.underlying / tree.Alpha, tree)
    
    'Création des liens voisins
    Set Me.future_mid.up = Me.future_up
    Set Me.future_mid.down = Me.future_down
    Set Me.future_down.up = Me.future_mid
    Set Me.future_up.down = Me.future_mid
    
    'Lien de parent mid sur le tronc
    Set Me.future_mid.parent_ = Me
End Sub
Function best_futur_mid(ByVal mk As Market, ByVal opt As opt, ByVal tree As tree, _
    ByVal up_position As Boolean, ByVal step As Long) As node
    'Fonction qui à partir du noeud au dessus ou en dessous du noeud_mid_tronc, trouve le futur mid idéal
    
    'Le noeud candidat pour être un mid, 'Noeud mid idéal s'il n'existait pas
    Dim candidate_node As node, new_node As node
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'ETAPE 1 : choisir comme candidate node : le futur mid que l'on aurait _
    eu s'il n'y avait pas de dividende, ou que l'on a en général s'il n'y a pas _
    de dividende
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Commencer par choisir le futur_mid comme celui dans le cas sans dividende
    If up_position Then
        Set candidate_node = Me.down.future_up
    Else
        Set candidate_node = Me.up.future_down
    End If
    
    'Regarder si ce candidate_node est le bon future_mid
    If Me.node_IsIdeal(candidate_node, mk, opt, step, tree) Then
        Set best_futur_mid = candidate_node
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'ETAPE 2 : Si le forward est bien supérieur au candidate_node qu'on a choisi(=futur_mid sans dividend
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    ElseIf (Me.fwd(mk, opt, step, tree) > candidate_node.underlying * (1 + (tree.Alpha - 1) / 2)) Then
        Set best_futur_mid = Find_BestMid_GoUp(candidate_node, mk, opt, step, tree)
        
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'ETAPE 3 : Si le forward est bien inférieur au candidate_node qu'on a
    'choisi(=futur_mid sans dividende)
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Effectuer le même processus qu'avant, en allant vers le bas
    Else
        Set best_futur_mid = Find_BestMid_GoDown(candidate_node, mk, opt, step, tree)
    End If
End Function
Function node_IsIdeal(ByVal candidate As node, ByVal mk As Market, _
    ByVal opt As opt, ByVal step As Long, ByVal tree As tree) As Boolean
    'Fonction booléenne qui montre si le forward est proche du futur_mid que l'on aurait eu sans dividende
    ' qui prend comme input le noeud candidat, les caractéristiques du marché, de l'option, le pas de l'abre, et l'arbre
    
    If ((candidate.underlying * (1 + (tree.Alpha - 1) / 2)) >= Me.fwd(mk, opt, step, tree) _
        And Me.fwd(mk, opt, step, tree) >= (candidate.underlying * (1 - (tree.Alpha - 1) / (2 * tree.Alpha)))) Then
        node_IsIdeal = True
    Else
        node_IsIdeal = False
    End If
End Function
Function Find_BestMid_GoUp(ByVal candidate_node As node, ByVal mk As Market, _
    ByVal opt As opt, ByVal step As Double, ByVal tree As tree) As node
    'Fonction qui permet de trouver le futur_mid en se déplaçant vers le haut
    'qui prend comme inputs le noeud candidat, les caractéristiques du marché, de l'option, le pas de l'arbre, et l'arbre

    Dim new_node As node
    
    'Boucler (passe au node_up) tant que le candidat_mid n'est pas idéal, et stopper si au dessus de ce candidat, il n'y a pas de noeud
    Do While Not (Me.node_IsIdeal(candidate_node, mk, opt, step, tree) Or candidate_node.up Is Nothing)
        Set candidate_node = candidate_node.up
    Loop
    
    'Si le noeud où la boucle s'est arrêté est le futur_mid idéal
    If Me.node_IsIdeal(candidate_node, mk, opt, step, tree) Then
        Set Find_BestMid_GoUp = candidate_node
        
    'Sinon candidate_node devient le down du mid ideal que l'on va créer
    Else
        Set new_node = make_node(candidate_node.underlying * tree.Alpha, tree)
        'Création des liens up et down
        Set new_node.down = candidate_node
        Set candidate_node.up = new_node
        
        'Le bon future mid est ce nouveau noeud créé
        Set Find_BestMid_GoUp = new_node
    End If
End Function
Function Find_BestMid_GoDown(ByVal candidate_node As node, ByVal mk As Market, ByVal opt As opt, _
    ByVal step As Double, ByVal tree As tree) As node
    Dim new_node As node
    'Fonction qui permet de trouver le futur mid en se déplaçant vers le bas
    'qui prend comme inputs le noeud candidat, les caractéristiques du marché, de l'option, le pas de l'arbre, et l'arbre
    
    'Boucler (passe au node_down) tant que le candidat_mid n'est pas idéal, et stopper si en dessous de ce candidat, il n'y a pas de noeud
    Do While Not (Me.node_IsIdeal(candidate_node, mk, opt, step, tree) _
        Or candidate_node.down Is Nothing)
        Set candidate_node = candidate_node.down
    Loop
    
    'Prendre le bon future_mid ou en créer un si besoin
    If Me.node_IsIdeal(candidate_node, mk, opt, step, tree) Then
        Set Find_BestMid_GoDown = candidate_node
    Else
        Set new_node = make_node(candidate_node.underlying / tree.Alpha, tree)
        Set new_node.up = candidate_node
        Set candidate_node.down = new_node
        Set Find_BestMid_GoDown = new_node
    End If
End Function




        
        
        
    

