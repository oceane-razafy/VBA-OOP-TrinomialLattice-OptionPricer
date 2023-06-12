VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Alpha As Double 'Alpha de l'abre
Public root As node 'Racine
Public nbSteps As Long 'Nombre de pas de l'arbre
Public Delta_t As Double 'Le delta_t

Const nb_root As Integer = 1
Const max_links As Integer = 3
Const min_parents As Integer = 1
Sub compute_alpha(ByVal mark As Market, ByVal opt As opt)
'Procédure permettant de calculer l'alpha de l'arbre
    'Plusieurs formules sont possibles :
    ' 1) Alpha = Exp(mark.InterestRate * Me.Delta_t + mark.Volatility * Sqr(3 * Me.Delta_t))
    '2)
    Alpha = 1 + Sqr(3 * (Exp(mark.Volatility ^ 2 * _
    (Me.Delta_t)) - 1)) * Exp(mark.InterestRate * (Me.Delta_t))
End Sub
Sub BuildCol(ByRef node_trunk As node, ByVal opt As opt, ByVal mk As Market, ByVal step As Long)
    'Procédure permettant de construire les colonnes de noeuds d'un arbre _
    qui prend comme inputs le noeud tronc, les charactérisques de l'option et du marché, le pas de l'arbre
    
    'Variables pour montrer et descendre dans la colonne
    Dim node_up As node
    Dim node_down As node
    
    'Création des liens fils
    Call node_trunk.creation_links(mk, opt, Me, True, False, step)
       
    'Initialisation des variables
    Set node_up = node_trunk.up
    Set node_down = node_trunk.down
    
    'Création pour la partie haute de la colonne
    Do
        Call node_up.creation_links(mk, opt, Me, False, True, step)
        Set node_up = node_up.up
    Loop Until node_up Is Nothing
    
    'Création pour la partie basse de la colonne
    Do
        Call node_down.creation_links(mk, opt, Me, False, False, step)
        Set node_down = node_down.down
    Loop Until node_down Is Nothing
End Sub
Sub TreeBuild(ByVal opt As opt, ByVal mk As Market)
    'Procédure permettant de créer l'arbre entièrement _
    qui prend comme inputs les caractéristiques de l'option et du marché

    Dim step As Long 'Pas de l'arbre
    Dim node_trunk As New node 'Noeud tronc
    Dim under As Double 'Sous-jacent
    
    'Prendre pour la racine, le S_0
    under = mk.StartPrice
    
    'Création de la racine
    Set Me.root = make_node(under, Me)
    Call Me.root.creation_links(mk, opt, Me.root.tree, True, False, 0)
    
    'Commencer par le 1er future_mid
    Set node_trunk = Me.root.future_mid

    'Boucle sur les pas de l'arbres
    'on commence à "mid 1" -> celui à côté de la racine
    For step = 1 To Me.nbSteps - 1
    
        'Création d'une colonne
        Call Me.BuildCol(node_trunk, opt, mk, step)
        
        'on passe au mid_trunk suivant
        Set node_trunk = node_trunk.future_mid
    Next step
End Sub
Sub Pricer(ByRef root As node, ByVal opt As opt, mk As Market)
    'Procédure pricant l'option (Mouvement Backward : en partant de la fin de l'arbre) _
    qui prend comme input le noeud racine, les caractéristiques de l'option et du marché
        
    'Variables pour se déplacer sur l'arbre
    Dim node_up As node, node_down As node, n_trunk As node, last_node_trunk As node
    
    'Pour ajuster le calcul du prix en fonction du type d'option
    Dim sign_ As Integer
    
    'Pas de l'arbre, prix de la racine
    Dim step As Integer, root_price As Double
     
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'I. Calcul du payoff sur la dernière colonne de l'arbre
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Si c'est un call ou un put
    sign_ = SignOption(opt)
    
    'On initalise last_node_trunk avec le noeud racine
    Set last_node_trunk = root
    
    'On boucle avec last_node_trunk pour arriver au dernier noeud du tronc
    Do
        Set last_node_trunk = last_node_trunk.future_mid
    Loop Until last_node_trunk.future_mid Is Nothing
       
    'Attribuer le payoff du dernier noeud tronc
    Let last_node_trunk.Value = WorksheetFunction.Max((last_node_trunk.underlying - opt.strike) * sign_, 0)
    Let last_node_trunk.IsValueAvailable = True
    Set node_up = last_node_trunk.up
    Set node_down = last_node_trunk.down
    
     'Attribuer tous les autres payoffs de la dernière colonne
     '______________Partie haute de la colonne_________________
      Call Me.Price_up(node_up, opt, mk, sign_, True)
     
     '______________Partie basse de la colonne_________________
     Call Me.Price_down(node_down, opt, mk, sign_, True)
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'II. Attribution des prix des toutes les autres colonnes, sauf la racine
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Attribution du  mid de l'avant derniere colonne
    Set n_trunk = last_node_trunk.parent_
    
    'Boucle jusqu'avant la racine
    For step = 1 To Me.nbSteps - 1
           '______________________Partie tronc_________________________
        Call Me.Price_Trunk(n_trunk, opt, mk, sign_)
        'Valeur de l'option au-dessus et en-dessous du tronc mid de la colonne sur laquelle on est
        Set node_up = n_trunk.up
        Set node_down = n_trunk.down
        '______________________Partie au dessus du tronc_________________________

        Call Me.Price_up(node_up, opt, mk, sign_)
        '______________________Partie en dessous du tronc_________________________
        Call Me.Price_down(node_down, opt, mk, sign_)
        'On recule sur le tronc d'une colonne
        Set n_trunk = n_trunk.parent_
    Next step
    
    'Attribution du prix de l'option dans la racine
    With n_trunk
         .Value = (.future_up.Value * .pup + .future_down.Value * .pdown + .future_mid.Value * .pmid) * mk.DF
         If opt.isAmerican Then
             .Value = WorksheetFunction.Max(.Value, (.underlying - opt.strike) * sign_)
         End If
         .IsValueAvailable = True
     End With
End Sub
Function SignOption(ByRef opt As opt) As Integer
    'Fonction permettant d'avoir d'avoir le signe permettant d'ajuster _
    le calcul du payoff en fonction du type d'option : Call ou Put
    'qui prend comme inputs les caractéristiques de l'option
    
    If opt.isCall Then
        SignOption = 1
    Else
        SignOption = -1
    End If
End Function
Sub Price_Trunk(ByRef n_trunk As node, ByRef opt As opt, ByRef mk As Market, sign_ As Integer)
    'Procédure permettant de pricer l'option sur le tronc de l'arbre
    'qui prend comme inputs le noeud sur le tronc, les caractéristiques de l'option, du marché, et le signe pour le calcul du payoff
    
    With n_trunk
        .Value = (.future_up.Value * .pup + .future_down.Value * .pdown + .future_mid.Value * .pmid) * mk.DF
        'Si l'option est Américaine ou non
        If opt.isAmerican Then
            .Value = WorksheetFunction.Max(.Value, (.underlying - opt.strike) * sign_)
        End If
        'Mettre qu'un prix de l'option a bien été calculé
        .IsValueAvailable = True
    End With
End Sub
Sub Price_up(ByRef node_up As node, ByRef opt As opt, ByRef mk As Market, ByRef sign_ As Integer, _
    Optional Last_col As Boolean = False)
    'Procédure permettant de pricer tous les noeuds au dessus du tronc
    'qui prend comme inputs le noeud, les caractéristiques de l'option, du marché, le signe pour le payoff _
    'un booléen permettant de savoir si on est à la dernière colonne ou non
    
    'Calcul pour la dernière colonne
    If Last_col Then
        Do
            node_up.Value = WorksheetFunction.Max((node_up.underlying - opt.strike) * sign_, 0)
            node_up.IsValueAvailable = True
            Set node_up = node_up.up
        Loop Until node_up Is Nothing
    'Calcul pour les autres colonnes
    Else
        Do
            With node_up
            'Valeur de l'option
                .Value = (.future_up.Value * .pup + .future_down.Value * .pdown + .future_mid.Value * .pmid) * mk.DF
                'Si l'option est Américaine ou non
                If opt.isAmerican Then
                    .Value = WorksheetFunction.Max(.Value, (.underlying - opt.strike) * sign_)
                End If
                'Mettre qu'un prix de l'option a bien été calculé
                .IsValueAvailable = True
                Set node_up = .up
            End With
        Loop Until node_up Is Nothing
    End If
End Sub
Sub Price_down(ByRef node_down As node, ByRef opt As opt, ByRef mk As Market, ByRef sign_ As Integer, _
    Optional Last_col As Boolean = False)
    'Procédure permettant de pricer tous les noeuds en dessous du tronc
    'qui prend comme inputs le noeud, les caractéristiques de l'option, du marché, le signe pour le payoff _
    'un booléen permettant de savoir si on est à la dernière colonne ou non
    
    'Calcul pour la dernière colonne
    If Last_col Then
        Do
            node_down.Value = WorksheetFunction.Max((node_down.underlying - opt.strike) * sign_, 0)
            node_down.IsValueAvailable = True
            Set node_down = node_down.down
        Loop Until node_down Is Nothing
    'Calcul pour les autres colonnes
    Else
        Do
            With node_down
            'Valeur de l'option
                .Value = (.future_up.Value * .pup + .future_down.Value * .pdown + .future_mid.Value * .pmid) * mk.DF
                'Si l'option est Américaine ou non
                If opt.isAmerican Then
                    .Value = WorksheetFunction.Max(.Value, (.underlying - opt.strike) * sign_)
                End If
                'Mettre qu'un prix de l'option a bien été calculé
                .IsValueAvailable = True
                Set node_down = .down
            End With
        Loop Until node_down Is Nothing
    End If
End Sub
Function LastColSizeUp(ByVal root As node) As Long
    'Fonction qui à partir du noeud racine, renvoie le nombre de noeuds au dessus du noeud tronc _
    à la dernière colonne pour pouvoir positionner les graphiques plus tard

    'Variable pour trouver le dernier noeud tronc
    Dim last_node_trunk As node
    
    'Variable pour montrer dans la colonne
    Dim node_up As node
    
    'Compteur du nombre de noeuds
    Dim nb_up As Long
    
    '................Trouver le dernier noeud du tronc.............
    'Commencer avec la racine
    Set last_node_trunk = root
    
    'Boucler jusqu'à la dernière colonne
    Do
        Set last_node_trunk = last_node_trunk.future_mid
    Loop Until last_node_trunk.future_mid Is Nothing
       
    Set node_up = last_node_trunk.up
    
    '...........Compter le nombre de noeuds au dessus du dernier noeud tronc.................
    nb_up = 1
    
    'Boucler jusqu'en haut de la colonne
    Do Until node_up Is Nothing
        Set node_up = node_up.up
        nb_up = nb_up + 1
    Loop
        
    nb_up = nb_up - 1
    
    Let LastColSizeUp = nb_up
End Function
Sub FreeTree(ByRef tree As tree)
    'Procédure permettant de libérer la mémoire pour faire plusieurs pricing
    'qui prend comme input l'arbre à libérer

    'Le pas de l'arbre
    Dim step As Long
    
    'Les variables noeuds pour bouger dans l'arbre
    Dim last_node_trunk As node, n As node, n_trunk As node
    
    'Initialiser last_node_trunk avec le noeud racine
    Set last_node_trunk = tree.root
    
    'Arriver au dernier noeud du tronc
    Do
        Set last_node_trunk = last_node_trunk.future_mid
    Loop Until last_node_trunk.future_mid Is Nothing
        
    'Attribution du  noeud mid (tronc) de l'avant dernière colonne
    Set n_trunk = last_node_trunk.parent_
    
    'Boucle sur les pas de l'arbre
    For step = 1 To tree.nbSteps - 1
    
        Set n = n_trunk.up
        
        'Arriver en haut de la colonne de noeud
        Call Me.GoUp(n)
        
        'Redescendre en libérant la mémoire jusq'au tronc
        Do
            Call Me.Free_n(n)
            Set n.up = Nothing
            Set n = n.down
        'Stopper quand on arrive au tronc
        Loop Until n Is n_trunk
        
        'Arriver en bas de la colonne
        Call Me.GoDown(n)
        
        'Remonter en libérant la mémoire jusq'au tronc
        Do
            Call Me.Free_n(n)
            Set n.down = Nothing
            Set n = n.up
        'Stopper quand on arrive au tronc
        Loop Until n Is n_trunk
        
        'Enlever les liens du noeud tronc
        Call Me.Free_n(n)
        Set n.up = Nothing
        Set n.down = Nothing
        'Passer au noeud tronc parent
        Set n_trunk = n_trunk.parent_
    Next step
    
    'Libère le noeud racine et ses liens
    Set n_trunk.future_mid.parent_ = Nothing
    Call Me.Free_n(n_trunk)
    Set tree.root = Nothing
    Set tree = Nothing
End Sub
Sub Free_n(ByRef n As node)
    'Procédure permettant de libérer la mémoire pour une partie de l'arbre
    Set n.future_up = Nothing
    Set n.future_mid.down = Nothing
    Set n.future_mid.up = Nothing
    Set n.future_mid = Nothing
    Set n.future_down = Nothing
End Sub
Sub GoUp(ByRef n As node)
    'Procédure qui à partir du n_trunk.up permet d'arriver en haut de la colonne
    Do While Not n.up Is Nothing
        Set n = n.up
    Loop
End Sub
Sub GoDown(ByRef n As node)
    'Procédure qui à partir du noeud tronc permet d'arriver en bas de la colonne
    Do While Not n.down Is Nothing
        Set n = n.down
    Loop
End Sub

