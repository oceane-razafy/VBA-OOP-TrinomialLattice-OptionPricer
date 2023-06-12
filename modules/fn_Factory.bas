Attribute VB_Name = "fn_Factory"
Option Explicit
Public Function make_node(ByVal underlying As Double, ByVal tree As tree) As node
'Création d'un objet node qui prend en input le sous-jacent et l'arbre
    Dim node As New node
    Let node.underlying = underlying
    Set node.tree = tree
    Set make_node = node
End Function


