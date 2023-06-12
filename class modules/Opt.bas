VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Opt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strike As Single
Public maturity As Single
Public time As Double
Public isAmerican As Boolean
Public isCall As Boolean
Sub FillOption(ByVal strike As Range, ByVal maturity As Range, ByVal time As Range, ByVal isAmerican As Range, isCall As Range)
    'Proc�dure permettant de remplir les valeurs des attributs de l'instance de la classe Option _
    qui prend en inputs les donn�es choisies sur la feuille "Pricer"

    Me.strike = strike.Value 'Strike
    Me.maturity = maturity.Value 'Maturit�
    Me.time = time.Value 'Temps en ann�es (entre start_date et maturity)
    Me.isAmerican = isAmerican.Value 'Si c'est option Am�ricaine ou Europ�enne
    Me.isCall = isCall.Value 'Si c'est option Call ou Put
End Sub
