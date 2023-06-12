VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Market"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public InterestRate As Single
Public Volatility As Single
Public Dividend As Double
Public StartPrice As Double
Public DF As Double
Public Div_date As Date
Public start_date As Date
Sub FillMarket(ByVal InterestRate As Range, ByVal Volatility As Range, ByVal Dividend As Range, _
    ByVal StartPrice As Range, ByVal DF As Range, ByVal start_date As Range, ByVal Div_date As Range)
    'Procédure permettant de remplir les valeurs des attributs de l'instance de la classe Market _
    qui prend en inputs les données choisies sur la feuille "Pricer"
    
    Me.InterestRate = InterestRate.Value 'Taux d'intérêt
    Me.Volatility = Volatility.Value 'Volatilité
    Me.Dividend = Dividend.Value 'Dividende
    Me.StartPrice = StartPrice.Value 'Prix de départ du sous-jacent
    Me.DF = DF.Value 'Discount Factor
    Me.start_date = start_date.Value 'La date de départ
    Me.Div_date = Div_date.Value 'La date ex-dividende
End Sub
