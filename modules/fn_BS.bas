Attribute VB_Name = "Fn_BS"
Option Explicit
Public Function Price_BS(ByVal opt As opt, ByVal mk As Market) As Double
    'Fonction permettant de calculer le prix Black & Scholes de l'option qui _
    prend en inputs les caractéristiques de l'option et du marché
    
    'Les paramètres du calcul
    Dim d_1 As Double
    Dim d_2 As Double
    Dim ln_SK As Double, arg_2 As Double
    Dim DF As Double
    
    'Calcul des parties de la formule de BS
    Let ln_SK = WorksheetFunction.Ln(mk.StartPrice / opt.strike)
    Let arg_2 = opt.time * (mk.InterestRate + (mk.Volatility ^ 2) / 2)
    Let d_1 = (ln_SK + arg_2) / ((opt.time ^ (1 / 2)) * mk.Volatility)
    Let d_2 = d_1 - ((opt.time ^ (1 / 2)) * mk.Volatility)
    Let DF = Exp(-mk.InterestRate * opt.time)
    
    'Prix en fonction du type de l'option : Call ou Put
    If opt.isCall Then
        Price_BS = mk.StartPrice * WorksheetFunction.Norm_Dist(d_1, 0, 1, True) _
        - opt.strike * DF * WorksheetFunction.Norm_Dist(d_2, 0, 1, True)
    Else
        Price_BS = -mk.StartPrice * WorksheetFunction.Norm_Dist(-d_1, 0, 1, True) _
        + opt.strike * DF * WorksheetFunction.Norm_Dist(-d_2, 0, 1, True)
    End If
End Function


