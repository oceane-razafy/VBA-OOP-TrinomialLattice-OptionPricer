# Option Pricing using Trinomial Lattice and VBA Code

A VBA project for implementing a trinomial lattice model to price options, developed as part of a master's degree coursework at Paris Dauphine University by another student and myself. Completed in November 2021 and received a grade of 18/20.

In order to make the VBA code visible on GitHub, it is stored in separate `.bas` files within the 'class modules' and 'modules' directories of this repository.

## Table of Contents

- [Project Description](#project-description)
- [Class Modules](#class-modules)
- [Usage](#usage)

## Project Description

This VBA code provides a practical implementation of the trinomial lattice method to price options. By using this code, you can study the convergence of the option price to the Black-Scholes formula, analyze the price differences between American and European options, and understand the impact of different parameters (strike price, maturity, volatility, interest rate) on option pricing.

## Class Modules
1. <u>Market Class Module</u>: This class module represents the market parameters for option pricing. It includes attributes such as InterestRate, Volatility, Dividend, StartPrice, DF (Discount Factor), start_date, and Div_date. The FillMarket procedure is used to fill the values of these attributes based on the input data from the **"Pricer"** sheet.
2. <u>Node Class Module</u>: This class module represents a single node in the trinomial lattice with its attributes.
3. <u>Opt Class Module</u>: This class module represents the option contract. It includes attributes such as strike, maturity, time, isAmerican, and isCall. The FillOption procedure is used to fill the values of these attributes based on the input data from the **"Pricer"** sheet.
4. <u>Tree Class Module</u>: This class module represents the trinomial lattice structure and provides methods to build the lattice and calculate option prices.

## Usage

1. Open the Excel file.
2. Fill in the input parameters in the **"Pricer"** sheet:
   - Set the option details:
     - **Strike**: Set the strike price.
     - **Maturity**: Set the maturity date.
     - **Time (Years)**: Specify the time to maturity in years.
     - **IsAmerican**: Check the box "FAUX" (False) for an European option or uncheck it for an American option.
     - **IsCall**: Check the box "VRAI" (True) for a Call option, or uncheck it for a Put option.
   - Set the market parameters:
		- **Interest Rate**: Enter the interest rate in percentage.
		- **Volatility**: Specify the volatility of the underlying asset in percentage.
		- **Dividend**: Enter the dividend yield of the underlying asset.
		- **Start Price**: Set the starting price of the underlying asset.
		- **Discount Factor**: Enter the discount factor. Its value should be entered as a decimal, such as 0.99920032.
		- **Start Date**: Specify the starting date for the option pricing calculations.
		- **Dividend Ex-Date**: Enter the ex-dividend date. This is important for accurately valuing options with dividend-paying underlying assets.
   - Set the number of steps (**Nb Steps**): Specify the number of steps in the trinomial lattice.
3. Click the **"Click here to price"** button in the **"Pricer"** sheet to execute the macro and generate the output.
4. In the **"Price"** field, review the calculated option price and compare it with the Black-Scholes Price.
4. The other sheets display the ouputs or provide key analyses: 
   - **Underlying Graph**: Shows the graphical representation of the trinomial lattice for the underlying asset.
   - **Option Graph**: Displays the graphical representation of the trinomial lattice for the option.
   - **Tree vs. BS Model**: Compares the option prices obtained from the trinomial lattice with the Black-Scholes model.
   - **European vs. American**: Compares the prices of European and American options.
   - **Impact of Interest Rate**: Analyzes the impact of changing interest rates on option pricing.
   - **Impact of Volatility**: Analyzes the impact of changing volatility on option pricing.
   - **Impact of Maturity**: Analyzes the impact of changing the time to maturity on option pricing.

_Note: Ensure that the necessary Excel calculations and macros are enabled for the code to execute properly. If prompted, enable macros and click "Enable Content" to proceed with the execution._