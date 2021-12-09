"""
Generalized RAS (GRAS) implementation in Python, following the guidelines of:
    Temurshoev, U., R.E. Miller and M.C. Bouwmeester (2013), A note on the GRAS method, Economic Systems Research, 25.
    Alves-Passoni, P., Freitas, F. (2020), Estimação de Matrizes Insumo-Produto anuais para o Brasil no
        Sistema de Contas Nacionais Referência 2010, Texto para Discussão 025, IE-UFRJ

In "main.py", the routine is specifically made for the uses in Alves-Passoni, Alves (from hereon, APF), which means
that the GRAS adjustment is made for each product. Therefore, the exogenously given row and column totals are:
  Row: TRU values for National Supply at Base Prices, Total Imports, Trade Margins, Transport Margins and Net Taxes.
  Columns: TRU values for each sector's Total Supply at Consumer Prices of the specific product.

However, the 'gras' function was made to be as flexible as possible, allowing for easy changes to the code
and for applications in any set of matrix, given that the row and column constraints are correctly supplied.
To conclude, all parts of the script were commented in order to make it easier to interpret the routine :)
----------------------------------------------------------------------------------------------------------------------
In terms of terminology, the function and its variables follows Temurshoev et al. (2013) (from hereon, TMB) as closely
as possible. When this doesn't happen, a comment will be made referring to the analogous variable in TMB's paper.
----------------------------------------------------------------------------------------------------------------------
Made by: Vinícius de Almeida Nery Ferreira (Ipea-DF); e-mail: vinicius.nery@ipea.gov.br or vnery5@gmail.com
Original Matlab function written by Umed Temurshoev; e-mail: umed.temurshoev@ec.europa.eu
R script made by Peter Horvat; e-mail: peter.horvat@oecd.org
The author acknowledges and thanks the help of Patieene Alves-Passoni.
----------------------------------------------------------------------------------------------------------------------
"""

## All calculations can be done using only numpy
import numpy as np

def gras(mA, vRowRestriction, vColRestriction, nEpsilon=10**(-8), nMaxIter=10000):
    """
    Estimates a new matrix (mX) with exogenously given row and column totals (restrictions) that is a close as possible
    to the given initial matrix (mA)
    :param mA: matrix that is to be adjusted
    :param vRowRestriction: vector containing row totals (addition across columns) that the new matrix has to match.
        Equivalent to 'u' in TMB's terminology.
    :param vColRestriction: vector containing columns totals (addition across rows) that the new matrix has to match.
        Equivalent to 'v' in TMB's terminology.
    :param nEpsilon: max error allowed for each column total. Defaults to 0.00000001.
    :param nMaxIter: max number of iterations to reach new matrix. Defaults to 1000.
    :return:
        mX: new matrix adjusted via GRAS whose row and column total match those of the restrictions;
        r: row multipliers (substitution effects)
        s: column multipliers (fabrication effects)
        nIterations: number of iterations necessary to achieve mX
    """

    ### Error handling ###########################################################################################
    ## Substituting names to match TMB's terminology and reshaping to a proper vector (if necessary)
    u = vRowRestriction.reshape(-1)
    v = vColRestriction.reshape(-1)

    ## Checking to see if vectors are of the same shape as the given matrix
    # Number of rows
    if mA.shape[0] != u.shape[0]:
        nRowsMA = mA.shape[0]
        nRowsU = u.shape[0]

        print(f"The vector of row restrictions u (size {nRowsU}) doesn't match the number of rows of mA ({nRowsMA}).")
        return None, None, None, None

    # Number of columns
    elif mA.shape[1] != v.shape[0]:
        nColsMA = mA.shape[1]
        nColsV = v.shape[0]

        print(f"The vector of col restrictions v (size {nColsV}) doesn't match the number of cols of mA ({nColsMA}).")
        return None, None, None, None

    """
    ## Checking to see if the constrains are 0, but the matrix contains non-zero elements
    # Rows
    for nRow in range(u.shape[0]):
        if u[nRow] == 0 and np.sum(mA[nRow, :]) != 0:
            print(f"Row {nRow} has a zero constrain, but initial matrix contains non-zero values in that row.")

    # Columns
    for nCol in range(v.shape[0]):
        if v[nCol] == 0 and np.sum(mA[:, nCol]) != 0:
            print(f"Column {nCol} has a zero constrain, but initial matrix contains non-zero values in that column.")
    """

    ### Iteration Procedure ###########################################################################################
    ## Defining function (inverts a vector and then diagonalizes it, replacing initial 0s with 1s)
    def inv_diag(vInput, nNaN=1):
        """
        Element by element inversion of a vector (1 / a_ij for every a_ij) and diagonalization into a square matrix
        :param vInput: vector to be element-wise inverted and diagonalized
        :param nNaN: substitute NaNs for what number? If None, don't substitute
        :return:
            mOutput: diagonal matrix of element-wise inverted vector
        """

        ## Inverting and substituting NaNs
        vOutput = 1 / vInput if nNaN is None else np.nan_to_num((1 / vInput), nan=nNaN, posinf=nNaN, neginf=nNaN)
        
        ## Returning diagonalized matrix
        return np.diagflat(vOutput)

    ## Decomposing initial matrix into positive and (the absolute values of) negative elements
    mP = np.where(mA > 0, mA, 0)
    mN = np.where(mA < 0, -mA, 0)

    ## Creating vector full of ones
    # Shape of rows
    vOnesRows = np.array([1] * mA.shape[0])

    # Shape of columns
    vOnesCols = np.array([1] * mA.shape[1])

    ## Initial guess for r (vector of 1s), as suggested by Junius and Oosterhaven (2003)
    r = vOnesRows

    ## Variables necessary for column multipliers calculation (see equation 7)
    pr = mP.T.dot(r)  # p_j(r) as defined in equation 7 (vector of length nColumns)
    nr = mN.T.dot(inv_diag(r)).dot(vOnesRows)  # n_j(r) as defined in equation 7 (vector of length nColumns)

    ## Calculating first step column multiplier (s1) - see equation 9b
    s1 = inv_diag(2 * pr).dot(v + np.sqrt(v ** 2 + 4 * pr * nr))
    # Solving for the values where column totals are negative
    s1 = np.where(pr == 0, -inv_diag(v).dot(nr), s1)

    ## Variables necessary for row multipliers calculation
    ps = mP.dot(s1)  # p_i(s) as defined in equation 7 (vector of length nRows)
    ns = mN.dot(inv_diag(s1)).dot(vOnesCols)  # n_i(s) as defined in equation 7 (vector of length nRows)

    # Calculating first step row multiplier - see equation 9a
    r = inv_diag(2 * ps).dot(u + np.sqrt(u ** 2 + 4 * ps * ns))
    # Solving for the values where row totals are negative
    r = np.where(ps == 0, -inv_diag(u).dot(ns), r)

    ## Calculating second step column multiplier (s)
    pr = mP.T.dot(r)  # p_j(r) as defined in equation 7
    nr = mN.T.dot(inv_diag(r)).dot(vOnesRows)  # n_j(r) as defined in equation 7

    ## Calculating second step column multiplier (s2) - see equation 9b
    s2 = inv_diag(2 * pr).dot(v + np.sqrt(v ** 2 + 4 * pr * nr))
    s2 = np.where(pr == 0, -inv_diag(v).dot(nr), s2)

    ## Compute difference between column multipliers s2 and s1 elements
    dif = s2 - s1
    # Max difference (will be compared to nEpsilon)
    M = np.max(np.abs(dif))

    ## Loop while the user-defined threshold for nEpsilon is not achieved
    # Creating number of iterations and starting loop
    nIterations = 1

    while M > nEpsilon:
        ## Error checking (matrix not converging)
        if nIterations == nMaxIter:
            print(f"Matrix is not converging :( (passed {nMaxIter} iterations).")
            return None, None, None, nMaxIter

        ## Initial column multiplier = column multiplier for last iteration
        s1 = s2

        ## Variables necessary for row multipliers calculation
        ps = mP.dot(s1)  # p_i(s) as defined in equation 7
        ns = mN.dot(inv_diag(s1)).dot(vOnesCols)  # n_i(s) as defined in equation 7

        # Calculating row multipliers
        r = inv_diag(2 * ps).dot(u + np.sqrt(u ** 2 + 4 * ps * ns))
        r = np.where(ps == 0, -inv_diag(u).dot(ns), r)

        ## Calculating second step column multiplier (s)
        pr = mP.T.dot(r)  # p_j(r) as defined in equation 7
        nr = mN.T.dot(inv_diag(r)).dot(vOnesRows)  # n_j(r) as defined in equation 7

        # Calculating column multipliers
        s2 = inv_diag(2 * pr).dot(v + np.sqrt(v ** 2 + 4 * pr * nr))
        s2 = np.where(pr == 0, -inv_diag(v).dot(nr), s2)

        ## Compute difference between s2 and s1 elements
        dif = s2 - s1

        ## Adding to the number of iterations
        nIterations += 1

        # Max difference between the multipliers
        M = np.max(np.abs(dif))

    ## Last-stage row multiplier
    s = s2

    # Variables necessary for the last row multiplier
    ps = mP.dot(s)  # p_i(s) as defined in equation 7
    ns = mN.dot(inv_diag(s)).dot(vOnesCols)  # n_i(s) as defined in equation 7

    # Calculating last step row multiplier
    r = inv_diag(2 * ps).dot(u + np.sqrt(u ** 2 + 4 * ps * ns))
    r = np.where(ps == 0, -inv_diag(u).dot(ns), r)

    ## Calculating updated matrix (see equations 4a and 4b)
    mX = np.diagflat(r).dot(mP).dot(np.diagflat(s)) - inv_diag(r).dot(mN).dot(inv_diag(s))

    return mX, r, s, nIterations
