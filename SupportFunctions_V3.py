"""
Auxiliary functions for Main.py (in order to keep the main code clean).

Annual estimation for Input Output Tables based on the methodology proposed by Guilhoto (2010) and Barry-Miller (2009).

Based on official Resources and Uses tables published by IBGE in the System of National Accounts (3rd edition).

Authors: João Maria de Oliveira and Vinícius de Almeida Nery Ferreira (Ipea-DF).

E-mails: joao.oliveira@ipea.gov.br and vinicius.nery@ipea.gov.br (or vnery5@gmail.com).
"""

## Importing pandas and numpy
import numpy as np
import pandas as pd

# ============================================================================================
def read_file_excel(sDirectory, sFileName, sSheetName):
    """
    Reads the content of an Excel file.
    :param sDirectory: directory of the spreadsheet file
    :param sFileName: name of the spreadsheet
    :param sSheetName: sheet name
    :return: dataframe containing the contents of the spreadsheet
    """

    sFile = sDirectory + sFileName
    mSheet = pd.read_excel(sFile, sheet_name=sSheetName, header=None)
    return mSheet

# ============================================================================================

def load_tru(sDirectoryInput, sFile, sSheetName, nRowIni, nColIni, nRows, nCols, bNames=False, bCodes=False):
    """
    Reads the respective tru and returns the correspondent numpy array
    :param sDirectoryInput: directory containing the files
    :param sFile: name of the Excel file
    :param sSheetName: sheet that contains the data
    :param nRowIni: initial row to read from (Excel - 1)
    :param nColIni: initial column to read from (Excel - 1)
    :param nRows: number of rows to read
    :param nCols: number of columns to read
    :param bNames: whether to get names of the rows/columns
    :param bCodes: whether to get codes of the rows/columns
    :return:
        mTRU: desired TRU array \n
        Optional: Names and Columns of the desired TRU array
    """

    ## Reading files as a dataframe
    dSheet = read_file_excel(sDirectoryInput, sFile, sSheetName)

    ## Getting only the necessary values
    mTRU = dSheet.values[nRowIni:nRowIni + nRows, nColIni:nColIni + nCols]
    # Transforming into floats
    mTRU = mTRU.astype(float)

    ### Reading names and codes (if requested)
    if bNames:
        ## Reading row's names
        vNameRows = list(dSheet.values[nRowIni:nRowIni + nRows, nColIni - 1])

        # Columns: we use the shape to determine whether they are sectors or other names and read accordingly
        if mTRU.shape[1] >= 12 and mTRU.shape[1] != 51:
            # Reading sector's names
            vNameCols = dSheet.values[nRowIni - 2, nColIni:nColIni + nCols]
            vNameCols = [s[5:] for s in vNameCols]
        else:
            vNameCols = list(dSheet.values[nRowIni - 2, nColIni:nColIni + nCols])

        # Removing \n and putting in spaces
        vNameCols = [s.replace("\n", " ") for s in vNameCols]

        ## Reading codes (if requested)
        if bCodes:
            # Rows (usually, products) codes (one column before the names)
            vCodeRows = list(dSheet.values[nRowIni:nRowIni + nRows, nColIni - 2])

            # Columns (usually, sectors) codes
            vCodeCols = dSheet.values[nRowIni - 2, nColIni:nColIni + nCols]
            vCodeCols = [s[:4] for s in vCodeCols]

            return mTRU, vNameRows, vNameCols, vCodeRows, vCodeCols
        ## If not requested, return only names
        else:
            return mTRU, vNameRows, vNameCols
    ## If neither names nor codes are requested, return only the names
    else:
        return mTRU

# ============================================================================================
def distribution_matrix_calcul(mIntermConsum, mDemand):
    """
    Calculates the percentage of the consumption of the product i by sector j
    relative to the total consumption of product i by all sectors.
    :param mIntermConsum: intermediate consumption matrix/array
    :param mDemand: final demand matrix/array
    :return:
        mDistribution: consumption array in relative/percentage terms
        mTotalConsum: mIntermConsum and mDemand concatenated horizontally
    """

    ## Disabling warnings for division by 0 (they are handled after)
    np.seterr(divide='ignore', invalid='ignore')

    ## Concatenating mIntermConsum and mDemand to get the total consumption matrix
    mTotalConsum = np.concatenate((mIntermConsum, mDemand), axis=1)

    ## for each product, calculate total consumption (intermediate and final) \
    # and assign it to a nProducts x 1 vector
    vTotalProducts = np.sum(mTotalConsum, axis=1)

    ## Dividing each element by the total production/demand of sector i
    mDistribution = mTotalConsum / vTotalProducts[:, None]
    # Checking for nan's and infs
    mDistribution = np.nan_to_num(mDistribution, nan=0, posinf=0, neginf=0)

    return mDistribution, mTotalConsum

# ============================================================================================
def calculation_margin(mAlpha, mMatrixInput, nColRef, vRowErase):
    """
    Returns the margin matrix/array (transport, trade, IPI, ICMS and other taxes).
    An important point: alongside other calculations, it assumes that margins are distributed
    in the same way as the market price consumption of each product.
    :param mAlpha: distribution matrix, calculated using the distribution_matrix_calc function
    :param mMatrixInput: supply matrix
    :param nColRef: number of the desired column (transport, trade...)
    :param vRowErase: rows to exclude
    :return: mMatrixOutput: nProduct x nSector array containing the margins of each product and sector
    """

    ## Creating reference vector
    vCol = mMatrixInput[:, nColRef]

    ## Creating range
    vRowErase = np.arange(vRowErase[0], vRowErase[1] + 1)

    ## Adding up all the margins of trade/transport (which are negative in the TRUs)
    nTotMargin = np.sum(vCol[vRowErase])

    ## Multiplying each element of the distribution matrix by the extended vector (column-wise)
    # This will create the distribution of the vector across all products and sectors
    mMatrixOutput = vCol[:, None] * mAlpha

    ## For each sector, get the total margins paid in all products EXCEPT those of trade/transport
    # Getting correspondence of product NOT in trade/transport
    vMarginSectors = np.ones(mMatrixOutput.shape[0], dtype=bool)
    vMarginSectors[vRowErase] = False

    # For each sector (column), add up the results of the above products (not included in trade or transport)
    vTotMarginSector = np.sum(mMatrixOutput[vMarginSectors, :], axis=0)

    ## Get percentage share of margins held by each trade/transport product (compared to the margin's total)
    vMultip = vCol[vRowErase] / nTotMargin
    vMultip = np.nan_to_num(vMultip, nan=0, posinf=0, neginf=0)

    # Forcing the sum of margins to be 0 for each sector, multiplying the share of trade/transport margins
    # of each trade/transport product by the total margins paid by sector i = j
    mMatrixOutput[vRowErase, :] = -vMultip[:, None] * vTotMarginSector[None, :]

    return mMatrixOutput

# ============================================================================================
def calculation_internal_matrix(mAlpha, mMatrixInput, nColRef):
    """
    Given the percentage of consumption/production of each product in each sector,
    estimate, assuming the same distribution, other distributions (import, taxes...)
    :param mAlpha:
        distribution matrix of the consumption of product i by sector j, relative
        to the total product i consumption by all sectors.
        Calculated using the distribution_matrix_calc function.
    :param mMatrixInput: vector to be estimated (or matrix containing one column to be estimated)
    :param nColRef: column number (0 for vectors, an integer for a matrix)
    :return: mMatrixOutput: estimated matrix (product x sector)
    """

    ## Creating reference vector
    # This will create the distribution of the vector across all products and sectors
    vCol = mMatrixInput[:, nColRef]

    ## Multiplying each element of the distribution matrix by the extended vector (column-wise)
    mMatrixOutput = vCol[:, None] * mAlpha

    return mMatrixOutput

# ============================================================================================
def payment_sector_total(mInput, nSectors):
    """
    For a given payment sector matrix (imports and taxes), calculates the total vector
    for intermediate consumption and final demand, as well as the total of both of the aforementioned components.

    :param mInput: matrix of the payment sector component estimated using the Alpha matrix
    :param nSectors: number of sectors
    :return:
        vTotOutput: (nSectors + 9 x 1) vector containing intermediate consumption (IC), total IC,
        final demand (FD), total FD and total IC + total FD
        vTotIC (nSectors x 1) vector containing intermediate consumption
        vTotFD (7 x 1) vector containing final demand components
    """

    ## Getting total across rows
    vTotInput = np.sum(mInput, axis=0)

    # Separating Intermediate Consumption and Final Demand
    vTotIC = vTotInput[:nSectors]
    vTotFD = vTotInput[nSectors:]

    # Total for both components
    nTotIC = np.sum(vTotIC)
    nTotFD = np.sum(vTotFD)
    nTot = nTotIC + nTotFD

    # Creating a new vector with the totals calculated above
    vTotOutput = np.hstack((vTotIC, [nTotIC], vTotFD, [nTotFD, nTot])).astype(float)

    return vTotOutput, vTotIC, vTotFD

# ============================================================================================
def load_number_disaggregations(sDirectoryInput, sFileAgregacao, sSheetNumeroAgregacoes):
    """
    Loads the number of aggregations/disaggregations of sectors or products, which can be changed
    in the "Agregação.xlsx" or "Desagregação.xlsx" Excel files.
    :param sDirectoryInput: directory of the file
    :param sFileAgregacao: name of the Excel file
    :param sSheetNumeroAgregacoes: sheet name
    :return:
        nNumAggregDisaggregSectors: number of sectors to be aggregated/disaggregated
        nNumAggregDisaggregProducts: number of products to be aggregated/disaggregated
    """

    mSheet = pd.read_excel(sDirectoryInput + sFileAgregacao, sheet_name=sSheetNumeroAgregacoes, header=None)
    nNumAggregDisaggregSectors = mSheet.values[0, 0]
    nNumAggregDisaggregProducts = mSheet.values[1, 0]
    return nNumAggregDisaggregSectors, nNumAggregDisaggregProducts

# ============================================================================================
def load_disaggregations(sDirectoryInput, sFileDesagregacao, sSheetDesagregacoes, nNumDisaggreg, nIndice):
    """
    Loads aggregations or disaggregations
    :param sDirectoryInput: directory of the file
    :param sFileDesagregacao: name of the Excel file
    :param sSheetDesagregacoes: sheet name
    :param nNumDisaggreg: number of sectors/products to be aggregated/disaggregated
        (see load_NumAggreg_Disaggreg function)
    :param nIndice: number of sectors/products in matrix (set by nDimension)
    :return:
        mPosDisaggreg: array (nNumDisaggreg x 1) containing the indices of the sectors/products
            to be disaggregated in the original IBGE matrix
        mMultipDisaggreg: array containing the multiples to be used for the disaggregation
            (the aggregated value presented in IBGE's matrix is going to be multiplied by these)
        vNameSectorDisaggreg: name of the disaggregated sectors/products
        nNewIndice: total number of sectors/products with the disaggregation
    """

    # Reading the Excel file and establishing the start of the actual data
    mSheet = pd.read_excel(sDirectoryInput + sFileDesagregacao, sheet_name=sSheetDesagregacoes, header=None)
    nLinIni = 1

    # Creating empty vectors (filled with 0s) to store function results
    mPosDisaggreg = np.zeros([nNumDisaggreg], dtype=int)
    mMultipDisaggreg = np.zeros([nNumDisaggreg], dtype=float)
    vNameSectorDisaggreg = []
    nAuxIndice = -1
    nNewIndice = nIndice

    # for each sector to be disaggregated,
    for i in range(nNumDisaggreg):
        # get position of the sector in the IBGE matrix
        mPosDisaggreg[i] = mSheet.values[nLinIni + i, 0]
        # get multiple of disaggregation (to be multiplied by the aggregated sector value presented in IBGE's matrix)
        mMultipDisaggreg[i] = mSheet.values[nLinIni + i, 1]
        # get name of the disaggregation
        vNameSectorDisaggreg.append(mSheet.values[nLinIni + i, 2])
        # and calculate the new numbers of sectors in the full matrix
        if nAuxIndice == mPosDisaggreg[i]:
            nNewIndice += 1

        nAuxIndice = mPosDisaggreg[i]

    return mPosDisaggreg, mMultipDisaggreg, vNameSectorDisaggreg, nNewIndice

# ============================================================================================
def column_sector_disaggregation(mArray, nNumDisaggreg, vPosDisaggreg, vMultipliers, nNewIndex, nIndex):
    """
    Changes the mArray structure to accommodate and create new disaggregated sectors where it is desired.
    :param mArray: matrix whose structure to be changed;
    :param nNumDisaggreg: number of sectoral disaggregations;
    :param vPosDisaggreg: vector containing the indexes of the products to be disaggregated;
    :param vMultipliers: array containing multiples in order to create the new sectors based on the aggregated one;
    :param nNewIndex: number of sectors WITH the disaggregations;
    :param nIndex: number of sectors WITHOUT the disaggregations;
    :return:
        mNewArray: mArray with the disaggregated sectors (and all the other ones)
    """

    ## Getting shape of the matrix and creating the structure of the new one
    nRow, nCol = np.shape(mArray)
    mNewArray = np.zeros([nRow, nNewIndex], dtype=float)

    ## Getting the index of the first sectors to be disaggregated
    nInitialPosition = vPosDisaggreg[0]

    ## For all the sectors up to the initial one - 1, just copy all values
    if nInitialPosition != 0:
        mNewArray[:, :nInitialPosition] = mArray[:, :nInitialPosition]

    # Auxiliary constants
    nDisaggreg = 0
    nCol = nInitialPosition

    ## Starting from the initial sectors to be disaggregated up until all disaggregations,
    while nDisaggreg < nNumDisaggreg:
        ## For each sector to be disaggregated,
        while nInitialPosition == vPosDisaggreg[nDisaggreg]:
            ## Multiply the original sector value by the disaggregation multiple
            mNewArray[:, nCol] = mArray[:, nInitialPosition] * vMultipliers[nDisaggreg]
            nCol += 1
            nDisaggreg += 1
            if nDisaggreg >= nNumDisaggreg:
                break

        # If there are multiple aggregations and the products aren't together,
        # populate the new matrix with the original values from array
        if nDisaggreg < nNumDisaggreg:
            nInitialPosition = vPosDisaggreg[nDisaggreg]
            if nInitialPosition - vPosDisaggreg[nDisaggreg - 1] > 1:
                nColEnd = nCol + nInitialPosition - vPosDisaggreg[nDisaggreg - 1]
                mNewArray[:, nCol:nColEnd] = mArray[:, vPosDisaggreg[nDisaggreg - 1] + 1:nInitialPosition]
                nCol = nColEnd

    ## For the rest of the matrix, copy the values from the original one
    mNewArray[:, nCol:nNewIndex] = mArray[:, vPosDisaggreg[nNumDisaggreg - 1] + 1:nIndex]

    return mNewArray

# ============================================================================================
def row_product_disaggregation(mArray, nNumDisaggreg, vPosDisaggreg, vMultipliers, nNewIndex, nIndex):
    """
    Changes the mArray structure to accommodate and create new disaggregated products where it is desired.
    :param mArray: matrix whose structure to be changed;
    :param nNumDisaggreg: number of product disaggregations to be made
    :param vPosDisaggreg: vector containing the indexes of the products to be disaggregated;
    :param vMultipliers: vector containing multipliers used to create the new products based on the aggregated one;
    :param nNewIndex: number of products WITH the disaggregations
    :param nIndex: number of products WITHOUT the disaggregations
    :return:
        mNewArray: mArray with the disaggregated products (and all the other ones)
    """

    ## Getting shape of the matrix and creating the structure of the new one
    nRow, nCol = np.shape(mArray)
    mNewArray = np.zeros([nNewIndex, nCol], dtype=float)

    ## Getting the index of the first product to be disaggregated
    nInitialPosition = vPosDisaggreg[0]

    ## For all the products up to the initial one - 1, just copy all values
    if nInitialPosition != 0:
        mNewArray[:nInitialPosition, :] = mArray[:nInitialPosition, :]

    ## Auxiliary constants
    nDisaggreg = 0
    nRow = nInitialPosition

    ## Starting from the initial product to be disaggregated up until all disaggregations,
    while nDisaggreg < nNumDisaggreg:
        ## For each product to be disaggregated,
        while nInitialPosition == vPosDisaggreg[nDisaggreg]:
            ## Multiply the original product value by the disaggregation multiple
            mNewArray[nRow, :] = mArray[nInitialPosition, :] * vMultipliers[nDisaggreg]
            nRow += 1
            nDisaggreg += 1
            if nDisaggreg >= nNumDisaggreg:
                break

        # If there are multiple aggregations and the products aren't together,
        # populate the new matrix with the original values from array
        if nDisaggreg < nNumDisaggreg:
            nInitialPosition = vPosDisaggreg[nDisaggreg]
            if nInitialPosition - vPosDisaggreg[nDisaggreg - 1] > 1:
                nRowEnd = nRow + nInitialPosition - vPosDisaggreg[nDisaggreg - 1]
                mNewArray[nRow:nRowEnd, :] = mArray[vPosDisaggreg[nDisaggreg - 1] + 1:nInitialPosition, :]
                nRow = nRowEnd

    ## For the rest of the matrix, copy the values from the original one
    mNewArray[nRow:nNewIndex, :] = mArray[vPosDisaggreg[nNumDisaggreg - 1] + 1:nIndex, :]

    return mNewArray

# ============================================================================================
def name_disaggregation(vNames, nNumDisaggreg, vPosDisaggreg, vNamesDisaggreg, nIndex):
    """
    Disaggregates the name's vector for products and sectors
    :param vNames: vector (1D array) containing the names of sectors/products (WITHOUT the disaggregations)
    :param nNumDisaggreg: number of sectors/products that were disaggregated;
    :param vPosDisaggreg: vector containing the indexes of the products to be disaggregated;
    :param vNamesDisaggreg: vector containing the names of the products to be disaggregated;
    :param nIndex: total number of sectors/products (WITHOUT the disaggregations);
    :return:
        vNewName: nNewSectors x 1 array containing all names of products/sectors (WITH the disaggregations)
    """

    ## Creating empty list to store the names
    vNewName = []
    nRow = 0
    nDisaggreg = 0

    ## Starting from the initial sector/product to be disaggregated up until all disaggregations,
    while nDisaggreg < nNumDisaggreg:
        ## Up until the initial product/sector to be disaggregated, just copy the values
        if nRow != vPosDisaggreg[nDisaggreg]:
            vNewName.append(vNames[nRow])
            nRow += 1
        ## If the products/sectors were created based on a disaggregation
        elif nRow == vPosDisaggreg[nDisaggreg]:
            ## Append its name to the list and add one to nDisaggreg in order to continue the loop
            vNewName.append(vNamesDisaggreg[nDisaggreg])
            nDisaggreg += 1

            ## Updating nRow for the loop (in cases where disaggregated products/sectors are not together)
            if nDisaggreg < nNumDisaggreg:
                nRow = vPosDisaggreg[nDisaggreg]

    ## For all product/sectors after the last disaggregation, just copy the values
    nRow += 1
    while nRow < nIndex:
        vNewName.append(vNames[nRow])
        nRow += 1

    return vNewName

# ============================================================================================
def gdp_calculation(mMIPGeral, nSector, nDimension):
    """
    Calculates Total GDP and each one of its components
    :param mMIPGeral: MIP matrix
    :param nSector: number of sectors (WITH the disaggregations)
    :param nDimension: dimension of the estimated MIP (0, 1, 2 or 3)
    :return:
        vGDP: 17x1 array containing the GDP and each of its components
            GDP is calculated in all three ways possible (production, income and expenditure)
        vNameGDP: 17x1 array containing the names of GDP components
        vNameColGDP: 1x1 array
    """

    ## Creating empty vectors
    vNameGDP = []
    vNameColGDP = ["Valores"]
    vGDP = np.zeros([16], dtype=float)
    nRowMIP, nColMIP = mMIPGeral.shape  # np.shape(mMIPGeral)

    ### GDP BY PRODUCTION
    ## Adding names
    vNameGDP.append('PIB pela ótica do produto')
    vNameGDP.append('Produção')
    vNameGDP.append('Impostos')
    vNameGDP.append('Consumo Intermediário')

    ## Calculating production, taxes, intermediate consumption
    nProducao = mMIPGeral[nSector, nColMIP - 1]
    nImpostos = mMIPGeral[nSector + 2, nColMIP - 1] + mMIPGeral[nSector + 3, nColMIP - 1] + \
        mMIPGeral[nSector + 4, nColMIP - 1] + mMIPGeral[nSector + 5, nColMIP - 1]
    nCI = mMIPGeral[nSector + 6, nSector]

    ## Calculating GDP by the Production way
    nPIBProduto = nProducao + nImpostos - nCI
    vGDP[0] = nPIBProduto
    vGDP[1] = nProducao
    vGDP[2] = nImpostos
    vGDP[3] = nCI

    ### GDP BY INCOME
    ## Adding names
    vNameGDP.append('PIB pela ótica da Renda')
    vNameGDP.append('Remuneração dos empregados')
    vNameGDP.append('EOB + RM')
    vNameGDP.append('Impostos líquidos sobre a produção e importação')

    ## Calculating salaries, capital and autonomous income and taxes on imports
    # n = 51 is different due to the presence of two export columns
    if nDimension == 2:
        nRemunEmpregados = mMIPGeral[nSector + 8, nSector]
        nEOB_RM = mMIPGeral[nSector + 14, nSector]
        nImpLiqProdImport = nImpostos + mMIPGeral[nSector + 15, nSector] + mMIPGeral[nSector + 16, nSector]
    else:
        nRemunEmpregados = mMIPGeral[nSector + 8, nSector]
        nEOB_RM = mMIPGeral[nSector + 14, nSector]
        nImpLiqProdImport = nImpostos + mMIPGeral[nSector + 17, nSector] + mMIPGeral[nSector + 18, nSector]

    ## Calculating GDP by the Income way
    nPIBrenda = nRemunEmpregados + nEOB_RM + nImpLiqProdImport
    vGDP[4] = nPIBrenda
    vGDP[5] = nRemunEmpregados
    vGDP[6] = nEOB_RM
    vGDP[7] = nImpLiqProdImport

    ### GDP BY EXPENDITURE
    ## Adding names
    vNameGDP.append('PIB pela ótica da despesa')
    vNameGDP.append('Consumo das Famílias')
    vNameGDP.append('Consumo do Governo')
    vNameGDP.append('Consumo das ISFLSF')
    vNameGDP.append('FBCF')
    vNameGDP.append('Variação do estoque')
    vNameGDP.append('exportação de bens e serviços')
    vNameGDP.append('importação de bens e serviços (-)')

    ## Calculating total exports, government consumption...
    # n = 51 is different due to the presence of two export columns
    if nDimension == 2:
        nExportTotal = mMIPGeral[nSector + 6, nSector + 1] + mMIPGeral[nSector + 6, nSector + 2]
        nGovernConsum = mMIPGeral[nSector + 6, nSector + 3]
        nISFLSFConsum = mMIPGeral[nSector + 6, nSector + 4]
        nFamilyConsum = mMIPGeral[nSector + 6, nSector + 5]
        nFBCF = mMIPGeral[nSector + 6, nSector + 6]
        nColEstockVar = mMIPGeral[nSector + 6, nSector + 7]
        nimporTotal = mMIPGeral[nSector + 1, nColMIP - 1]
    else:
        nExportTotal = mMIPGeral[nSector + 6, nSector + 1]
        nGovernConsum = mMIPGeral[nSector + 6, nSector + 2]
        nISFLSFConsum = mMIPGeral[nSector + 6, nSector + 3]
        nFamilyConsum = mMIPGeral[nSector + 6, nSector + 4]
        nFBCF = mMIPGeral[nSector + 6, nSector + 5]
        nColEstockVar = mMIPGeral[nSector + 6, nSector + 6]
        nimporTotal = mMIPGeral[nSector + 1, nColMIP - 1]

    ## Calculating GDP by the Expenditure way
    nPIBDespesa = nExportTotal + nGovernConsum + nISFLSFConsum + nFamilyConsum + nFBCF + nColEstockVar - nimporTotal
    vGDP[8] = nPIBDespesa
    vGDP[9] = nFamilyConsum
    vGDP[10] = nGovernConsum
    vGDP[11] = nISFLSFConsum
    vGDP[12] = nFBCF
    vGDP[13] = nColEstockVar
    vGDP[14] = nimporTotal
    vGDP[15] = nExportTotal

    return vGDP, vNameGDP, vNameColGDP

# ============================================================================================
def write_data_excel(FileName, lSheetName, lDataSheet, lRowsLabel, lColsLabel):
    """
    Writes data to an Excel file (with multiple sheets) in the "Output" directory.
    :param FileName: name of the desired file
    :param lSheetName: array containing the desired names of the sheets
    :param lDataSheet: array containing the data to be written
    :param lRowsLabel: array containing the row labels
    :param lColsLabel: array containing the column labels
    :return: Nothing; an Excel file is written in the "Output" directory.
    """

    ## Creating Writer object (allows multiple sheets into a single file)
    Writer = pd.ExcelWriter('./Output/' + FileName, engine='openpyxl')
    # List that will contain the dataframes
    lDataFrames = []

    ## For each sheet...
    for nSheet in range(len(lSheetName)):
        # Create DataFrame object
        dfData = pd.DataFrame(lDataSheet[nSheet], index=lRowsLabel[nSheet], columns=lColsLabel[nSheet], dtype=float)

        # Append a dataframe to the list and export to Writer object
        lDataFrames.append(dfData)

        ## Determining float format
        # Getting max value within DataFrame
        nMax = np.amax(dfData.values)
        # Defining float format (show more decimal columns when values are smaller)
        # sFloatFormat = "%.6f" if nMax <= 4 else "%.2f"

        ## Writing to Excel
        lDataFrames[nSheet].to_excel(Writer, lSheetName[nSheet], freeze_panes=(1, 1))

    ## Saving file
    Writer.save()

# ============================================================================================
