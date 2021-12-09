"""
Annual estimation for Input Output Tables based on the methodology proposed by Guilhoto (2010) and Barry-Miller (2009).

Based on official Resources and Uses tables published by IBGE in the System of National Accounts (3rd edition).

Authors: João Maria de Oliveira and Vinícius de Almeida Nery Ferreira (Ipea-DF).

E-mails: joao.oliveira@ipea.gov.br and vinicius.nery@ipea.gov.br (or vnery5@gmail.com).
"""

## Import numpy, time and support file
import numpy as np
import time
import SupportFunctions_V3 as Support

## Only run if it's the main file (don't run on import)
if __name__ == '__main__':
    # ==================================================================================================================
    # Parameters for MIP Estimation
    # ==================================================================================================================

    ## Constants that determine...
    # ... size of MIP (possible values: 0 - 12x12; 1 - 20x20; 2 - 107X51; 3 - 128x68)
    nDimension = 3
    # ... year to be estimated
    nYear = 2019

    ## Constants that identify software version and software config version
    sVersionConfigSoft = "3.0"
    sVersionSoft = "3.0"

    ## Disaggregate or aggregate products/sectors?
    bAggregateDisaggregate = False  # True or False

    ## Lists that contain the correspondent parameters for each MIP Dimension
    vProducts = [12, 20, 107, 128]  # number of products
    vSectors = [12, 20, 51, 68]  # number of sectors
    vRowsTrade = [[5, 5], [6, 6], [88, 88], [92, 93]]  # indices (base 0) of the trade products (initial and final rows)
    vRowsTransp = [[6, 6], [7, 7], [89, 90], [94, 97]]  # indices of the transport products (initial and final rows)
    vColsTrade = [[5, 5], [6, 6], [36, 36], [40, 41]]  # indices of the trade sectors (initial and final columns)
    vColsTransp = [[6, 6], [7, 7], [37, 37], [42, 44]]  # indices of the transport sectors (initial and final columns)

    # Getting values that match the desired dimension
    nProducts = vProducts[nDimension]
    nSectors = vSectors[nDimension]
    vRowsTradeElim = vRowsTrade[nDimension]
    vRowsTranspElim = vRowsTransp[nDimension]
    vColsTradeElim = vColsTrade[nDimension]
    vColsTranspElim = vColsTransp[nDimension]

    ## Adjust trade/transport margins to only one product/activity?
    lAdjustMargins = False  # True or False
    sAdjustMargins = "_Agreg" if lAdjustMargins else ""
    nAdjust = 0

    ## Number of final demand columns in IBGE's TRU demand table (in 107x68, there are two export columns)
    nColsDemand = 6 if nDimension != 2 else 7
    ## Number of supply columns in IBGE's TRU table
    nColsOffer = 7
    ## Number of Added Value (AV) rows in IBGE's TRU table (in 107x68, there are no separate EOB/RM lines)
    nRowsAV = 14 if nDimension != 2 else 12
    ## Total Production Row
    nRowTotalProduction = nRowsAV - 2

    ## Indices of the final demand components (in 107x68, there are two export columns)
    nColExport = 0 if nDimension != 2 else [0, 1]
    nColFBCF = 4 if nDimension != 2 else 5
    nColEstockVar = 5 if nDimension != 2 else 6

    ## Indices that determine from where to start reading the spreadsheets
    # General initial column
    nColIni = 2 if nDimension != 2 else 1
    # Supply initial column
    nColIniOffer = 2

    # ==================================================================================================================
    # General Parameters
    # ==================================================================================================================

    ## Defining entry and out directories
    sDirectoryBaseInput = './Input/'
    if nDimension == 3:
        sDirectoryInput = './Input/Nível68/'
        sDirectoryOutput = './Output/'
    elif nDimension == 1:
        sDirectoryInput = './Input/Nível20/'
        sDirectoryOutput = './Output/'
    elif nDimension == 0:
        sDirectoryInput = './Input/Nível12/'
        sDirectoryOutput = './Output/'
    else:
        sDirectoryInput = './InputRetro/Nível51/'
        sDirectoryOutput = './Output/Nível51/'

    ## String that identifies the Uses spreadsheet file
    sFileUses = f"{nSectors}_tab2_{nYear}.xls"
    # Sheet Names
    sSheetIntermedConsum = 'CI'  # Intermediate Consumption
    sSheetDemand = 'demanda'  # Final Demand
    sSheetAddedValue = 'VA'  # Added Value

    ## String that identifies the Resources spreadsheet file
    sFileResources = f"{nSectors}_tab1_{nYear}.xls"
    sSheetOffer = 'oferta'  # Supply Components (taxes, margins and base prices)
    sSheetProduction = 'producao'  # Production (products x sectors)
    sSheetImport = 'importacao'  # Imports (products x 1 vector)

    # ==================================================================================================================
    # Parameters for aggregation and disaggregation
    # ==================================================================================================================

    sFileAgregacao = "Agregação.xlsx"
    sSheetNumeroAgregacoes = "NumeroAgregacoes"
    sSheetAgregacaoSetor = "AgregaçãoSetor"
    sSheetAgregacaoProduto = "AgregaçãoProduto"

    sFileDesagregacao = "Desagregação.xlsx"
    sSheetNumeroDesagregacoes = "NumeroDesagregacoes"
    sSheetDesagregacaoSetor = "DesagregaçãoSetor"
    sSheetDesagregacaoProduto = "DesagregaçãoProduto"

    # ==================================================================================================================
    # STARTING ESTIMATION
    # ==================================================================================================================

    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print(f"+ Estimação da Matriz Insumo Produto Nacional - Versão {sVersionSoft}")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    ## Timers
    nBeginModel = time.perf_counter()
    sTimeBeginModel = time.localtime()

    # ==================================================================================================================
    # Import values from TRUs
    # ==================================================================================================================

    # Intermediate Consumption
    mIntermConsum, vNameProduct, vNameSector, vCodeProduct, vCodeSector = \
        Support.load_tru(sDirectoryInput, sFileUses, sSheetIntermedConsum,
                         nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors, bNames=True, bCodes=True)
    # Final Demand
    mDemand, vNameProduct1, vNameDemand = \
        Support.load_tru(sDirectoryInput, sFileUses, sSheetDemand,
                         nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nColsDemand, bNames=True)
    # Added Value
    mAddedValue, vNameAddedValue, vNameSector1 = \
        Support.load_tru(sDirectoryInput, sFileUses, sSheetAddedValue,
                         nRowIni=5, nColIni=1, nRows=nRowsAV, nCols=nSectors, bNames=True)
    # Supply
    mOffer, vNameProduct2, vNameOffer = \
        Support.load_tru(sDirectoryInput, sFileResources, sSheetOffer,
                         nRowIni=5, nColIni=nColIniOffer, nRows=nProducts, nCols=nColsOffer, bNames=True)
    # Production
    mProduction = Support.load_tru(sDirectoryInput, sFileResources, sSheetProduction,
                                   nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors)
    # Imports
    vImport = Support.load_tru(sDirectoryInput, sFileResources, sSheetImport,
                               nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=1)

    # ==================================================================================================================
    # Import values of disaggregation
    # ==================================================================================================================

    if bAggregateDisaggregate:
        nNumDisaggregSectors, nNumDisaggregProducts = Support.load_number_disaggregations(sDirectoryBaseInput,
                                                                                          sFileDesagregacao,
                                                                                          sSheetNumeroDesagregacoes)

        if (nNumDisaggregSectors > 0) and (nDimension == 3):
            mPosDisaggreg, mMultipDisaggreg, vNameSectorDisaggreg, nNewSectors = \
                Support.load_disaggregations(sDirectoryBaseInput, sFileDesagregacao, sSheetDesagregacaoSetor,
                                             nNumDisaggregSectors, nSectors)

            ## Disaggregating the necessary sectors (columns) in each matrix
            mIntermConsum = Support.column_sector_disaggregation(mIntermConsum, nNumDisaggregSectors, mPosDisaggreg,
                                                                 mMultipDisaggreg, nNewSectors, nSectors)
            mAddedValue = Support.column_sector_disaggregation(mAddedValue, nNumDisaggregSectors, mPosDisaggreg,
                                                               mMultipDisaggreg, nNewSectors, nSectors)
            mProduction = Support.column_sector_disaggregation(mProduction, nNumDisaggregSectors, mPosDisaggreg,
                                                               mMultipDisaggreg, nNewSectors, nSectors)
            vNameSector = Support.name_disaggregation(vNameSector, nNumDisaggregSectors, mPosDisaggreg,
                                                      vNameSectorDisaggreg, nSectors)
            nSectors = nNewSectors

        if (nNumDisaggregProducts > 0) and (nDimension == 3):
            mPosDisaggreg, mMultipDisaggreg, vNameProductDisaggreg, nNewProducts = \
                Support.load_disaggregations(sDirectoryBaseInput, sFileDesagregacao, sSheetDesagregacaoProduto,
                                             nNumDisaggregProducts, nProducts)

            ## Disaggregating the necessary products (rows) in each matrix
            mIntermConsum = Support.row_product_disaggregation(mIntermConsum, nNumDisaggregProducts, mPosDisaggreg,
                                                               mMultipDisaggreg, nNewProducts, nProducts)
            mProduction = Support.row_product_disaggregation(mProduction, nNumDisaggregProducts, mPosDisaggreg,
                                                             mMultipDisaggreg, nNewProducts, nProducts)
            mDemand = Support.row_product_disaggregation(mDemand, nNumDisaggregProducts, mPosDisaggreg,
                                                         mMultipDisaggreg, nNewProducts, nProducts)
            mOffer = Support.row_product_disaggregation(mOffer, nNumDisaggregProducts, mPosDisaggreg,
                                                        mMultipDisaggreg, nNewProducts, nProducts)
            vImport = Support.row_product_disaggregation(vImport, nNumDisaggregProducts, mPosDisaggreg,
                                                         mMultipDisaggreg, nNewProducts, nProducts)
            vNameProduct = Support.name_disaggregation(vNameProduct, nNumDisaggregProducts, mPosDisaggreg,
                                                       vNameProductDisaggreg, nProducts)
            nProducts = nNewProducts
    # ==================================================================================================================
    # Adjusting Trade and Transport for Products and for Sectors
    # ==================================================================================================================

    while lAdjustMargins:
        if nAdjust == 0:
            nRowIni = vRowsTradeElim[0]
            nRowFim = vRowsTradeElim[1]
            vNameProduct[nRowIni] = 'Comércio'
            nColIni = vColsTradeElim[0]
            nColFim = vColsTradeElim[1]
            vNameSector[nColIni] = 'Comércio'
        else:
            nRowIni = vRowsTranspElim[0]
            nRowFim = vRowsTranspElim[1]
            vNameProduct[nRowIni] = 'Transporte'
            nColIni = vColsTranspElim[0]
            nColFim = vColsTranspElim[1]
            vNameSector[nColIni] = 'Transporte'

        for nElim in range(nRowIni + 1, nRowFim + 1):
            vNameProduct[nElim] = 'x'

        for nElim in range(nRowIni + 1, nRowFim + 1):
            vImport[nRowIni] += vImport[nElim]
            vImport[nElim] = 0.0

        for i in range(nColsOffer):
            for nElim in range(nRowIni + 1, nRowFim + 1):
                mOffer[nRowIni, i] += mOffer[nElim, i]
                mOffer[nElim, i] = 0.0

        for i in range(nSectors + 1):
            for nElim in range(nRowIni + 1, nRowFim + 1):
                mProduction[nRowIni, i] += mProduction[nElim, i]
                mProduction[nElim, i] = 0.0
                mIntermConsum[nRowIni, i] += mIntermConsum[nElim, i]
                mIntermConsum[nElim, i] = 0.0

        for i in range(nColsDemand):
            for nElim in range(nRowIni + 1, nRowFim + 1):
                mDemand[nRowIni, i] += mDemand[nElim, i]
                mDemand[nElim, i] = 0.0

        for nElim in range(nColIni + 1, nColFim + 1):
            vNameSector[nElim] = 'x'

        for i in range(nRowsAV):
            for nElim in range(nColIni + 1, nColFim + 1):
                mAddedValue[i, nColIni] += mAddedValue[i, nElim]
                mAddedValue[i, nElim] = 0.0

        for i in range(nProducts + 1):
            for nElim in range(nColIni + 1, nColFim + 1):
                mProduction[i, nColIni] += mProduction[i, nElim]
                mProduction[i, nElim] = 0.0
                mIntermConsum[i, nColIni] += mIntermConsum[i, nElim]
                mIntermConsum[i, nElim] = 0.0

        nAdjust += 1
        if nAdjust == 2:
            lAdjustMargins = False

    # ==================================================================================================================
    # Calculating coefficients without stock variation
    # ==================================================================================================================

    ## Copying demand
    mDemandWithoutEstock = np.copy(mDemand)

    # Excluding ∆ stock column (giving it all 0s in order to maintain the number of columns)
    mDemandWithoutEstock[:, nColEstockVar] = 0.0

    # Calculating the distribution/alpha matrix (see function documentation for more details)
    mDistribution, mTotalConsum = Support.distribution_matrix_calcul(mIntermConsum, mDemandWithoutEstock)

    # ==================================================================================================================
    # Calculating arrays internally distributed by alphas
    # For each product/sector pair, estimates the margin of trade/taxes paid of product i in sector j,
    # under the assumption that the margins/taxes follow the same distribution observed in production
    # ==================================================================================================================

    ## Trade margins
    nColMarginTrade = 1 if nDimension != 2 else 0
    mMarginTrade = Support.calculation_margin(mDistribution, mOffer, nColMarginTrade, vRowsTradeElim)

    ## Transport margins
    nColMarginTransport = 2 if nDimension != 2 else 1
    mMarginTransport = Support.calculation_margin(mDistribution, mOffer, nColMarginTransport, vRowsTranspElim)

    ## Taxes
    nColIPI = 4 if nDimension != 2 else 3
    mIPI = Support.calculation_internal_matrix(mDistribution, mOffer, nColIPI)

    nColICMS = 5 if nDimension != 2 else 4
    mICMS = Support.calculation_internal_matrix(mDistribution, mOffer, nColICMS)

    nColOtherTaxes = 6 if nDimension != 2 else 5
    mOtherTaxes = Support.calculation_internal_matrix(mDistribution, mOffer, nColOtherTaxes)

    # ==================================================================================================================
    # Calculating coefficients without exports and stock variation
    # This will be used to calculate the distribution of imports and import taxes
    # ==================================================================================================================

    # Copying demand without stock
    mDemandWithoutExport = np.copy(mDemandWithoutEstock)

    # Excluding export column (giving it all 0s in order to maintain the number of columns)
    mDemandWithoutExport[:, nColExport] = 0

    # Calculating the distribution/alpha matrix (see function documentation for more details)
    mDistributionWithoutExport, mTotalConsumWithoutExport = \
        Support.distribution_matrix_calcul(mIntermConsum, mDemandWithoutExport)

    # ==================================================================================================================
    # Calculating Arrays internally distributed by alphas without exports
    # ==================================================================================================================

    ## For each product/sector pair, estimates the import of product i by sector j, as well as the import taxes involved
    nColImport = 0
    mImport = Support.calculation_internal_matrix(mDistributionWithoutExport, vImport, nColImport)

    nColImportTax = 3 if nDimension != 2 else 2
    mImportTax = Support.calculation_internal_matrix(mDistributionWithoutExport, mOffer, nColImportTax)

    # ==================================================================================================================
    # Calculating the Matrix of Consumption with base prices
    # ==================================================================================================================

    ## Creating total consumption matrix
    mTotalConsum = np.concatenate((mIntermConsum, mDemand), axis=1)

    ## Subtracting all margins in taxes in order to arrive at consumption in base prices (not market ones)
    mConsumBasePrice = \
        mTotalConsum - mMarginTrade - mMarginTransport - mIPI - mICMS - mOtherTaxes - mImport - mImportTax

    ### Creating E Matrix with basic price
    ## Getting final demand in base prices
    mE = mConsumBasePrice[:, nSectors:]

    ## Getting intermediate consumption in base prices
    mIntermConsumBasePrice = mConsumBasePrice[:, :nSectors]

    # ==================================================================================================================
    # Calculating Intersectorial Demand (Intermediate Consumption and Final Demand)
    # ==================================================================================================================

    ## Disabling warnings for division by 0 (they are handled after)
    np.seterr(divide='ignore', invalid='ignore')

    ### Creating D Matrix
    ## Transposing production matrix (nSectors x nProducts)
    mProductionTrans = mProduction.T

    ## Adding all sector's production in order to get totals by product
    vRowTotProductionTrans = np.sum(mProductionTrans, axis=0)

    # For each product, calculate the proportion of production made by each sector
    # relative to the total production of product j (∑mD rows = 1)
    mD = mProductionTrans / vRowTotProductionTrans[None, :]
    # Checking for NaNs and infinities
    mD = np.nan_to_num(mD, nan=0, posinf=0, neginf=0)

    # Creating X Vector (total production by sector in consumer prices)
    vX = mAddedValue[nRowTotalProduction, 0:nSectors]

    # Creating Matrix of National Coefficients - Bn Matrix
    # (proportion of the product i that sector j consumed relative to the total production of sector j)
    mBn = mConsumBasePrice[:, :nSectors] / vX[None, :]
    # Checking for NaNs and infinities
    mBn = np.nan_to_num(mBn, nan=0, posinf=0, neginf=0)

    # Creating Matrix of Imports Coefficients - Bm Matrix
    # (proportion of the imports of product i that sector j imported relative to the total production of sector j)
    mBm = mImport[:, :nSectors] / vX[None, :]
    # Checking for NaNs and infinities
    mBm = np.nan_to_num(mBm, nan=0, posinf=0, neginf=0)

    ### Creating A Matrix (Technical Coefficients) - Sector x Sector
    ## mD: proportion of production made by each sector i relative to the total of product j
    ## mBn: proportion of product i that sector j consumed relative to the total production of sector j
    # Sector by sector
    mA = np.dot(mD, mBn)
    # Product by Product
    mProd = np.dot(mBn, mD)

    ## Creating Z Matrix (Intermediate Consumption) - Sector x Sector
    # For each sector, multiply the direct technical coefficient by the sector's total production,
    # which equals the intermediate consumption by sector j from the production of sector i
    mZ = mA * vX[None, :]

    # Creating Y Matrix (nSectors x 6 (components of demand))
    # mD: proportion of production made by each sector i relative to the total production of product j
    # mE: final demand for each product
    mY = np.dot(mD, mE)

    # Calculating Leontief Matrix (I-A)^(-1) (Sector x Sector)
    mI = np.eye(nSectors)
    mLeontief = np.linalg.inv(mI - mA)

    # ==================================================================================================================
    # Creating I-O Structure and checking consistency
    # ==================================================================================================================

    ## Adding all product's imports, import taxes and other taxes in order to get sectorial and final demand totals
    vRowImports, vImports_IC, vImports_FD = Support.payment_sector_total(mImport, nSectors)
    vRowImportTax, vITax_IC, vITax_FD = Support.payment_sector_total(mImportTax, nSectors)
    vRowIPI, vIPI_IC, vIPI_FD = Support.payment_sector_total(mIPI, nSectors)
    vRowICMS, vICMS_IC, vICMS_FD = Support.payment_sector_total(mICMS, nSectors)
    vRowOtherTaxes, vOtherTaxes_IC, vOtherTaxes_FD = Support.payment_sector_total(mOtherTaxes, nSectors)

    # Payment Sector Total
    vTotSP = vRowImports + vRowImportTax + vRowIPI + vRowICMS + vRowOtherTaxes

    ## Checking total supply
    vTotalSupply = np.sum(mZ, axis=0) + vImports_IC + vICMS_IC + vIPI_IC + vOtherTaxes_IC + vITax_IC + mAddedValue[0, :]
    # Difference from IBGE's values
    vDiff = vX - vTotalSupply
    nDiff = np.sum(vDiff)
    vDiff = np.hstack((vDiff, nDiff, [0]*(nColsDemand + 1), nDiff))

    ### Creating upper part of the MIP
    ## Column Totals
    # Intermediate Consumption
    vZ_Tot_Col = np.sum(mZ, axis=1).reshape((nSectors, 1))
    # Final Demand
    vY_Tot_Col = np.sum(mY, axis=1).reshape((nSectors, 1))
    # Added Value
    vAddedValue_Tot_Col = np.sum(mAddedValue, axis=1).reshape((nRowsAV, 1))

    ## Upper part of the MIP
    mMIP_Upper = np.concatenate((
        mZ,  # Intermediate Consumption
        vZ_Tot_Col,  # Total Intermediate Consumption of inputs produced by sector i
        mY,  # Final Demand for sector i products
        vY_Tot_Col,  # Total final demand for sector i products
        vZ_Tot_Col + vY_Tot_Col  # Total demand for sector i products
    ), axis=1)

    # Finding total national consumption at base prices and appending to end of the matrix
    vNatConsumption = np.sum(mMIP_Upper, axis=0)
    mMIP_Upper = np.vstack((mMIP_Upper, vNatConsumption))

    ## Lower part of the MIP (payment sector - imports and taxes - and added value)
    # Adjusting shape of the AV matrix
    mZerosAV = np.zeros([nRowsAV, nColsDemand + 1])
    mAddedValue_Full = np.concatenate((mAddedValue, vAddedValue_Tot_Col, mZerosAV, vAddedValue_Tot_Col), axis=1)

    # Concatenating all vectors to form lower part of the MIP
    mMIP_Lower = np.vstack((
        vRowImports, vRowImportTax, vRowIPI, vRowICMS, vRowOtherTaxes,  # payment sector
        vTotSP + vNatConsumption,  # total consumption (national at base prices + imports and taxes)
        mAddedValue_Full,  # full added value
        vDiff  # difference from IBGE's values
    ))

    ## Joining both parts
    mMIPGeral = np.vstack((mMIP_Upper, mMIP_Lower)).astype(float)

    ## Creating names for index and columns
    # Payment Sector components
    vNamesSP = ["Importação", "Impostos sobre Importação", "IPI", "ICMS", "Outros Impostos Líquidos"]

    # Vectors
    vNamesMIP_Cols = np.hstack((vNameSector, ["Total de Consumo Intermediário"],
                                vNameDemand, ["Demanda Final", "Demanda Total"]))
    vNamesMIP_Rows = np.hstack((vNameSector, ["Consumo Nacional"], vNamesSP, ["CI Total"],
                                vNameAddedValue, ["Diferença"]))

    vGDP, vNameGDP, vNameColGDP = Support.gdp_calculation(mMIPGeral, nSectors, nDimension)

    # ==================================================================================================================
    # Writing Excel file (with multiple sheets)
    # ==================================================================================================================

    print(f"Estimation complete! Writing data to Excel... ({time.strftime('%d/%b/%Y - %H:%M:%S', time.localtime())})\n")

    ## List that contains the data to be written in each sheet
    vDataSheet = [mAddedValue, mDemand, mIntermConsum, mOffer, mProduction, vImport,
                  mDistribution, mMarginTrade, mMarginTransport, mIPI, mICMS, mOtherTaxes,
                  mDistributionWithoutExport, mImport, mImportTax, mConsumBasePrice,
                  mBn, mBm, mD, mA, mZ, mY, mLeontief,
                  mMIPGeral, vGDP
                  ]

    ## Sheet Names
    vSheetName = ["VA", "Demanda", "CI", "Oferta", "Produção", "Importação",
                  "Distribuição", "MGC", "MGT", "IPI", "ICMS", "OILL",
                  "Distribuição_2", "Importação_2", "II", "Usos Pb",
                  "Matriz_Bn", "Matriz_Bm", "Matriz_D", "Matriz_A", "Matriz_Z", "Matriz_Y", "Leontief",
                  "MIP", "PIB"
                  ]

    ## Row (index) labels
    vRowsLabel = [vNameAddedValue, vNameProduct, vNameProduct, vNameProduct, vNameProduct, vNameProduct,
                  vNameProduct, vNameProduct, vNameProduct, vNameProduct,  vNameProduct, vNameProduct,
                  vNameProduct, vNameProduct, vNameProduct, vNameProduct,
                  vNameProduct, vNameProduct, vNameSector, vNameSector, vNameSector, vNameSector, vNameSector,
                  vNamesMIP_Rows, vNameGDP
                  ]

    ## Column labels
    # Concatenating sector and final demand names
    vNameCIDemand = np.hstack((vNameSector + vNameDemand))
    # Import Name
    vNameImport = ["Importação"]

    vColsLabel = [vNameSector, vNameDemand, vNameSector, vNameOffer, vNameSector, vNameImport,
                  vNameCIDemand, vNameCIDemand, vNameCIDemand, vNameCIDemand, vNameCIDemand, vNameCIDemand,
                  vNameCIDemand, vNameCIDemand, vNameCIDemand, vNameCIDemand,
                  vNameSector, vNameSector, vNameProduct, vNameSector, vNameSector, vNameDemand, vNameSector,
                  vNamesMIP_Cols, vNameColGDP
                  ]

    # vDataSheet.append(mBasePriceUses)
    # vSheetName.append('Usos PB')
    # vRowsLabel.append(vNameProduct + vNameComplemenBP)
    # vColsLabel.append(vNameCIDemand)

    ## Writing Excel file to Output directory
    # Indicators of number of sectors
    nSectorsOutputFile = f"{nSectors}" if nSectors <= 68 else "68+"

    # Title string and calling function to write data
    sFileSheet = f"MIP_{nYear}_{nSectorsOutputFile}{sAdjustMargins}.xlsx"
    Support.write_data_excel(sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel)

    ## Calculating total execution time
    nEndModel = time.perf_counter()
    nElapsedTime = (nEndModel - nBeginModel) / 60.

    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print(f"Begun at {time.strftime('%d/%b/%Y - %H:%M:%S', sTimeBeginModel)}")
    print(f"Ended at {time.strftime('%d/%b/%Y - %H:%M:%S', time.localtime())}")
    print(f"Time spent: {round(nElapsedTime, 3)} mins ({round(nElapsedTime*60,  1)} seconds).")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
