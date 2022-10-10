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
from GRAS import gras

## Only run if it's the main file (don't run on import)
if __name__ == '__main__':
    # ==================================================================================================================
    # Parameters for MIP Estimation
    # ==================================================================================================================

    ## Constants that determine...
    # ... size of MIP (possible values: 0 - 12x12; 1 - 20x20; 2 - 107X51; 3 - 128x68)
    nDimension = 3
    # ... years to be estimated
    nFirstYear = 2010  # 2010: base year, normal MIP suffices
    nLastYear = 2019
    lYears = np.arange(nFirstYear, nLastYear + 1)

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
    sDirectoryOutput = './Output/'
    if nDimension == 3:
        sDirectoryInput = './Input/Nível68/'
    elif nDimension == 1:
        sDirectoryInput = './Input/Nível20/'
    elif nDimension == 0:
        sDirectoryInput = './Input/Nível12/'
    else:
        sDirectoryInput = './InputRetro/Nível51/'

    # ==================================================================================================================
    # STARTING ESTIMATION
    # ==================================================================================================================

    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("+ Estimação & Deflação da Matriz Insumo Produto Nacional")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    ## Timers
    nBeginModel = time.perf_counter()
    sTimeBeginModel = time.localtime()

    ## We need to estimate the MIP in this and last year's prices + deflated ones at the end
    arrConsumBasePrices = np.zeros((3, len(lYears), nProducts, nSectors + nColsDemand), dtype=float)
    arrProductionTrans = np.zeros((3, len(lYears), nSectors, nProducts), dtype=float)
    arrValueAdded = np.zeros((3, len(lYears), nRowsAV, nSectors), dtype=float)
    arr_vX = np.zeros((3, len(lYears), nSectors), dtype=float)
    arr_vX_Prod = np.zeros((3, len(lYears), nProducts), dtype=float)
    arr_Demand = np.zeros((3, len(lYears), nSectors + nColsDemand), dtype=float)
    arr_VPB = np.zeros((3, len(lYears), 1), dtype=float)

    # Creating skeletons for price indexes and yearly deflators
    ## Price indexes
    arrConsumBasePrices_Index = np.zeros((len(lYears), nProducts, nSectors + nColsDemand), dtype=float)
    arrProductionTrans_Index = np.zeros((len(lYears), nSectors, nProducts), dtype=float)
    arr_vX_Index = np.zeros((len(lYears), nSectors), dtype=float)
    arr_vX_Prod_Index = np.zeros((len(lYears), nProducts), dtype=float)
    arr_Demand_Index = np.zeros((len(lYears), nSectors + nColsDemand), dtype=float)
    arr_VPB_Index = np.zeros((len(lYears)), dtype=float)

    # 2010 = 100 (base year)
    arrConsumBasePrices_Index[0, :, :] = 100
    arrProductionTrans_Index[0, :, :] = 100
    arr_vX_Index[0, :] = 100
    arr_vX_Prod_Index[0, :] = 100
    arr_Demand_Index[0, :] = 100
    arr_VPB_Index[0] = 100

    # Loop
    for nYear in lYears:
        # Strings that identifies the Uses' spreadsheet file
        sFileUses = f"{nSectors}_tab2_{nYear}.xls"
        sFileResources = f"{nSectors}_tab1_{nYear}.xls"
        sFileUsesLY = f"{nSectors}_tab4_{nYear}.xls" if nYear > nFirstYear else f"{nSectors}_tab2_{nYear}.xls"
        sFileResourcesLY = f"{nSectors}_tab3_{nYear}.xls" if nYear > nFirstYear else f"{nSectors}_tab1_{nYear}.xls"

        # Sheet Names
        sSheetIntermedConsum = 'CI'  # Intermediate Consumption
        sSheetDemand = 'demanda'  # Final Demand
        sSheetAddedValue = 'VA'  # Added Value (just in this year's prices)

        sSheetOffer = 'oferta'  # Supply Components (taxes, margins and base prices)
        sSheetProduction = 'producao'  # Production (products x sectors)
        sSheetImport = 'importacao'  # Imports (products x 1 vector)

        # ==========================================================================================================
        # Import values from TRUs
        # ==========================================================================================================

        # Intermediate Consumption
        ## This year's prices
        mIntermConsum, vNameProduct, vNameSector, vCodeProduct, vCodeSector = \
            Support.load_tru(sDirectoryInput, sFileUses, sSheetIntermedConsum,
                             nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors, bNames=True, bCodes=True)

        ## Last year's prices
        mIntermConsumLY = Support.load_tru(sDirectoryInput, sFileUsesLY, sSheetIntermedConsum,
                                           nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors)
        # Final Demand
        ## This year's prices
        mDemand, vNameProduct1, vNameDemand = \
            Support.load_tru(sDirectoryInput, sFileUses, sSheetDemand,
                             nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nColsDemand, bNames=True)

        ## Last year's prices
        mDemandLY = Support.load_tru(sDirectoryInput, sFileUsesLY, sSheetDemand,
                                     nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nColsDemand)

        # Added Value (only this year's prices)
        mAddedValue, vNameAddedValue, vNameSector1 = \
            Support.load_tru(sDirectoryInput, f"{nSectors}_tab2_{nYear}.xls", sSheetAddedValue,
                             nRowIni=5, nColIni=1, nRows=nRowsAV, nCols=nSectors, bNames=True)
        # Supply
        ## This year's prices
        mOffer, vNameProduct2, vNameOffer = \
            Support.load_tru(sDirectoryInput, sFileResources, sSheetOffer,
                             nRowIni=5, nColIni=nColIniOffer, nRows=nProducts, nCols=nColsOffer, bNames=True)

        ## Last year's prices
        mOfferLY = Support.load_tru(sDirectoryInput, sFileResourcesLY, sSheetOffer,
                                    nRowIni=5, nColIni=nColIniOffer, nRows=nProducts, nCols=nColsOffer)

        # Production
        ## This year's prices
        mProduction = Support.load_tru(sDirectoryInput, sFileResources, sSheetProduction,
                                       nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors)

        ## Last year's prices
        mProductionLY = Support.load_tru(sDirectoryInput, sFileResourcesLY, sSheetProduction,
                                         nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors)

        # Imports (three columns until 2009 for nDimension == 2)
        if nYear < 2010 and nDimension == 2:
            vImport = np.sum(Support.load_tru(sDirectoryInput, sFileResources, sSheetImport,
                                              nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=3),
                             axis=1, keepdims=True)
            vImportLY = np.sum(Support.load_tru(sDirectoryInput, sFileResourcesLY, sSheetImport,
                                                nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=3),
                               axis=1, keepdims=True)

        else:
            ## This year's prices
            vImport = Support.load_tru(sDirectoryInput, sFileResources, sSheetImport,
                                       nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=1)
            ## Last year's prices
            vImportLY = Support.load_tru(sDirectoryInput, sFileResourcesLY, sSheetImport,
                                         nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=1)

        # ==========================================================================================================
        # Calculating coefficients without stock variation
        # ==========================================================================================================

        ## Copying demand
        mDemandWithoutEstock = np.copy(mDemand)
        mDemandWithoutEstockLY = np.copy(mDemandLY)

        # Excluding ∆ stock column (giving it all 0s in order to maintain the number of columns)
        mDemandWithoutEstock[:, nColEstockVar] = 0.0
        mDemandWithoutEstockLY[:, nColEstockVar] = 0.0

        # Calculating the distribution/alpha matrix (see function documentation for more details)
        mDistribution, mTotalConsum = Support.distribution_matrix_calcul(mIntermConsum, mDemandWithoutEstock)
        mDistributionLY, mTotalConsumLY = Support.distribution_matrix_calcul(mIntermConsumLY, mDemandWithoutEstockLY)

        # ==========================================================================================================
        # Calculating arrays internally distributed by alphas
        # For each product/sector pair, estimates the margin of trade/taxes paid of product i in sector j,
        # under the assumption that the margins/taxes follow the same distribution observed in production
        # ==========================================================================================================

        ## Trade margins
        nColMarginTrade = 1 if nDimension != 2 else 0
        mMarginTrade = Support.calculation_margin(mDistribution, mOffer, nColMarginTrade, vRowsTradeElim)
        mMarginTradeLY = Support.calculation_margin(mDistributionLY, mOfferLY, nColMarginTrade, vRowsTradeElim)

        ## Transport margins
        nColMarginTransport = 2 if nDimension != 2 else 1
        mMarginTransport = Support.calculation_margin(mDistribution, mOffer, nColMarginTransport, vRowsTranspElim)
        mMarginTransportLY = Support.calculation_margin(mDistributionLY, mOfferLY, nColMarginTransport, vRowsTranspElim)

        ## Taxes
        nColIPI = 4 if nDimension != 2 else 3
        mIPI = Support.calculation_internal_matrix(mDistribution, mOffer, nColIPI)
        mIPILY = Support.calculation_internal_matrix(mDistributionLY, mOfferLY, nColIPI)

        nColICMS = 5 if nDimension != 2 else 4
        mICMS = Support.calculation_internal_matrix(mDistribution, mOffer, nColICMS)
        mICMSLY = Support.calculation_internal_matrix(mDistributionLY, mOfferLY, nColICMS)

        nColOtherTaxes = 6 if nDimension != 2 else 5
        mOtherTaxes = Support.calculation_internal_matrix(mDistribution, mOffer, nColOtherTaxes)
        mOtherTaxesLY = Support.calculation_internal_matrix(mDistributionLY, mOfferLY, nColOtherTaxes)

        # ==========================================================================================================
        # Calculating coefficients without exports and stock variation
        # This will be used to calculate the distribution of imports and import taxes
        # ==========================================================================================================

        # Copying demand without stock
        mDemandWithoutExport = np.copy(mDemandWithoutEstock)
        mDemandWithoutExportLY = np.copy(mDemandWithoutEstockLY)

        # Excluding export column (giving it all 0s in order to maintain the number of columns)
        mDemandWithoutExport[:, nColExport] = 0
        mDemandWithoutExportLY[:, nColExport] = 0

        # Calculating the distribution/alpha matrix (see function documentation for more details)
        mDistributionWithoutExport, mTotalConsumWithoutExport = \
            Support.distribution_matrix_calcul(mIntermConsum, mDemandWithoutExport)
        mDistributionWithoutExportLY, mTotalConsumWithoutExportLY = \
            Support.distribution_matrix_calcul(mIntermConsumLY, mDemandWithoutExportLY)

        # ==========================================================================================================
        # Calculating Arrays internally distributed by alphas without exports
        # ==========================================================================================================

        ## For each product/sector pair, estimates the import of product i by sector j + the import taxes involved
        nColImport = 0
        mImport = Support.calculation_internal_matrix(mDistributionWithoutExport, vImport, nColImport)
        mImportLY = Support.calculation_internal_matrix(mDistributionWithoutExportLY, vImportLY, nColImport)

        nColImportTax = 3 if nDimension != 2 else 2
        mImportTax = Support.calculation_internal_matrix(mDistributionWithoutExport, mOffer, nColImportTax)
        mImportTaxLY = Support.calculation_internal_matrix(mDistributionWithoutExportLY, mOfferLY, nColImportTax)

        # ==========================================================================================================
        # Calculating the Matrix of Consumption with base prices
        # ==========================================================================================================

        ## Creating total consumption matrix
        mTotalConsum = np.concatenate((mIntermConsum, mDemand), axis=1)
        mTotalConsumLY = np.concatenate((mIntermConsumLY, mDemandLY), axis=1)

        ## Subtracting all margins in taxes in order to arrive at NATIONAL consumption in base prices (not market)
        mConsumBasePrice = \
            mTotalConsum - mMarginTrade - mMarginTransport - mIPI - mICMS - mOtherTaxes - mImport - mImportTax
        mConsumBasePriceLY = mTotalConsumLY - mMarginTradeLY - mMarginTransportLY - \
            mIPILY - mICMSLY - mOtherTaxesLY - mImportLY - mImportTaxLY

        # Transposing production matrix (nSectors x nProducts)
        mProductionTrans = mProduction.T
        mProductionTransLY = mProductionLY.T

        # Creating X Vector (total production by sector in consumer prices)
        vX = np.sum(mProduction, axis=0)
        vX_Products = np.sum(mProduction, axis=1)

        vXLY = np.sum(mProductionLY, axis=0)
        vX_ProductsLY = np.sum(mProductionLY, axis=1)

        # Total final demand
        vDemand = np.sum(mDemand, axis=0)
        vDemandLY = np.sum(mDemandLY, axis=0)

        # Total demand at base prices
        vTotalDemandBasePrices_Sectors = np.sum(mConsumBasePrice, axis=0)
        vTotalDemandBasePrices_Products = np.sum(mConsumBasePrice, axis=1)

        vTotalDemandBasePrices_SectorsLY = np.sum(mConsumBasePriceLY, axis=0)
        vTotalDemandBasePrices_ProductsLY = np.sum(mConsumBasePriceLY, axis=1)

        # ==========================================================================================================
        # Storing results
        # ==========================================================================================================

        # Populating arrays
        arrConsumBasePrices[0, nYear - nFirstYear, :, :] = mConsumBasePrice
        arrProductionTrans[0, nYear - nFirstYear, :, :] = mProductionTrans
        arrValueAdded[0, nYear - nFirstYear, :, :] = mAddedValue
        arr_vX[0, nYear - nFirstYear, :] = vX
        arr_vX_Prod[0, nYear - nFirstYear, :] = vX_Products
        arr_Demand[0, nYear - nFirstYear, :] = vTotalDemandBasePrices_Sectors
        arr_VPB[0, nYear - nFirstYear, 0] = np.sum(vX)

        arrConsumBasePrices[1, nYear - nFirstYear, :, :] = mConsumBasePriceLY
        arrProductionTrans[1, nYear - nFirstYear, :, :] = mProductionTransLY
        arrValueAdded[1, nYear - nFirstYear, :, :] = mAddedValue  # no information at last year's prices
        arr_vX[1, nYear - nFirstYear, :] = vXLY
        arr_vX_Prod[1, nYear - nFirstYear, :] = vX_ProductsLY
        arr_Demand[1, nYear - nFirstYear, :] = vTotalDemandBasePrices_SectorsLY
        arr_VPB[1, nYear - nFirstYear, 0] = np.sum(vXLY)

        arrValueAdded[2, nYear - nFirstYear, :, :] = mAddedValue  # no information at last year's prices

        # ==========================================================================================================
        # Deflators and TRU (last year's prices) adjustments (see Alves-Passoni (2019, p. 192-193)
        # ==========================================================================================================

        if nYear > nFirstYear:
            # Auxiliary variable for array indexing
            nArrayPosition = nYear - nFirstYear

            # Cell-specific indexes
            ## Uses at base prices
            mConsumBasePrices_Deflator = \
                arrConsumBasePrices[0, nArrayPosition, :, :] / arrConsumBasePrices[1, nArrayPosition, :, :]

            ## Production matrix
            mProductionTrans_Deflator = \
                arrProductionTrans[0, nArrayPosition, :, :] / arrProductionTrans[1, nArrayPosition, :, :]

            ## Production vector
            v_vX_Deflator = arr_vX[0, nArrayPosition, :] / arr_vX[1, nArrayPosition, :]
            v_vX_Prod_Deflator = arr_vX_Prod[0, nArrayPosition, :] / arr_vX_Prod[1, nArrayPosition, :]

            ## Demand
            v_Demand_Deflator = arr_Demand[0, nArrayPosition, :] / arr_Demand[1, nArrayPosition, :]

            # General deflator
            n_VPB_Deflator = np.sum(arr_VPB[0, nArrayPosition, :] / arr_VPB[1, nArrayPosition, :])

            # Adjustments
            ## nan: 0/0 --> 1
            ## 0: 0/something --> adjustment via markdown
            ## inf: something / 0 --> adjustment via markdown

            ## Markdown adjustments
            ### Make matrix
            mMarkdown_ProductionTrans = mProductionTrans / vX[:, None]

            ### Uses matrix
            mMarkdown_ConsumBasePrice = mConsumBasePrice / np.sum(mConsumBasePrice, 0)[None, :]

            ### Adjustment
            mProductionTransLY_Adjusted = np.where(
                (np.isinf(mProductionTrans_Deflator)) | (mProductionTrans_Deflator == 0),
                mMarkdown_ProductionTrans * vXLY[:, None],
                mProductionTransLY
            )
            mConsumBasePriceLY_Adjusted = np.where(
                (np.isinf(mConsumBasePrices_Deflator)) | (mConsumBasePrices_Deflator == 0),
                mMarkdown_ConsumBasePrice * np.sum(mConsumBasePriceLY, 0)[None, :],
                mConsumBasePriceLY
            )

            ### GRAS
            mProductionTransLY_Adjusted, r1, s1, nIter1 = gras(
                mA=mProductionTransLY_Adjusted,
                vRowRestriction=vXLY,
                vColRestriction=vX_ProductsLY
            )
            mConsumBasePriceLY_Adjusted, r2, s2, nIter2 = gras(
                mA=mConsumBasePriceLY_Adjusted,
                vRowRestriction=vTotalDemandBasePrices_ProductsLY,
                vColRestriction=vTotalDemandBasePrices_SectorsLY
            )

            # Recalculating deflators
            ## Uses at base prices
            mConsumBasePrices_Deflator = \
                arrConsumBasePrices[0, nArrayPosition, :, :] / mConsumBasePriceLY_Adjusted

            ## Production matrix
            mProductionTrans_Deflator = \
                arrProductionTrans[0, nArrayPosition, :, :] / mProductionTransLY_Adjusted

            ## nan adjustment
            mConsumBasePrices_Deflator = np.where(np.isnan(mConsumBasePrices_Deflator), 1, mConsumBasePrices_Deflator)
            mProductionTrans_Deflator = np.where(np.isnan(mProductionTrans_Deflator), 1, mProductionTrans_Deflator)
            v_Demand_Deflator = np.where(np.isnan(v_Demand_Deflator), 1, v_Demand_Deflator)

            # ==========================================================================================================
            # Chained Price Indexes
            # ==========================================================================================================

            if nYear > nFirstYear:
                ## Uses at base prices
                arrConsumBasePrices_Index[nArrayPosition, :, :] = \
                    mConsumBasePrices_Deflator * arrConsumBasePrices_Index[nArrayPosition - 1, :, :]

                ## Production matrix
                arrProductionTrans_Index[nArrayPosition, :, :] = \
                    mProductionTrans_Deflator * arrProductionTrans_Index[nArrayPosition - 1, :, :]

                ## Production vector
                arr_vX_Index[nArrayPosition, :] = v_vX_Deflator * arr_vX_Index[nArrayPosition - 1, :]
                arr_vX_Prod_Index[nArrayPosition, :] = v_vX_Prod_Deflator * arr_vX_Prod_Index[nArrayPosition - 1, :]

                ## Demand
                arr_Demand_Index[nArrayPosition, :] = v_Demand_Deflator * arr_Demand_Index[nArrayPosition - 1, :]

                ## General deflator
                arr_VPB_Index[nArrayPosition] = n_VPB_Deflator * arr_VPB_Index[nArrayPosition - 1]

    # ==================================================================================================================
    # Double-Deflation (see Alves-Passoni (2019, p.81-85))
    # ==================================================================================================================

    # Populating matrices with data valued at current year prices
    arrConsumBasePrices_Deflated = np.copy(arrConsumBasePrices[0, :, :, :])
    arrProductionTrans_Deflated = np.copy(arrProductionTrans[0, :, :, :])
    arr_vX_Deflated = np.copy(arr_vX[0, :, :])
    arr_vX_Prod_Deflated = np.copy(arr_vX_Prod[0, :, :])
    arr_Demand_Deflated = np.copy(arr_Demand[0, :, :])
    arr_VPB_Deflated = np.sum(np.copy(arr_VPB[0, :, :]), axis=1)

    # Deflating
    for nYear in lYears:
        # Auxiliary variable for array indexing
        nArrayPosition = nYear - nFirstYear

        # Deflation 1: adjusting for inflation (cell-specific)
        ## Uses at base prices
        arrConsumBasePrices_Deflated[nArrayPosition, :, :] = \
            arrConsumBasePrices_Deflated[nArrayPosition, :, :] / arrConsumBasePrices_Index[nArrayPosition, :, :] * 100

        ## Production matrix
        arrProductionTrans_Deflated[nArrayPosition, :, :] = \
            arrProductionTrans_Deflated[nArrayPosition, :, :] / arrProductionTrans_Index[nArrayPosition, :, :] * 100

        ## Total sectoral production
        arr_vX_Deflated[nArrayPosition, :] = \
            arr_vX_Deflated[nArrayPosition, :] / arr_vX_Index[nArrayPosition, :] * 100

        arr_vX_Prod_Deflated[nArrayPosition, :] = \
            arr_vX_Prod_Deflated[nArrayPosition, :] / arr_vX_Prod_Index[nArrayPosition] * 100

        ## Total sectoral/final demand
        arr_Demand_Deflated[nArrayPosition, :] = \
            arr_Demand_Deflated[nArrayPosition, :] / arr_Demand_Index[nArrayPosition] * 100

        ## Total production
        arr_VPB_Deflated[nArrayPosition] = \
            arr_VPB_Deflated[nArrayPosition] / arr_VPB_Index[nArrayPosition] * 100

        # Deflation 2: adjusting for changes in real purchaser power of each industry
        ## Uses at base prices
        arrConsumBasePrices_Deflated[nArrayPosition, :, :] = \
            arrConsumBasePrices_Deflated[nArrayPosition, :, :] * arrConsumBasePrices_Index[nArrayPosition, :, :] / \
            arr_VPB_Index[nArrayPosition]

        ## Production matrix
        arrProductionTrans_Deflated[nArrayPosition, :, :] = \
            arrProductionTrans_Deflated[nArrayPosition, :, :] * arrProductionTrans_Index[nArrayPosition, :, :] / \
            arr_VPB_Index[nArrayPosition]

        ## Total sectoral/product production
        arr_vX_Deflated[nArrayPosition, :] = \
            arr_vX_Deflated[nArrayPosition, :] * arr_vX_Index[nArrayPosition, :] / arr_VPB_Index[nArrayPosition]

        arr_vX_Prod_Deflated[nArrayPosition, :] = \
            arr_vX_Prod_Deflated[nArrayPosition, :] * arr_vX_Prod_Index[nArrayPosition] / arr_VPB_Index[nArrayPosition]

        ## Total sectoral/final demand
        arr_Demand_Deflated[nArrayPosition, :] = \
            arr_Demand_Deflated[nArrayPosition, :] * arr_Demand_Index[nArrayPosition] / arr_VPB_Index[nArrayPosition]

    # Checking consistency (total production in each year)
    a0 = np.sum(arrProductionTrans_Deflated, axis=(1, 2)) / 1000
    a1 = np.sum(arrConsumBasePrices_Deflated, axis=(1, 2)) / 1000
    a2 = np.sum(arr_vX_Deflated, axis=1) / 1000
    a3 = np.sum(arr_vX_Prod_Deflated, axis=1) / 1000
    a4 = np.sum(arr_Demand_Deflated, axis=1) / 1000
    a5 = arr_VPB_Deflated / 1000

    # Adding deflated components to arrays
    arrConsumBasePrices[2, :, :, :] = arrConsumBasePrices_Deflated
    arrProductionTrans[2, :, :, :] = arrProductionTrans_Deflated
    arr_vX[2, :, :] = arr_vX_Deflated
    arr_vX_Prod[2, :, :] = arr_vX_Prod_Deflated
    arr_Demand[2, :, :] = arr_Demand_Deflated
    arr_VPB[2, :, 0] = arr_VPB_Deflated

    # ==================================================================================================================
    # Transforming to sector x sector format
    # ==================================================================================================================

    # Disabling warnings for division by 0 (they are handled after)
    np.seterr(divide='ignore', invalid='ignore')

    # Creating E matrix and intermediate consumtpion in basic prices
    arr_mE = arrConsumBasePrices[:, :, :, nSectors:]
    arr_mIC = arrConsumBasePrices[:, :, :, :nSectors]

    # Creating D Matrix: for each product, calculate the proportion of production made by each sector
    # relative to the total production of product j (∑mD cols (across rows) = 1)
    arr_mD = arrProductionTrans[:, :, :, :] / arr_vX_Prod[:, :, None, :]
    # Checking for NaNs and infinities
    arr_mD = np.nan_to_num(arr_mD, nan=0, posinf=0, neginf=0)

    # Creating Matrix of National Coefficients - Bn Matrix
    # (proportion of the product i that sector j consumed relative to the total production of sector j)
    arr_mBn = arr_mIC / arr_vX[:, :, None, :]
    # Checking for NaNs and infinities
    arr_mBn = np.nan_to_num(arr_mBn, nan=0, posinf=0, neginf=0)

    # Creating A Matrix (Technical Coefficients) - Sector x Sector
    ## mD: proportion of production made by each sector i relative to the total of product j
    ## mBn: proportion of product i that sector j consumed relative to the total production of sector j
    ### Creating skeleton
    arr_mA = np.zeros((3, len(lYears), nSectors, nSectors), dtype=float)

    ### Looping in order to matrix multiplicate
    for nDim in range(3):
        for nYear in lYears:
            arr_mA[nDim, nYear - nFirstYear, :, :] = np.dot(arr_mD[nDim, nYear - nFirstYear, :, :],
                                                            arr_mBn[nDim, nYear - nFirstYear, :, :])

    # Creating Z Matrix (Intermediate Consumption) - Sector x Sector
    ## For each sector, multiply the direct technical coefficient by the sector's total production,
    ## which equals the intermediate consumption by sector j from the production of sector i
    arr_mZ = arr_mA * arr_vX[:, :, None, :]

    # Creating Y Matrix (nSectors x 6 (components of demand))
    # mD: proportion of production made by each sector i relative to the total production of product j
    # mE: final demand for each product
    ### Creating skeleton
    arr_mY = np.zeros((3, len(lYears), nSectors, nColsDemand), dtype=float)

    ### Looping in order to matrix multiplicate
    for nDim in range(3):
        for nYear in lYears:
            arr_mY[nDim, nYear - nFirstYear, :, :] = np.dot(arr_mD[nDim, nYear - nFirstYear, :, :],
                                                            arr_mE[nDim, nYear - nFirstYear, :, :])

    # Calculating Leontief Matrix (I-A)^(-1) (Sector x Sector)
    ## Identity matrix
    mI = np.eye(nSectors)

    ## Skeleton
    arr_mLeontief = np.zeros((3, len(lYears), nSectors, nSectors), dtype=float)

    ### Looping in order to invert matrices
    for nDim in range(3):
        for nYear in lYears:
            arr_mLeontief[nDim, nYear - nFirstYear, :, :] = np.linalg.inv(mI - arr_mA[nDim, nYear - nFirstYear, :, :])

    ## Concatenating mZ and mY
    arr_MIP = np.zeros((3, len(lYears), nSectors, nSectors + nColsDemand), dtype=float)
    for nDim in range(3):
        for nYear in lYears:
            arr_MIP[nDim, nYear - nFirstYear, :, :] = np.concatenate(
                (
                    arr_mZ[nDim, nYear - nFirstYear, :, :],
                    arr_mY[nDim, nYear - nFirstYear, :, :]
                ),
                axis=1
            )

    # ==================================================================================================================
    # Writing Excel file (with multiple sheets)
    # ==================================================================================================================

    print(f"Estimation complete! Writing data to Excel... ({time.strftime('%d/%b/%Y - %H:%M:%S', time.localtime())})\n")

    # Defining years to be written (Excel não aguenta todos)
    lYearsToWrite = [2010, 2013, 2016, 2019]
    # lYearsToWrite = lYears

    # List that contains the data to be written in each sheet
    ## 2010
    vDataSheet = [arr_MIP[0, 0, :, :]]
    ## 2011-19
    for nYear in range(nFirstYear + 1, nLastYear):
        for nDim in range(3):
            if nYear in lYearsToWrite:
                vDataSheet.append(arr_MIP[nDim, nYear - nFirstYear, :, :])

    # Sheet names
    ## 2010
    vSheetNames = [f"MIP_{nFirstYear}"]
    ## 2011-19
    vMIPNames = ["PCorrentes", "PAno_Anterior", "Deflacionada"]
    for nYear in range(nFirstYear + 1, nLastYear):
        for sNome in vMIPNames:
            if nYear in lYearsToWrite:
                vSheetNames.append(f"MIP_{nYear}_{sNome}")

    # Column Labels
    lColumnLabels = np.hstack((vNameSector, vNameDemand))
    vColumnLabels = [lColumnLabels]
    for nYear in range(nFirstYear + 1, nLastYear):
        for nDim in range(3):
            if nYear in lYearsToWrite:
                vColumnLabels.append(lColumnLabels)

    # Row Labels
    vRowLabels = [vNameSector]
    for nYear in range(nFirstYear + 1, nLastYear):
        for nDim in range(3):
            if nYear in lYearsToWrite:
                vRowLabels.append(vNameSector)

    ## Writing Excel file to Output directory
    # Indicators of number of sectors
    nSectorsOutputFile = f"{nSectors}" if nSectors <= 68 else "68+"

    # Title string and calling function to write data
    sFileSheet = f"MIPs_Deflacionadas_{nSectorsOutputFile}.xlsx"
    Support.write_data_excel(sFileSheet, vSheetNames, vDataSheet, vRowLabels, vColumnLabels)

    ## Calculating total execution time
    nEndModel = time.perf_counter()
    nElapsedTime = (nEndModel - nBeginModel) / 60.

    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print(f"Begun at {time.strftime('%d/%b/%Y - %H:%M:%S', sTimeBeginModel)}")
    print(f"Ended at {time.strftime('%d/%b/%Y - %H:%M:%S', time.localtime())}")
    print(f"Time spent: {round(nElapsedTime, 3)} mins ({round(nElapsedTime * 60,  1)} seconds).")
    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
