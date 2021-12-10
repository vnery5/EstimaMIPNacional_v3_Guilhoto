"""
Reads and compiles product data from retropolated TRUs (2000 - 2019).
"""

## Import numpy, time and support file
import numpy as np
import pandas as pd
import time
import SupportFunctions_V3 as Support

## Only run if it's the main file (don't run on import)
if __name__ == '__main__':
    # ==================================================================================================================
    # Parameters for MIP Estimation
    # ==================================================================================================================

    ## Constants that determine...
    # ... size of MIP (possible values: 0 - 12x12; 1 - 20x20; 2 - 107X51; 3 - 128x68)
    nDimension = 2
    # First year
    nYear0 = 2000
    
    # Last year of available data + 1
    nYear1 = 2019 + 1

    ## Estimate at this year's or last year's prices?
    bThisYearPrices = True

    ## TRUs indexes and starting year
    nIndexUses = 2 if bThisYearPrices else 4
    nIndexResources = 1 if bThisYearPrices else 3
    nYear0 = nYear0 if bThisYearPrices else nYear0 + 1

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
    for nYear in range(nYear0, nYear1):
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
    
        ## String that identifies the Uses' spreadsheet file
        sFileUses = f"{nSectors}_tab{nIndexUses}_{nYear}.xls"
        # Sheet Names
        sSheetIntermedConsum = 'CI'  # Intermediate Consumption
        sSheetDemand = 'demanda'  # Final Demand
        sSheetAddedValue = 'VA'  # Added Value
    
        ## String that identifies the Resources' spreadsheet file
        sFileResources = f"{nSectors}_tab{nIndexResources}_{nYear}.xls"
        sSheetOffer = 'oferta'  # Supply Components (taxes, margins and base prices)
        sSheetProduction = 'producao'  # Production (products x sectors)
        sSheetImport = 'importacao'  # Imports (products x 1 vector)
    
        # ==============================================================================================================
        # Parameters for aggregation and disaggregation
        # ==============================================================================================================
    
        sFileAgregacao = "Agregação.xlsx"
        sSheetNumeroAgregacoes = "NumeroAgregacoes"
        sSheetAgregacaoSetor = "AgregaçãoSetor"
        sSheetAgregacaoProduto = "AgregaçãoProduto"
    
        sFileDesagregacao = "Desagregação.xlsx"
        sSheetNumeroDesagregacoes = "NumeroDesagregacoes"
        sSheetDesagregacaoSetor = "DesagregaçãoSetor"
        sSheetDesagregacaoProduto = "DesagregaçãoProduto"
    
        # ==============================================================================================================
        # STARTING ESTIMATION
        # ==============================================================================================================
    
        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
        print(f"+ Compilação de TRUs - {nYear} +")
        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    
        ## Timers
        nBeginModel = time.perf_counter()
        sTimeBeginModel = time.localtime()
    
        # ==============================================================================================================
        # Import values from TRUs
        # ==============================================================================================================
        
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
            Support.load_tru(sDirectoryInput, f"{nSectors}_tab2_{nYear}.xls", sSheetAddedValue,
                             nRowIni=5, nColIni=1, nRows=nRowsAV, nCols=nSectors, bNames=True)
        # Supply
        mOffer, vNameProduct2, vNameOffer = \
            Support.load_tru(sDirectoryInput, sFileResources, sSheetOffer,
                             nRowIni=5, nColIni=nColIniOffer, nRows=nProducts, nCols=nColsOffer, bNames=True)
        # Production
        mProduction = Support.load_tru(sDirectoryInput, sFileResources, sSheetProduction,
                                       nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=nSectors)
        # Imports (three columns until 2009 for nDimension == 2)
        if nYear < 2010 and nDimension == 2:
            vImport = np.sum(Support.load_tru(sDirectoryInput, sFileResources, sSheetImport,
                                              nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=3), axis=1, keepdims=True)

        else:
            vImport = Support.load_tru(sDirectoryInput, sFileResources, sSheetImport,
                                       nRowIni=5, nColIni=nColIni, nRows=nProducts, nCols=1)

        if nYear == nYear0:
            mExportacao = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                       columns=[f"{i}" for i in range(nYear0, nYear1)])
            mGoverno = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                    columns=[f"{i}" for i in range(nYear0, nYear1)])
            mConsumo = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                    columns=[f"{i}" for i in range(nYear0, nYear1)])
            mInvestimento = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                         columns=[f"{i}" for i in range(nYear0, nYear1)])
            mImpostos = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                     columns=[f"{i}" for i in range(nYear0, nYear1)])
            mImportacao = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                       columns=[f"{i}" for i in range(nYear0, nYear1)])
            mOfertaNacPB = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                        columns=[f"{i}" for i in range(nYear0, nYear1)])
            mOfertaTotPB = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                        columns=[f"{i}" for i in range(nYear0, nYear1)])
            mOfertaTotPC = pd.DataFrame(np.zeros([nProducts, nYear1 - nYear0]), index=vNameProduct,
                                        columns=[f"{i}" for i in range(nYear0, nYear1)])
    
        mExportacao.values[:, nYear - nYear0] = np.sum(mDemand[:, nColExport], axis=1)
        mGoverno.values[:, nYear - nYear0] = mDemand[:, 2]
        mConsumo.values[:, nYear - nYear0] = mDemand[:, 3] + mDemand[:, 4]
        mInvestimento.values[:, nYear - nYear0] = mDemand[:, nColFBCF]
        mImpostos.values[:, nYear - nYear0] = mOffer[:, 6]
        mImportacao.values[:, nYear - nYear0] = vImport[:, 0]
        mOfertaNacPB.values[:, nYear - nYear0] = np.sum(mProduction, axis=1)
        mOfertaTotPB.values[:, nYear - nYear0] = np.sum(mProduction, axis=1) + vImport[:, 0]
        mOfertaTotPC.values[:, nYear - nYear0] = np.sum(mProduction, axis=1) + vImport[:, 0] + np.sum(mOffer, axis=1)

    vNomes = ["Exportação", "Consumo do Governo", "Consumo das Famílias_ISFLSF",
              "Investimento_FBCF", "Impostos_Subsidios", "Importação", "OfertaNacPB",
              "mOfertaTotPB", "mOfertaTotPC"]

    vDados = [mExportacao, mGoverno, mConsumo, mInvestimento, mImpostos, mImportacao,
              mOfertaNacPB, mOfertaTotPB, mOfertaTotPC]

    ## Creating Writer object (allows to write multiple sheets into a single file)
    sPAIndicator = "" if bThisYearPrices else "PA"
    Writer = pd.ExcelWriter(f"{sDirectoryOutput}Compilacao_Tru_Retro{sPAIndicator}.xlsx", engine='openpyxl')
    # Lists to store dataframe
    lDataFrames = []

    ## For each dataframe,
    for nSheet in range(len(vNomes)):
        ## Append to list
        lDataFrames.append(vDados[nSheet])

        ## Writing to Excel
        lDataFrames[nSheet].to_excel(Writer, vNomes[nSheet], freeze_panes=(1, 1))

    ## Saving file
    Writer.save()
    print("Fim!")
