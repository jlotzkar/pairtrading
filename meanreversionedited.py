# -*- coding: utf-8 -*-

"""

Created on Fri May 28 12:42:52 2021

 

@author: Lotzkar

"""

 

#%% Importing packages

 

import numpy as np #numpy package for stat functions

 

import pandas as pd #pandas package for scraping

pd.__version__

 

#pd is the common "alias" used in Python community

#http://pandas.pydata.org/pandas-docs/stable/

 

#df is the name for DataFrame used here

 

import matplotlib.pyplot as plt #package for plotting and visualizing data

 

import matplotlib.patches as mpatches #for matplot legend

 

from statsmodels.tsa.stattools import coint #for cointegration testing

 

from scipy.stats import shapiro #shapiro-wilk test for normality

 

import seaborn #for cointegration heatmap

 

import os

 

import streamlit as st #for web visualization

 

import win32com.client #for updating excel sheets where data are being read in from

 

import time #waiting for Excel refresh

 

import plotly.graph_objects as go

 

from plotly import tools

 

import plotly.offline as py

 

import plotly.express as px

# =============================================================================

# import dash

# import dash_core_components as dcc

# import dash_html_components as html

# =============================================================================

 

#plotly

#import plotly.plotly as py

#import plotly.express as px

# =============================================================================

# import plotly.tools as tls

# from plotly.graph_objs import *

# =============================================================================

#%% PART 1 - DATA ANALYSIS

#Importing Excel data

 

#dictionary mapping proposed dataframes to Excel workbooks

dictExcel = {

    "dfCommServices": "SPTSX_CommServices.xlsm",

    "dfConsDiscretionary": "SPTSX_ConsumerDiscretionary.xlsm",

    "dfConsStaples": "SPTSX_ConsumerStaples.xlsm",

    "dfEnergy": "SPTSX_Energy.xlsm",

    "dfFinancials": "SPTSX_Financials.xlsm",

    "dfHealth": "SPTSX_HealthCare.xlsm",

    "dfIndustrials": "SPTSX_Industrials.xlsm",

    "dfIT": "SPTSX_IT.xlsm",

    "dfMaterials": "SPTSX_Materials.xlsm",

    "dfRealEstate": "SPTSX_RealEstate.xlsm",

    "dfUtilities": "SPTSX_Utilities.xlsm"}

 

folder = "C:/Users/lotzkar/Desktop/Mean Reversion Project/StockData/"

 

#test code

# =============================================================================

# dictExcel2 = {

#     "dfEnergy": "SPTSX_Energy.xlsm",

#     "dfIT": "SPTSX_IT.xlsm"

#     }

# =============================================================================

 

#function to automatically refresh Excel sheets' price data

def refr_sheet(filename):

    app_global = win32com.client.Dispatch("Excel.Application")

    global_sprdsheet = app_global.Workbooks.open(folder+filename)

    global_sprdsheet.RefreshAll()

    global_sprdsheet.Application.Run(filename+"!refresh_bbg")

    time.sleep(10)

    global_sprdsheet.Save()

    global_sprdsheet.Close()

    #app_global.Quit()

   

 

#first, refresh data to get historical three year period stock prices from today

#=============================================================================

#filepath = "C:/Users/lotzkar/Desktop/Mean Reversion Project/StockData/"

for key, value in dictExcel.items():

    refr_sheet(value)

   

#%%

#=============================================================================

#creating dataframes from updated Excel information

dfCommServices = pd.read_excel("StockData/SPTSX_CommServices.xlsm")

dfConsDiscretionary = pd.read_excel("StockData/SPTSX_ConsumerDiscretionary.xlsm")

dfConsStaples = pd.read_excel("StockData/SPTSX_ConsumerStaples.xlsm")

dfEnergy = pd.read_excel("StockData/SPTSX_Energy.xlsm")

dfFinancials = pd.read_excel("StockData/SPTSX_Financials.xlsm")

dfHealth = pd.read_excel("StockData/SPTSX_HealthCare.xlsm")

dfIndustrials = pd.read_excel("StockData/SPTSX_Industrials.xlsm")

dfIT = pd.read_excel("StockData/SPTSX_IT.xlsm")

dfMaterials = pd.read_excel("StockData/SPTSX_Materials.xlsm")

dfRealEstate = pd.read_excel("StockData/SPTSX_RealEstate.xlsm")

dfUtilities = pd.read_excel("StockData/SPTSX_Utilities.xlsm")

 

#deleting extra row

dfCommServices = dfCommServices.iloc[1:]

dfConsDiscretionary = dfConsDiscretionary.iloc[1:]

dfConsStaples = dfConsStaples.iloc[1:]

dfEnergy = dfEnergy.iloc[1:]

dfFinancials = dfFinancials.iloc[1:]

dfHealth = dfHealth.iloc[1:]

dfIndustrials = dfIndustrials.iloc[1:]

dfIT = dfIT.iloc[1:]

dfMaterials = dfMaterials.iloc[1:]

dfRealEstate = dfRealEstate.iloc[1:]

dfUtilities = dfUtilities.iloc[1:]

 

#renaming first column to 'Date' for ease

dfCommServices.rename(columns={dfCommServices.columns[0]: 'Date'}, inplace = True)

dfConsDiscretionary.rename(columns={dfConsDiscretionary.columns[0]: 'Date'}, inplace = True)

dfConsStaples.rename(columns={dfConsStaples.columns[0]: 'Date'}, inplace = True)

dfEnergy.rename(columns={dfEnergy.columns[0]: 'Date'}, inplace = True)

dfFinancials.rename(columns={dfFinancials.columns[0]: 'Date'}, inplace = True)

dfHealth.rename(columns={dfHealth.columns[0]: 'Date'}, inplace = True)

dfIndustrials.rename(columns={dfIndustrials.columns[0]: 'Date'}, inplace = True)

dfIT.rename(columns={dfIT.columns[0]: 'Date'}, inplace = True)

dfMaterials.rename(columns={dfMaterials.columns[0]: 'Date'}, inplace = True)

dfRealEstate.rename(columns={dfRealEstate.columns[0]: 'Date'}, inplace = True)

dfUtilities.rename(columns={dfUtilities.columns[0]: 'Date'}, inplace = True)

 

#setting dates column to be the index

dfCommServices.set_index(['Date'], drop = True, inplace = True)

dfConsDiscretionary.set_index(['Date'], drop = True, inplace = True)

dfConsStaples.set_index(['Date'], drop = True, inplace = True)

dfEnergy.set_index(['Date'], drop = True, inplace = True)

dfFinancials.set_index(['Date'], drop = True, inplace = True)

dfHealth.set_index(['Date'], drop = True, inplace = True)

dfIndustrials.set_index(['Date'], drop = True, inplace = True)

dfIT.set_index(['Date'], drop = True, inplace = True)

dfMaterials.set_index(['Date'], drop = True, inplace = True)

dfRealEstate.set_index(['Date'], drop = True, inplace = True)

dfUtilities.set_index(['Date'], drop = True, inplace = True)

 

#configuring 'N/A' data in dataframe

dfCommServices = dfCommServices.fillna(0)

dfConsDiscretionary = dfConsDiscretionary.fillna(0)

dfConsStaples = dfConsStaples.fillna(0)

dfEnergy = dfEnergy.fillna(0)

dfFinancials = dfFinancials.fillna(0)

dfHealth = dfHealth.fillna(0)

dfIndustrials = dfIndustrials.fillna(0)

dfIT = dfIT.fillna(0)

dfMaterials = dfMaterials.fillna(0)

dfRealEstate = dfRealEstate.fillna(0)

dfUtilities = dfUtilities.fillna(0)

 

#%% Multiprocessing to speed up function calls

# =============================================================================

# import multiprocessing

# import time

# =============================================================================

 

#%% Testing for cointegration to find relevant security pairs before doing mean reversion analysis

#cointegration: to see if a linear combo of the variables is stationary https://blog.quantinsti.com/pairs-trading-basics/#:~:text=The%20most%20common%20test%20for,of%20the%20variables%20is%20stationary.&text=If%20A%20and%20B%20are,this%20equation%20above%20is%20stationary.)

#i.e. how strongly a price ratio between two securities varies around a mean

 

def cointCalc(df):

    n = df.shape[1]

    pval_matrix = np.ones((n,n)) #creating a matrix of 1's to be later populated with p-values

    keys = df.keys()

    pairs = [] #this will be the array containing suitable pairs to be analyzed

    for i in range(n):

        for j in range(i+1,n):

            security1 = df[keys[i]]

            security2 = df[keys[j]]

            result = coint(security1, security2) #calculating cointegration of two securities

            pval_matrix[i,j] = result[1]

           

            #null hypothesis is series is non-stationary; less than 0.02 and we reject null hypothesis; series is stationary

            if result[1] < 0.02: #certain p-value threshold

                pairs.append((keys[i], keys[j]))

    return pval_matrix, pairs

 

#%% Multiprocessing to speed up function calls

# =============================================================================

# if __name__ == '__main__':

#     pool = multiprocessing.Pool()

#     pool = multiprocessing.Pool(processes=4)

#    

#     inputs = [dfCommServices, dfConsDiscretionary, dfConsStaples, dfEnergy, dfFinancials, dfHealth, dfIndustrials, dfIT, dfMaterials, dfRealEstate, dfUtilities]

#    

#     outputs = pool.map(cointCalc, inputs)

#    

#     print("Output: {}".format(outputs))

# =============================================================================

#calling above function to give p-value matrix and pairs

pvaluesCSe, listOfPairsCSe = cointCalc(dfCommServices)

pvaluesCD, listOfPairsCD = cointCalc(dfConsDiscretionary)

pvaluesCSt, listOfPairsCSt = cointCalc(dfConsStaples)

pvaluesEnergy, listOfPairsEnergy = cointCalc(dfEnergy)

pvaluesFin, listOfPairsFin = cointCalc(dfFinancials)

pvaluesHealth, listOfPairsHealth = cointCalc(dfHealth)

pvaluesInd, listOfPairsInd = cointCalc(dfIndustrials)

pvaluesIT, listOfPairsIT = cointCalc(dfIT)

pvaluesMat, listOfPairsMat = cointCalc(dfMaterials)

pvaluesRE, listOfPairsRE = cointCalc(dfRealEstate)

pvaluesUtil, listOfPairsUtil = cointCalc(dfUtilities)

 

#print(listOfPairsCSe)

 

#cointegration heatmap to see which pairs to investigate

#list of tickers

#names = list(dfCommServices)

 

#seaborn.heatmap(pvaluesCSe, xticklabels = names, yticklabels = names, cmap = 'RdYlGn_r', mask = (pvaluesCSe >= 0.98))

#plt.tight_layout() #to ensure axes are included properly in the plot

#plt.show()

 

 

#%% Getting total combined pairs list thus far

#combined pairs list

combinedListOfPairs = listOfPairsCSe + listOfPairsCD + listOfPairsCSt + listOfPairsEnergy + listOfPairsFin + listOfPairsHealth + listOfPairsInd + listOfPairsIT + listOfPairsMat + listOfPairsRE + listOfPairsUtil

 

#print(combinedListOfPairs)

 

#%%#%% Testing pairs for normality to further filter

#we will use Shapiro-Wilk test due to small (<tens of thousands) sample size, and consider price ratio of pairs already calculated

 

#Price ratio dataframes

 

#merging previous dataframes so we can search specific columns for price histories

combinedf = pd.concat([dfCommServices, dfConsDiscretionary, dfConsStaples, dfEnergy, dfFinancials, dfHealth, dfIndustrials, dfIT, dfMaterials, dfRealEstate, dfUtilities], axis = 1)

 

#drop NANs

combinedf = combinedf.dropna()

 

# =============================================================================

# #remove duplicates

# combinedf = combinedf.T.drop_duplicates().T

#

# #set date as index

# combinedf['Date'] = pd.to_datetime(combinedf['Date'])

# combinedf.set_index(['Date'], drop=True, inplace = True)

# =============================================================================

 

#finding pairs in combinedListOfPairs, getting historical values from combinedf dataframe

combinedListOfPairs = dict(combinedListOfPairs)

priceRatioData = [] #appending ratios to list, will later be converted into dataframe

names=[]

for key, value in combinedListOfPairs.items():

    ratioValue = combinedf[''+key].div(combinedf[''+value].where(combinedf[''+value] != 0, np.nan))

    priceRatioData.append(ratioValue)

    ratioName = key + '/' + value

    names += [ratioName] #compiling names of price ratios so dataframe can be renamed

dfPriceRatios = pd.DataFrame(priceRatioData) #converting list to dataframe

dfPriceRatios = dfPriceRatios.transpose()

 

#renaming dataframe columns

dfPriceRatios.columns = names

 

#testing normality of each column of pair price ratios using Shapiro-Wilk Test

#null hypothesis = normal sample of data, if p-score less than alpha we reject null hypothesis; sample of data not normal. 

#Otherwise, we fail to reject null hypothesis; data is of normal distribution.

#more info: https://www.spss-tutorials.com/spss-shapiro-wilk-test-for-normality/

 

finalPairList1 = [] #appending suitable pairs based on normality test; this will contain our final pairs to study

alpha = 0.04 #alpha value for Shapiro test, 80% of standard 0.05 due to anomalies

for column in dfPriceRatios:

    stat, p = shapiro(dfPriceRatios[column])

    if p > alpha:

        finalPairList1.append(column)

    else:

        continue

#print(finalPairList1)

 

#second normal test to confirm results

#Anderson-Darling Test

#more info: https://machinelearningmastery.com/a-gentle-introduction-to-normality-tests-in-python/

np.seterr(divide='ignore', invalid='ignore') #to ensure no invalid values are divided

from scipy.stats import anderson

finalPairList2 = [] #appending suitable pairs based on normality test; this will contain our final pairs to study

p = 0

for column in dfPriceRatios:

    result = anderson(dfPriceRatios[column])

    for i in range(len(result.critical_values)):

        sl, cv = result.significance_level[i], result.critical_values[i]

        if result.statistic < result.critical_values[i]:

            finalPairList2.append(column)

        else:

            continue

#removing duplicates from the list (finalPairList2)

finalPairList2 = list(set(finalPairList2))

#print(finalPairList2)

 

#Collecting pairs to analyze

#First, check lengths of lists.  If too small, use larger subset.  Otherwise, keep lists.

# =============================================================================

# if len(finalPairList1) < 20:

#     finalAnalysisPairsList = finalPairList2

# elif len(finalPairList2) < 20:

#     finalAnalysisPairsList = finalPairList1

# elif len(finalPairList1) < 20 and len(finalPairList2) < 20:

#     finalPairList1 = [] #appending suitable pairs based on normality test; this will contain our final pairs to study

#     alpha = 0.02 #alpha value for Shapiro test, 80% of standard 0.05 due to anomalies

#     for column in dfPriceRatios:

#         stat, p = shapiro(dfPriceRatios[column])

#         if p > alpha:

#             finalPairList1.append(column)

#         else:

#             continue

#     finalAnalysisPairsList = finalPairList1

# else:

# =============================================================================

#Getting overlapping pairs between two lists from two normality tests

finalAnalysisPairsList = set(finalPairList1).intersection(finalPairList2)

 

#number of elements in list too few; expanding search criteria in finalPairList1 (adjusting Shapiro's test parameters)

if len(finalAnalysisPairsList) < 5:

    finalPairList1 = [] #appending suitable pairs based on normality test; this will contain our final pairs to study

    alpha = 0.02 #alpha value edited for Shapiro test

    for column in dfPriceRatios:

        stat, p = shapiro(dfPriceRatios[column])

        if p > alpha:

            finalPairList1.append(column)

        else:

            continue

    finalAnalysisPairsList = set(finalPairList1).intersection(finalPairList2)

#%% Exporting pairs to Excel

 

pd.DataFrame(finalAnalysisPairsList).to_excel('outputPairsFinal.xlsx', header=False, index=False)

 

#%% PARTS 2 & 3 - PAIR ANALYSIS, PLOTTING ON DASHBOARD

#Reading in Excel pairs that were generated

 

#reading in StockData PairAnalysis_Updated sheet which contains two pairs to be studied

dfPairs = pd.read_excel("StockData/PairAnalysis_Updated.xlsx", sheet_name="Pair Input")

dfPairs = dfPairs.rename(columns={ dfPairs.columns[3]: dfPairs.columns[2] })

col_list = list(dfPairs)

col_list[0], col_list[1], col_list[2] = 0, 'Date', col_list[1]

dfPairs.columns = col_list

 

#drop first column

dfPairs.drop(dfPairs.columns[[0]], axis = 1, inplace = True)

 

#drop first three rows

dfPairs = dfPairs.iloc[3:]

 

#make dates as index by first renaming first column, then transforming

dfPairs = dfPairs.rename(columns={ dfPairs.columns[0]: 'Date' })

dfPairs.set_index(['Date'], drop=True, inplace = True)

 

#plt.figure()

dfPairs.plot()

plt.grid()

plt.title("Historical Prices of Selected Securities")

plt.xlabel("Date")

plt.ylabel("Price")

#plt.close()

 

#saving plot 1 - price histories

plt.savefig('plot1.png')

 

#%% Z-Score function


# #z score = (x - mean) / standard deviation, given price distributions between ratios of selected pairs, we want

# #distribution to be normal with mean 0.  Creates normal data with appropriate threshold levels (sigma).

#

def zScore(series):

    return (series - series.mean()) / (np.std(series))

#

# #%% dfPair manipulation

# #user is able to choose a particular row (with one pair), with the historical dates and see mean reversion analysis

#

#import xbbg import blp

#

# # =============================================================================

# # #user input for dates

# # startdate = input("Enter start date (YYYYMMDD): ")

# # enddate = input("Enter end date (YYYYMMDD: ")

# # firstSecurity = input("Enter first security: ")

# # secondSecurity = input("Enter second security: ")

# #

# # #creating list of dataframes to extract specific pair price histories

# # list_of_df = [dfCommServices, dfConsDiscretionary, dfConsStaples, dfEnergy, dfFinancials, dfHealth, dfIndustrials, dfIT, dfMaterials, dfRealEstate]

# #

# # #merging dataframes to filter down to only selected pair necessary

# # dfFinalPair = pd.concat(list_of_df, axis=1)

# #

# # dfFinalPair = dfFinalPair.dropna(how = 'all', axis = 1, inplace = True)

# # =============================================================================

#

#price ratio of first equity / second equity

ratios = dfPairs[dfPairs.columns[0]]/dfPairs[dfPairs.columns[1]]

# #print(len(ratios))

#

#z-score of price ratio

zScorePair = zScore(ratios)

#

#rolling z-score calculations

#for more accurate spread, use moving average.  This is an included feature in the rolling z-score.  Taking avg of dataset over time.

#here, we using 30 day (1 month) moving average of ratio for the rolling mean,

#5 day moving average of ratio for the current mean

#and 30 day standard deviation

#

movingAverage = 30

currentAverage = 5

stdVal = 30

#

# #train = ratios[:703]

# #test = ratios[703:]

# #ratiosMoving = train.rolling(window=movingAverage,center=False).mean()

rollingMean = ratios.rolling(movingAverage).mean()

currentMean = ratios.rolling(currentAverage).mean()

stdDev = ratios.rolling(stdVal).std(ddof=0) #ddof=0 provides a maximum likelihood estimate of the variance for normally distributed variables

#

# #rolling z-score w/ plot and standard deviation lines

dfPairs['zScoreRolling'] = (currentMean - rollingMean)/(stdDev)

 

#plt.clf()

plt.figure()

dfPairs['zScoreRolling'].plot()

plt.grid()

plt.title("Rolling Ratio Z-Score Plot")

plt.xlabel("Date")

plt.ylabel("Ratio")

 

# zScoreRolling.plot()

plt.axhline(0, color = 'black')

plt.axhline(2, color = 'red', linestyle = '--')

plt.axhline(-2, color = 'green', linestyle = '--')

plt.legend(['Rolling Ratio Z-Score', 'Mean', '+2 STD', '-2 STD'])

 

#saving plot 2 - z-score

plt.savefig('plot2.png')

 

# plt.show()

#

# #%% Graphing the original price histories of the two securities with buy/sell signals

# #Buy is when rolling z-score below -2 SD (expect to get back to 0, ratio to increase)

# #Sell is when rolling z-score above 2 SD (expect to get back to 0, ratio to decrease)

#

# #analysis of when to buy vs sell

# # =============================================================================

# # dfPairs['Buy-Stock1'] = dfPairs.apply(lambda x: x.iloc[:] if x['zScoreRolling'] < -2 else np.nan, axis=1)

# # dfPairs['Sell-Stock1'] = dfPairs.apply(lambda x: x[x.columns[0]] if x['zScoreRolling'] > 2 else np.nan, axis=1)

# #

# # dfPairs['Sell-Stock2'] = dfPairs.apply(lambda x: x[x.columns[1]] if x['zScoreRolling'] < -2 else np.nan, axis=1)

# # dfPairs['Buy-Stock2'] = dfPairs.apply(lambda x: x[x.columns[1]] if x['zScoreRolling'] > 2 else np.nan, axis=1)

# # =============================================================================

firstCol = dfPairs.columns[0]

secondCol = dfPairs.columns[1]

#

plt.figure()

s1 = dfPairs[dfPairs.columns[0]]

s2 = dfPairs[dfPairs.columns[1]]

s1[30:].plot(color='blue')

s2[30:].plot(color='grey')

#

dfPairs['Buy-Stock1'] = dfPairs.apply(lambda x: x[firstCol] if x['zScoreRolling'] < -2 else np.nan, axis=1)

dfPairs['Sell-Stock1'] = dfPairs.apply(lambda x: x[firstCol] if x['zScoreRolling'] >= 2 else np.nan, axis=1)

#

dfPairs['Sell-Stock2'] = dfPairs.apply(lambda x: x[secondCol] if x['zScoreRolling'] < -2 else np.nan, axis=1)

dfPairs['Buy-Stock2'] = dfPairs.apply(lambda x: x[secondCol] if x['zScoreRolling'] >= 2 else np.nan, axis=1)

#

plt.plot(dfPairs.index, dfPairs['Buy-Stock1'],marker='^',linestyle='None',color='g')

plt.plot(dfPairs.index, dfPairs['Sell-Stock1'],marker='^',linestyle='None',color='r')

#

plt.plot(dfPairs.index, dfPairs['Buy-Stock2'],marker='^',linestyle='None',color='g')

plt.plot(dfPairs.index, dfPairs['Sell-Stock2'],marker='^',linestyle='None',color='r')

 

plt.grid()

plt.title("Historical Prices of Selected Securities with Buy/Sell Signals")

plt.xlabel("Date")

plt.ylabel("Price")

 

#legend names for buy/sell signals plot

security1 = dfPairs.columns[0]

security2 = dfPairs.columns[1]

 

plt.legend([security1, security2])

 

#saving plot 2 - z-score

plt.savefig('plot3.png')

#

# plt.show()

#

# #TO DO: Fix legends and plotting

#

# #daily notification of buy or sell (if zScoreRolling either >=2 or <-2)

# # =============================================================================

# # if (np.isnan(dfPairs['Buy-Stock1'].iloc[-1]) and np.isnan(dfPairs['Sell-Stock2'].iloc[-1]):

# #     print("No buying Stock 1 and selling Stock 2 opportunity")

# # elif (np.isnan(dfPairs['Sell-Stock1'].iloc[-1]) and np.isnan(dfPairs['Buy-Stock2'].iloc[-1]):

# #     print("No selling Stock 1 and buying Stock 2 opportunity")

# # else:

# #     print("Opportunity")

# #    

# # if(dfPairs['Buy-Stock1'].iloc[-1] == 'nan'):

# #     print("True")

# # else:

# #     print("False")

# #    

# # if (pd.isnull(dfPairs.at[3,-1]):

# #     print("true")

# # else:

# #     print("false")

# # =============================================================================

    

import base64

 

if((dfPairs['Buy-Stock1'].iloc[-1]) > 0):

     #print("Buy Stock 1 and sell Stock 2.")

     plt.figure()

     plt.text(0.5,0.5,"Buy Stock 1 and sell Stock 2.")

     plt.savefig('plot4.png', dpi=100)

 

elif((dfPairs['Sell-Stock1'].iloc[1]) > 0):

     #print("Sell Stock 1 and buy Stock 2.")

     plt.figure()

     plt.text(0.25,0.5,"Sell Stock 1 and buy Stock 2.")

     plt.savefig('plot4.png', dpi=100)

 

else:

     #print("No buying/selling opportunity today.")

     plt.figure()

     plt.text(0.25,0.5,"No buying/selling opportunity today.")

     plt.savefig('plot4.png', dpi=100)

 

# =============================================================================

#end