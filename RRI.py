#///////////////////////////////////////////////////////////////////////////////
# @file		RRI.py
#
# @brief    This file creates a ranked list of given teams using a modified
#           version of the USAU RRI algorithm. 
#
#           This script was created for a local ultimate org and is tailored to
#           their league needs. This RRI implementation excludes the date weight.
#			In the original algorithm, date weight would weight the first week at 0.5
#			and the most recent week at 1.0, with exponential growth in between. 
#
# @inputs	GameData.xlsx- A spreadsheet containing all played games to be calculated.
#			This sheet must have a tab titled "RESULTS" with the following column headers:
#			WEEK | TEAM_A | TEAM_A_SCORE | TEAM_B | TEAM_B_SCORE
#		    The GameData spreadsheet can have as many additional sheets as desired, 
#           so long as the RESULTS sheet exists in the correct format.
#
# @outputs	Rankings.xlsx- A spreadsheet containing the list of existing teams ranked
#           by their calculated RRI
#
#///////////////////////////////////////////////////////////////////////////////

import sys
import os
import fileinput
import shutil
import pandas as pd
import numpy as np
from datetime import datetime
from math import comb

#Global variables
df_Games = pd.DataFrame
df_TeamScoresIter = pd.DataFrame
df_FinalRankings = pd.DataFrame

#///////////////////////////////////////////////////////////////////
# @brief Calculate rankings
#
#   This function is called by main to calculate the rankings.
#
# @inputs   There are no direct inputs to the function. However,
#			the GameData file is found using the path
#
#///////////////////////////////////////////////////////////////////
def calculate_rankings():

	#Limit traceback information to only display relevant information
	sys.tracebacklimit=0

	#Get GameData.xlsx file from path
	path = os.getcwd()
	gameDataFile = path + r'\\GameData.xlsx'

	if ( os.path.exists(gameDataFile) != True) :
		raise Exception("Path error. GameData.xlsx file not found.") from None

	#Extract the game data
	print('\nExtracting the game data.')
	extract_game_data(gameDataFile)
	#Write output to the file
	print('Creating Rankings.xlsx.')
	create_rankings_file()
	print('\nSuccess! Rankings.xlsx can be found in the current directory.')



#///////////////////////////////////////////////////////////////////
# @brief Extract the game data
#
#	This function is called to extract the game data.
#
# @inputs	gameDataFile- the GameData.xlsx file path
#
#///////////////////////////////////////////////////////////////////
def extract_game_data(gameDataFile):

	global df_Games
	global df_TeamScoresIter
	global df_FinalRankings

	#Set setpoints
	seed = 700 #initial score for all teams on iteration 0

	#Create the data frame
	sheetName = 'RESULTS'
	columnsToImport = ['WEEK', 'TEAM_A', 'TEAM_A_SCORE', 'TEAM_B', 'TEAM_B_SCORE']
	df_Games = pd.read_excel(gameDataFile, sheet_name = sheetName, usecols = columnsToImport)
	df_Games.head()

	#Initialize columns to hold calculated scores for each game
	df_Games["TeamA_GameScore"] = [0.0] * len(df_Games)
	df_Games["TeamB_GameScore"] = [0.0] * len(df_Games)
	df_Games["GameWeight"] = [0.0] * len(df_Games)
	#df_Games["dateWeight"] = [0] * len(df_Games)

	#Cast created columns to float
	df_Games["TeamA_GameScore"] = df_Games["TeamA_GameScore"].astype(float)
	df_Games["TeamB_GameScore"] = df_Games["TeamB_GameScore"].astype(float)
	df_Games["GameWeight"] = df_Games["GameWeight"].astype(float)
	#df_Games["DateWeight"] = df_Games["DateWeight"].astype(float)

	#Create the list of teams
	teamListAll = df_Games["TEAM_A"] #Full list of TEAM_A teams, may contain duplicates
	teams = []
	for i in teamListAll:  #Remove duplicates
		if i not in teams:
			teams.append(i)
	teamListAll = df_Games["TEAM_B"] #Full list of TEAM_B teams, may contain duplicates
	for i in teamListAll:  #Remove duplicates
		if i not in teams:
			teams.append(i)

	#Create the data frame that will hold averaged team scores for each iteration
	df_TeamScoresIter = pd.DataFrame(columns = teams)
	df_TeamScoresIter.loc[0] = [seed] * len(teams) #Initialize team scores at starting setpoint

	#Create the data frame that will hold each game score and its weight
	dummyArray = np.empty((100000,len(teams)))
	dummyArray[:] = np.nan
	df_TeamGameScores = pd.DataFrame(dummyArray, columns=teams)

	#Count how many games each team has had
	teamGameCount = pd.Series(index=teams)
	for i in teams:
		teamGameCount[i] = int(0)

	#Initialize the metric to test whether iterations have converged
	#Use SSE of the difference between the current and previous iteration
	#SSE initialized to something close to what we expect the average to be
	#Alternatively, while loop can be replaced with a for loop for explicit number of iterations
	SSE = 1000
	SSE_log = []
	iter = 0
	print('Calculating the RRI. This may take a few minutes.')
	while SSE > 1:
		iter += 1

		#Calculate the game scores of each game
		for j in range(0,len(df_Games)):
			# pull the data
			teamA = df_Games["TEAM_A"].iloc[j]
			teamB = df_Games["TEAM_B"].iloc[j]
			scoreA = int(df_Games["TEAM_A_SCORE"].iloc[j])
			scoreB = int(df_Games["TEAM_B_SCORE"].iloc[j])
			prevScoreA = df_TeamScoresIter[teamA].iloc[iter-1]
			prevScoreB = df_TeamScoresIter[teamB].iloc[iter-1]
			week = df_Games["WEEK"].loc[j]

			#Assign winner and loser
			if scoreA > scoreB:
				winner = teamA
				W = scoreA
				L = scoreB
			elif scoreB > scoreA:
				winner = teamB
				W = scoreB
				L = scoreA
			else:
				winner = "TIE"

			#Calculate the difference. Difference will be negative if B is the winner.
			r = L/(W-1)
			if winner == teamA:
				diff = 125 + 475 * np.sin((min(1, (1 - r) * 2)) * 0.4 * np.pi) / np.sin(0.4 * np.pi)
			elif winner == teamB:
				diff = -(125 + 475 * np.sin((min(1,(1-r)*2)) * 0.4 * np.pi)/np.sin(0.4 * np.pi))
			elif winner == "TIE":
				diff = 0

			#Calculate the game weight
			gameWeight = np.sqrt((W + max(L,np.sqrt((W-1)/2)))/19)

			#Calculate the game scores the each team and store to the data frame
			df_TeamGameScores.loc[int(teamGameCount[teamA]), teamA] = (prevScoreB + diff) * gameWeight 
			df_TeamGameScores.loc[int(teamGameCount[teamB]), teamB] = (prevScoreA - diff) * gameWeight 
			
			#Exclude game in the case of a tie or if difference between scores is greater than the diff at 600
			if diff == 0:
				df_TeamGameScores.loc[int(teamGameCount[teamA]), teamA]	 = np.NaN
				df_TeamGameScores.loc[int(teamGameCount[teamB]), teamB]	  = np.NaN
			if abs(diff) == 600 and prevScoreA - prevScoreB > 600:
				df_TeamGameScores.loc[int(teamGameCount[teamA]), teamA]	 = np.NaN
				df_TeamGameScores.loc[int(teamGameCount[teamB]), teamB]	  = np.NaN

			#Write data to data frame for later data validation
			df_Games.loc[j, "TeamA_GameScore"] = df_TeamGameScores.loc[int(teamGameCount[teamA]), teamA]
			df_Games.loc[j, "TeamB_GameScore"] = df_TeamGameScores.loc[int(teamGameCount[teamB]), teamB]
			df_Games.loc[j, "GameWeight"] = gameWeight

			teamGameCount[teamA] += 1
			teamGameCount[teamB] += 1

		#Calculate the final team scores for this iteration
		teamScores = pd.Series(index=teams)
		for i in teams:
			teamScores[i] = np.nanmean(df_TeamGameScores[i]) #Ignore cells set to nan
		df_TeamScoresIter.loc[iter] = teamScores


		#Calculate updated SSE metric
		SSE = sum((df_TeamScoresIter.iloc[iter].subtract(df_TeamScoresIter.iloc[iter-1]))**2)	 #Matrix subtraction, itemwise square(**2), then sum
		SSE_log.append(SSE)

	#Write final converged scores
	iterations = len(df_TeamScoresIter)
	df_FinalRankings = df_TeamScoresIter.iloc[iterations-1].sort_values(ascending=False)

	#Print summary
	pd.set_option('display.max_rows', None)
	print("NUMBER OF ITERATIONS = ", end="")
	print(iterations)
	print("FINAL RANKINGS:")
	print(df_FinalRankings)



#///////////////////////////////////////////////////////////////////
# @brief Create the Rankings.xlsx file
#
#	This function is called to create the .xlsx file itself. It
#	uses the dynamic ranking data created earlier to write the
#	output into one file.
#
# @outputs	Rankings.xlsx- a spreadsheet containing the rankings,
#			as well as the score summaries by game and iterations.
#			The latter two are provided for data validation only.
#
#			The file is generated into the current working directory
#
#///////////////////////////////////////////////////////////////////
def create_rankings_file():

	global df_Games
	global df_TeamScoresIter
	global df_FinalRankings

	#File will be generated into current working directory
	path = os.getcwd()
	rankingsFile = path + r'\\Rankings.xlsx'

    #Append timestamp to generated sheets so same file can be used without overwriting previous week results
    #Note that same day runs will still be overwritten unless generated sheets are manually renamed
	dateObj = datetime.now()
	suffix = dateObj.strftime("%Y-%m-%d") 
	fileName = rankingsFile
	gameSheetName = 'ByGame_' + suffix
	teamSheetName = 'ByTeam_' + suffix
	rankingSheetName = 'Ranking_' + suffix
	
	if os.path.exists(rankingsFile) :
		writerMode = 'a'
		writer = pd.ExcelWriter(rankingsFile, engine='openpyxl', mode=writerMode, if_sheet_exists='replace')
	else :
		writerMode = 'w'
		writer = pd.ExcelWriter(rankingsFile, engine='openpyxl', mode=writerMode)
	df_FinalRankings.to_excel(writer, sheet_name=rankingSheetName)
	df_Games.to_excel(writer, sheet_name=gameSheetName)
	df_TeamScoresIter.to_excel(writer, sheet_name=teamSheetName)
	writer.close()



#- MAIN ------------------------------------------------------------------------------

#///////////////////////////////////////////////////////////////////
# @brief Main
#
#		The main function of the script. Main has no inputs or
#		outputs.
#
#///////////////////////////////////////////////////////////////////
def main():
	calculate_rankings()


#///////////////////////////////////////////////////////////////////
# @brief Only run script directly
#
#		This ensures the script will only execute when intended
#
#///////////////////////////////////////////////////////////////////
if __name__ == "__main__":
	try:
		main()
	except Exception as e:
		print("Error: ", end="")
		print(e)
		sys.exit(1)