'''Author: Robert Shea
Code Purpose: The purpose of this code is to predict the Most Valuable Player for a season in the National Basketball Association. The program will parse statistics from 
various Excel files that are organized by season. The sheets in the Excel files are organized according to different types of statistics to be parsed. The goal is to create
the most accurate model at predicting a players value using statistics that are available from basketball-reference.com'''

import pandas as pd #Libraries that are imported to read and organize data
from pandas import ExcelWriter
from pandas import ExcelFile
'''import seaborn as sns
sns.set(color_codes=True)''' #There were issues in getting this library to work properly and output a scatter chart

class Excel_Data_Reader:
        '''This class will read an excel file, parse each individual player to be evaluated, then parse the respective statistics for each player. Then the statistics/data will be
        used as variables in various equations to understand a players Value when determining whether an NBA player should win the MVP award. Lastly, various tables and graphs will
        be created to display the results of calculations'''

        def __init__(self, path, sheet_name):

                self.path = path #Excel file
                self.sheet_name = sheet_name
                self.players = dict() #Stores each instance of a player to be evaluated using the model
                self.season = []
                self.analyze_excel_file()
                self.table() #This function is for parsing Excel files by a single season
                #self.decade_table() #This function will parse multiple Excel files and output a comprehensive table of Value over the last decade
                #self.chart()

        def file_reader(self, path, sheet_name):
                '''This function will try to read an excel file and return an error if the Excel file can not be found or if it does not exist.'''
                try:
                        df = pd.read_excel(path, sheet_name= sheet_name)
                        return df
                except FileNotFoundError:
                        raise FileNotFoundError(f"Can not open file {path}!")
    
        def analyze_excel_file(self):
                '''This function starts the analysis of an Excel file'''
                player = ''
                for files in self.path: #Useful for when there is more than one file to be analyzed
                        df = self.file_reader(files, 'Per Game')
                        self.season.append(files.strip('.xslx')) #The current season will be appended and added to a list. The excel files are conveniently named according to season
                        for n in df.columns: #Reads each of the first columns in the excel sheet to find the column labeled 'Player'
                                if n == 'Player':
                                        for i in df.index: #Goes through the values of the column that is labelled 'Player' to collect the player names of each player to be analyzed and create a new dictionary instance
                                                player = df[n][i] + ' ' + files.strip('.xslx')
                                                self.players[player] = Players() #A new dictionary key is created for each player
                                                self.players[player].player_name = (df[n][i]).strip('*') #Removes an asterisk for instances where the player is a Hall of Famer
                                                self.players[player].player_index = i
                                                for sheet in self.sheet_name: #Parses the stats by going from sheet to sheet in the excel file and determining which stats are needed from each sheet
                                                        self.parse_stats(files, sheet, player)

                                        '''These are the different functions for calculating different results using the parsed statistics as variables'''
                                                self.fantasy_basketball_stats_totals(player)
                                                self.fantasy_basketball_stats_per_100_poss(player)
                                                self.fantasy_basketball_stats_per_36_min(player)
                                                self.fantasy_basketball_stats_average(player)
                                                self.game_score(player)
                                                self.total_stats(player)
                                                self.net_rating(player)
                                                self.quality_of_impact(player)
                                                self.level_of_impact(player)
                                                self.win_contribution(player)
                                                self.value(player)



        def parse_stats(self, files, sheet, player):
                '''The function for parsing statistics for each player. Each instance of player is evaluated using this function. Also, this takes into account
                that more than one sheet should be analyzed and sometimes more than one file may be analyzed. Each statistic, except the Season, will be stored as a float
                so that it may be used as a variable in an equation.'''
                df = self.file_reader(files, sheet)
                if sheet in ['Per Game', 'Totals', 'Advanced','Per 100 Poss', 'Per 36 Min']: #These sheets contain specific individual statistics that are to be stored as values for each instance in the players dictionary
                        for n in df.columns:
                                if n == 'Season': #Stores the season for the player
                                        self.players[player].season = str(df[n][self.players[player].player_index])

                                elif n == 'G': #Stores the total games played by the player
                                        self.players[player].games_played = float(df[n][self.players[player].player_index])

                                elif n == 'MP' and sheet == 'Per Game': #Stores the minutes played per game by the player
                                        self.players[player].minutes_per_game = float(df[n][self.players[player].player_index])

                                elif n == 'PER': #Stores each players Player Efficiency Rating as found from the 'Advanced' sheet
                                        self.players[player].player_efficiency_rating = float(df[n][self.players[player].player_index])

                                elif n == 'TS%': #Stores each players True Shooting Percentage as found from the 'Advanced' sheet
                                        self.players[player].true_shooting_percentage = float(df[n][self.players[player].player_index])
                
                                elif n == 'USG%': #Stores each players Usage Rate as found from the 'Advanced' sheet
                                        self.players[player].usage = float(df[n][self.players[player].player_index])

                                elif n == 'WS': #Stores each players Win Shares as found from the 'Advanced' sheet
                                        self.players[player].Win_Share = float(df[n][self.players[player].player_index])

                                elif n == 'VORP': #Stores each players Value Over Replacement Player as found from the 'Advanced' sheet
                                        self.players[player].VORP = float(df[n][self.players[player].player_index])

                                elif n == 'ORtg': #Stores each players Offensive Rating as found from the 'Per 100 Poss' sheet
                                        self.players[player].offensive_rating = float(df[n][self.players[player].player_index])

                                elif n == 'DRtg': #Stores each players Defensive Rating as found from the 'Per 100 Poss' sheet
                                        self.players[player].defensive_rating = float(df[n][self.players[player].player_index])

                                elif n == 'FG' and sheet == 'Per Game': #The field goals made are taken from the 'Per Game' sheet for the Game Score calculation
                                        self.players[player].FGM_per_game = float(df[n][self.players[player].player_index])
        
                                elif n == 'FGA' and sheet == 'Per Game': #The field goals attempted are taken from the 'Per Game' sheet for the Game Score calculation
                                        self.players[player].FGA_per_game = float(df[n][self.players[player].player_index])

                                elif n == 'FT' and sheet == 'Per Game': #The free throws made are taken from the 'Per Game' sheet for the Game Score calculation
                                        self.players[player].FTM_per_game = float(df[n][self.players[player].player_index])

                                elif n == 'FTA' and sheet == 'Per Game': #The free throws attempted are taken from the 'Per Game' sheet for the Game Score calculation
                                        self.players[player].FTA_per_game = float(df[n][self.players[player].player_index])

                                elif n == 'ORB' and sheet == 'Per Game': #The offensive rebounds are taken from the 'Per Game' sheet for the Game Score calculation
                                        self.players[player].ORB_per_game = float(df[n][self.players[player].player_index])
        
                                elif n == 'DRB' and sheet == 'Per Game': #The defensive rebounds are taken from the 'Per Game' sheet for the Game Score calculation
                                        self.players[player].DRB_per_game = float(df[n][self.players[player].player_index])

                                elif n == 'TRB' and sheet == 'Totals': #The total rebounds (offensive rebounds + defensive rebounds) are taken from three different sheets for the Fantasy Basketball Stats Calculation
                                        self.players[player].total_rebounds = float(df[n][self.players[player].player_index])
                                elif n == 'TRB' and sheet == 'Per 100 Poss':
                                        self.players[player].rebounds_per_100_poss = float(df[n][self.players[player].player_index])
                                elif n == 'TRB' and sheet == 'Per 36 Min':
                                        self.players[player].rebounds_per_36_min = float(df[n][self.players[player].player_index])

                                elif n == 'AST' and sheet == 'Per Game': #The assists are taken from four different sheets for the Fantasy Basketball Stats Calculation and Game Score calculation
                                        self.players[player].assists_per_game = float(df[n][self.players[player].player_index]) 
                                elif n == 'AST' and sheet == 'Totals':
                                        self.players[player].total_assists = float(df[n][self.players[player].player_index])
                                elif n == 'AST' and sheet == 'Per 100 Poss':
                                        self.players[player].assists_per_100_poss = float(df[n][self.players[player].player_index]) 
                                elif n == 'AST' and sheet == 'Per 36 Min':
                                        self.players[player].assists_per_36_min = float(df[n][self.players[player].player_index]) 

                                elif n == 'STL' and sheet == 'Per Game': #The steals are taken from four different sheets for the Fantasy Basketball Stats Calculation and Game Score calculation
                                        self.players[player].steals_per_game = float(df[n][self.players[player].player_index])
                                elif n == 'STL' and sheet == 'Totals':
                                        self.players[player].total_steals = float(df[n][self.players[player].player_index])
                                elif n == 'STL' and sheet == 'Per 100 Poss':
                                        self.players[player].steals_per_100_poss = float(df[n][self.players[player].player_index])
                                elif n == 'STL' and sheet == 'Per 36 Min':
                                        self.players[player].steals_per_36_min = float(df[n][self.players[player].player_index])

                                elif n == 'BLK' and sheet == 'Per Game': #The blocks are taken from four different sheets for the Fantasy Basketball Stats Calculation and Game Score calculation
                                        self.players[player].blocks_per_game = float(df[n][self.players[player].player_index])
                                elif n == 'BLK' and sheet == 'Totals':
                                        self.players[player].total_blocks = float(df[n][self.players[player].player_index])
                                elif n == 'BLK' and sheet == 'Per 100 Poss':
                                        self.players[player].blocks_per_100_poss = float(df[n][self.players[player].player_index])
                                elif n == 'BLK' and sheet == 'Per 36 Min':
                                        self.players[player].blocks_per_36_min = float(df[n][self.players[player].player_index])

                                elif n == 'TOV' and sheet == 'Per Game': #The turnovers are taken from four different sheets for the Fantasy Basketball Stats Calculation and Game Score calculation
                                        self.players[player].turnovers_per_game = float(df[n][self.players[player].player_index]) 
                                elif n == 'TOV' and sheet == 'Totals':
                                        self.players[player].total_turnovers = float(df[n][self.players[player].player_index])
                                elif n == 'TOV' and sheet == 'Per 100 Poss':
                                        self.players[player].turnovers_per_100_poss = float(df[n][self.players[player].player_index])
                                elif n == 'TOV' and sheet == 'Per 36 Min':
                                        self.players[player].turnovers_per_36_min = float(df[n][self.players[player].player_index])

                                elif n == 'PF' and sheet == 'Per Game': #The personal fouls are taken from four different sheets for the Fantasy Basketball Stats Calculation and Game Score calculation
                                        self.players[player].fouls_per_game = float(df[n][self.players[player].player_index])
                                elif n == 'PF' and sheet == 'Totals':
                                        self.players[player].total_fouls = float(df[n][self.players[player].player_index])
                                elif n == 'PF' and sheet == 'Per 100 Poss':
                                        self.players[player].fouls_per_100_poss = float(df[n][self.players[player].player_index])
                                elif n == 'PF' and sheet == 'Per 36 Min':
                                        self.players[player].fouls_per_36_min = float(df[n][self.players[player].player_index])

                                elif n == 'PTS' and sheet == 'Per Game': #The points scored are taken from four different sheets for the Fantasy Basketball Stats Calculation and Game Score calculation     
                                        self.players[player].points_per_game= float(df[n][self.players[player].player_index])
                                elif n == 'PTS' and sheet == 'Totals':             
                                        self.players[player].total_points= float(df[n][self.players[player].player_index])
                                elif n == 'PTS' and sheet == 'Per 100 Poss':             
                                        self.players[player].points_per_100_poss = float(df[n][self.players[player].player_index])
                                elif n == 'PTS' and sheet == 'Per 36 Min':             
                                        self.players[player].points_per_36_min = float(df[n][self.players[player].player_index])
                                
                elif sheet == 'MVP Voting':
                        '''The sheet named 'MVP Voting' contains the results of the MVP voting for the season and each player that received a vote will have their ranking (according to votes received)
                        stored as a value for each instance of player'''
                        for n in df.columns:
                                for i in df.index:
                                        if n == 'Player' and self.players[player].player_name in df[n][i]: #The 'Player' column in the sheet is first found and the value for players name is cross checked with the list of players to detect if that player received MVP votes
                                                if str(df['Rank'][i]).endswith('T'): #If the Rank ends in T it means the player tied and this will remove the T
                                                        self.players[player].MVP_vote_standing = int(df['Rank'][i].strip('T'))
                                                else:
                                                        self.players[player].MVP_vote_standing = int(df['Rank'][i])
                                
                elif sheet == 'MVP Tracker':
                        '''The sheet named 'MVP Tracker' contains the current odds for players that are in the MVP conversation.
                The players are ranked on probability of winning MVP and the ranking is parsed to detect if the player with the highest Value calculated from the model is the same
                player that is the frontrunner for MVP.'''
                        for n in df.columns:
                                for i in df.index:
                                        if n == 'Player' and self.players[player].player_name in df[n][i]:
                                                if str(df['Rk'][i]).endswith('T'):
                                                        self.players[player].MVP_vote_standing = int(df['Rk'][i].strip('T'))
                                                else:
                                                        self.players[player].MVP_vote_standing = int(df['Rk'][i])

                elif sheet in ['ATL', 'BKN', 'BOS', 'CHA', 'CHI', 'CLE', 'DAL', 'DEN', 'DET','GS', 'HOU', 'IND', 'LAC', 'LAL', 'MEM', 'MIA', 'MIL', 'MIN', 'NO', 'NY', 'OKC', 'ORL', 'PHI', 'PHX', 'POR', 'SA', 'SAC', 'TOR', 'UTAH', 'WSH']:
                        '''Used to figure out how many wins a player had in a season. The number of wins is used for the Level of Impact calculation. 
                        This is not the most efficienct way of determining how many wins a player has in a season but it is easy to understand and implement.'''
                        for n in df.columns:
                                for i in df.index:
                                        if n == 'Player' and self.players[player].player_name in df[n][i]: '''Each team sheet has a roster of the players on that team for the season. If the player is on that roster then the wins associated with that team will be added as a value for that instance of player'''
                                                self.players[player].wins = int(df['Wins'][0])


        
        def fantasy_basketball_stats_totals(self, player):
                '''Function to calculate Fantasy Basketball Stats Totals based on a players total box score statistics'''
                player = self.players[player]
                player.fantasy_basketball_stats_totals = ((player.total_points)*player.true_shooting_percentage) + 1.5*(player.total_assists) + 1.2*(player.total_rebounds) + 3*(player.total_blocks) + 3*(player.total_steals) - player.total_fouls - player.total_turnovers
                return player.fantasy_basketball_stats_totals

        def fantasy_basketball_stats_per_36_min(self, player):
                '''Function to calculate Fantasy Basketball Stats Per 36 Minutes based on a players per 36 minutes box score statistics'''
                player = self.players[player]
                player.fantasy_basketball_stats_per_36_min = ((player.points_per_36_min)*player.true_shooting_percentage) + 1.5*(player.assists_per_36_min) + 1.2*(player.rebounds_per_36_min) + 3*(player.blocks_per_36_min) + 3*(player.steals_per_36_min) - player.fouls_per_36_min - player.turnovers_per_36_min
                return player.fantasy_basketball_stats_per_36_min

        def fantasy_basketball_stats_per_100_poss(self, player):
                '''Function to calculate Fantasy Basketball Stats Per 100 Possessions based on a players per 100 possessions box score statistics'''
                player = self.players[player]
                player.fantasy_basketball_stats_per_100_poss = ((player.points_per_100_poss)*player.true_shooting_percentage) + 1.5*(player.assists_per_100_poss) + 1.2*(player.rebounds_per_100_poss) + 3*(player.blocks_per_100_poss) + 3*(player.steals_per_100_poss) - player.fouls_per_100_poss - player.turnovers_per_100_poss
                return player.fantasy_basketball_stats_per_100_poss

        def fantasy_basketball_stats_average(self, player):
                '''Takes an average for the three different calculations of Fantasy Basketball Statistics. The calculations have different weightings because
                the total box score statistics would heavily skew the results if it was weighted significantly.'''
                player = self.players[player]
                player.fantasy_basketball_stats_average = .275*(player.fantasy_basketball_stats_totals)+ .3625*(player.fantasy_basketball_stats_per_100_poss) + .3625*(player.fantasy_basketball_stats_per_36_min)
                return player.fantasy_basketball_stats_average
                
        def game_score(self, player):
                '''Calculation  of Game Score, a statistic that was created by John Hollinger to give an idea fo a player's productivity for a single game.
                The weightings are widely accepted so the equation was used with per game box score statistics'''
                player = self.players[player]
                player.game_score = (player.points_per_game + .4*(player.FGM_per_game) - .7*(player.FGA_per_game) - .4*(player.FTA_per_game - player.FTM_per_game) + .7*(player.ORB_per_game) + .3*(player.DRB_per_game) + player.steals_per_game + .7*(player.assists_per_game) + .7*(player.blocks_per_game) - .4*(player.fouls_per_game) - player.turnovers_per_game)
                return player.game_score

        def total_stats(self, player):
                '''Calculation of Total Stats: the first component of determining a players Value'''
                player = self.players[player]
                player.total_stats = (.3*(player.fantasy_basketball_stats_average) + .3*(player.game_score) + .4*(player.player_efficiency_rating))
                return player.total_stats

        def net_rating(self, player):
                '''Calculation of Net Rating by subtracting a players offensive rating from their defensive rating. Used in the calculation of Quality of Impact'''
                player = self.players[player]
                player.net_rating = player.offensive_rating - player.defensive_rating
                return player.net_rating

        def quality_of_impact(self, player):
                '''Calculation of Quality of Impact: the first component of Win Contribution'''
                player = self.players[player]
                player.quality_of_impact = .35*(player.Win_Share) + .35*(player.VORP) + .3*(player.net_rating)
                return player.quality_of_impact

        def level_of_impact(self, player):
                '''Calculation of Level of Impact: the second component of Win Contribution'''
                player = self.players[player]
                player.level_of_impact = player.wins*(player.games_played/82)*(player.minutes_per_game/48)*(player.usage/100)
                return player.level_of_impact
        
        def win_contribution(self, player):
                '''Calculation of Win Contribution: The second component of Value'''
                player = self.players[player]
                player.win_contribution = player.quality_of_impact * player.level_of_impact
                return player.win_contribution
        
        def value(self, player):
                '''Calculation of a players Value'''
                player = self.players[player]
                player.value = .5*(player.total_stats) + .5*(player.win_contribution)
                return player.value
        
        def table(self):
                '''This function will create tables and a scatter chart for whatever season is being evaluated. The method of creating the tables is not the most efficient but
                it is simple and works properly.'''
                for season in self.season:
                        '''These lists will store values according to each instance of player in the order that they are evaluated. The values are appended to the end
                        of the list so that when the lists need to be evaluated, the players statistics are in the same index spot in each list'''
                        total_player_name = []
                        total_season = []
                        total_fantasy_basketball_stats_average = []
                        total_game_score = []
                        total_player_efficiency_rating = []
                        total_total_stats = []
                        total_VORP = []
                        total_quality_of_impact = []
                        total_level_of_impact = []
                        total_win_contribution = []
                        total_value = []
                        total_MVP_vote_standing = []
                        for player in self.players:
                                candidate = self.players[player]
                                total_player_name.append(candidate.player_name)
                                total_season.append(candidate.season)
                                total_fantasy_basketball_stats_average.append(candidate.fantasy_basketball_stats_average)
                                total_game_score.append(candidate.game_score)
                                total_player_efficiency_rating.append(candidate.player_efficiency_rating)
                                total_total_stats.append(candidate.total_stats)
                                total_VORP.append(candidate.VORP)
                                total_quality_of_impact.append(candidate.quality_of_impact)
                                total_level_of_impact.append(candidate.level_of_impact)
                                total_win_contribution.append(candidate.win_contribution)
                                total_value.append(candidate.value)
                                total_MVP_vote_standing.append(candidate.MVP_vote_standing)
                        MVP_candidates = {'Player': total_player_name, 'Season': total_season,'Fantasy Basketball Stats Average': total_fantasy_basketball_stats_average, 'Game Score': total_game_score,'Player Efficiency Rating': total_player_efficiency_rating, 'Total Stats': total_total_stats, 'VORP': total_VORP, 'Quality of Impact': total_quality_of_impact, 'Level of Impact': total_level_of_impact, 'Win Contribution': total_win_contribution, 'Value': total_value, 'MVP Voting Standing': total_MVP_vote_standing}
                        table = pd.DataFrame(MVP_candidates) #Creates a table using the Pandas library that contains columns that are the same as listed in line 316
                        graph = pd.DataFrame({'Player': total_player_name, 'Season': total_season, 'Total Stats': total_total_stats, 'Win Contribution': total_win_contribution}) #Less data is needed when creating a scatter chart
                        with pd.ExcelWriter(str(season) + '_Results.xlsx') as writer: #Writes the chart into a new excel file that is named for the season and _Results
                                table.sort_values('Value',ascending=False).to_excel(writer, sheet_name='Value Descending') #A table is created in the first sheet named 'Value Descending' that is organized by each players value in descending order
                                table.sort_values('Predicted MVP Voting Standing',ascending=True).to_excel(writer, sheet_name='MVP Vote Standing') #A table is created in the second sheet named 'MVP Vote Standing' that organizes the table by each players MVP rank in ascending order
                                table.sort_values('VORP',ascending=False).to_excel(writer, sheet_name='VORP') #A table is created in the third sheet named 'VORP' that is organized by VORP of the players in descending order
                                graph.to_excel(writer, sheet_name='Total Stats vs Win Contribution') #A blank graph is created in the fourth sheet named 'Total Stats vs Win Contribution'. The graph will be filled with series which correspond to players values
                                workbook = writer.book
                                worksheet=writer.sheets['Total Stats vs Win Contribution']
                                chart = workbook.add_chart({'type': 'scatter'}) #Adds a scatter chart
                                max_row = len(total_total_stats) #Figures out the length of the total_total_stats list to determine how many players, or series, are to be evaluated
                                for i in range(len(total_player_name)):
                                        col = i + 1
                                        chart.add_series({'name': ['Total Stats vs Win Contribution', col, 1], 
                                        'categories': ['Total Stats vs Win Contribution', col, 3, col, 3], #The x values in the scatter chart are the Total Stats
                                        'values': ['Total Stats vs Win Contribution', col, 4, col, 4], #The y values in the scatter chart are the Win Contribution
                                        'marker': {'type': 'circle', 'size': 7},})
                                chart.add_series({'categories': ['Total Stats vs Win Contribution', 1, 3, max_row, 3], #Adds a regression line or trendline to the scatter chart
                                        'values': ['Total Stats vs Win Contribution', 1, 4, max_row, 4],
                                        'marker': {'type': 'none'}, 
                                        'trendline': {'type': 'linear'}})
                                chart.set_title({'name': 'Total Stats vs Win Contribution'})                
                                chart.set_x_axis({'name': 'Total Stats', 'min': 80})
                                chart.set_y_axis({'name': 'Win Contribution',
                                'major_gridlines': {'visible': False}})
                                worksheet.insert_chart('K2', chart, {'x_offset': 25, 'y_offset': 10})
                                writer.save()

        '''def decade_table(self): #Similar funtion to the table() function except this is used for when more than one Excel file, or season, is to be evaluated to create a comprehensive table or scatter chart
                total_player_name = []
                total_season = []
                total_fantasy_basketball_stats_average = []
                total_game_score = []
                total_player_efficiency_rating = []
                total_total_stats = []
                total_VORP = []
                total_quality_of_impact = []
                total_level_of_impact = []
                total_win_contribution = []
                total_value = []
                total_MVP_vote_standing = []
                for player in self.players:
                        candidate = self.players[player]
                        total_player_name.append(candidate.player_name)
                        total_season.append(candidate.season)
                        total_fantasy_basketball_stats_average.append(candidate.fantasy_basketball_stats_average)
                        total_game_score.append(candidate.game_score)
                        total_player_efficiency_rating.append(candidate.player_efficiency_rating)
                        total_total_stats.append(candidate.total_stats)
                        total_VORP.append(candidate.VORP)
                        total_quality_of_impact.append(candidate.quality_of_impact)
                        total_level_of_impact.append(candidate.level_of_impact)
                        total_win_contribution.append(candidate.win_contribution)
                        total_value.append(candidate.value)
                        total_MVP_vote_standing.append(candidate.MVP_vote_standing)
                MVP_candidates = {'Player': total_player_name, 'Season': total_season,'Fantasy Basketball Stats Average': total_fantasy_basketball_stats_average, 'Game Score': total_game_score,'Player Efficiency Rating': total_player_efficiency_rating, 'Total Stats': total_total_stats, 'VORP': total_VORP, 'Quality of Impact': total_quality_of_impact, 'Level of Impact': total_level_of_impact, 'Win Contribution': total_win_contribution, 'Value': total_value, 'MVP Voting Standing': total_MVP_vote_standing}
                table = pd.DataFrame(MVP_candidates)
                graph = pd.DataFrame({'Player': total_player_name, 'Season': total_season, 'Total Stats': total_total_stats, 'Win Contribution': total_win_contribution})
                with pd.ExcelWriter('Decade_Results.xlsx') as writer:
                        table.sort_values('Value',ascending=False).to_excel(writer, sheet_name='Value Descending')
                        table.sort_values('MVP Voting Standing',ascending=True).to_excel(writer, sheet_name='MVP Vote Standing')
                        table.sort_values('VORP',ascending=False).to_excel(writer, sheet_name='VORP')
                        graph.to_excel(writer, sheet_name='Total Stats vs Win Contribution')
                        workbook = writer.book
                        worksheet=writer.sheets['Total Stats vs Win Contribution']
                        chart = workbook.add_chart({'type': 'scatter'})
                        max_row = len(total_total_stats)
                        for i in range(len(total_player_name)):
                                col = i + 1
                                chart.add_series({'name': ['Total Stats vs Win Contribution', col, 1], 
                                'categories': ['Total Stats vs Win Contribution', col, 3, col, 3],
                                'values': ['Total Stats vs Win Contribution', col, 4, col, 4],
                                'marker': {'type': 'circle', 'size': 7},})
                        chart.add_series({'categories': ['Total Stats vs Win Contribution', 1, 3, max_row, 3],
                                'values': ['Total Stats vs Win Contribution', 1, 4, max_row, 4],
                                'marker': {'type': 'none'}, 
                                'trendline': {'type': 'linear'}})
                        chart.set_title({'name': 'Total Stats vs Win Contribution'})                
                        chart.set_x_axis({'name': 'Total Stats', 'min': 80})
                        chart.set_y_axis({'name': 'Win Contribution',
                        'major_gridlines': {'visible': False}})
                        worksheet.insert_chart('K2', chart, {'x_offset': 25, 'y_offset': 10})
                        writer.save()'''

        '''def graph(self): #Attempted to create a scatter chart using seaborn library and MatPlotLib library but was unsuccessful. May look into it when adjusting model in future
                total_total_stats = []
                total_win_contribution = []
                for player in self.players:
                        candidate = self.players[player]
                        total_total_stats.append(candidate.total_stats)
                        total_win_contribution.append(candidate.win_contribution)
                MVP_candidates = {'Total Stats': total_total_stats, 'Win Contribution': total_win_contribution}
                table = pd.DataFrame(MVP_candidates)
                graph = sns.load_dataset(table)
                ax = sns.regplot(x='Total Stats', y='Win Contribution', data=graph)'''

                       

class Players:
        '''This class stores the values for each instance of player in the self.player dictionary. Each value corresponds to a statistic or result of a calculation'''
        def __init__(self, player_index=None, player_name=None, season=None, games_played=None, minutes_per_game=None, player_efficiency_rating= None, true_shooting_percentage=None, usage= None, Win_Share=None, VORP=None, offensive_rating=None, defensive_rating= None, FGA_per_game=None, FGM_per_game=None, FTM_per_game=None, FTA_per_game= None, ORB_per_game=None, DRB_per_game=None, total_rebounds=None, rebounds_per_100_poss=None, rebounds_per_36_min=None,assists_per_game=None, total_assists=None, assists_per_100_poss=None, assists_per_36_min=None, steals_per_game=None, total_steals=None, steals_per_100_poss=None, steals_per_36_min=None, blocks_per_game=None, total_blocks=None, blocks_per_100_poss=None, blocks_per_36_min=None, turnovers_per_game=None, total_turnovers=None, turnovers_per_100_poss=None, turnovers_per_36_min=None, fouls_per_game=None, total_fouls=None, fouls_per_100_poss=None, fouls_per_36_min=None, points_per_game=None, total_points=None, points_per_100_poss=None, points_per_36_min=None, wins=None, fantasy_basketball_stats_totals=None, fantasy_basketball_stats_per_100_poss=None, fantasy_basketball_stats_per_36_min=None, fantasy_basketball_stats_average=None, game_score=None, total_stats=None, net_rating=None, quality_of_impact=None, level_of_impact=None, win_contribution=None, value=None, MVP_vote_standing=None):
                self.player_name = player_name
                self.player_index = player_index
                self.season = season
                self.games_played= games_played
                self.minutes_per_game= minutes_per_game
                self.player_efficiency_rating= player_efficiency_rating
                self.true_shooting_percentage= true_shooting_percentage
                self.usage= usage
                self.Win_Share= Win_Share
                self.VORP= VORP
                self.offensive_rating= offensive_rating
                self.defensive_rating= defensive_rating
                self.FGM_per_game= FGM_per_game
                self.FGA_per_game= FGA_per_game
                self.FTM_per_game= FTM_per_game
                self.FTA_per_game= FTA_per_game
                self.ORB_per_game= ORB_per_game
                self.DRB_per_game= DRB_per_game
                self.total_rebounds= total_rebounds
                self.rebounds_per_100_poss= rebounds_per_100_poss
                self.rebounds_per_36_min= rebounds_per_36_min
                self.assists_per_game= assists_per_game
                self.total_assists = total_assists
                self.assists_per_100_poss = assists_per_100_poss
                self.assists_per_36_min = assists_per_36_min
                self.steals_per_game= steals_per_game
                self.total_steals= total_steals
                self.steals_per_100_poss= steals_per_100_poss
                self.steals_per_36_min= steals_per_36_min
                self.blocks_per_game= blocks_per_game
                self.total_blocks=total_blocks
                self.blocks_per_100_poss= blocks_per_100_poss
                self.blocks_per_36_min= blocks_per_36_min
                self.turnovers_per_game= turnovers_per_game
                self.total_turnovers = total_turnovers
                self.turnovers_per_100_poss=turnovers_per_100_poss
                self.turnovers_per_36_min=turnovers_per_36_min
                self.fouls_per_game= fouls_per_game
                self.total_fouls=total_fouls
                self.fouls_per_100_poss=fouls_per_100_poss
                self.fouls_per_36_min=fouls_per_36_min
                self.points_per_game = points_per_game
                self.total_points= total_points
                self.points_per_100_poss= points_per_100_poss
                self.points_per_36_min= points_per_36_min
                self.wins= wins
                self.fantasy_basketball_stats_totals = fantasy_basketball_stats_totals
                self.fantasy_basketball_stats_per_100_poss = fantasy_basketball_stats_per_100_poss
                self.fantasy_basketball_stats_per_36_min = fantasy_basketball_stats_per_36_min
                self.fantasy_basketball_stats_average = fantasy_basketball_stats_average
                self.game_score = game_score
                self.total_stats = total_stats
                self.net_rating = net_rating
                self.quality_of_impact = quality_of_impact
                self.level_of_impact = level_of_impact
                self.win_contribution = win_contribution
                self.value = value
                self.MVP_vote_standing = MVP_vote_standing

def main():
    '''This runs the program.'''

    path = ["2019-2020.xlsx"] #The list of files to be evaluated. Some iterations may evaluate a single file or more than one file
    sheet_name = ['Per Game', 'Totals', 'Advanced', 'Per 100 Poss', 'Per 36 Min', 'MVP Tracker', 'ATL', 'BKN', 'BOS', 'CHA', 'CHI', 'CLE', 'DAL', 'DEN', 'DET','GS', 'HOU', 'IND', 'LAC', 'LAL', 'MEM', 'MIA', 'MIL', 'MIN', 'NO', 'NY', 'OKC', 'ORL', 'PHI', 'PHX', 'POR', 'SA', 'SAC', 'TOR', 'UTAH', 'WSH']
    Excel_Data_Reader(path, sheet_name)
    

if __name__ == '__main__':
    main()



