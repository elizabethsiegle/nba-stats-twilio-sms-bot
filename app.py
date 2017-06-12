from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
from openpyxl import load_workbook

app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])

def parseExcel():
	msg = request.form['Body'].lower() #check that it's lowercase
	list_of_players = []
	list_of_stats = []
	stat_dict = {
	"min":"B", "pts":"C","fgm":"D","fga":"E","fg%":"F","3pm":"G",
	"3pa":"H","3p%":"I","ftm":"J","fta":"K","ft%":"L","oreb":"M","dreb":"N",
	"reb":"O","ast":"P","stl":"Q","blk":"R","tov":"S","eff":"T"
	}
	reg_season = True # set boolean, could change depending on user input
	if msg == "play":
		ret = MessagingResponse().message("type \'a\' for regular season or \'b\' for playoffs")
	elif msg == 'a':
		ret = MessagingResponse().message("Type a player or two (first + last name, players separated by a space), then a stat(MIN,PTS,FGM,FGA,FG%,3PM,3PA,3P%,FTM,FTA,FT%,OREB,DREB,REB,AST,STL,BLK,TOV,EFF)")
	elif msg == 'b':
		reg_season = False
		ret = MessagingResponse().message("Type a player or two (first + last name, players separated by a space), then a stat(MIN,PTS,FGM,FGA,FG%,3PM,3PA,3P%,FTM,FTA,FT%,OREB,DREB,REB,AST,STL,BLK,TOV,EFF)")
	else:
		if reg_season:
			wkbk = 'nba_reg_season_16_17.xlsx'
			season_var = "regular season"
		else:
			wkbk = 'nba_playoff_stats_2017.xlsx'
			season_var = "playoffs"
		wb = load_workbook(wkbk) # playoffs or regular season, take your pick!!
		ws = wb[wb.sheetnames[0]]
		player_and_stat = msg.split() #split 
		if len(player_and_stat) == 3: # 1 player
			full_name = player_and_stat[0] + " " + player_and_stat[1]
			stat = player_and_stat[2]
			for row in range(1, ws.max_row+1): #need +1 to get last row!
				for col in "A": #A gets players for texted season
					cell_name="{}{}".format(col, row)
					list_of_players.append(ws[cell_name].value.lower())
				for col in stat_dict[stat]: #pts
					cell_name="{}{}".format(col, row)
					list_of_stats.append(ws[cell_name].value)
			player_stat_map = dict(zip(list_of_players, list_of_stats))
			if full_name in player_stat_map.keys():
				ret = MessagingResponse().message(full_name + " averaged a " + str(stat) + " of: " +str(player_stat_map[full_name]))
			else:
				ret = MessagingResponse().message("send first and last name(s) and stat. to check, or check for typos!")
		elif len(player_and_stat) == 5: #2 players
			player1 = player_and_stat[0] + " " + player_and_stat[1]
			player2 = player_and_stat[2] + " " + player_and_stat[3]
			stat = player_and_stat[4]
			for row in range(1, ws.max_row+1): #need +1 to get last row!
				for col in "A": #A gets players for texted season
					cell_name="{}{}".format(col, row)
					list_of_players.append(ws[cell_name].value.lower())
				for col in stat_dict[stat]: #pts
					cell_name="{}{}".format(col, row)
					list_of_stats.append(ws[cell_name].value)
			player_stat_map = dict(zip(list_of_players, list_of_stats))
			if player1 in player_stat_map.keys() and player2 in player_stat_map.keys():
				if player_stat_map[player1] > player_stat_map[player2]:
					ret = MessagingResponse().message(player1 + " averaged " + str(player_stat_map[player1]) + ", higher than " + player2 + "\'s " + str(player_stat_map[player2]))
				else:
					ret = MessagingResponse().message(player2 + " averaged " + str(player_stat_map[player2]) + ", higher than " + player1 + "\'s " + str(player_stat_map[player1]))
		else: #idk how many players
			ret = MessagingResponse().message("send first and last name(s), or check for typos!")
	return str(ret)

if __name__ == "__main__":
	app.run(debug=True)
