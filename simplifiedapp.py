from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
from openpyxl import load_workbook, Workbook 

def parse_data(data):
     cols = ['player', "age", "gp", "w", "l","min" ,"pts", "fgm", "fga",
     "fg%", "3pm", "3pa", "3p%", "ftm", "fta", "ft%" ,"oreb", "dreb", "reb",
     "ast", "tov",  "stl", "blk", "pf", "fp", "dd2", "td3"]
  
     stat_col = cols.index(data)
     player_col = 0
  
     wb = load_workbook("nbastats2018.xlsx")
     ws = wb['Sheet1']
     for row in ws.values:
         yield row[player_col].lower(), row[stat_col]

app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def send_sms():
    msg = request.values.get("Body").lower() # convert to lowercase
    player_and_stat = msg.split() #split 
    res = MessagingResponse()
    if len(player_and_stat) == 5: # check input: 2 players + stat
        stat = player_and_stat.pop()
        player1 = " ".join(player_and_stat[:2])
        player2 = " ".join(player_and_stat[2:])
        print( "player 1 ", player1, "player 2 ", player2)
        player_stat_map = dict(parse_data(stat))
        
        player1_stats = player_stat_map.get(player1)
        player2_stats = player_stat_map.get(player2)
        print("player1stats", player1_stats, "player2stats", player2_stats)
        ret = ''
        if player1_stats and player2_stats:
            ret = "In the 2017-2018 regular NBA season, {0}'s total {2} of {3} is higher than {1}'s of {4}"
            if player2_stats > player1_stats:
                ret = "In the 2017-2018 regular NBA season, {1}'s total {2} of {4} is higher than {0}'s of {3}"
                print("ret ", ret)
            ret = ret.format(player1, player2, stat, player1_stats, player2_stats)

        else: #check
            ret = "send 1st + last names of 2 players followed by a stat (GP,W,L,MIN,PTS,FG%,3P%,FT%,REB,AST,STL,BLK). Check for typos!"
    res.message(ret)
    return str(res)

if __name__ == "__main__":
    app.run(debug=True)
