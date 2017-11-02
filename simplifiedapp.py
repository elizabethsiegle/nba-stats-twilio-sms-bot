from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
from openpyxl import load_workbook, Workbook 


def parse_data(data):
     cols = ['name', "age", "gp", "w", "l","min" ,"pts", "fgm", "fga",
     "fg%", "3pm", "3pa", "ftm", "fta", "ft%" ,"oreb", "dreb", "reb",
     "ast", "tov",      "stl", "blk", "pf", "dd2", "td3"]
  
     stat_col = cols.index(data)
     player_col = 0
  
     wb = load_workbook("nbastats.xlsx")
     ws = wb['Sheet1']
     for row in ws.values:
         yield row[player_col].lower(), row[stat_col]


app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def send_sms():
    msg = request.form['Body'].lower() # convert to lowercase
    msg = "send 1st + last names of 2 players followed by a stat (GP,W,L,MIN,PTS,FG%,3P%,FT%,REB,AST,STL,BLK). Check for typos!"
    player_and_stat = msg.split() #split 

    if len(player_and_stat) == 5: # check input: 2 players + stat
        stat = player_and_stat.pop()
        player1 = " ".join(player_and_stat[:2])
        player1 = " ".join(player_and_stat[2:])

        player_stat_map = dict(parse_data(stat))
        
        player1_stats = player_stat_map.get(player1)
        player2_stats = player_stat_map.get(player2)
        
        if player1_stats and player2_stats:
            msg = "{0}'s total, higher than {1}'s"
            if player2_stats > player1_stats:
                msg = "{1}'s total, higher than {0}'s"
            msg = msg.format([player1_stats, player2_stats])

        else: #check
            msg = "check both players' names (first and last!)"

    return MessagingResponse().message(msg)

if __name__ == "__main__":
    app.run(debug=True)
