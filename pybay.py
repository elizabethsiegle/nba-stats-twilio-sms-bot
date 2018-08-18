from twilio.twiml.messaging_response import MessagingResponse
from flask import Flask, request
from openpyxl import load_workbook, Workbook 

def parse_data(stat):
    cols = ['player', "gp", "w", "l","min" ,"pts", "fgm", "fga", "fg%", "3pm", "3pa", "3p%", "ftm", "fta", "ft%" ,"oreb", "dreb", "reb",
    "ast", "tov",  "stl", "blk", "pf", "fp", "dd2", "td3"]
    stat_col = cols.index(stat)
    player_col = 0

    wb = load_workbook("nbastats2018.xlsx")
    ws = wb["Sheet1"]

    for row in ws.values:
        yield row[player_col].value, row[stat_col]

app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def send_sms():
    #incoming message
    msg = request.values.get("Body").lower()
    player_and_stat = msg.split()

    #outgoing messages
    res = MessagingResponse()
    ret = "send 1st + last names of 2 players followed by a stat (GP,W,L,MIN,PTS,FG%,3P%,FT%,REB,AST,STL,BLK). Check for typos!"
    #check input: 2 players + stat
    if len(player_and_stat) ==5:
        #parse input
        stat = player_and_stat.pop()
        player1 = " ".join(player_and_stat[:2])
        player2 = " ".join(player_and_stat[2:])
       
        #dictionaries
        player_and_stat_dict = dict(parse_data(stat))
        player1_stats = player_and_stat_dict.get(player1)
        player2_stats = player_and_stat_dict.get(player2)
    
        if player1_stats and player2_stats:
            ret = "In the 2017-2018 regular NBA season, {0}'s avg. {2}/game of {3} is higher than {1}'s per game avg. of {4}"
            if player2_stats > player1_stats:
                ret = "In the 2017-2018 regular NBA season, {1}'s avg. {2}/game of {4} is higher than {0}'s per game avg. of {3}"
            ret = ret.format(player1, player2, stat, player1_stats, player2_stats)
    res.message(ret)
    return str(res)

if __name__ == "__main__":
    app.run(debug=True)

