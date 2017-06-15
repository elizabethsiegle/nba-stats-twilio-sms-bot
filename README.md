Compare NBA stats with Openpyxl and Twilio!
<img src = "https://user-images.githubusercontent.com/8932430/27196835-0f338ff0-51da-11e7-8faf-93c606b65895.png">

<a href = "https://github.com/elizabethsiegle/twilio-sms-bot-compare-nba-stats/blob/master/app.py">app.py</a>: Text a Twilio number "play," then text "a" or "b" for regular season 2016-1017 versus playoffs, and then text either one or two players separated by a space and a statistic (based on the text you got.)

<a href = "https://github.com/elizabethsiegle/twilio-sms-bot-compare-nba-stats/blob/master/simplifiedapp.py">simplifiedapp.py</a>: Text a Twilio number two NBA players (first and last names) separated by a space and a statistic to compare.

We all know Excel sheets hold useful information, but it can be hard to read the data we want, and especially hard to compare two specific datasets within all that data. Bots, on the other hand, can do that hard work of finding the information for us. 

We will read some NBA statistics from Excel sheets in Python using the Openpyxl library. How will we know which statistics to look for and return? Incoming SMS is two players and a type of NBA statistic, and then the outgoing SMS looks up the statistics of the corresponding players.

<h3>Data</h3>
The data we use is about specific NBA players from <a href = "stats.nba.com/players"target="_blank">this past season</a>.
There, you can specify which season, season type (playoffs versus regular season), data type (game average versus total for the season), which dates, and more. Fun, right? 

No Microsoft Excel? No problem! You can copy and paste the data directly into Google Sheets and export it as a .xlsx file. Here is an excerpt from the the Excel sheet (lots more data not shown to the right and below.) 

<img src = "https://user-images.githubusercontent.com/8932430/27196938-6cbfa208-51da-11e7-9469-98d30a62ac92.png">
Some statistics above include age, games played, wins, losses, minutes, points, field goal percentage, three-point shot percentage, and more.

<h3>Setup your Developer Environment</h3>
Make sure your Python and Flask development environment is set up, like <a href = "https://www.twilio.com/docs/guides/how-to-set-up-your-python-and-flask-development-environment#create-a-simple-flask-application">this</a>. If you don’t have a Twilio number to send and receive SMS messages, let’s do that <a href = "https://www.twilio.com/console">here</a>.


Once your environment is up and running, run the following command in the directory your python file will live in. 
<img src = "https://user-images.githubusercontent.com/8932430/27196974-8e4d2148-51da-11e7-89a5-7cea43505961.png"/>


<h3>Building the Flask app</h3>
Make a file called app.py, and import these libraries at the top.
<img src = "https://user-images.githubusercontent.com/8932430/27197002-ae6e5dde-51da-11e7-8ba4-7961b2475760.png"/>

And then make our Flask object:
https://user-images.githubusercontent.com/8932430/27197789-9ad16264-51dd-11e7-8fec-2bf0295342a1.png

Don't forget to run <a href = "ngrok.com" target="_blank">Ngrok</a> http 5000 in terminal! In your terminal in the same directory, run ngrok http 5000.
<img src ="https://user-images.githubusercontent.com/8932430/27196971-8a1114b8-51da-11e7-821a-b61939f8b597.png"/>

Now onto some fun stuff.

<h3>Parse the Data with Openpyxl</h3>
<a href = "https://openpyxl.readthedocs.io/en/default/">Openpyxl</a> is an open source Python library that reads and writes Microsoft Excel 2010 files.
 
The higher-order function below takes in our entire Excel file of NBA data returns a dictionary of the data in our Excel file. Players are the keys, and the specific statistic data per each player as the values. 

<img src ="https://user-images.githubusercontent.com/8932430/27197009-b18fce58-51da-11e7-8070-be60cabb2644.png">

The data structures we use are two separate lists of players and their corresponding statistics we want to search (ie. just games played, wins, losses, minutes, points, field goal percentage, etc.) We then map those statistics to different columns of the Excel sheet, represented by letters, in stat_dict.
 
Then, we need to load the Excel file full of data with load_workbook and create a worksheet. More complex apps or data may have different sheets (with NBA data, one could be Regular Season while another could be for the Playoffs.) Since we only have one worksheet, we just want the one at index zero. Then, beginning with the for loop, we loop through each item in our Excel spreadsheet. 
 
Finally, we loop through our Excel spreadsheet. The “A” column is for players, so each player in the column (and thus the sheet) are added to our list. Then, we loop through the column which is the value of our dictionary of statistics and columns and add each to the separate statistics list.
 
Say you want to read from columns. Each column is represented by a letter (that’s why we made the dictionary above, but the dictionary values match the columns in the Excel sheets.) To search multiple columns and just the front row, you could search something like this:  
<img src = "https://user-images.githubusercontent.com/8932430/27197012-b43cf73e-51da-11e7-838c-77eb6d819cd1.png">

With our data, this returns the name of the player in the first row (Russell Westbrook) and the statistics from columns D, E, and F (wins, losses, and minutes played.)
 
If you want to access an individual cell, the following code are two different ways to return whatever is in the B column in the second row. (Without using value, you would just get “<Cell u'Sheet1'.B2>”.)
<img src ="https://user-images.githubusercontent.com/8932430/27197029-c494b70c-51da-11e7-8c9a-15696598675c.png">

(another way of writing this same line of code is: <em>b2v2 = ws.cell(row=2, column=2)</em>)

What does our code do after searching through the cells and rows creating our two distinct lists? These two lists are then zipped together into one dictionary with players as keys and the corresponding statistic numbers as values. This dictionary will be returned in our SendSms() function so we can check if the inputSMS message is in the dictionary (and thus in the Excel sheet.)

Now, onto the core of the code! 

<img src = "https://user-images.githubusercontent.com/8932430/27197031-c7165166-51da-11e7-8bcd-30478e7fdaaa.png">

Let’s break this down. First, we get our input SMS and convert it to lowercase so it’s easier to check. Then, we break it up by whitespace and add each piece to a string array. If that array has a length of five (which is what we expect, because input should be two players (first and last names) and a statistic), then we assign variables to the two players and the statistic.
 
Next, we call our higher-order function parseDataIntoDict, and use the dictionary it returns to check that the variables are in it. If they are, we check if the data of one player is greater than the other. Depending on that, a different message is returned. If one or both of the players are not in the dictionary, we return an error message. 
 
Lastly, we run our Flask app!

<img src ="https://user-images.githubusercontent.com/8932430/27197035-cb3202d6-51da-11e7-8fd2-73d8467cf2c6.png">

Wow! You just used Openpyxl to read an Excel spreadsheet. Isn’t it a very handy tool? Who knew reading and writing an Excel spreadsheet could be done this way?
 
So what’s next? Think of the possibilities! You can use Openpyxl for financial data, for baseball data, etc. 

Questions? Comments? Tweet at me <a href = "twitter.com/lizziepika" target="_blank">@lizziepika</a>.

