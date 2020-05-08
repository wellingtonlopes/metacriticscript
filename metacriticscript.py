import requests
from openpyxl import Workbook

text = open("gameslist.txt", "r")
scorebook = Workbook()
scoresheet = scorebook.active
scoresheet.title = "Game Pass for PC Metascores"
scoresheet.append(["Game Title", "Metascore"])


def convert_to_games(text):
    games = text.read().splitlines()
    games_in_url = []
    for game in games:
        if game.endswith("A)"):
            game = game[:-5]
        elif game.endswith(")"):
            game = game[:-6]
        game = "%20".join(game.split())
        games_in_url.append(game)
    return games_in_url


games = convert_to_games(text)

urls = ["https://chicken-coop.p.rapidapi.com/games/" + game for game in games]

querystring = { 
                "platform":"pc",
                }

headers = {
    'x-rapidapi-host': "chicken-coop.p.rapidapi.com",
    'x-rapidapi-key': "85c0481f7amsh909e100ef71325fp14744cjsn98dffd5fa4dc"
    }


def write_scores(list_of_urls):
        for url in list_of_urls:
            response = requests.request("GET", url, headers=headers, params=querystring)
            keys = ['title', 'score']
            game_info = response.json().get('result')
            game_score = [game_info.get(key) for key in keys] if game_info != 'No result' else ["No Data", "No Data"]
            scoresheet.append(game_score)
            print(game_score)
        
        scorebook.save("gamescores.xlsx")


write_scores(urls)