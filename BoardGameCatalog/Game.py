from libbgg.apiv2 import BGG as BGG2
from html import unescape
import util

def get_collection_data(user):
    # Connect to BGG API v2
    conn2 = BGG2()

    # pull dict of collection info
    results = conn2.get_collection(user, own=1, stats=True)

    # Local image directory template
    image_template = 'C:\\\\Users\\\\teddk\\\\Desktop\\\\Other\\\\BGCatalog\\\\Images\\\\{}.png'
    game_dicts = []

    # For each game in collection, get data
    for game in results["items"]["item"][:7]:

        bggid = game["objectid"]
        print(game['name']['TEXT'])

        # use search for extra stats (weight, description, etc.)
        # Game expansions don't have this feature, so they will be skipped
        try:
            game_plus = conn2.boardgame(bggid, stats=True)["items"]["item"]
        except KeyError:
            continue

        category, mechanisms, designer, artist = [], [], [], []
        for stat in game_plus["link"]:
            if stat["type"] == "boardgamecategory":
                category.append(stat["value"])
            elif stat["type"] == "boardgamemechanic":
                mechanisms.append(stat["value"])
            elif stat["type"] == "boardgamedesigner":
                designer.append(stat["value"])
            elif stat["type"] == "boardgameartist":
                artist.append(stat["value"])

        data = {
            'Image': game['thumbnail']['TEXT'],
            'Name': game['name']['TEXT'],
            'Year': game['yearpublished']['TEXT'],
            'PlayerCount': game["stats"]["minplayers"] if (game["stats"]["minplayers"] == game["stats"]["maxplayers"])
                                                       else game["stats"]["minplayers"]+"-"+game["stats"]["maxplayers"],
            'PlayTime': game["stats"]["minplaytime"] if (game["stats"]["minplaytime"] == game["stats"]["playingtime"])
                                                       else game["stats"]["minplaytime"]+"-"+game["stats"]["playingtime"],
            'Weight': game_plus["statistics"]["ratings"]['averageweight']["value"][:4],
            'Category': ", ".join(category),
            'Mechanisms': ", ".join(mechanisms),
            'Designer': ", ".join(designer),
            'Artist': ", ".join(artist),
            'Description': unescape(game_plus["description"]["TEXT"]),
            'QR': util.get_qr(bggid)
        }
        game_dicts.append(data)
    return game_dicts
    # print(game_dicts)

    # __slots__ = [
    #     'designers', 'artists', 'playingtime', 'thumbnail',
    #     'image', 'description', 'minplayers', 'maxplayers',
    #     'categories', 'mechanics', 'families', 'publishers',
    #     'website', 'year', 'names', 'bgid'
    # ]


if __name__ == "__main__":
    user = 'Teddsticle'
    game_dicts = get_collection_data(user)

    template_file = 'test_template_online.docx'
    output_file = 'test-output-2.docx'
    util.build_word_file(game_dicts,
                         template_file,
                         output_file)
