from mailmerge import MailMerge
from libbgg.apiv1 import BGG
# You can also use version 2 of the api:
from libbgg.apiv2 import BGG as BGG2
import qrcode
from html import unescape

conn = BGG()

user = "Teddsticle"

# V2
conn2 = BGG2()
results = conn2.get_collection(user, own=1, stats=True)
# print(json.dumps(results, indent=4, sort_keys=True))

# Link for website
link_template = "https://boardgamegeek.com/boardgame/{}/"
image_template = 'C:\\\\Users\\\\teddk\\\\Desktop\\\\Other\\\\BGCatalog\\\\Images\\\\{}.png'
game_dicts = []

for game in results["items"]["item"][:5]:



    bggid = game["objectid"]
    print(game['name']['TEXT'], image_template.format(game['name']['TEXT']))

    #use search for extra stats (weight, description, etc.)
    try:
        game_plus = conn2.boardgame(bggid, stats=True)["items"]["item"]
    except KeyError:
        continue
    # print(json.dumps(game_plus, indent=4, sort_keys=True))

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

    f = open("imagetesting.png", "rb")
    img_data = f.read()
    f.close()


    qr = qrcode.QRCode(
        version=1,
        box_size=10,
        border=5)
    qr.add_data(link_template.format(bggid))
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    # img.save('test_qrcode.png')

    data = {
        'Image': image_template.format(game['name']['TEXT']),
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
        'Description': unescape(game_plus["description"]["TEXT"])
    }
    game_dicts.append(data)
# print(game_dicts)

# TODO:
def  qr(bgid):
    filename = "TODO"
    return filename

# __slots__ = [
#     'designers', 'artists', 'playingtime', 'thumbnail',
#     'image', 'description', 'minplayers', 'maxplayers',
#     'categories', 'mechanics', 'families', 'publishers',
#     'website', 'year', 'names', 'bgid'
# ]

with MailMerge('test_template_online.docx') as document:
    print(document.get_merge_fields())

    document.merge_templates(game_dicts, 'page_break')
    document.write('test-output-2.docx')