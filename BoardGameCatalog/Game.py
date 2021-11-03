from mailmerge import MailMerge
import json
from libbgg.apiv1 import BGG
# You can also use version 2 of the api:
from libbgg.apiv2 import BGG as BGG2
import qrcode

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

for game in results["items"]["item"]:



    bggid = game["objectid"]
    print(game['name']['TEXT'])

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
    }
    game_dicts.append(data)
print(game_dicts)

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

with MailMerge('test_template.docx') as document:
    print(document.get_merge_fields())

    cust_1 = {
        'Image': 'C:\\\\Users\\\\teddk\\\\Desktop\\\\Other\\\\BGCatalog\\\\Images\\\\Keyforge.png',
        'Year': '1234',
        'Name': 'kEyFoRgE'
    }

    cust_2 = {
        'Image': 'C:\\\\Users\\\\teddk\\\\Desktop\\\\Other\\\\BGCatalog\\\\Images\\\\Cryptid.png',
        'Year': '4321',
        'Name': 'CrYpTiD'
    }

    document.merge_templates(game_dicts, 'page_break')
    document.write('test-output-2.docx')

# replacements = {
#     "\"^p": "\"",
#     "^p\"": "\"",
#     "1234": "nibba"
#     }
#
# def paragraph_replace_text(paragraph, key, replacement):
#     """Replace first occurence of `key` in `paragraph` with `replacement`.
#
#     Where a newline ("\n") appears in the replacement text, insert a new
#     paragraph such that the newline is interpreted as a paragraph break.
#     """
#     new_text = paragraph.text.replace(key, replacement)
#     lines = new_text.split("\n")
#     for line in lines[:-1]:
#         paragraph.insert_paragraph_before(text=line)
#     paragraph.text = lines[-1]
#
# doc = Doc('test-output-2.docx')
#
# for key, replacement in replacements.items():
#     for p in list(doc.paragraphs):
#         if key in p.text:
#             paragraph_replace_text(p, key, replacement)
#
# # --- save changed document ---
# doc.save('./replace-test-1.docx')