from mailmerge import MailMerge
import json
from libbgg.apiv1 import BGG
# You can also use version 2 of the api:
from libbgg.apiv2 import BGG as BGG2

conn = BGG()

user = "Teddsticle"

# V2
conn2 = BGG2()
results = conn2.get_collection(user, own=1)
# print(json.dumps(results, indent=4, sort_keys=True))

game_dicts = []
for game in results["items"]["item"]:
    data = {
        'Image': 'C:\\\\Users\\\\teddk\\\\Desktop\\\\Other\\\\BGCatalog\\\\Images\\\\Keyforge.png',
        'Name': game['name']['TEXT'],
        'Year': game['yearpublished']['TEXT']
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