from boardgamegeek import BGGClient
import time

bgg = BGGClient()

collection = bgg.collection(user_name='Teddsticle', version=True, prev_owned=False)

for g in collection:
    if g.version:
        print(g.name, " - ", g.version.name)
    else:
        print(g.name)
