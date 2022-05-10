import qrcode
from mailmerge import MailMerge

def get_qr(bggid: int):
    """
    Generate a QR code image containing a game's BGG link.

    :param bggid: (int) BGG game ID
    :return: (PilImage) QR image containing BGG game link
    """
    # BGG game page template
    link_template = "https://boardgamegeek.com/boardgame/{}/"
    qr = qrcode.QRCode(
        version=1,
        box_size=10,
        border=5)
    qr.add_data(link_template.format(bggid))
    qr.make(fit=True)
    qr_img = qr.make_image(fill='black', back_color='white')
    return qr_img


def build_word_file(game_dicts, template_file, output_file):
    """
    Build Microsoft Word Docx Catalog from game data dict.

    :param game_dicts: (dict) game data dict
    :param template_file: (string) full filename of mail merge template Word file
    :param output_file: full filename of output Word file
    :return:
    """
    with MailMerge(template_file) as document:
        print(document.get_merge_fields())
        document.merge_templates(game_dicts, 'page_break')
        document.write('output_file')
    return True
