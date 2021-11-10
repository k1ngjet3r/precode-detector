import re

def matcher_slice(keywords, cell_data):
    sen = str(cell_data).lower()
    # sen = str(cell_data).replace('"', '').lower()
    for key in keywords:
        if re.search(key, sen):
            return True
    return False


def matcher_split(keywords, cell_data):
    clean_sentance = re.sub(r'[^\w]', ' ', str(cell_data).lower())
    word_list = clean_sentance.split()
    for key in keywords:
        if key in word_list:
            return True
    return False
