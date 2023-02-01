import json
import re
import urllib
import urllib.request

MAX_ITEMS = 1_000_000


def load_list(code: str) -> list:
    url = f'https://arxiv.org/list/{code}/pastweek'
    text = urllib.request.urlopen(url).read().decode('utf8')
    pattern = 'total of (\d+) entries'
    matches = re.findall(pattern, text)
    num_items = int(matches[0]) if len(set(matches)) == 1 else -1
    if num_items < 0:
        print('number of items not item')

    url = f'https://arxiv.org/list/{code}/pastweek?show={num_items if num_items >= 0 else MAX_ITEMS}'
    text = urllib.request.urlopen(url).read().decode('utf8')

    # open(code, 'w').write(text)

    # text = open(code, 'r').read()

    patterns = {
        'title':
        '<span class="descriptor">Title:</span>(.*)</div><div class="list-authors">',
        'link': '<a href="([/\w\d\.]+)" title="Abstract">',
        'pdf': '<a href="([/\w\d\.]+)" title="Download PDF">',
    }

    text = text.replace('\n', '')
    start = '<span class="list-identifier">'
    end = '<span class="descriptor">Authors:</span>'
    items = []
    for _ in range(MAX_ITEMS):
        i = text.find(start)
        if i < 0:
            break
        text = text[(i + len(start)):]
        i = text.find(end)
        if i < 0:
            break
        part = text[:i]
        text = text[(i + len(end)):]

        items.append(part)

    summary = []
    for item in items:
        item = {
            key: re.findall(pattern, item)
            for key, pattern in patterns.items()
        }

        for key_link in ('link', 'pdf'):
            item[
                key_link] = f'=HYPERLINK("https://arxiv.org{item[key_link][0]}", "goto")' if bool(
                    item[key_link]) else ''
        item['title'] = item['title'][0]

        title = item['title']
        for c in ',.-/":;_+':
            title = title.replace(c, ' ')
        parts = set(part.lower() for part in title.split(' ') if bool(part))
        item['keywords'] = parts

        summary.append(item)

    return summary


def highlight(data, mask):
    format = data.copy()
    format.loc[:, :] = ''
    format.loc[mask, :] = 'background-color: lime'
    return format


def export():
    import pandas
    configs = json.loads(open('config.json').read())
    with pandas.ExcelWriter('export.xlsx', engine='xlsxwriter') as writer:
        for config in configs:
            summary = load_list(config['id'])
            is_interesting = [
                any(kw_config in kw_title for kw_config in config['keywords']
                    for kw_title in item['keywords']) for item in summary
            ]

            to_write = {
                key: [item[key] for item in summary]
                for key in ('title', 'link', 'pdf')
            }

            pandas.DataFrame(to_write).style.apply(
                lambda x: highlight(x, mask=is_interesting),
                axis=None).to_excel(writer,
                                    sheet_name=config['name'],
                                    index=False)
            writer.sheets[config['name']].set_column(0, 0, 96)
            writer.sheets[config['name']].set_column(1, len(to_write) - 1, 4)


export()
