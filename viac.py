import datetime
import json
import sys
import os
import re
import urllib
import urllib.request
import yaml

MAX_ITEMS = 1_000_000
DEBUG_WRITE = False
DEBUG_READ = False
DEBUG_FILENAME = "temp.txt"

PARAMS = {'proxies': {}}


def clean_text(text: str) -> str:
    return ''.join(line.strip()
                   for line in text.split('\n')).replace('href =', 'href=')


def split_text(text: str, start: str, end: str) -> tuple:
    id_s = text.find(start)
    id_e = text.find(end)
    if id_s >= 0 and id_e >= 0 and id_e > id_s:
        return text[(id_s + len(start)):id_e], text[(id_e + len(end)):]
    return None, None


def split_text_all(text: str, start: str, end: str) -> list:
    parts = []

    while text is not None:
        part, text = split_text(text, start, end)
        if part is not None:
            parts.append(part)

    return parts


def load_list(code: str) -> list:
    root = 'https://arxiv.org'
    url = f'{root}/list/{code}/pastweek'
    proxy_handler = urllib.request.ProxyHandler(PARAMS['proxies'])
    opener = urllib.request.build_opener(proxy_handler)
    text = open(DEBUG_FILENAME, 'r', encoding='utf-8').read(
    ) if DEBUG_READ else opener.open(url).read().decode('utf8')
    pattern = 'Total of (\d+) entries'
    matches = re.findall(pattern, text)
    num_items = int(matches[0]) if len(set(matches)) == 1 else -1
    if num_items < 0:
        print('number of items not positive')

    url = f'{root}/list/{code}/pastweek?show={num_items if num_items >= 0 else MAX_ITEMS}'
    print(url)
    text = open(DEBUG_FILENAME, 'r', encoding='utf-8').read(
    ) if DEBUG_READ else opener.open(url).read().decode('utf8')

    if DEBUG_WRITE:
        open(DEBUG_FILENAME, 'w', encoding='utf-8').write(text)

    text = clean_text(text)
    links = split_text_all(text, '<dt>', '</dt>')
    descs = split_text_all(text, '<dd>', '</dd>')

    patterns = {
        'link': '<a href="([/\w\d\.]+)" title="Abstract"',
        'pdf': '<a href="([/\w\d\.]+)" title="Download PDF"',
    }

    summary = []
    for text_link, text_desc in zip(links, descs):
        item = {
            key: re.findall(pattern, text_link)
            for key, pattern in patterns.items()
        }
        item['title'], _ = split_text(text_desc, '</span>', '</div>')

        for key_link in ('link', 'pdf'):
            item[
                key_link] = f'=HYPERLINK("{root}{item[key_link][0]}", "goto")' if bool(
                    item[key_link]) else ''

        title = item['title']
        for c in ',.-/":;_+':
            title = title.replace(c, ' ')
        parts = set(part.lower() for part in title.split(' ') if bool(part))
        item['keywords'] = parts

        summary.append(item)

    print(f'added {len(summary)} items')

    return summary


def highlight(data, mask):
    format = data.copy()
    format.loc[:, :] = ''
    format.loc[mask, :] = 'background-color: lime'
    return format


def get_name() -> str:
    now = datetime.datetime.now()
    return 'viac_output_{:04}_{:02}_{:02}_{:02}.xlsx'.format(
        now.year, now.month, now.day, now.hour)


def export():
    import pandas
    configs = json.loads(open('config.json').read())
    with pandas.ExcelWriter(get_name(), engine='xlsxwriter') as writer:
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
            writer.sheets[config['name']].set_column(0, 0, 128)
            writer.sheets[config['name']].set_column(1, len(to_write) - 1, 4)


if __name__ == '__main__':
    path_proxy = os.path.join(os.path.dirname(sys.argv[0]), '.proxy.yaml')
    if os.path.isfile(path_proxy):
        PARAMS['proxies'] = yaml.safe_load(open(path_proxy).read())
    export()
