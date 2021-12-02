import multiprocessing
import re
from pathlib import Path

import openpyxl.cell.read_only
import pypandoc


def process(data):
    source, target, content = data
    if target.exists():
        return
    if not source.exists():
        with open(source, encoding='utf-8', mode='w') as file:
            file.write(content)
    try:
        pypandoc.convert_file(str(source), format='html', to='docx',
                              outputfile=str(target))
    except:
        print(f'Failed: {target.name}')


def main():
    book = openpyxl.load_workbook('chainnews-archive-508-966.xlsx', read_only=True)
    sheet = book.get_sheet_by_name(book.sheetnames[0])
    headers = [cell.value for cell in sheet['A1:Z1'][0]]
    title_index = headers.index('title')
    url_index = headers.index('article-page-href')
    content_index = headers.index('html')

    target_dir = Path(__file__).parent / 'documents'
    target_dir.mkdir(exist_ok=True, parents=True)

    directory = Path(__file__).parent / 'html'
    directory.mkdir(exist_ok=True, parents=True)

    params = []
    for row in sheet['A2:Z10000']:
        url = row[url_index].value
        title = row[title_index].value
        match = re.match(r'https://chainnews-archive.org/posts/(\d+)/', url)
        assert match
        slug = match.group(1)
        content = row[content_index].value
        title_in_filename = re.sub(r'[\\/:*?"<>|]', '-', title)[:30]
        file_name = f'{title_in_filename} {slug}.docx'

        source = directory / f'{slug}.html'
        target = target_dir / file_name

        params.append((source, target, content))

    with multiprocessing.Pool(20) as p:
        print(p.map(process, params))


if __name__ == '__main__':
    main()
