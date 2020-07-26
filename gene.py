# 標準ライブラリが先
from collections import namedtuple
import string

# pipで後から入れたものは後に置く
import openpyxl
from openpyxl import Workbook
from Bio import Entrez

# read は名前が衝突するので読み込まない。Entrez.read() で使う
# コピペしてきたらならコピペ元が悪い
# from Bio.Entrez import efetch, read

Entrez.email = 'endaaman@eis.hokudai.ac.jp'


# 本当は argparse を使って python gene.py --file hoge.xlsx とかでファイルをしているするけど
# ごちゃごちゃになるので頭に定数として定義しておく
# ここをいじるだけで全体の挙動を調節できるようにする
YEAR_RANGE = [2001, 2019]
GENE_LIST = ['A2M', 'CGNL1', 'BAALC', 'RNF112', 'TMTC1', 'GPR37L1', 'SLC4A4', 'CXCL12', 'RGS2', 'EGR1', 'LRRC73']
ADDITINAL_WORDS = ['meningioma']
DATE_RANGE = ['2001/01/01', '2020/07/26']


YEARS = [str(y) for y in list(range(*YEAR_RANGE))]


# 何度も同じルールで書き換えるので関数にしておく
def tokenize(s):
    s = s.translate(str.maketrans('', '', string.punctuation))
    s = s.replace(' ', '')
    s = s.lower()
    return s

# namedtupleは変更がなくメソッドを持たいない純粋に複数のデータを入れるのに便利
# 配列のようにもクラスのようにも扱える
#   r = SearchResult('PRB2', 10, ['123', '456'])
#   print(r.count)
# みたいに使える
JounalInfo = namedtuple('JounalInfo', ['year', 'title', 'iso', 'issn', 'impact_factor', 'raw_title', 'raw_iso'])
SearchResult = namedtuple('SearchResult', ['gene', 'needle', 'count', 'ids'])

# classは大文字ではじめる(HogeFuga など)
# またあまりにも一般的な名前にするとかぶるのでできるだけ自明な名前にする
class JournalListByYear:
    def __init__(self, year):
        self.year = year

        # f-strings という新し目の書き方
        filename = f'data/{year}.xlsx'
        self.wb = openpyxl.load_workbook(filename)

        # 目的の―という意味のプリフィクスには target_ が普通
        # 関数ないで使い捨ての変数なら tmp_ も可(これを嫌がる人もいる)
        target_sheet = self.wb[str(year)]

        # シートの構成をコメントでメモっておく
        # 1 B: title
        # 2 C: iso(abbr title)
        # 3 D: issn
        # 4 E: cites
        # 5 F: impact factor

        self.journals = []

        # 行ごとに読み込む
        for row in target_sheet.rows:
            # リスト内包表記で6番目までのcellのvalueの配列にして、分割代入で変数に展開
            values = [cell.value for cell in row[:6]]
            rank, title, iso, issn, cites, imp = values
            # rankが数字のときだけ有効な行と見做す
            if not isinstance(rank, int):
                continue

            # 行情報をtokenizeしながら保存
            self.journals.append(JounalInfo(
                year=year,
                title=tokenize(title) if title else '', # 空白で埋める
                iso=tokenize(iso) if title else '', # 空白で埋める
                issn=issn,
                impact_factor=imp,
                raw_title=title,
                raw_iso=iso,
            ))

        print(f'Loaded {filename}.')

    # その年のジャーナル一覧から title と iso と issn で検索し、ジャーナル情報とマッチした条件を返す
    def match(self, title, iso, issn):
        for journal in self.journals:
            if issn == journal.issn:
                return journal, 'issn'

            if iso == journal.iso:
                return journal, 'iso'

            if title == journal.title:
                return journal, 'title'
        return None, ''


class JournalLists:
    def __init__(self, years):
        # dict内包表記
        # 年ごとのエクセルファイルを一個ずつ読み取る
        self.lists = [JournalListByYear(year) for year in YEARS]
        print('All sheets loaded')

    # その年のジャーナル一覧から title と iso と issn で検索し、ジャーナル情報とマッチした条件を返す
    def match(self, title, iso, issn):
        # 新しい順に探す
        for l in reversed(self.lists):
            journal_info, condition = l.match(title, iso, issn)
            if journal_info:
                return journal_info, condition
        return None, ''


class ArticleData:
    def __init__(self, data):
        self.data = data

    def get_title(self):
        return self.data.get('ArticleTitle', '')

    def get_year(self):
        # read year
        year = None
        try:
            year = self.data['Journal']['JournalIssue']['PubDate']['MedlineDate']
        except:
            try:
                year = self.data['Journal']['JournalIssue']['PubDate']['Year']
            except:
                return ''
        return str(year)[:4]

    def get_journal_issn(self):
        try:
            issn = article_data['Journal']['ISSN']
        except:
            return ''
        return str(issn)

    def get_journal_title(self):
        journal_title = None
        try:
            al_title_data = article_data['Journal']['Title']
        except:
            return ''
        return str(journal_title)

    def get_journal_iso(self):
        iso = None
        try:
            iso = article_data['Journal']['ISOAbbreviation']
        except:
            return ''
        return str(iso_abbr_data)


def search_pubmed_by_gene(gene_list, additinal_words):
    results = []
    for gene in GENE_LIST:
        # AND区切りの検索ワードを作る
        # 1. join()関数: ','.join(['A', 'B', 'C') -> 'A,B,C'
        # 2. * はリストを引数やリストの要素展開する。下の省略表現
        #    'AND'.join([gene, ADDITINAL_WORDS[0], ADDITINAL_WORDS[1]])
        # 3. 検索ワードには慣習的に needle という名前が使用される
        needle = ' AND '.join([gene, *ADDITINAL_WORDS])
        handle = Entrez.esearch(db='pubmed', term=needle, mindate=DATE_RANGE[0], maxdate=DATE_RANGE[1], retmax=100000)
        record = Entrez.read(handle)
        # 遺伝子名をキー、ヒット数を値としたdict
        results.append(SearchResult(gene, needle, record['Count'], record['IdList']))
    return results


# エクセルファイルを読みはじめる
journal_lists = JournalLists(YEARS)

print(f'Additinal words: {ADDITINAL_WORDS}')
search_results = search_pubmed_by_gene(GENE_LIST, ADDITINAL_WORDS)
# ヒット数を表示
for v in search_results:
    print(f'[{v.gene}]: {v.count}')



OutputRow = namedtuple('OutputRow', [
    'article_title',
    'year',
    'journal_title',
    'journal_iso',
    'journal_year',
    'journal_impact_factor',
    'pubmed_id',
    'url',
])



rows_and_if_by_gene = {}
for search_result in search_results:
    print(f'searching {search_result.gene}')
    rows = []
    impact_factor = 0.0

    # id は予約語なので避ける
    for _id in search_result.ids:
        url = f'https://pubmed.ncbi.nlm.nih.gov/{_id}/'

        handle = Entrez.efetch(db='pubmed', id=_id, retmode='xml')
        xml_data = Entrez.read(handle)
        article_data = ArticleData(xml_data['PubmedArticle'][0]['MedlineCitation']['Article'])

        journal_info, match_confition = journal_lists.match(
            tokenize(article_data.get_journal_title()),
            tokenize(article_data.get_journal_iso()),
            article_data.get_journal_issn(),
        )
        impact_factor += float(journal_info.impact_factor)

        rows.append(OutputRow(
            article_data.get_title(),
            article_data.get_year(),
            journal_info.raw_title,
            journal_info.raw_iso,
            journal_info.year,
            journal_info.impact_factor,
            _id,
            url,
        ))

    rows_and_if_by_gene[search_result.gene] = [rows, impact_factor]


wb = Workbook()
ws = wb.active

# ヘッダ
ws.append([
    'Search word',
    'Hit No.',
    'Article title',
    'Year',
    'Full journal title',
    'J. abbrev.',
    'Impact factor',
    'PubMed ID',
    'URL',
])


for search_result in search_results:
    # 配列を変数に展開
    [rows, impact_factor] = rows_and_if_by_gene[search_result.gene]

    # geneごとの見出し行
    ws.append([
        search_result.needle,  # 'Search word',
        search_result.count,  # 'Hit No.',
        '',  # 'Article title',
        '',  # 'Year',
        '',  # 'Full journal title',
        '',  # 'J. abbrev.',
        impact_factor,  # 'Impact factor',
        '',  # 'PubMed ID',
        '',  # 'URL',
    ])

    for row in rows:
        ws.append([
            '',  # 'Search word',
            '',  # 'Hit No.',
            row.article_title,  # 'Article title',
            row.journal_year,  # 'Year',
            row.journal_title,  # 'Full journal title',
            row.journal_iso,  # 'J. abbrev.',
            row.journal_impact_factor,  # 'Impact factor',
            row.pubmed_id,  # 'PubMed ID',
            row.url,  # 'URL',
        ])

wb.save(filename='output.xlsx')

