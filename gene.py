# 標準ライブラリが先
from collections import namedtuple
import string

# pipで後から入れたものは後に置く
from tqdm import tqdm
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

# 読み取るエクセルファイルの最初と最後
# YEAR_RANGE = [2001, 2020]
YEAR_RANGE = [2018, 2020]

# 検索ワード
# GENE_LIST = ['A2M', 'CGNL1', 'BAALC', 'RNF112', 'TMTC1', 'GPR37L1', 'SLC4A4', 'CXCL12', 'RGS2', 'EGR1', 'LRRC73']
GENE_LIST = ['NPWT']

# 追加検索ワード
ADDITINAL_WORDS = ['negative', 'pressure', 'wound', 'therapy']
# ADDITINAL_WORDS = ['meningioma']

# 検索範囲
DATE_RANGE = ['2001/01/01', '2020/07/26']



# int連番を作って内包表記で全部strに変換する
YEARS = [str(y) for y in list(range(*YEAR_RANGE))]

def is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
    return False

# 何度も同じルールで書き換えるので関数にしておく
def tokenize(s):
    s = s.translate(str.maketrans('', '', string.punctuation))
    s = s.replace(' ', '')
    s = s.lower()
    return s


# a['hoge']['fuga']['piyo'] とアクセスするとKeyErrorが補足できないので、
# 安全に階層を下るためのヘルパー関数を用意する
def get_recursively(element, keys, default_value=None):
    # 0からキー配列を探索開始
    i = 0
    e = element
    while i < len(keys):
        # 現在のキーを取得
        key = keys[i]
        try:
            # 現在のキーの値を取得
            e = e[key]
        except KeyError:
            # 現在のキーの値がなければその時点で脱出
            return default_value
        # 配列にも使えるように
        except IndexError:
            return default_value
        i += 1

    # 最後まで抜けて来られれば最後に取得した値が、最後のキーに対応する値になる
    return e

# namedtupleは変更がなくメソッドを持たいない純粋に複数のデータを入れるのに便利
# 配列のようにもクラスのようにも扱える
#   Hoge = namedtuple('Hoge', ['fuga', 'piyo'])
#   hoge = Hoge(123, 'foo')
#   print(hoge.fuga, hoge.piyo)
# みたいに使える
JournalInfo = namedtuple('JournalInfo', ['year', 'title', 'iso', 'issn', 'impact_factor', 'raw_title', 'raw_iso'])

# classは大文字ではじめる(HogeFuga など)
# またあまりにも一般的な名前にするとかぶるのでできるだけ自明な名前にする
class JournalListByYear:
    def __init__(self, year):
        self.year = year

        # f-strings という新し目の書き方。変数を埋め込める
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
            rank, title, iso, issn, cites, imp = [cell.value for cell in row[:6]]
            # rankが数字のときだけ有効な行と見做す
            if not isinstance(rank, int):
                continue

            # 行情報をtokenizeしながら、上で用意したJournalInfoとして作成して配列に収める
            self.journals.append(JournalInfo(
                year=year,
                title=tokenize(title) if title else '', # 空白で埋める
                iso=tokenize(iso) if title else '', # 空白で埋める
                issn=issn,
                impact_factor=imp,
                raw_title=title,
                raw_iso=iso,
            ))

        print(f'Loaded {filename} ({len(self.journals)})')

    def compare(self, a, b):
        return a != '' and b != '' and a == b

    # その年のジャーナル一覧から title と iso と issn で検索し、ジャーナル情報とマッチした条件を返す
    def match(self, title, iso, issn):
        for journal in self.journals:
            if self.compare(issn, journal.issn):
                return journal, 'issn'

            if self.compare(iso, journal.iso):
                return journal, 'iso'

            if self.compare(title, journal.title):
                return journal, 'title'

        # 見つからなければここまで落ちてくる
        return None, ''


class JournalLists:
    def __init__(self, years):
        # 年ごとのエクセルファイルを一個ずつ読み取る
        self.lists = [JournalListByYear(year) for year in years]
        print('All sheets loaded')

    # その年のジャーナル一覧から title と iso と issn で検索し、ジャーナル情報とマッチした条件を返す
    def match(self, title, iso, issn):
        # 新しい順に 2019 -> 2018 ... と該当するものを探す
        for l in reversed(self.lists):
            journal_info, condition = l.match(title, iso, issn)

            # 新しい順にループしてるので最新のものがマッチする
            if journal_info:
                return journal_info, condition

        # 全部さらって一個もマッチしなかったらここまで来る
        return None, ''

# xmlのデータは複雑なので、読み取り部分をクラスに閉じ込める。外から見たときに
# 「細かいことは知らんけど get_journal_title() をすればジャーナルのタイトルが取れるぜ！」
# という状態に落とし込むのが構造化プログラミングの基本
class ArticleData:
    def __init__(self, data):
        self.data = data

    def get_title(self):
        return self.data.get('ArticleTitle', '')

    def get_year(self):
        return get_recursively(self.data, ['ArticleDate', 0, 'Year'], '')

    def get_journal_issn(self):
        return get_recursively(self.data, ['Journal', 'ISSN'], '')

    def get_journal_title(self):
        return get_recursively(self.data, ['Journal', 'Title'], '')

    def get_journal_iso(self):
        return get_recursively(self.data, ['Journal', 'ISOAbbreviation'], '')

# 検索結果（遺伝子 実際の検索語句 ヒット数 論文ID配列）を格納する
SearchResult = namedtuple('SearchResult', ['gene', 'needle', 'count', 'ids'])

# 遺伝子名と追加の語句を合体させながら検索して、検索結果をSearchResultの配列で返す
def search_pubmed_by_gene(gene_list, additinal_words):
    results = []
    for gene in gene_list:
        # AND区切りの検索ワードを作る
        # 1. join()関数: ','.join(['A', 'B', 'C']) -> 'A,B,C'
        # 2. * はリストを引数やリストの要素展開する。下の省略表現
        #    'AND'.join([gene, ADDITINAL_WORDS[0], ADDITINAL_WORDS[1]])
        # 3. 検索ワードには慣習的に needle という名前が使用される
        needle = ' AND '.join([gene, *additinal_words])
        handle = Entrez.esearch(db='pubmed', term=needle, mindate=DATE_RANGE[0], maxdate=DATE_RANGE[1], retmax=100000)
        record = Entrez.read(handle)
        # 遺伝子名をキー、ヒット数を値としたdict
        results.append(SearchResult(gene, needle, record['Count'], record['IdList']))
    return results




# 〜〜〜ここまではパーツを用意する処理。ここから下でそれらを使いながら実際の読み取りしていく 〜〜〜




# 検索する
search_results = search_pubmed_by_gene(GENE_LIST, ADDITINAL_WORDS)

# ヒット数を表示
total_count = 0
for v in search_results:
    total_count += int(v.count)
    print(f'[{v.gene}]: {v.count}')

answer = input(f'estimated time {total_count//60}min (1s/itr). OK to start search?(y/n) ')
if len(answer) > 0 and answer[0] == 'n':
    exit(1)

# エクセルファイルを読みはじめる
journal_lists = JournalLists(YEARS)

print(f'Additinal words: {ADDITINAL_WORDS}')


# 出力する行のデータ(論文ごとの)
OutputRow = namedtuple('OutputRow', [
    'article_title',
    'pubmed_id',
    'year',
    'journal_title',
    'journal_iso',
    'journal_issn',
    'journal_year',
    'journal_impact_factor',
    'match_condition',
    'url',
])

# 遺伝子名をキーとして、出力する行の配列とそのインパクトファクターの和のdict
rows_and_if_by_gene = {}

for search_result in search_results:
    print(f'searching {search_result.gene}')
    rows = []
    impact_factor = 0.0

    # id は予約語なので避ける
    for _id in tqdm(search_result.ids):
        handle = Entrez.efetch(db='pubmed', id=_id, retmode='xml')
        xml_data = Entrez.read(handle)

        # xmlのまま扱うとごちゃごちゃにになるので上で作ったクラスを経由する
        data = get_recursively(xml_data, ['PubmedArticle', 0, 'MedlineCitation', 'Article'])
        if not data:
            continue
        article_data = ArticleData(data)

        # エクセルのデータの配列を持ってるクラスの検索関数を呼ぶ
        # 上で関数にまとめたので細かいことを気にせずジャーナルのタイトルとか読み取って引数にできる
        article_journal_title = article_data.get_journal_title()
        article_journal_iso = article_data.get_journal_iso()
        article_journal_issn = article_data.get_journal_issn()

        journal_info, match_condition = journal_lists.match(
            tokenize(article_journal_title),
            tokenize(article_journal_iso),
            article_journal_issn,
        )

        rows.append(OutputRow(
            article_data.get_title(),
            _id,
            article_data.get_year(),
            # ジャーナ一覧の一致があればそのタイトルを使う。なければ論文データについてるジャーナル名で埋めておく
            journal_info.raw_title if journal_info else article_journal_title,
            journal_info.raw_iso if journal_info else article_journal_iso,
            journal_info.issn if journal_info else article_journal_issn,
            journal_info.year if journal_info else '',
            journal_info.impact_factor if journal_info else 0.0,
            match_condition,
            f'https://pubmed.ncbi.nlm.nih.gov/{_id}/',
        ))

        if journal_info and is_float(journal_info.impact_factor):
            impact_factor += float(journal_info.impact_factor)

    rows_and_if_by_gene[search_result.gene] = [rows, impact_factor]


wb = Workbook()
ws = wb.active

# ヘッダ
ws.append([
    'Search word',
    'Hit count',
    'Article title',
    'PubMed ID',
    'Year',
    'Journal title',
    'Journal abbr',
    'Journal issn',
    'Impact factor',
    'Match condition',
    'URL',
])

for search_result in search_results:
    # 配列を変数に展開
    [rows, impact_factor] = rows_and_if_by_gene[search_result.gene]

    # geneごとの見出し行
    ws.append([
        search_result.needle, # 'Search word',
        search_result.count,  # 'Hit count',
        '',  # 'Article title',
        '',  # 'PubMed ID',
        '',  # 'Year',
        '',  # 'Journal title',
        '',  # 'Journal abbr',
        impact_factor,  # 'Impact factor',
        '',  # 'PubMed ID',
        '',  # 'URL',
    ])

    for row in rows:
        ws.append([
            '',  # 'Search word',
            '',  # 'Hit No.',
            row.article_title,  # 'Article title',
            row.pubmed_id,  # 'PubMed ID',
            row.year,  # 'Year',
            row.journal_title,  # 'Journal title',
            row.journal_iso,  # 'Journal abbr',
            row.journal_issn,  # 'Journal issn',
            row.journal_impact_factor,  # 'Impact factor',
            row.match_condition,  # 'Match condition',
            row.url,  # 'URL',
        ])

wb.save(filename=OUTPUT_FILE)
