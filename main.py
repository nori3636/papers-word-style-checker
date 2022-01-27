import win32com.client
import numpy
import re


hiragana = numpy.array(['挙げる', '当たり', '扱う', '新たな', '表す', '現れる',
                        '合わせて', '併せて', '言う', '幾らか', '至る', '一層', '一旦', '受ける', '得る',
                        '起きる', '及び', '極めて', '事とする', '様々な', '様だ', '更に', '従って', '全て',
                        '出来る', 'する時', '共に', '伴い', '後に', '半ば', '初め', '巡る', '稀に', '分かる',
                        '僅かに', '為'])
bad_word = numpy.array(['とても', 'すごく', 'だいたい', 'だから', 'でも,',
                        'だけど', 'けれど', 'けど', 'どうやっても', 'どうしても', 'と思', 'かもしれない',
                        'と感じる', 'おもしろい', 'を知りたい', 'の意味がわかりません', 'いっぱい',
                        'した方がいい', '無駄', '嫌い', 'なかった', '間違いだ', '事実はない', 'みんなが',
                        '教科書に', 'は読まなかった', 'じゃないか', 'かもしれ', '興味深い', '知りたい',
                        'ない方が', 'しないで', 'なくて', 'あって', 'わかっていて', 'かなり', '値段が高い',
                        '難しい', '大事なこと', 'だんだん', '良くな', 'ですから', 'がわかる', '見る', '見た',
                        '見ていく', 'を含む', 'であるとしている', '同様の', 'あるが', '可能性を示唆'
                        ])
avoid_word = numpy.array(['考えられる', 'ような', '今回の研究', 'この結果は',
                          'ことができる'])

# 以下は研究室の方針によって書き換える
lab_word = numpy.array(['\,', '\.', '１', '２', '３', '４', '５', '６', '７', '８',
                        '９', '０', 'ｍｍ', '超伝導'])


def colorcode_to_int(colorcode):
    hex = colorcode[1:7]
    r = int(hex[0:2], 16)
    g = int(hex[2:4], 16)
    b = int(hex[4:6], 16)
    return r + g*256 + b*256*256


def check(app, item, color):
    word = item[0]
    for i in range(1, item.shape[-1]):
        word = word + '|' + item[i]
    for para in app.Paragraphs:
        match = re.search(word, str(para))
        if match:
            print('match', match)
            start, end = match.span()
            print('start', para.Range.Start)
            app.Range(para.Range.Start+start,
                      para.Range.Start+end).Shading.BackgroundPatternColor = colorcode_to_int(color)


def main():
    #Wordを起動する : Applicationオブジェクトを生成する
    wd_app = win32com.client.Dispatch("Word.Application")
    #Wordを画面表示する : VisibleプロパティをTrueにする
    wd_app.Visible = True
    word = wd_app.Documents.Open(
        r"C:\Users\A\Desktop\修了論文_ver1.docx")  # ←ここにチェックしたいファイル名を入れる

    #　ひらがなで書くべき表現のチェック
    check(word, hiragana, '#006400')
    # 悪い表現のチェック
    check(word, bad_word, '#ff0000')
    # できる限りさけるべき表現のチェック
    check(word, avoid_word, '#ffd700')
    # 研究室用
    check(word, lab_word, '#ff1493')


if __name__ == '__main__':
    main()
