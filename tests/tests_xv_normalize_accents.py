import unittest
from core import UnitTestCreator
from core import TestSerializer
from core import debugger
from excel_enhancer import xv_normalize_accents

debugger.testing = True


class MyTestForXvNormalizeAccents(unittest.TestCase):

    def test_xv_normalize_accents0(self):
        # get arguments
        arg0 = TestSerializer.deserialize(([0], [['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                 ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                 ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'], ['Webdesigner'],
                                                 ['Garment/textile technologist'], ['Chirurgien'],
                                                 ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                 ['Ingenieur structures'], ['Research officer, government'],
                                                 ['Produtor de eventos'], ['Charity officer']]), 'DataFrame')

        # get expected result
        ret_cmp = (TestSerializer.deserialize(([0], [['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                     ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                     ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                     ['Webdesigner'], ['Garment/textile technologist'], ['Chirurgien'],
                                                     ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                     ['Ingenieur structures'], ['Research officer, government'],
                                                     ['Produtor de eventos'], ['Charity officer']]), 'DataFrame'), [
                       TestSerializer.deserialize(([0], [['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                         ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                         ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                         ['Webdesigner'], ['Garment/textile technologist'],
                                                         ['Chirurgien'], ['Cancerologista ou Oncologista'],
                                                         ['Youth worker'], ['Ingenieur structures'],
                                                         ['Research officer, government'], ['Produtor de eventos'],
                                                         ['Charity officer']]), 'DataFrame')])
        # evaluate function
        ret = xv_normalize_accents(arg0)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xv_normalize_accents1(self):
        # get arguments
        arg0 = TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                 ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                 ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'], ['Webdesigner'],
                                                 ['Garment/textile technologist'], ['Chirurgien'],
                                                 ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                 ['Ingenieur structures'], ['Research officer, government'],
                                                 ['Produtor de eventos'], ['Charity officer'], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None], [None], [None], [None],
                                                 [None]]), 'DataFrame')

        # get expected result
        ret_cmp = (TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                     ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                     ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                     ['Webdesigner'], ['Garment/textile technologist'], ['Chirurgien'],
                                                     ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                     ['Ingenieur structures'], ['Research officer, government'],
                                                     ['Produtor de eventos'], ['Charity officer'], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None], [None]]), 'DataFrame'), [
                       TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                         ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                         ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                         ['Webdesigner'], ['Garment/textile technologist'],
                                                         ['Chirurgien'], ['Cancerologista ou Oncologista'],
                                                         ['Youth worker'], ['Ingenieur structures'],
                                                         ['Research officer, government'], ['Produtor de eventos'],
                                                         ['Charity officer'], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None], [None],
                                                         [None]]), 'DataFrame')])
        # evaluate function
        ret = xv_normalize_accents(arg0)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xv_normalize_accents2(self):
        # get arguments
        arg0 = TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                 ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                 ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'], ['Webdesigner'],
                                                 ['Garment/textile technologist'], ['Chirurgien'],
                                                 ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                 ['Ingenieur structures'], ['Research officer, government'],
                                                 ['Produtor de eventos'], ['Charity officer'], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None], [None], [None], [None],
                                                 [None], [None], [None], [None], [None], [None]]), 'DataFrame')

        # get expected result
        ret_cmp = (TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                     ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                     ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                     ['Webdesigner'], ['Garment/textile technologist'], ['Chirurgien'],
                                                     ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                     ['Ingenieur structures'], ['Research officer, government'],
                                                     ['Produtor de eventos'], ['Charity officer'], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None], [None], [None], [None], [None], [None],
                                                     [None], [None], [None]]), 'DataFrame'), [
                       TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                         ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                         ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                         ['Webdesigner'], ['Garment/textile technologist'],
                                                         ['Chirurgien'], ['Cancerologista ou Oncologista'],
                                                         ['Youth worker'], ['Ingenieur structures'],
                                                         ['Research officer, government'], ['Produtor de eventos'],
                                                         ['Charity officer'], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None], [None],
                                                         [None], [None], [None], [None], [None], [None], [None],
                                                         [None]]), 'DataFrame')])
        # evaluate function
        ret = xv_normalize_accents(arg0)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xv_normalize_accents3(self):
        # get arguments
        arg0 = TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                 ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                 ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'], ['Webdesigner'],
                                                 ['Garment/textile technologist'], ['Chirurgien'],
                                                 ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                 ['Ingenieur structures'], ['Research officer, government'],
                                                 ['Produtor de eventos'], ['Charity officer'], [None], [None], [None]]),
                                          'DataFrame')

        # get expected result
        ret_cmp = (TestSerializer.deserialize(([0], [['job'], ['Dentist'], ['Educateur'], ['寿司職人'], ['アニメーター'],
                                                     ['Garment/textile technologist'], ['Hydrologist'], ['アナウンサー'],
                                                     ['ウェティンクフランナー'], ['Petroleiro'], ['アナウンサー'], ['行政書士'],
                                                     ['Webdesigner'], ['Garment/textile technologist'], ['Chirurgien'],
                                                     ['Cancerologista ou Oncologista'], ['Youth worker'],
                                                     ['Ingenieur structures'], ['Research officer, government'],
                                                     ['Produtor de eventos'], ['Charity officer'], [None], [None],
                                                     [None]]), 'DataFrame'), [TestSerializer.deserialize(([0], [['job'],
                                                                                                                [
                                                                                                                    'Dentist'],
                                                                                                                [
                                                                                                                    'Educateur'],
                                                                                                                [
                                                                                                                    '寿司職人'],
                                                                                                                [
                                                                                                                    'アニメーター'],
                                                                                                                [
                                                                                                                    'Garment/textile technologist'],
                                                                                                                [
                                                                                                                    'Hydrologist'],
                                                                                                                [
                                                                                                                    'アナウンサー'],
                                                                                                                [
                                                                                                                    'ウェティンクフランナー'],
                                                                                                                [
                                                                                                                    'Petroleiro'],
                                                                                                                [
                                                                                                                    'アナウンサー'],
                                                                                                                [
                                                                                                                    '行政書士'],
                                                                                                                [
                                                                                                                    'Webdesigner'],
                                                                                                                [
                                                                                                                    'Garment/textile technologist'],
                                                                                                                [
                                                                                                                    'Chirurgien'],
                                                                                                                [
                                                                                                                    'Cancerologista ou Oncologista'],
                                                                                                                [
                                                                                                                    'Youth worker'],
                                                                                                                [
                                                                                                                    'Ingenieur structures'],
                                                                                                                [
                                                                                                                    'Research officer, government'],
                                                                                                                [
                                                                                                                    'Produtor de eventos'],
                                                                                                                [
                                                                                                                    'Charity officer'],
                                                                                                                [None],
                                                                                                                [None],
                                                                                                                [
                                                                                                                    None]]),
                                                                                                         'DataFrame')])
        # evaluate function
        ret = xv_normalize_accents(arg0)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xv_normalize_accents4(self):
        # get arguments
        arg0 = TestSerializer.deserialize(([0, 1, 2, 3], [['job', 'company', 'name', 'address'],
                                                          ['Dentist', 'Vlckova', 'Zdenek Vlcek',
                                                           'Tylisovska 1/4\n071 25 Rabi'],
                                                          ['Educateur', 'Blanchard Guillon S.A.R.L.',
                                                           'Charlotte Marques',
                                                           '532, rue de Thierry\n45580 Perret-la-Foret'],
                                                          ['寿司職人', '笹田銀行株式会社', '宮沢 康弘',
                                                           '群馬県横浜市神奈川区北青山2丁目12番16号 ハレス箭坪676'],
                                                          ['アニメーター', '合同会社高橋食品', '山本 涼平', '鹿児島県大島町大中33丁目20番2号'],
                                                          ['Garment/textile technologist', 'Maskova',
                                                           'Zaneta Novotna Th.D.', 'Zinkovska 6\n314 29 Jevicko'],
                                                          ['Hydrologist', 'Vanek', 'Richard Jelinek',
                                                           'Hnevkovskeho 2/2\n932 20 Cernovice'],
                                                          ['アナウンサー', '小林印刷株式会社', '鈴木 春香', '秋田県新宿区横林8丁目1番8号'],
                                                          ['ウェティンクフランナー', '合同会社村山保険', '大垣 淳',
                                                           '大分県千代田区中鉢石町20丁目2番15号 大京町シャルム209'],
                                                          ['Petroleiro', 'Peixoto', 'Mariane Mendes',
                                                           'Distrito Lais Fernandes, 27\nErnesto Nascimento\n07248627 Viana / GO'],
                                                          ['アナウンサー', '有限会社笹田電気', '加納 健一', '奈良県稲城市東三島16丁目18番4号'],
                                                          ['行政書士', '杉山食品株式会社', '渚 加奈', '福岡県香取郡神崎町横林20丁目21番13号'],
                                                          ['Webdesigner', 'Dias', 'Beatriz Mendes',
                                                           'Passarela Cardoso, 912\nCidade Jardim\n29741044 Moraes / MA'],
                                                          ['Garment/textile technologist', 'Kopecka Jelinek a.s.',
                                                           'Kristina Krejcova', 'V Padoline 5/9\n350 26 Upice'],
                                                          ['Chirurgien', 'Chauveau', 'Maurice Gomes',
                                                           '671, avenue Gerard\n73740 Maillarddan'],
                                                          ['Cancerologista ou Oncologista', 'Viana Viana e Filhos',
                                                           'Pedro Goncalves',
                                                           'Morro Eloah Rocha, 13\nVila Antena\n09226492 Gomes / AP'],
                                                          ['Youth worker', 'Kolarova', 'Rudolf Blazek',
                                                           'Prujezdna 683\n094 45 Dobrichovice'],
                                                          ['Ingenieur structures', 'Moreno', 'Raymond Gay',
                                                           'chemin Coste\n17650 Petitjean-la-Foret'],
                                                          ['Research officer, government', 'Polak o.s.',
                                                           'Drahomira Cermakova', 'Kostkova 2/7\n758 73 Raspenava'],
                                                          ['Produtor de eventos', 'da Costa', 'Esther da Cunha',
                                                           'Rua Martins, 3\nVila Nova Cachoeirinha 3ª Secao\n97347-525 Porto do Amparo / PA'],
                                                          ['Charity officer', 'Zemanova', 'Daniela Dolezalova',
                                                           'Na Pricne Mezi 7\n870 99 Mimon']]), 'DataFrame')

        # get expected result
        ret_cmp = (TestSerializer.deserialize(([0, 1, 2, 3], [['job', 'company', 'name', 'address'],
                                                              ['Dentist', 'Vlckova', 'Zdenek Vlcek',
                                                               'Tylisovska 1/4\n071 25 Rabi'],
                                                              ['Educateur', 'Blanchard Guillon S.A.R.L.',
                                                               'Charlotte Marques',
                                                               '532, rue de Thierry\n45580 Perret-la-Foret'],
                                                              ['寿司職人', '笹田銀行株式会社', '宮沢 康弘',
                                                               '群馬県横浜市神奈川区北青山2丁目12番16号 ハレス箭坪676'],
                                                              ['アニメーター', '合同会社高橋食品', '山本 涼平', '鹿児島県大島町大中33丁目20番2号'],
                                                              ['Garment/textile technologist', 'Maskova',
                                                               'Zaneta Novotna Th.D.', 'Zinkovska 6\n314 29 Jevicko'],
                                                              ['Hydrologist', 'Vanek', 'Richard Jelinek',
                                                               'Hnevkovskeho 2/2\n932 20 Cernovice'],
                                                              ['アナウンサー', '小林印刷株式会社', '鈴木 春香', '秋田県新宿区横林8丁目1番8号'],
                                                              ['ウェティンクフランナー', '合同会社村山保険', '大垣 淳',
                                                               '大分県千代田区中鉢石町20丁目2番15号 大京町シャルム209'],
                                                              ['Petroleiro', 'Peixoto', 'Mariane Mendes',
                                                               'Distrito Lais Fernandes, 27\nErnesto Nascimento\n07248627 Viana / GO'],
                                                              ['アナウンサー', '有限会社笹田電気', '加納 健一', '奈良県稲城市東三島16丁目18番4号'],
                                                              ['行政書士', '杉山食品株式会社', '渚 加奈', '福岡県香取郡神崎町横林20丁目21番13号'],
                                                              ['Webdesigner', 'Dias', 'Beatriz Mendes',
                                                               'Passarela Cardoso, 912\nCidade Jardim\n29741044 Moraes / MA'],
                                                              ['Garment/textile technologist', 'Kopecka Jelinek a.s.',
                                                               'Kristina Krejcova', 'V Padoline 5/9\n350 26 Upice'],
                                                              ['Chirurgien', 'Chauveau', 'Maurice Gomes',
                                                               '671, avenue Gerard\n73740 Maillarddan'],
                                                              ['Cancerologista ou Oncologista', 'Viana Viana e Filhos',
                                                               'Pedro Goncalves',
                                                               'Morro Eloah Rocha, 13\nVila Antena\n09226492 Gomes / AP'],
                                                              ['Youth worker', 'Kolarova', 'Rudolf Blazek',
                                                               'Prujezdna 683\n094 45 Dobrichovice'],
                                                              ['Ingenieur structures', 'Moreno', 'Raymond Gay',
                                                               'chemin Coste\n17650 Petitjean-la-Foret'],
                                                              ['Research officer, government', 'Polak o.s.',
                                                               'Drahomira Cermakova', 'Kostkova 2/7\n758 73 Raspenava'],
                                                              ['Produtor de eventos', 'da Costa', 'Esther da Cunha',
                                                               'Rua Martins, 3\nVila Nova Cachoeirinha 3ª Secao\n97347-525 Porto do Amparo / PA'],
                                                              ['Charity officer', 'Zemanova', 'Daniela Dolezalova',
                                                               'Na Pricne Mezi 7\n870 99 Mimon']]), 'DataFrame'), [
                       TestSerializer.deserialize(([0, 1, 2, 3], [['job', 'company', 'name', 'address'],
                                                                  ['Dentist', 'Vlckova', 'Zdenek Vlcek',
                                                                   'Tylisovska 1/4\n071 25 Rabi'],
                                                                  ['Educateur', 'Blanchard Guillon S.A.R.L.',
                                                                   'Charlotte Marques',
                                                                   '532, rue de Thierry\n45580 Perret-la-Foret'],
                                                                  ['寿司職人', '笹田銀行株式会社', '宮沢 康弘',
                                                                   '群馬県横浜市神奈川区北青山2丁目12番16号 ハレス箭坪676'],
                                                                  ['アニメーター', '合同会社高橋食品', '山本 涼平', '鹿児島県大島町大中33丁目20番2号'],
                                                                  ['Garment/textile technologist', 'Maskova',
                                                                   'Zaneta Novotna Th.D.',
                                                                   'Zinkovska 6\n314 29 Jevicko'],
                                                                  ['Hydrologist', 'Vanek', 'Richard Jelinek',
                                                                   'Hnevkovskeho 2/2\n932 20 Cernovice'],
                                                                  ['アナウンサー', '小林印刷株式会社', '鈴木 春香', '秋田県新宿区横林8丁目1番8号'],
                                                                  ['ウェティンクフランナー', '合同会社村山保険', '大垣 淳',
                                                                   '大分県千代田区中鉢石町20丁目2番15号 大京町シャルム209'],
                                                                  ['Petroleiro', 'Peixoto', 'Mariane Mendes',
                                                                   'Distrito Lais Fernandes, 27\nErnesto Nascimento\n07248627 Viana / GO'],
                                                                  ['アナウンサー', '有限会社笹田電気', '加納 健一', '奈良県稲城市東三島16丁目18番4号'],
                                                                  ['行政書士', '杉山食品株式会社', '渚 加奈', '福岡県香取郡神崎町横林20丁目21番13号'],
                                                                  ['Webdesigner', 'Dias', 'Beatriz Mendes',
                                                                   'Passarela Cardoso, 912\nCidade Jardim\n29741044 Moraes / MA'],
                                                                  ['Garment/textile technologist',
                                                                   'Kopecka Jelinek a.s.', 'Kristina Krejcova',
                                                                   'V Padoline 5/9\n350 26 Upice'],
                                                                  ['Chirurgien', 'Chauveau', 'Maurice Gomes',
                                                                   '671, avenue Gerard\n73740 Maillarddan'],
                                                                  ['Cancerologista ou Oncologista',
                                                                   'Viana Viana e Filhos', 'Pedro Goncalves',
                                                                   'Morro Eloah Rocha, 13\nVila Antena\n09226492 Gomes / AP'],
                                                                  ['Youth worker', 'Kolarova', 'Rudolf Blazek',
                                                                   'Prujezdna 683\n094 45 Dobrichovice'],
                                                                  ['Ingenieur structures', 'Moreno', 'Raymond Gay',
                                                                   'chemin Coste\n17650 Petitjean-la-Foret'],
                                                                  ['Research officer, government', 'Polak o.s.',
                                                                   'Drahomira Cermakova',
                                                                   'Kostkova 2/7\n758 73 Raspenava'],
                                                                  ['Produtor de eventos', 'da Costa', 'Esther da Cunha',
                                                                   'Rua Martins, 3\nVila Nova Cachoeirinha 3ª Secao\n97347-525 Porto do Amparo / PA'],
                                                                  ['Charity officer', 'Zemanova', 'Daniela Dolezalova',
                                                                   'Na Pricne Mezi 7\n870 99 Mimon']]), 'DataFrame')])
        # evaluate function
        ret = xv_normalize_accents(arg0)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xv_normalize_accents5(self):
        # get arguments
        arg0 = TestSerializer.deserialize(([0], [['# https://en.wikipedia.org/wiki/Diacritic'], [
            'Among the types of diacritic used in alphabets based on the Latin script are:'], [None], [
                                                     'accents (so called because the acute, grave, and circumflex were originally used to indicate different types of pitch accents in the polytonic transcription of Greek)'],
                                                 ['◌ – acute (Latin: apex)'], ['◌ – grave'], ['◌ – circumflex'],
                                                 ['◌ – caron, wedge'], ['◌ – double acute'], ['◌ – double grave'],
                                                 ['◌ – tilde'], ['dots'], ['◌ – overdot (Indic anusvara)'], [
                                                     '◌ – an underdot is used in Rheinische Dokumenta and in Hebrew, Indic and Arabic transcription'],
                                                 ['◌·◌ – interpunct'],
                                                 ['tittle, the superscript dot of the modern lowercase Latin i and j'],
                                                 ['◌ – umlaut or diaeresis'],
                                                 ['◌ː – triangular colon, used in the IPA to mark long vowels.'],
                                                 ['curves'], ['◌ – breve'], ['◌ – inverted breve'], [
                                                     '◌ – sicilicus, a palaeographic diacritic similar to a caron or breve'],
                                                 ['◌ – tilde'], ['◌ – titlo'], ['vertical stroke'], [
                                                     '◌ – syllabic a subscript vertical stroke is used in IPA to mark syllabicity and in Rheinische Dokumenta to mark a schwa'],
                                                 ['macron or horizontal line'], ['◌ – macron'], ['◌ – underbar'],
                                                 ['overlays'], ['◌ – vertical bar through the character'],
                                                 ['◌ – slash through the character'],
                                                 ['◌ – crossbar through the character'], ['ring'], ['◌ – overring'],
                                                 ['superscript curls'], ['◌ – apostrophe'], ['◌ – inverted apostrophe'],
                                                 ['◌ – reversed apostrophe'], ['◌ – hoi (Vietnamese: dau hoi)'],
                                                 ['◌ – horn (Vietnamese: dau moc)'], ['subscript curls'],
                                                 ['◌ – undercomma'], ['◌ – cedilla'],
                                                 ['◌ ◌ – hook, left or right, sometimes superscript'], ['◌ – ogonek'],
                                                 ['double marks (over or under two base characters)'],
                                                 ['◌◌ – double breve'], ['◌◌ – tie bar or top ligature'],
                                                 ['◌◌ – double circumflex'], ['◌◌ – longum'], ['◌◌ – double tilde'],
                                                 ['double sub/superscript diacritics'], ['◌  – double cedilla'],
                                                 ['◌  – double ogonek'], ['◌  – double diaeresis'],
                                                 ['◌ͺ – double ypogegrammeni']]), 'DataFrame')

        # get expected result
        ret_cmp = (TestSerializer.deserialize(([0], [['# https://en.wikipedia.org/wiki/Diacritic'], [
            'Among the types of diacritic used in alphabets based on the Latin script are:'], [None], [
                                                         'accents (so called because the acute, grave, and circumflex were originally used to indicate different types of pitch accents in the polytonic transcription of Greek)'],
                                                     ['◌ – acute (Latin: apex)'], ['◌ – grave'], ['◌ – circumflex'],
                                                     ['◌ – caron, wedge'], ['◌ – double acute'], ['◌ – double grave'],
                                                     ['◌ – tilde'], ['dots'], ['◌ – overdot (Indic anusvara)'], [
                                                         '◌ – an underdot is used in Rheinische Dokumenta and in Hebrew, Indic and Arabic transcription'],
                                                     ['◌·◌ – interpunct'], [
                                                         'tittle, the superscript dot of the modern lowercase Latin i and j'],
                                                     ['◌ – umlaut or diaeresis'],
                                                     ['◌ː – triangular colon, used in the IPA to mark long vowels.'],
                                                     ['curves'], ['◌ – breve'], ['◌ – inverted breve'], [
                                                         '◌ – sicilicus, a palaeographic diacritic similar to a caron or breve'],
                                                     ['◌ – tilde'], ['◌ – titlo'], ['vertical stroke'], [
                                                         '◌ – syllabic a subscript vertical stroke is used in IPA to mark syllabicity and in Rheinische Dokumenta to mark a schwa'],
                                                     ['macron or horizontal line'], ['◌ – macron'], ['◌ – underbar'],
                                                     ['overlays'], ['◌ – vertical bar through the character'],
                                                     ['◌ – slash through the character'],
                                                     ['◌ – crossbar through the character'], ['ring'], ['◌ – overring'],
                                                     ['superscript curls'], ['◌ – apostrophe'],
                                                     ['◌ – inverted apostrophe'], ['◌ – reversed apostrophe'],
                                                     ['◌ – hoi (Vietnamese: dau hoi)'],
                                                     ['◌ – horn (Vietnamese: dau moc)'], ['subscript curls'],
                                                     ['◌ – undercomma'], ['◌ – cedilla'],
                                                     ['◌ ◌ – hook, left or right, sometimes superscript'],
                                                     ['◌ – ogonek'],
                                                     ['double marks (over or under two base characters)'],
                                                     ['◌◌ – double breve'], ['◌◌ – tie bar or top ligature'],
                                                     ['◌◌ – double circumflex'], ['◌◌ – longum'], ['◌◌ – double tilde'],
                                                     ['double sub/superscript diacritics'], ['◌  – double cedilla'],
                                                     ['◌  – double ogonek'], ['◌  – double diaeresis'],
                                                     ['◌ͺ – double ypogegrammeni']]), 'DataFrame'), [
                       TestSerializer.deserialize(([0], [['# https://en.wikipedia.org/wiki/Diacritic'], [
                           'Among the types of diacritic used in alphabets based on the Latin script are:'], [None], [
                                                             'accents (so called because the acute, grave, and circumflex were originally used to indicate different types of pitch accents in the polytonic transcription of Greek)'],
                                                         ['◌ – acute (Latin: apex)'], ['◌ – grave'], ['◌ – circumflex'],
                                                         ['◌ – caron, wedge'], ['◌ – double acute'],
                                                         ['◌ – double grave'], ['◌ – tilde'], ['dots'],
                                                         ['◌ – overdot (Indic anusvara)'], [
                                                             '◌ – an underdot is used in Rheinische Dokumenta and in Hebrew, Indic and Arabic transcription'],
                                                         ['◌·◌ – interpunct'], [
                                                             'tittle, the superscript dot of the modern lowercase Latin i and j'],
                                                         ['◌ – umlaut or diaeresis'], [
                                                             '◌ː – triangular colon, used in the IPA to mark long vowels.'],
                                                         ['curves'], ['◌ – breve'], ['◌ – inverted breve'], [
                                                             '◌ – sicilicus, a palaeographic diacritic similar to a caron or breve'],
                                                         ['◌ – tilde'], ['◌ – titlo'], ['vertical stroke'], [
                                                             '◌ – syllabic a subscript vertical stroke is used in IPA to mark syllabicity and in Rheinische Dokumenta to mark a schwa'],
                                                         ['macron or horizontal line'], ['◌ – macron'],
                                                         ['◌ – underbar'], ['overlays'],
                                                         ['◌ – vertical bar through the character'],
                                                         ['◌ – slash through the character'],
                                                         ['◌ – crossbar through the character'], ['ring'],
                                                         ['◌ – overring'], ['superscript curls'], ['◌ – apostrophe'],
                                                         ['◌ – inverted apostrophe'], ['◌ – reversed apostrophe'],
                                                         ['◌ – hoi (Vietnamese: dau hoi)'],
                                                         ['◌ – horn (Vietnamese: dau moc)'], ['subscript curls'],
                                                         ['◌ – undercomma'], ['◌ – cedilla'],
                                                         ['◌ ◌ – hook, left or right, sometimes superscript'],
                                                         ['◌ – ogonek'],
                                                         ['double marks (over or under two base characters)'],
                                                         ['◌◌ – double breve'], ['◌◌ – tie bar or top ligature'],
                                                         ['◌◌ – double circumflex'], ['◌◌ – longum'],
                                                         ['◌◌ – double tilde'], ['double sub/superscript diacritics'],
                                                         ['◌  – double cedilla'], ['◌  – double ogonek'],
                                                         ['◌  – double diaeresis'], ['◌ͺ – double ypogegrammeni']]),
                                                  'DataFrame')])
        # evaluate function
        ret = xv_normalize_accents(arg0)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))


if __name__ == '__main__':
    unittest.main()
