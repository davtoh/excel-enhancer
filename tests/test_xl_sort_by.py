
import unittest
from core import UnitTestCreator
from core import TestSerializer
from core import debugger
from excel_enhancer import xl_sort_by

debugger.testing = True


class MyTestForXlSortBy(unittest.TestCase):

    def test_xl_sort_by0(self):
        # get arguments
        arg0 = TestSerializer.deserialize((['job'], [['Patent attorney'], ['Engineer, automotive'], ['Domoticien'], ['Technicien des industries du verre']]), 'DataFrame')
        arg1 = 'job'
        arg2 = True
        arg3 = 'first'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job'], [['Domoticien'], ['Engineer, automotive'], ['Patent attorney'], ['Technicien des industries du verre']]), 'DataFrame'), [TestSerializer.deserialize((['job'], [['Patent attorney'], ['Engineer, automotive'], ['Domoticien'], ['Technicien des industries du verre']]), 'DataFrame')])
        # evaluate function
        ret = xl_sort_by(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xl_sort_by1(self):
        # get arguments
        arg0 = TestSerializer.deserialize((['job', 'company'], [['Patent attorney', 'Anderson, Rivera and Shaffer'], ['Engineer, automotive', 'Font Inc'], ['Domoticien', 'Charles S.A.R.L.'], ['Technicien des industries du verre', 'Petitjean']]), 'DataFrame'
)
        arg1 = 'job'
        arg2 = True
        arg3 = 'first'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job', 'company'], [['Domoticien', 'Charles S.A.R.L.'], ['Engineer, automotive', 'Font Inc'], ['Patent attorney', 'Anderson, Rivera and Shaffer'], ['Technicien des industries du verre', 'Petitjean']]), 'DataFrame'), [TestSerializer.deserialize((['job', 'company'], [['Patent attorney', 'Anderson, Rivera and Shaffer'], ['Engineer, automotive', 'Font Inc'], ['Domoticien', 'Charles S.A.R.L.'], ['Technicien des industries du verre', 'Petitjean']]), 'DataFrame')])
        # evaluate function
        ret = xl_sort_by(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xl_sort_by2(self):
        # get arguments
        arg0 = TestSerializer.deserialize((['job', 'company'], [['Patent attorney', 'Anderson, Rivera and Shaffer'], ['Engineer, automotive', 'Font Inc'], ['Domoticien', 'Charles S.A.R.L.'], ['Technicien des industries du verre', 'Petitjean']]), 'DataFrame'
)
        arg1 = None
        arg2 = True
        arg3 = 'first'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job', 'company'], [['Domoticien', 'Charles S.A.R.L.'], ['Engineer, automotive', 'Font Inc'], ['Patent attorney', 'Anderson, Rivera and Shaffer'], ['Technicien des industries du verre', 'Petitjean']]), 'DataFrame'), [TestSerializer.deserialize((['job', 'company'], [['Patent attorney', 'Anderson, Rivera and Shaffer'], ['Engineer, automotive', 'Font Inc'], ['Domoticien', 'Charles S.A.R.L.'], ['Technicien des industries du verre', 'Petitjean']]), 'DataFrame')])
        # evaluate function
        ret = xl_sort_by(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xl_sort_by3(self):
        # get arguments
        arg0 = TestSerializer.deserialize((['job', 'company', 'ssn', 'residence', 'current_location', 'blood_group', 'website', 'username', 'name', 'sex', 'address', 'mail', 'birthdate'], [['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313', '921 Francis Estate Apt. 373\nLake Jessicaville, TX 80346', "('82.603046', '-87.345006')", 'O-', "['https://www.campbell.org/', 'https://stewart.com/', 'https://www.oliver-mccullough.com/']", 'anthony37', 'Richard Brown', 'M', '06914 Shannon Ramp\nLake Adrian, NJ 74012', 'steelelisa@yahoo.com', 18910.0], ['Engineer, automotive', 'Font Inc', '558-10-8908', 'Calle de Aitana Hoz 74\nToledo, 28554', "('-6.020216', '-130.454104')", 'AB+', "['https://www.valentin-hidalgo.org/']", 'peralesconsuela', 'Godofredo Casas Camacho', 'M', 'Paseo Salomé Vergara 51 Apt. 20 \nGuadalajara, 51555', 'petrona33@hotmail.com', 40859.0], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521', '66, rue de Masse\n97936 Delaunay', "('-2.832681', '-99.509790')", 'A-', "['https://chauvin.com/', 'https://www.toussaint.fr/']", 'ahoareau', 'Émilie de Faure', 'F', '302, boulevard de Laurent\n03117 Bruneldan', 'margauxpetit@live.com', 24169.0], ['Technicien des industries du verre', 'Petitjean', '601-97-4157', '76, boulevard Alice Millet\n01441 Bigot', "('0.6079185', '-2.943182')", 'B+', "['https://www.duval.com/']", 'nicolasda-costa', 'Auguste Potier', 'M', 'avenue de Guillot\n27068 Sainte Carolineboeuf', 'vboulanger@free.fr', 40954.0]]), 'DataFrame')
        arg1 = None
        arg2 = True
        arg3 = 'first'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job', 'company', 'ssn', 'residence', 'current_location', 'blood_group', 'website', 'username', 'name', 'sex', 'address', 'mail', 'birthdate'], [['Domoticien', 'Charles S.A.R.L.', '601-32-6521', '66, rue de Masse\n97936 Delaunay', "('-2.832681', '-99.509790')", 'A-', "['https://chauvin.com/', 'https://www.toussaint.fr/']", 'ahoareau', 'Émilie de Faure', 'F', '302, boulevard de Laurent\n03117 Bruneldan', 'margauxpetit@live.com', 24169.0], ['Engineer, automotive', 'Font Inc', '558-10-8908', 'Calle de Aitana Hoz 74\nToledo, 28554', "('-6.020216', '-130.454104')", 'AB+', "['https://www.valentin-hidalgo.org/']", 'peralesconsuela', 'Godofredo Casas Camacho', 'M', 'Paseo Salomé Vergara 51 Apt. 20 \nGuadalajara, 51555', 'petrona33@hotmail.com', 40859.0], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313', '921 Francis Estate Apt. 373\nLake Jessicaville, TX 80346', "('82.603046', '-87.345006')", 'O-', "['https://www.campbell.org/', 'https://stewart.com/', 'https://www.oliver-mccullough.com/']", 'anthony37', 'Richard Brown', 'M', '06914 Shannon Ramp\nLake Adrian, NJ 74012', 'steelelisa@yahoo.com', 18910.0], ['Technicien des industries du verre', 'Petitjean', '601-97-4157', '76, boulevard Alice Millet\n01441 Bigot', "('0.6079185', '-2.943182')", 'B+', "['https://www.duval.com/']", 'nicolasda-costa', 'Auguste Potier', 'M', 'avenue de Guillot\n27068 Sainte Carolineboeuf', 'vboulanger@free.fr', 40954.0]]), 'DataFrame'), [arg0])
        # evaluate function
        ret = xl_sort_by(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xl_sort_by4(self):
        # get arguments
        arg0 = TestSerializer.deserialize((['job', 'company', 'ssn', 'residence', 'current_location', 'blood_group', 'website', 'username', 'name', 'sex', 'address', 'mail', 'birthdate'], [['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313', '921 Francis Estate Apt. 373\nLake Jessicaville, TX 80346', "('82.603046', '-87.345006')", 'O-', "['https://www.campbell.org/', 'https://stewart.com/', 'https://www.oliver-mccullough.com/']", 'anthony37', 'Richard Brown', 'M', '06914 Shannon Ramp\nLake Adrian, NJ 74012', 'steelelisa@yahoo.com', 18910.0], ['Engineer, automotive', 'Font Inc', '558-10-8908', 'Calle de Aitana Hoz 74\nToledo, 28554', "('-6.020216', '-130.454104')", 'AB+', "['https://www.valentin-hidalgo.org/']", 'peralesconsuela', 'Godofredo Casas Camacho', 'M', 'Paseo Salomé Vergara 51 Apt. 20 \nGuadalajara, 51555', 'petrona33@hotmail.com', 40859.0], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521', '66, rue de Masse\n97936 Delaunay', "('-2.832681', '-99.509790')", 'A-', "['https://chauvin.com/', 'https://www.toussaint.fr/']", 'ahoareau', 'Émilie de Faure', 'F', '302, boulevard de Laurent\n03117 Bruneldan', 'margauxpetit@live.com', 24169.0], ['Technicien des industries du verre', 'Petitjean', '601-97-4157', '76, boulevard Alice Millet\n01441 Bigot', "('0.6079185', '-2.943182')", 'B+', "['https://www.duval.com/']", 'nicolasda-costa', 'Auguste Potier', 'M', 'avenue de Guillot\n27068 Sainte Carolineboeuf', 'vboulanger@free.fr', 40954.0]]), 'DataFrame')
        arg1 = None
        arg2 = False
        arg3 = 'first'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job', 'company', 'ssn', 'residence', 'current_location', 'blood_group', 'website', 'username', 'name', 'sex', 'address', 'mail', 'birthdate'], [['Technicien des industries du verre', 'Petitjean', '601-97-4157', '76, boulevard Alice Millet\n01441 Bigot', "('0.6079185', '-2.943182')", 'B+', "['https://www.duval.com/']", 'nicolasda-costa', 'Auguste Potier', 'M', 'avenue de Guillot\n27068 Sainte Carolineboeuf', 'vboulanger@free.fr', 40954.0], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313', '921 Francis Estate Apt. 373\nLake Jessicaville, TX 80346', "('82.603046', '-87.345006')", 'O-', "['https://www.campbell.org/', 'https://stewart.com/', 'https://www.oliver-mccullough.com/']", 'anthony37', 'Richard Brown', 'M', '06914 Shannon Ramp\nLake Adrian, NJ 74012', 'steelelisa@yahoo.com', 18910.0], ['Engineer, automotive', 'Font Inc', '558-10-8908', 'Calle de Aitana Hoz 74\nToledo, 28554', "('-6.020216', '-130.454104')", 'AB+', "['https://www.valentin-hidalgo.org/']", 'peralesconsuela', 'Godofredo Casas Camacho', 'M', 'Paseo Salomé Vergara 51 Apt. 20 \nGuadalajara, 51555', 'petrona33@hotmail.com', 40859.0], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521', '66, rue de Masse\n97936 Delaunay', "('-2.832681', '-99.509790')", 'A-', "['https://chauvin.com/', 'https://www.toussaint.fr/']", 'ahoareau', 'Émilie de Faure', 'F', '302, boulevard de Laurent\n03117 Bruneldan', 'margauxpetit@live.com', 24169.0]]), 'DataFrame'), [arg0])
        # evaluate function
        ret = xl_sort_by(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))
