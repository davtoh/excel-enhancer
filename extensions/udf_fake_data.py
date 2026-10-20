from faker import Faker
# from faker.config import AVAILABLE_LOCALES
import pandas as pd
import xlwings as xw
from core import helper_correct_to_str

PREFERRED_LOCALES = ['es_ES', 'pt_BR', 'it_IT', 'fr_FR', 'en_US', 'ja_JP']


@xw.func
@xw.arg('no_rows', doc='number of rows to create')
@xw.arg('locale', doc="locales like ['es_ES', 'pt_BR', 'it_IT', 'fr_FR', 'en_US', 'ja_JP']")
@xw.arg('header', doc='True to print header; False to not print')
@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xf_fake_profiles(no_rows, locale=None, header=True):

    # DONE NO_CUENTA, ID, NOMBRE, CONCEPTO, CREDITO, DEBITO
    if locale is None:
        locale = PREFERRED_LOCALES
    fake = Faker(locale)

    rows = []
    for _ in range(int(no_rows)):
        profile = fake.profile()
        profile["current_location"] = str(tuple([str(i) for i in profile["current_location"]]))
        profile["website"] = str(profile["website"])
        rows.append(profile)

    profiles = pd.DataFrame(rows)

    if header:
        return [list(profiles.columns)] + profiles.values.tolist()
    return profiles.values.tolist()
