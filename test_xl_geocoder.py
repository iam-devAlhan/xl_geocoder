import unittest
from xl_geocoder import parse_street_name


class test_parse_street_name(unittest.TestCase):

    num_first_cases = {
        'positive': {
            '11-go listopada 17':      '17, 11-go listopada',
            r'11-go listopada 17\23':  r'17\23, 11-go listopada',
            '11-go listopada 17-23':   '17-23, 11-go listopada',
            '11-go listopada 17/1243': '17/1243, 11-go listopada',
            '3 Maja 23':    '23, 3 Maja',
            '3 Maja 23a':   '23a, 3 Maja',
            '3 Maja 2 a':   '2a, 3 Maja',
            '3 Maja 2B':    '2B, 3 Maja',
            '3 Maja 2 B':   '2B, 3 Maja',
            '3 Maja 5/a':   '5a, 3 Maja',
            '3 Maja 23/4a': '23/4a, 3 Maja',
            '3 Maja 2-6':   '2-6, 3 Maja',
            'grudnia 1970 43': '43, grudnia 1970'},
        'negative': {
            '3 Maja 23aa': '3 Maja 23aa',
            '3 Maja 2 aa': '3 Maja 2 aa',
            '3 Maja 2 Aa': '3 Maja 2 Aa',
            '3 Maja 2 AA': '3 Maja 2 AA',
            '3 Maja 2-A': '3 Maja 2-A',
            '3 Maja -2': '3 Maja -2',
           r'11-go listopada \32': r'11-go listopada \32',
            '11-go listopada /32': '11-go listopada /32',
           r'11-go listopada 17\\23': r'11-go listopada 17\\23',
           r'11-go listopada 17//23': r'11-go listopada 17//23',
            '3 Maja12': '3 Maja12'}
    }

    def test_name_filter(self):
        name = 'ul. Dworcowa 35'

        self.assertFalse(parse_street_name(name, name_filter=['Dworcowa']))
        self.assertFalse(parse_street_name(name, name_filter=['Krucza', 'ul.']))
        self.assertFalse(parse_street_name(name, name_filter=['Krucza', 'Dworcowa']))
        self.assertEqual(parse_street_name(name, name_filter=['Krucza']), name)
        self.assertEqual(parse_street_name(name, name_filter=''), name)

    def test_expand_abbrev(self):
        name = 'ul. św. Jerzego 20'

        self.assertEqual(parse_street_name(name, expand_abbrev={'św.':'świętego'}), 
                        'ul. świętego Jerzego 20')
        self.assertEqual(parse_street_name(name,
                         expand_abbrev={'św.': 'świętego', 'ul.': 'ulica'}),
                         'ulica świętego Jerzego 20')
        self.assertEqual(parse_street_name(name, expand_abbrev={'Św.': 'świętego'}),
                         'ul. świętego Jerzego 20')
        self.assertEqual(parse_street_name(name, expand_abbrev={'św.':'Świętego'}), 
                        'ul. Świętego Jerzego 20')

    def test_remove_abbrev(self):
        name = 'ul. św. Jerzego 20'

        self.assertEquals(parse_street_name(name, remove_abbrev=None), name)
        self.assertEquals(parse_street_name(name, remove_abbrev=True), 'Jerzego 20')

    def test_building_num_first(self):
        for case, answer in self.num_first_cases['positive'].items():
            self.assertEquals(parse_street_name(case, building_number_first=True), answer)

        for case, answer in self.num_first_cases['negative'].items():
            self.assertEquals(parse_street_name(case, building_number_first=True), answer)


if __name__ == "__main__":
    unittest.main()
