'''Regression tests for IAMC hydro-capacity weight selection.'''

import unittest

import pandas as pd

from iamc_utils import (
    HYDRO_BARE_VARIABLE,
    HYDRO_RESERVOIR_VARIABLE,
    HYDRO_ROR_VARIABLE,
    select_hydro_capacity_weights,
)


class HydroWeightTests(unittest.TestCase):
    def setUp(self):
        self.data = pd.DataFrame([
            {'region': 'AA01', 'variable': HYDRO_BARE_VARIABLE,
             'unit': 'MW', 'value': 10.0},
            {'region': 'AA01', 'variable': HYDRO_RESERVOIR_VARIABLE,
             'unit': 'MW', 'value': 40.0},
            {'region': 'AA01', 'variable': HYDRO_ROR_VARIABLE,
             'unit': 'MW', 'value': 5.0},
            {'region': 'AA02', 'variable': HYDRO_RESERVOIR_VARIABLE,
             'unit': 'GW', 'value': 0.06},
            {'region': 'AA02', 'variable': HYDRO_ROR_VARIABLE,
             'unit': 'MW', 'value': 15.0},
            {'region': 'BB01', 'variable': HYDRO_BARE_VARIABLE,
             'unit': 'GW', 'value': 0.1},
            {'region': 'AA>parent', 'variable': HYDRO_RESERVOIR_VARIABLE,
             'unit': 'MW', 'value': 999.0},
            {'region': 'AA01',
             'variable': 'Capacity|Electricity|Hydro|Pumped Storage',
             'unit': 'MW', 'value': 999.0},
        ])
        self.groups = {'AA': ['AA01', 'AA02'], 'BB': ['BB01', 'BB02']}

    def test_annual_sums_all_non_pumped_components(self):
        weights, _ = select_hydro_capacity_weights(
            self.data, self.groups, 'annual_total')
        self.assertEqual(weights['AA01'], 55.0)
        self.assertEqual(weights['AA02'], 75.0)
        self.assertEqual(weights['BB01'], 100.0)
        self.assertEqual(weights['BB02'], 0.0)

    def test_seasonal_prefers_reservoir_and_falls_back_to_bare(self):
        weights, sources = select_hydro_capacity_weights(
            self.data, self.groups, 'seasonal_reservoir')
        self.assertEqual(weights['AA01'], 40.0)
        self.assertEqual(weights['AA02'], 60.0)
        self.assertEqual(sources['AA'], 'reservoir capacity')
        self.assertEqual(weights['BB01'], 100.0)
        self.assertEqual(weights['BB02'], 0.0)
        self.assertIn('legacy fallback', sources['BB'])

    def test_rejects_unknown_capacity_units(self):
        bad = self.data.iloc[:1].copy()
        bad['unit'] = 'TJ'
        with self.assertRaisesRegex(ValueError, 'Unsupported hydro-capacity unit'):
            select_hydro_capacity_weights(
                bad, {'AA': ['AA01']}, 'seasonal_reservoir')


if __name__ == '__main__':
    unittest.main()
