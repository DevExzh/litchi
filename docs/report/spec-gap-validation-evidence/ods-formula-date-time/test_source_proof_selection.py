"""A proof selection must cover its exact requirements without borrowing others."""
import copy
import importlib.util
from pathlib import Path
import unittest

path = Path(__file__).with_name('coverage_negative_cases.py')
spec = importlib.util.spec_from_file_location('coverage_fixtures', path)
fixtures = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fixtures)


class ProofSelectionTests(unittest.TestCase):
    def setUp(self):
        self.verifier = fixtures.load_verifier()

    def append_unselected(self, fixture):
        proof = copy.deepcopy(fixture['receipt_value']['proofs'][0])
        proof['id'] = 'proof.other'
        proof['requirements'] = ['another requirement']
        fixture['receipt_value']['proofs'].append(proof)
        fixtures.write_source_review_receipt(self.verifier, fixture)

    def test_independent_unselected_proof_does_not_invalidate_selection(self):
        self.assertEqual(fixtures.source_review_case(self.verifier, self.append_unselected), [])

    def test_missing_selected_proof_is_refused(self):
        def mutate(fixture):
            self.append_unselected(fixture)
            fixture['binding']['identifiers'] = ['proof.missing']
        with self.assertRaisesRegex(self.verifier.VerificationError, 'absent from receipt'):
            fixtures.source_review_case(self.verifier, mutate)

    def test_unselected_proof_cannot_supply_missing_coverage(self):
        def mutate(fixture):
            self.append_unselected(fixture)
            fixture['requirements'].add('another requirement')
            fixture['binding']['requirements'].append('another requirement')
        with self.assertRaisesRegex(self.verifier.VerificationError, 'do not exactly cover'):
            fixtures.source_review_case(self.verifier, mutate)

    def test_unselected_proofs_still_require_valid_structure(self):
        def mutate(fixture):
            self.append_unselected(fixture)
            fixture['receipt_value']['proofs'][1].pop('source')
            fixtures.write_source_review_receipt(self.verifier, fixture)
        with self.assertRaisesRegex(self.verifier.VerificationError, 'unexpected fields'):
            fixtures.source_review_case(self.verifier, mutate)


if __name__ == '__main__':
    unittest.main()
