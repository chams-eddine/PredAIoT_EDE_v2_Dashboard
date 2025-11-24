import unittest
from ede_core import make_decision

class TestEDE(unittest.TestCase):
    def test_execute(self):
        data = {"maintenance_cost": 100, "financial_loss_without": 200}
        self.assertEqual(make_decision(data)["decision"], "EXECUTE")

    def test_postpone(self):
        data = {"maintenance_cost": 100, "financial_loss_without": 60}
        self.assertEqual(make_decision(data)["decision"], "POSTPONE")

    def test_no_action(self):
        data = {"maintenance_cost": 100, "financial_loss_without": 40}
        self.assertEqual(make_decision(data)["decision"], "NO_ACTION")

if __name__ == '__main__':
    unittest.main()