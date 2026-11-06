from __future__ import annotations

import unittest

from app.network_filter import build_ip_condition, parse_ip_targets, summarize_ip_targets
from xiaoxin_assistant import NetworkWorker, ensure_windivert_impostor_guard


class NetworkFilterTests(unittest.TestCase):
    def _condition(self, text):
        return build_ip_condition(parse_ip_targets(text))

    def test_empty_targets_build_no_condition(self):
        self.assertEqual("", self._condition(" , ； "))
        self.assertEqual("全部 IP", summarize_ip_targets(""))

    def test_single_ip_keeps_previous_rule(self):
        self.assertEqual("(ip.SrcAddr == 1.1.1.1 or ip.DstAddr == 1.1.1.1)", self._condition("1.1.1.1"))

    def test_multiple_separators_and_duplicates(self):
        targets = parse_ip_targets("1.1.1.1, 8.8.8.8；9.9.9.9\n1.1.1.1，2.2.2.2")
        self.assertEqual(4, len(targets))
        self.assertTrue(summarize_ip_targets("1.1.1.1, 8.8.8.8").startswith("2 个"))

    def test_cidr_and_reversed_range(self):
        condition = self._condition("10.0.0.5/24 1.1.1.20-1.1.1.10")
        self.assertIn("ip.SrcAddr >= 10.0.0.0 and ip.SrcAddr <= 10.0.0.255", condition)
        self.assertIn("ip.DstAddr >= 1.1.1.10 and ip.DstAddr <= 1.1.1.20", condition)

    def test_mixed_ipv4_ipv6(self):
        condition = self._condition("2001:db8::1, 1.1.1.1")
        self.assertIn("ipv6.DstAddr == 2001:db8::1", condition)
        self.assertIn("ip.DstAddr == 1.1.1.1", condition)

    def test_invalid_items_name_the_entry(self):
        for text, item in (("1.1.1.1, abc", "abc"), ("1.1.1.1-2001:db8::1", "1.1.1.1-2001:db8::1"), ("1.1.1.1/99", "1.1.1.1/99")):
            with self.assertRaises(ValueError) as caught:
                parse_ip_targets(text)
            self.assertIn(item, str(caught.exception))

    def test_generated_rules_pass_windivert_check(self):
        for text in ("1.1.1.1, 8.8.8.8", "10.0.0.0/24, 1.1.1.1-1.1.1.20", "2001:db8::1, 1.1.1.1"):
            self.assertTrue(NetworkWorker.validate_filter(ensure_windivert_impostor_guard(self._condition(text))))

    def test_too_many_targets_give_clear_message(self):
        text = ", ".join(f"10.1.{i // 250}.{i % 250 + 1}" for i in range(400))
        with self.assertRaises(ValueError) as caught:
            NetworkWorker.validate_filter(ensure_windivert_impostor_guard(self._condition(text)))
        self.assertIn("目标 IP 数量", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
