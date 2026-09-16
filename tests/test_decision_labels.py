from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "environment"))

from runtime.grading.hybrid_semantic import apply_semantic_judgments  # noqa: E402


class DecisionLabelTests(unittest.TestCase):
    def test_labels_identify_the_actual_decision_source_without_changing_verdicts(self) -> None:
        criteria = [
            {
                "id": criterion_id,
                "description": criterion_id,
                "value": 1,
                "weight": 1,
                "category": "structure",
                "semantic": True,
            }
            for criterion_id in ("deterministic_pass", "deterministic_fail", "llm_pass")
        ]
        result = {
            "criteria": criteria,
            "semantic_review": {
                "task_id": "test_task",
                "criteria": [
                    {
                        "criterion_id": "deterministic_pass",
                        "hard_gate_met": True,
                        "hard_gate_evidence": "exact fact present",
                    },
                    {
                        "criterion_id": "deterministic_fail",
                        "hard_gate_met": False,
                        "hard_gate_evidence": "conflicting value present",
                    },
                    {
                        "criterion_id": "llm_pass",
                        "hard_gate_met": True,
                        "hard_gate_evidence": "exact prerequisite present",
                    },
                ],
            },
        }
        judgments = {
            "deterministic_pass": {
                "met": True,
                "reason": "deterministic professional association matched",
                "response_id": None,
            },
            "deterministic_fail": {
                "met": False,
                "reason": "deterministic hard gate failed: conflicting value present",
                "response_id": None,
            },
            "llm_pass": {
                "met": True,
                "reason": "equivalent professional wording is correctly associated",
                "response_id": "resp_test",
            },
        }

        updated = apply_semantic_judgments(result, judgments, judge_model="judge-test")
        by_id = {row["id"]: row for row in updated["criteria"]}
        audit_by_id = {
            row["criterion_id"]: row
            for row in updated["semantic_review_result"]["criteria"]
        }

        self.assertEqual(
            {criterion_id: row["value"] for criterion_id, row in by_id.items()},
            {"deterministic_pass": 1, "deterministic_fail": 0, "llm_pass": 1},
        )
        for criterion_id in ("deterministic_pass", "deterministic_fail"):
            self.assertIn("Deterministic check:", by_id[criterion_id]["evidence"])
            self.assertNotIn("semantic_judge", by_id[criterion_id]["evidence"])
            self.assertNotIn("judge-test", by_id[criterion_id]["evidence"])
            self.assertFalse(audit_by_id[criterion_id]["judge_invoked"])
            self.assertIsNone(audit_by_id[criterion_id]["response_id"])

        self.assertIn("LLM semantic review:", by_id["llm_pass"]["evidence"])
        self.assertIn("judge-test", by_id["llm_pass"]["evidence"])
        self.assertTrue(audit_by_id["llm_pass"]["judge_invoked"])
        self.assertEqual(audit_by_id["llm_pass"]["response_id"], "resp_test")

    def test_missing_semantic_verdict_is_not_mislabeled_as_an_llm_call(self) -> None:
        result = {
            "criteria": [
                {
                    "id": "missing",
                    "description": "missing",
                    "value": 1,
                    "weight": 1,
                    "category": "structure",
                    "semantic": True,
                }
            ],
            "semantic_review": {
                "task_id": "test_task",
                "criteria": [
                    {
                        "criterion_id": "missing",
                        "hard_gate_met": True,
                        "hard_gate_evidence": "exact prerequisite present",
                    }
                ],
            },
        }

        updated = apply_semantic_judgments(result, {}, judge_model="judge-test")
        criterion = updated["criteria"][0]
        audit = updated["semantic_review_result"]["criteria"][0]
        self.assertEqual(criterion["value"], 0)
        self.assertIn("Semantic review: UNMET", criterion["evidence"])
        self.assertNotIn("judge-test", criterion["evidence"])
        self.assertEqual(audit["decision_source"], "semantic_review_no_verdict")
        self.assertFalse(audit["judge_invoked"])


if __name__ == "__main__":
    unittest.main()
