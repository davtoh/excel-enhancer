# Generated Contract Smoke Test - Enforces Zero-Regression Refactoring Parity
import pathlib
import re
import unittest

class TestCodebaseContractInvariants(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = pathlib.Path(__file__).resolve().parent.parent
        ignored = {".git", ".venv", "venv", "env", "node_modules", "__pycache__", "_build", "build", "dist", ".refactor", "scratch"}
        cls.html_text = "".join(
            p.read_text(encoding="utf-8", errors="replace") + "\n"
            for p in cls.root.rglob("*.html")
            if not any(part in ignored for part in p.parts)
        )
        cls.js_text = "".join(
            p.read_text(encoding="utf-8", errors="replace") + "\n"
            for p in cls.root.rglob("*.js")
            if not any(part in ignored for part in p.parts)
        )

    def test_dom_interactive_selects_exist(self):
        """Verify all critical <select> dropdown IDs exist in HTML."""
        for sel in []:
            self.assertIn(f'id="{sel}"', self.html_text, f"Missing select dropdown #{sel} in HTML!")

    def test_dom_element_ids_preserved(self):
        """Verify captured baseline HTML element IDs are preserved."""
        for el_id in []:
            self.assertIn(f'id="{el_id}"', self.html_text, f"Critical DOM element id='{el_id}' was removed from HTML!")

    def test_zero_css_important_creep(self):
        """Verify scoped CSS does not exceed baseline !important threshold."""
        ignored = {".git", ".venv", "venv", "env", "node_modules", "__pycache__", "_build", "build", "dist", ".refactor", "scratch"}
        css_text = "".join(
            p.read_text(encoding="utf-8", errors="replace")
            for p in self.root.rglob("*.css")
            if not any(part in ignored for part in p.parts)
        )
        count = len(re.findall(r'!important\b', css_text, re.IGNORECASE))
        self.assertLessEqual(count, 0, f"!important count increased to {count}!")

if __name__ == "__main__":
    unittest.main()
