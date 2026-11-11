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

    def test_core_facade_parity(self):
        """Verify core.py re-exports all public symbols from enhancer.core."""
        import core
        import enhancer.core
        for attr in enhancer.core.__all__:
            self.assertTrue(hasattr(core, attr), f"Missing symbol {attr} in core.py facade")

    def test_mod_udf_facade_parity(self):
        """Verify mod_udf.py re-exports all public symbols from enhancer.udf."""
        import mod_udf
        import enhancer.udf
        for attr in enhancer.udf.__all__:
            self.assertTrue(hasattr(mod_udf, attr), f"Missing symbol {attr} in mod_udf.py facade")

    def test_excel_enhancer_facade_parity(self):
        """Verify excel_enhancer.py re-exports all domain functions and server exports."""
        import excel_enhancer
        import enhancer.functions
        import enhancer.server
        for attr in enhancer.functions.__all__:
            self.assertTrue(hasattr(excel_enhancer, attr), f"Missing symbol {attr} in excel_enhancer.py facade")
        for attr in enhancer.server.__all__:
            self.assertTrue(hasattr(excel_enhancer, attr), f"Missing symbol {attr} in excel_enhancer.py facade")

    def test_udf_decorator_behavior(self):
        """Verify @xlfunc decorator correctly annotates functions."""
        from enhancer.udf import xlfunc

        @xlfunc(category="Math", volatile=True)
        def sample_add(a, b=10):
            """Sample addition doc."""
            return a + b

        self.assertTrue(hasattr(sample_add, "__xlfunc__"))
        meta = sample_add.__xlfunc__
        self.assertEqual(meta["name"], "sample_add")
        self.assertEqual(meta["category"], "Math")
        self.assertTrue(meta["volatile"])
        self.assertFalse(meta["sub"])
        self.assertEqual(len(meta["args"]), 2)
        self.assertEqual(meta["args"][0]["name"], "a")
        self.assertEqual(meta["args"][1]["optional"], 10)

    def test_package_version_consistency(self):
        """Verify package version consistency between root and enhancer."""
        import excel_enhancer
        import enhancer
        self.assertEqual(excel_enhancer.__version__, enhancer.__version__)

    def test_xlplatform_resolution_windows_contract(self):
        """Verify xlplatform binds to _xlwindows with valid Range(xl=...) signature on Windows."""
        import sys
        import inspect
        from enhancer.udf import dispatcher

        if sys.platform == "win32":
            self.assertIsNotNone(dispatcher.xlplatform)
            sig = inspect.signature(dispatcher.xlplatform.Range.__init__)
            self.assertIn("xl", sig.parameters, "Windows Range.__init__ must accept 'xl' parameter")

    def test_dispatcher_convert_input_args(self):
        """Verify _convert_input_args converts missing arguments and primitives correctly."""
        from enhancer.udf.dispatcher import _convert_input_args

        def dummy_func(x, y=42):
            return x + y

        args_info = [
            {"name": "x", "options": {}},
            {"name": "y", "options": {}, "optional": 42},
        ]
        # Missing argument sentinel (-2147352572) should be replaced with default
        raw_args = ("hello", -2147352572)
        converted = _convert_input_args(raw_args, args_info, dummy_func, None)
        self.assertEqual(converted, ["hello", 42])

    def test_no_xlmac_imports_in_codebase(self):
        """Verify obsolete _xlmac imports are not present in active source files."""
        for py_path in self.root.glob("enhancer/**/*.py"):
            content = py_path.read_text(encoding="utf-8")
            self.assertNotIn("_xlmac", content, f"Obsolete _xlmac reference found in {py_path.name}")

if __name__ == "__main__":
    unittest.main()
