import importlib.util
import math
from pathlib import Path
import sys
import unittest


MODULE_PATH = Path(__file__).with_name("SEM_Position_Relocation_Calculator.py")
SPEC = importlib.util.spec_from_file_location("relocation_calculator", MODULE_PATH)
calculator = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = calculator
SPEC.loader.exec_module(calculator)


class TransformTests(unittest.TestCase):
    def test_translation_and_ninety_degree_rotation(self):
        transform = calculator.calculate_transform(
            calculator.Point(0, 0), calculator.Point(10, 0),
            calculator.Point(100, 200), calculator.Point(100, 210),
        )
        result = transform.apply(calculator.Point(4, 3))
        self.assertAlmostEqual(result.x, 97)
        self.assertAlmostEqual(result.y, 204)
        self.assertAlmostEqual(transform.rotation_degrees, 90)
        self.assertEqual(transform.scale, 1)

    def test_optional_uniform_scale(self):
        transform = calculator.calculate_transform(
            calculator.Point(0, 0), calculator.Point(10, 0),
            calculator.Point(5, 5), calculator.Point(25, 5),
            use_scale=True,
        )
        result = transform.apply(calculator.Point(2, 3))
        self.assertAlmostEqual(result.x, 9)
        self.assertAlmostEqual(result.y, 11)
        self.assertAlmostEqual(transform.scale, 2)

    def test_zero_length_reference_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Old reference points"):
            calculator.calculate_transform(
                calculator.Point(1, 1), calculator.Point(1, 1),
                calculator.Point(2, 2), calculator.Point(3, 3),
            )

    def test_reference_b_maps_exactly_in_rigid_mode_when_distances_match(self):
        old_a = calculator.Point(-2.5, 8.5)
        old_b = calculator.Point(4.0, -3.0)
        new_a = calculator.Point(20.0, 30.0)
        angle = math.radians(-17)
        dx, dy = old_b.x - old_a.x, old_b.y - old_a.y
        new_b = calculator.Point(
            new_a.x + dx * math.cos(angle) - dy * math.sin(angle),
            new_a.y + dx * math.sin(angle) + dy * math.cos(angle),
        )
        result = calculator.calculate_transform(old_a, old_b, new_a, new_b).apply(old_b)
        self.assertAlmostEqual(result.x, new_b.x)
        self.assertAlmostEqual(result.y, new_b.y)


if __name__ == "__main__":
    unittest.main()
