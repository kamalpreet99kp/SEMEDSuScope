"""MM4 coordinate regressions; all generated files stay in temporary folders."""

import csv
import math
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import CreateTSVfromPycroOutput_App_MM4 as crops
import CorrectSEMPositionsFromTemplateV3_MM4 as correction


class MM4CoordinatesTests(unittest.TestCase):
    def test_roi_center_and_ordered_bounds(self):
        record = crops._parse_crop_image(
            "Pos1_1_100_200_F1_Roi10_20_30_60.jpg",
            crops.PycroTsvSettings(um_per_pixel=0.5),
        )
        self.assertEqual((record.x, record.y), (110, 180))
        self.assertEqual(record.bbox, (105, 170, 115, 190))

    def test_negative_roi_and_stage_coordinates(self):
        record = crops._parse_crop_image(
            "Pos1_1_-100_-200_F1_Roi-30_-60_-10_-20.jpeg",
            crops.PycroTsvSettings(um_per_pixel=0.5),
        )
        self.assertEqual((record.x, record.y), (-110, -180))
        self.assertEqual(record.bbox, (-115, -190, -105, -170))

    def test_rotation_matches_independent_cartesian_reference(self):
        a, b = (3, -7), (18, 4)
        for degrees in (0, 90, -90, 31, 180):
            for units in (1, 1000):
                with self.subTest(degrees=degrees, unit_factor=units):
                    angle, scale = math.radians(degrees), 2.5

                    def expected(point):
                        x, y = point[0] - a[0], point[1] - a[1]
                        return (
                            (100 + scale * (x * math.cos(angle) - y * math.sin(angle))) / units,
                            (-80 + scale * (x * math.sin(angle) + y * math.cos(angle))) / units,
                        )

                    transform = correction.make_transform(a, b, expected(a), expected(b), units)
                    # Include points sharing just one axis with B: these must rotate.
                    for point in (a, b, (18, 8), (23, 4), (-5, 21)):
                        actual = transform(point)
                        target = expected(point)
                        self.assertAlmostEqual(actual[0], target[0], places=10)
                        self.assertAlmostEqual(actual[1], target[1], places=10)

    def test_rotation_preserves_orientation(self):
        transform = correction.make_transform((0, 0), (10, 0), (100, 200), (100, 220), 1)
        self.assertAlmostEqual(transform((4, 3))[0], 94)
        self.assertAlmostEqual(transform((4, 3))[1], 208)
        a, b, c = (transform(p) for p in ((0, 0), (1, 0), (0, 1)))
        signed_area = (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])
        self.assertGreater(signed_area, 0)

    def test_degenerate_fiducials_still_rejected(self):
        with self.assertRaisesRegex(ValueError, "Microscope"):
            correction.make_transform((0, 0), (0, 0), (1, 2), (3, 4), 1)
        with self.assertRaisesRegex(ValueError, "SEM"):
            correction.make_transform((0, 0), (1, 0), (1, 2), (1, 2), 1)

    def test_all_project_lists_and_csv_correction_leave_legacy_outputs_untouched(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy = {}
            for folder in ("mark", "grains", "Mineral Check"):
                subdir = root / folder
                subdir.mkdir()
                (subdir / "Pos1_1_100_200_F1_Roi10_20_30_60.jpg").write_bytes(b"fixture")
                for name in ("list.csv", "list_sem_corrected.csv"):
                    path = subdir / name
                    path.write_bytes(b"legacy output")
                    legacy[path] = path.read_bytes()
                (subdir / "image_list").mkdir()
                path = subdir / "image_list" / "001.jpg"
                path.write_bytes(b"legacy image")
                legacy[path] = path.read_bytes()

            for _ in range(2):  # Repeated runs must use only the MM4 output paths.
                result = crops.generate_mark_and_grain_lists(root, crops.PycroTsvSettings(0.5))
                for key in ("mark", "grains", "mineral_check"):
                    data = result[key]
                    self.assertEqual(data["csv"].name, "list_mm4.csv")
                    self.assertEqual(data["image_list"].name, "image_list_mm4")
                    output = correction.default_output_path(data["csv"])
                    count, _ = correction.write_corrected_csv(
                        data["csv"], output, (100, 200), (110, 200),
                        (1, 2), (1, 2.02), "mm", 5,
                    )
                    self.assertEqual(count, 1)
                    with output.open(newline="") as stream:
                        row = next(csv.DictReader(stream))
                    self.assertEqual((row["X"], row["Y"]), ("1.04000", "2.02000"))
                    self.assertEqual((row["Microscope_X"], row["Microscope_Y"]), ("110.0", "180.0"))
                    self.assertIn("Roi10_20_30_60.jpg", row["Image Paths"])
            for path, content in legacy.items():
                self.assertEqual(path.read_bytes(), content)

    def test_csv_refuses_input_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.csv"
            path.write_text("X,Y\n4,3\n")
            before = path.read_bytes()
            with self.assertRaisesRegex(ValueError, "overwrite"):
                correction.write_corrected_csv(path, path, (0, 0), (10, 0), (0, 0), (10, 0), "um", 2)
            self.assertEqual(path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
