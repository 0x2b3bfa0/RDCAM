# coding=utf-8
"""Clipping geometry and gradient/mask rasterising."""

import io
import math
import os
import re
import struct
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

import inkex  # noqa: E402

import rdlaser_geometry as geo  # noqa: E402
import rdlaser_send as send  # noqa: E402
from test_rdlaser import INKSCAPE, HandoffFormatTest, laserworks_decode_K  # noqa: E402


def csp(d):
    return [s for s in inkex.Path(d).to_superpath() if len(s) >= 2]


def square_clip(x0, y0, x1, y1, rule="nonzero"):
    return geo.ClipShape([geo.Region([[(x0, y0), (x1, y0), (x1, y1), (x0, y1)]], rule)])


def length(points):
    return sum(math.dist(a, b) for a, b in zip(points, points[1:]))


def polys(parts):
    return [p for p in parts if p[0] == "poly"]


class ClipGeometryTest(unittest.TestCase):
    def test_open_line_stays_open(self):
        parts = geo.clip_outline(csp("M 0,50 L 100,50"), [square_clip(25, 25, 75, 75)])
        self.assertEqual(len(parts), 1)
        kind, points, closed = parts[0]
        self.assertEqual((kind, closed), ("poly", False))
        self.assertAlmostEqual(points[0][0], 25)
        self.assertAlmostEqual(points[-1][0], 75)

    def test_curve_is_cut_at_the_clip_edge(self):
        parts = geo.clip_outline(csp("M 0,50 C 30,0 70,100 100,50"), [square_clip(25, 0, 75, 100)])
        (_, points, closed), = parts
        self.assertFalse(closed)
        self.assertAlmostEqual(points[0][0], 25, places=6)
        self.assertAlmostEqual(points[-1][0], 75, places=6)
        for x, y in points:  # the flattened curve stays on the Bezier
            self.assertTrue(24.99 <= x <= 75.01)

    def test_fully_inside_keeps_beziers_and_outside_is_dropped(self):
        inside = geo.clip_outline(csp("M 40,40 C 45,30 55,30 60,40"), [square_clip(0, 0, 100, 100)])
        self.assertEqual([p[0] for p in inside], ["bezier"])
        outside = geo.clip_outline(csp("M 140,40 L 160,40"), [square_clip(0, 0, 100, 100)])
        self.assertEqual(outside, [])

    def test_stroked_rectangle_keeps_only_visible_edges(self):
        parts = geo.clip_outline(csp("M 0,0 H 50 V 50 H 0 Z"), [square_clip(25, 25, 75, 75)])
        (_, points, closed), = parts
        self.assertFalse(closed)  # the clip edge is not cut: there is no fill
        self.assertAlmostEqual(length(points), 50)

    def test_filled_rectangle_gets_one_closed_outline(self):
        parts = geo.clip_outline(
            csp("M 0,0 H 50 V 50 H 0 Z"), [square_clip(25, 25, 75, 75)], filled=True
        )
        (_, points, closed), = parts
        self.assertTrue(closed)
        self.assertAlmostEqual(length(points), 100)
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        self.assertAlmostEqual(min(xs), 25)
        self.assertAlmostEqual(max(xs), 50)
        self.assertAlmostEqual(min(ys), 25)
        self.assertAlmostEqual(max(ys), 50)

    def test_coincident_edges_are_cut_once(self):
        parts = geo.clip_outline(
            csp("M 0,0 H 50 V 50 H 0 Z"), [square_clip(0, 0, 50, 50)], filled=True
        )
        total = sum(length(p[1]) for p in polys(parts)) + sum(
            inkex.CubicSuperPath([p[1]]).to_path().bounding_box().width * 4
            for p in parts
            if p[0] == "bezier"
        )
        self.assertAlmostEqual(total, 200)

    def test_partly_coincident_clip_closes_the_outline(self):
        # clip = left half, sharing three edges with the filled rectangle
        parts = geo.clip_outline(
            csp("M 0,0 H 20 V 10 H 0 Z"), [square_clip(0, 0, 10, 10)], filled=True
        )
        (_, points, closed), = parts
        self.assertTrue(closed)
        self.assertAlmostEqual(length(points), 40)

    def test_many_segments_is_fast(self):
        import time

        circle = csp("M 80,50 A 30,30 0 1 1 20,50 A 30,30 0 1 1 80,50 Z")  # crosses every spike
        star = [
            (50 + (45 if i % 2 else 20) * math.cos(i * math.pi / 1000),
             50 + (45 if i % 2 else 20) * math.sin(i * math.pi / 1000))
            for i in range(2000)
        ]
        clip = geo.ClipShape([geo.Region([star])])
        start = time.perf_counter()
        parts = geo.clip_outline(circle, [clip], filled=True, tolerance=0.001)
        elapsed = time.perf_counter() - start
        # every spike crosses the circle: the visible area is one jagged region
        self.assertEqual(len(parts), 1)  # one closed jagged outline
        self.assertTrue(parts[0][2])
        self.assertGreater(len(parts[0][1]), 2000)
        self.assertLess(elapsed, 10)

    def test_evenodd_clip_with_hole(self):
        clip = geo.ClipShape(
            [
                geo.Region(
                    [
                        [(0, 0), (100, 0), (100, 100), (0, 100)],
                        [(40, 40), (60, 40), (60, 60), (40, 60)],
                    ],
                    "evenodd",
                )
            ]
        )
        parts = geo.clip_outline(csp("M -10,50 L 110,50"), [clip])
        lengths = sorted(round(length(p[1]), 6) for p in parts)
        self.assertEqual(lengths, [40, 40])

    def test_nested_clips_intersect(self):
        parts = geo.clip_outline(
            csp("M 0,50 L 100,50"), [square_clip(10, 0, 70, 100), square_clip(30, 0, 90, 100)]
        )
        (_, points, _), = parts
        self.assertAlmostEqual(points[0][0], 30)
        self.assertAlmostEqual(points[-1][0], 70)

    def test_empty_clip_hides_everything(self):
        self.assertEqual(geo.clip_outline(csp("M 0,0 L 1,1"), [geo.ClipShape([])]), [])


CLIP_SVG = """<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
  width="100mm" height="100mm" viewBox="0 0 100 100">
 <defs>
  <clipPath id="cpUser" transform="translate(10,0)"><rect x="10" y="0" width="40" height="100"/></clipPath>
  <clipPath id="cpBox" clipPathUnits="objectBoundingBox"><rect x="0" y="0" width="0.5" height="1"/></clipPath>
  <linearGradient id="lg"><stop offset="0" style="stop-color:#000"/><stop offset="1" style="stop-color:#fff"/></linearGradient>
  <pattern id="pat" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="2" height="2" fill="#000"/></pattern>
  <mask id="mk"><rect x="0" y="80" width="100" height="20" fill="#fff"/></mask>
 </defs>
 <g clip-path="url(#cpUser)">
  <path id="cutline" d="M 0,10 H 100" style="fill:none;stroke:#ff0000"/>
 </g>
 <rect id="halfbox" x="60" y="20" width="20" height="10" clip-path="url(#cpBox)" style="fill:#0000ff;stroke:none"/>
 <rect id="grad" x="10" y="40" width="30" height="10" style="fill:url(#lg);stroke:#00ff00;stroke-width:0.5"/>
 <rect id="patt" x="50" y="40" width="20" height="10" style="fill:url(#pat);stroke:none"/>
 <g id="masked" mask="url(#mk)"><rect x="10" y="85" width="20" height="10" style="fill:#000000"/></g>
</svg>"""


@unittest.skipUnless(INKSCAPE and os.path.exists(INKSCAPE), "Inkscape binary not found")
class ClipEndToEndTest(HandoffFormatTest):
    def test_clips_gradients_patterns_masks(self):
        tmp = self.tmp()
        exe = os.path.join(tmp, "LaserWorkV5.exe")
        open(exe, "wb").close()
        svg = os.path.join(tmp, "in.svg")
        with open(svg, "w") as f:
            f.write(CLIP_SVG)
        os.environ["INKSCAPE_COMMAND"] = INKSCAPE
        send.SendToLaserWorks().run([f"--exe={exe}", "--launch=false", svg], output=io.BytesIO())
        temp = os.path.join(tmp, "temp")
        text = open(os.path.join(temp, "rd.ai"), encoding="ascii").read()

        blocks = re.findall(r"([\d.]+ [\d.]+ [\d.]+ [\d.]+) K\n(?:.*\n){4}((?:[-\d. ]+ [mLc]\n)+)([sS])", text)
        by_color = {}
        for cmyk, body, close in blocks:
            rgb = laserworks_decode_K(*map(float, cmyk.split()))
            pts = [tuple(float(v) / send.MM_TO_PT for v in line.split()[-3:-1]) for line in body.splitlines()]
            by_color.setdefault(rgb, []).append((pts, close))

        # Red cut line: clip is x 20..60 (rect x=10 + clip translate 10), y up = 100-10.
        (pts, close), = by_color[(255, 0, 0)]
        self.assertEqual(close, "S")
        self.assertAlmostEqual(pts[0][0], 20, places=3)
        self.assertAlmostEqual(pts[-1][0], 60, places=3)
        self.assertAlmostEqual(pts[0][1], 90, places=3)

        # Blue box clipped to the left half of its own bbox: x 60..70, y 70..80 (Y up)
        (pts, close), = by_color[(0, 0, 255)]
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        self.assertEqual(close, "s")
        self.assertAlmostEqual(min(xs), 60, places=3)
        self.assertAlmostEqual(max(xs), 70, places=3)
        self.assertAlmostEqual(min(ys), 70, places=3)
        self.assertAlmostEqual(max(ys), 80, places=3)

        # Gradient rect: solid green stroke stays a vector, the fill becomes a bitmap
        self.assertIn((0, 255, 0), by_color)
        self.assertEqual(set(by_color), {(255, 0, 0), (0, 0, 255), (0, 255, 0)})

        # Bitmaps: gradient fill, pattern fill, masked group (document order)
        data = open(os.path.join(temp, "dest.tmp"), "rb").read()
        self.assertEqual(struct.unpack_from("<hh", data), (0, 3))
        boxes = [struct.unpack_from("<4d", data, 4 + 32 * i) for i in range(3)]
        expected = [
            (30.0, 10.0, 25.0, 55.0),  # gradient fill only (stroke stripped)
            (20.0, 10.0, 60.0, 55.0),  # pattern
            (20.0, 10.0, 20.0, 10.0),  # masked group
        ]
        for got, want in zip(boxes, expected):
            for g, w in zip(got, want):
                self.assertAlmostEqual(g, w, places=2)

        from PIL import Image

        with Image.open(os.path.join(temp, "1.bmp")) as img:
            self.assertEqual(img.size, (354, 118))  # 30 x 10 mm at 300 dpi
            row = [img.getpixel((x, 59)) for x in (2, 176, 351)]
            self.assertLess(row[0], 30)  # black -> white gradient
            self.assertTrue(100 < row[1] < 160)
            self.assertGreater(row[2], 225)
        for name in ("2.bmp", "3.bmp"):
            self.assertTrue(os.path.isfile(os.path.join(temp, name)))


if __name__ == "__main__":
    unittest.main()
