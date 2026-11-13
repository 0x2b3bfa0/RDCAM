# coding=utf-8
"""Tests for the LaserWorks V5 Inkscape extension.

Run with a Python that has inkex >= 1.2 (e.g. ``pip install inkex pytest pillow``)::

    INKSCAPE_COMMAND=/path/to/inkscape pytest tests

Tests that need the Inkscape binary are skipped when it is not available.
"""

import base64
import io
import os
import re
import shutil
import struct
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import inkex  # noqa: E402
from inkex.tester.inx import InxFile, InxMixin  # noqa: E402

import rdlaser_embroidery as emb  # noqa: E402
import rdlaser_send as send  # noqa: E402

ROOT = os.path.dirname(HERE)


def laserworks_decode_K(c, m, y, k):
    """LaserWorks' own decoding of 'c m y k K' (ExFileMgr.dll, 0x1002a063)."""

    def chan(v):
        return int((1.0 - min(1.0, v + k)) * 255.0)  # truncation like _ftol

    return chan(c), chan(m), chan(y)


class HandoffFormatTest(unittest.TestCase):
    def test_cmyk_roundtrip_through_laserworks_truncation(self):
        for v in range(256):
            ops = send.ai_stroke_cmyk((v, 255 - v, v // 2))
            ops = tuple(float(f"{o:.4f}") for o in ops)  # as written in the file
            self.assertEqual(laserworks_decode_K(*ops), (v, 255 - v, v // 2))

    def test_dest_layout(self):
        path = os.path.join(self.tmp(), "dest.tmp")
        send.write_dest(path, [(10.0, 20.0, 30.0, 40.0), (1, 2, 3, 4)])
        data = open(path, "rb").read()
        self.assertEqual(len(data), 4 + 2 * 32)
        self.assertEqual(struct.unpack_from("<hh", data), (0, 2))
        self.assertEqual(struct.unpack_from("<4d", data, 4), (10, 20, 30, 40))

        send.write_dest(path, [])
        self.assertEqual(open(path, "rb").read(), b"\0\0\0\0")

    def test_dest2_layout(self):
        path = os.path.join(self.tmp(), "dest2.tmp")
        send.write_dest2(path, (1.5, 2.5, 3.5, 4.5))
        self.assertEqual(struct.unpack("<4d", open(path, "rb").read()), (1.5, 2.5, 3.5, 4.5))

    def test_clear_handoff_only_removes_handoff_files(self):
        tmp = self.tmp()
        for name in ("dest.tmp", "DEST2.TMP", "rd.ai", "rd.dxf", "1.bmp", "12.bmp", "logo.bmp", "keep.txt"):
            open(os.path.join(tmp, name), "w").close()
        send.clear_handoff(tmp)
        self.assertEqual(sorted(os.listdir(tmp)), ["keep.txt", "logo.bmp"])

    def tmp(self):
        path = os.path.join(HERE, "_tmp", self.id().rsplit(".", 1)[-1])
        shutil.rmtree(path, ignore_errors=True)
        os.makedirs(path)
        self.addCleanup(shutil.rmtree, path, True)
        return path


class R5DllTest(HandoffFormatTest):
    def test_written_exactly(self):
        tmp = self.tmp()
        appdata = os.path.join(tmp, "AppData", "Roaming")
        target = send.ensure_r5_dll(appdata)
        self.assertEqual(target, os.path.join(appdata, "r5.dll"))
        self.assertEqual(open(target, "rb").read(), b"aaaaaaaaaa")

    def test_wrong_content_is_replaced(self):
        tmp = self.tmp()
        with open(os.path.join(tmp, "r5.dll"), "wb") as f:
            f.write(b"aaaaaaaaaa\r\n")
        send.ensure_r5_dll(tmp)
        self.assertEqual(open(os.path.join(tmp, "r5.dll"), "rb").read(), b"aaaaaaaaaa")

    @unittest.skipIf(sys.platform == "win32", "Wine layout only")
    def test_wine_appdata_from_prefix(self):
        tmp = self.tmp()
        roaming = os.path.join(tmp, "drive_c", "users", "alice", "AppData", "Roaming")
        os.makedirs(roaming)
        env = {"WINEPREFIX": tmp, "USER": "alice"}
        self.assertEqual(send.appdata_dir(["wine"], env), roaming)


class BitmapTest(HandoffFormatTest):
    def test_png_to_bmp_matches_pillow(self):
        import random

        from PIL import Image

        rnd = random.Random(1)
        tmp = self.tmp()
        for mode in ("L", "LA", "RGB", "RGBA", "P"):
            for size in ((1, 1), (7, 5), (33, 17)):
                img = Image.new("RGB", size)
                img.putdata([tuple(rnd.randrange(256) for _ in range(3)) for _ in range(size[0] * size[1])])
                img = img.convert(mode)
                png = os.path.join(tmp, f"{mode}.png")
                img.save(png, optimize=False)
                bmp = os.path.join(tmp, f"{mode}.bmp")
                send.png_to_gray_bmp(png, bmp)
                with Image.open(bmp) as got:
                    self.assertEqual(got.size, size)
                    want = img.convert("RGB").convert("L")
                    diffs = [abs(a - b) for a, b in zip(got.convert("L").getdata(), want.getdata())]
                    self.assertLessEqual(max(diffs), 1, (mode, size))

    def test_image_pixel_width(self):
        from PIL import Image

        for fmt in ("PNG", "JPEG", "GIF", "BMP", "WEBP"):
            buf = io.BytesIO()
            Image.new("RGB", (123, 45)).save(buf, fmt)
            self.assertEqual(send.image_pixel_width(buf.getvalue()), 123, fmt)


class AIWriterTest(unittest.TestCase):
    def test_structure(self):
        out = io.BytesIO()
        writer = send.AIWriter(out)
        writer.header((0, 0, 100, 50))
        writer.path((255, 0, 0), [("M", (0, 0)), ("L", (10, 0)), ("C", (1, 2), (3, 4), (5, 6))], True)
        writer.path((0, 0, 0), [("M", (0, 0)), ("L", (10, 0))], False)
        writer.footer()
        text = out.getvalue().decode("ascii")
        self.assertTrue(text.startswith("%!PS-Adobe-3.0\n"))
        self.assertIn("0.0000 0.9980 0.9980 0.0000 K\n", text)
        self.assertIn("1.0000 2.0000 3.0000 4.0000 5.0000 6.0000 c\ns\n", text)
        self.assertIn("10.0000 0.0000 L\nS\n", text)
        self.assertTrue(text.endswith("%%EOF\n"))


class InxTest(InxMixin, unittest.TestCase):
    def test_inx_files(self):
        for name in sorted(os.listdir(ROOT)):
            if name.endswith(".inx"):
                with self.subTest(inx=name):
                    inx = InxFile(os.path.join(ROOT, name))
                    self.assertInxSchemaValid(inx.filename)
                    self.assertTrue(inx.extension_class, "script/class not found")
                    self.assertParams(inx, inx.extension_class)


INKSCAPE = os.environ.get("INKSCAPE_COMMAND") or shutil.which("inkscape")

def _png_4x2():
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (4, 2), (255, 128, 0)).save(buf, "PNG")
    return base64.b64encode(buf.getvalue()).decode()


PNG_4x2 = _png_4x2()

TEST_SVG = f"""<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:inkscape="http://www.inkscape.org/namespaces/inkscape"
  width="200mm" height="100mm" viewBox="0 0 400 200">
 <g id="layer1" inkscape:groupmode="layer" transform="translate(20,10)">
  <rect id="r1" x="0" y="0" width="40" height="20" style="fill:none;stroke:#ff0000;stroke-width:1"/>
  <text id="t1" x="100" y="40" style="font-size:20px;font-family:sans-serif;fill:#0000ff">Hi</text>
  <use id="u1" xlink:href="#r1" transform="translate(0,100)"/>
  <circle id="c1" cx="300" cy="100" r="10" style="fill:#00ff00;stroke:none"/>
  <path id="nopaint" d="M 0,0 L 50,50" style="fill:none;stroke:none"/>
  <image id="img1" x="200" y="0" width="80" height="40" xlink:href="data:image/png;base64,{PNG_4x2}"/>
 </g>
 <g id="hidden" inkscape:groupmode="layer" style="display:none">
  <rect id="r2" x="0" y="0" width="10" height="10" style="stroke:#000000"/>
 </g>
</svg>"""


@unittest.skipUnless(INKSCAPE and os.path.exists(INKSCAPE), "Inkscape binary not found")
class EndToEndTest(HandoffFormatTest):
    def run_ext(self, *args):
        tmp = self.tmp()
        exe = os.path.join(tmp, "LaserWorkV5.exe")
        open(exe, "wb").close()
        svg = os.path.join(tmp, "in.svg")
        with open(svg, "w") as f:
            f.write(TEST_SVG)
        os.environ["INKSCAPE_COMMAND"] = INKSCAPE
        out = io.BytesIO()
        send.SendToLaserWorks().run([f"--exe={exe}", "--launch=false", *args, svg], output=out)
        self.assertEqual(out.getvalue(), b"", "document must not be modified")
        return os.path.join(tmp, "temp")

    def ai_paths(self, temp):
        text = open(os.path.join(temp, "rd.ai"), encoding="ascii").read()
        colors = re.findall(r"^([\d.]+ [\d.]+ [\d.]+ [\d.]+) K$", text, re.M)
        return text, [laserworks_decode_K(*map(float, c.split())) for c in colors]

    def test_whole_drawing(self):
        temp = self.run_ext()
        self.assertEqual(
            sorted(os.listdir(temp)), ["1.bmp", "dest.tmp", "dest2.tmp", "rd.ai"]
        )
        text, colors = self.ai_paths(temp)
        # rect + clone (red), "Hi" glyph outlines (blue), circle (green)
        self.assertEqual(colors.count((255, 0, 0)), 2)
        self.assertGreaterEqual(colors.count((0, 0, 255)), 2)
        self.assertEqual(colors.count((0, 255, 0)), 1)
        self.assertNotIn((0, 0, 0), colors)  # hidden layer / unpainted path skipped

        # Vector bbox in mm, Y up from the page bottom (page = 200 x 100 mm,
        # 1 user unit = 0.5 mm). Leftmost: rect at x=20uu -> 10 mm. Rightmost:
        # circle edge at 20+310=330uu -> 165 mm. Top: rect at y=10uu ->
        # 100-5 = 95 mm. Bottom: clone bottom at 10+120=130uu -> 100-65 = 35 mm.
        w, h, cx, cy = struct.unpack("<4d", open(os.path.join(temp, "dest2.tmp"), "rb").read())
        self.assertAlmostEqual(w, 155.0, places=3)
        self.assertAlmostEqual(h, 60.0, places=3)
        self.assertAlmostEqual(cx, 87.5, places=3)
        self.assertAlmostEqual(cy, 65.0, places=3)

        # Image: x 220..300uu -> 110..150 mm, y 10..50uu -> 95..75 mm (Y up)
        data = open(os.path.join(temp, "dest.tmp"), "rb").read()
        self.assertEqual(struct.unpack_from("<hh", data), (0, 1))
        box = struct.unpack_from("<4d", data, 4)
        for got, want in zip(box, (40.0, 20.0, 130.0, 85.0)):
            self.assertAlmostEqual(got, want, places=3)

        from PIL import Image

        with Image.open(os.path.join(temp, "1.bmp")) as img:
            self.assertEqual(img.format, "BMP")
            self.assertEqual(img.mode, "L")
            self.assertEqual(img.size, (4, 2))  # exported at native resolution
            # uniform orange (255, 128, 0) -> one gray level (Inkscape uses Rec.709 luma: 146)
            self.assertEqual(set(img.getdata()), {146})

        # AI coordinates are points: rect corner (10 mm, 95 mm)
        self.assertIn(f"{10 * send.MM_TO_PT:.4f} {95 * send.MM_TO_PT:.4f} m", text)

    def test_selection_only(self):
        temp = self.run_ext("--id=c1", "--scope=auto")
        self.assertEqual(sorted(os.listdir(temp)), ["dest.tmp", "dest2.tmp", "rd.ai"])
        _, colors = self.ai_paths(temp)
        self.assertEqual(colors, [(0, 255, 0)])
        w, h, cx, cy = struct.unpack("<4d", open(os.path.join(temp, "dest2.tmp"), "rb").read())
        self.assertAlmostEqual(w, 10.0, places=3)
        self.assertAlmostEqual(h, 10.0, places=3)
        self.assertAlmostEqual(cx, 160.0, places=3)
        self.assertAlmostEqual(cy, 45.0, places=3)
        self.assertEqual(open(os.path.join(temp, "dest.tmp"), "rb").read(), b"\0\0\0\0")

    def test_no_bitmaps_option(self):
        temp = self.run_ext("--bitmaps=false")
        self.assertNotIn("1.bmp", os.listdir(temp))
        self.assertEqual(open(os.path.join(temp, "dest.tmp"), "rb").read(), b"\0\0\0\0")


def encode_dst(dx, dy, flags=0b00000011):
    """Tiny DST record encoder (inverse of the reader) for |dx|,|dy| <= 121."""

    def ternary(v):
        digits = []
        for _ in range(5):
            r = v % 3
            v //= 3
            if r == 2:
                r = -1
                v += 1
            digits.append(r)
        return digits  # weights 1, 3, 9, 27, 81

    b = [0, 0, flags]
    xbits = {0: (0, 0, 1), 1: (1, 0, 1), 2: (0, 2, 3), 3: (1, 2, 3), 4: (2, 2, 3)}
    ybits = {0: (0, 7, 6), 1: (1, 7, 6), 2: (0, 5, 4), 3: (1, 5, 4), 4: (2, 5, 4)}
    for i, d in enumerate(ternary(dx)):
        byte, plus, minus = xbits[i]
        if d:
            b[byte] |= 1 << (plus if d > 0 else minus)
    for i, d in enumerate(ternary(-dy)):  # file stores Y up
        byte, plus, minus = ybits[i]
        if d:
            b[byte] |= 1 << (plus if d > 0 else minus)
    return bytes(b)


class EmbroideryTest(unittest.TestCase):
    def test_dst_decoder_roundtrip(self):
        for dx in range(-121, 122, 7):
            for dy in range(-121, 122, 11):
                rec = encode_dst(dx, dy)
                got = list(emb.decode_dst(b"\0" * 512 + rec))
                self.assertEqual(got, [("stitch", dx, dy)])

    def test_dst_blocks_and_import(self):
        header = b"LA:TEST DESIGN     \r" + b" " * 492
        body = (
            encode_dst(10, 0) + encode_dst(0, 10)  # two stitches
            + encode_dst(50, 50, 0b10000011)  # jump
            + encode_dst(10, 0)  # stitch
            + encode_dst(0, 0, 0b11000011)  # colour change
            + encode_dst(0, -10)
            + bytes([0, 0, 0xF3])  # end
        )
        blocks = emb.stitch_blocks(emb.decode_dst(header + body))
        self.assertEqual(
            blocks,
            [
                [0, [(0, 0), (10, 0), (10, 10)]],
                [0, [(60, 60), (70, 60)]],
                [1, [(70, 60), (70, 50)]],
            ],
        )
        out = io.BytesIO()
        emb.EmbroideryInput().run([self.write(header + body)], output=out)
        svg = inkex.load_svg(out.getvalue()).getroot()
        paths = svg.xpath("//svg:path")
        self.assertEqual(len(paths), 3)
        self.assertEqual(svg.xpath("//svg:g[@inkscape:groupmode='layer']")[0].label, "TEST DESIGN")
        # 1 unit = 0.1 mm, 5 mm margin
        self.assertEqual(str(paths[0].path), "M 5 5 L 6 5 L 6 6")

    def test_dsb(self):
        body = bytes([0x00, 10, 20, 0x60, 5, 5, 0x01, 0, 0, 0xE9, 0, 0, 0x00, 1, 1, 0xF8, 0, 0])
        recs = list(emb.decode_dsb(b"\0" * 512 + body))
        self.assertEqual(
            recs,
            [("stitch", 20, -10), ("stitch", -5, 5), ("jump", 0, 0), ("color", 0, 0), ("stitch", 1, -1)],
        )

    def write(self, data):
        path = os.path.join(HERE, "_tmp_test.dst")
        with open(path, "wb") as f:
            f.write(data)
        self.addCleanup(os.remove, path)
        return path


if __name__ == "__main__":
    unittest.main()
