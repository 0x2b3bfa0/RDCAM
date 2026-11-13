#!/usr/bin/env python3
# coding=utf-8
#
# Embroidery import -- replaces "Import Dst/Dsb Data" (ImportDstFile) from
# RLaserV5.bas, which needed the Windows-only drIptBdrFe.dll.
#
# Record decoding follows pyembroidery's DstReader/DsbReader (MIT licence):
#   https://github.com/EmbroidePy/pyembroidery
# Both formats start with a 512-byte Tajima-style text header followed by
# 3-byte stitch records. Units are 0.1 mm.
"""Import Tajima .dst / Barudan .dsb embroidery files as stitch outlines."""

import inkex

HEADER_SIZE = 512
UNIT_MM = 0.1


def _bit(byte, pos):
    return (byte >> pos) & 1


def decode_dst(data):
    """Yield (kind, dx, dy) with dy pointing down; kind is stitch/jump/color."""
    for i in range(HEADER_SIZE, len(data) - 2, 3):
        b0, b1, b2 = data[i], data[i + 1], data[i + 2]
        dx = (
            81 * (_bit(b2, 2) - _bit(b2, 3))
            + 27 * (_bit(b1, 2) - _bit(b1, 3))
            + 9 * (_bit(b0, 2) - _bit(b0, 3))
            + 3 * (_bit(b1, 0) - _bit(b1, 1))
            + (_bit(b0, 0) - _bit(b0, 1))
        )
        dy_up = (
            81 * (_bit(b2, 5) - _bit(b2, 4))
            + 27 * (_bit(b1, 5) - _bit(b1, 4))
            + 9 * (_bit(b0, 5) - _bit(b0, 4))
            + 3 * (_bit(b1, 7) - _bit(b1, 6))
            + (_bit(b0, 7) - _bit(b0, 6))
        )
        if b2 & 0b11110011 == 0b11110011:
            return
        if b2 & 0b11000011 == 0b11000011:
            yield "color", dx, -dy_up
        elif b2 & 0b01000011 == 0b01000011:
            yield "jump", dx, -dy_up  # sequin mode toggle: no thread drawn
        elif b2 & 0b10000011 == 0b10000011:
            yield "jump", dx, -dy_up
        else:
            yield "stitch", dx, -dy_up


def decode_dsb(data):
    """Yield (kind, dx, dy) with dy pointing down for Barudan .dsb records."""
    for i in range(HEADER_SIZE, len(data) - 2, 3):
        ctrl, by, bx = data[i], data[i + 1], data[i + 2]
        dy = -by if ctrl & 0x40 == 0 else by
        dx = -bx if ctrl & 0x20 else bx
        low = ctrl & 0b11111
        if low == 0:
            yield "stitch", dx, dy
        elif low == 1:
            yield "jump", dx, dy
        elif ctrl == 0xF8:
            return
        elif ctrl in (0xE7, 0xE8):  # trim, stop
            yield "jump", 0, 0
        elif 0xE9 <= ctrl < 0xF8:  # needle change
            yield "color", 0, 0
        else:
            return


def stitch_blocks(records):
    """Group absolute stitch points into polylines: [[color_index, [pts...]]...]."""
    x = y = 0
    color = 0
    blocks = []
    current = None
    for kind, dx, dy in records:
        x += dx
        y += dy
        if kind == "stitch":
            if current is None:
                current = [color, [(x - dx, y - dy)]]
                blocks.append(current)
            current[1].append((x, y))
        else:
            current = None
            if kind == "color":
                color += 1
    return [block for block in blocks if len(block[1]) >= 2]


def header_label(data):
    """Design name from the 'LA:' header field, if any."""
    text = data[:HEADER_SIZE].split(b"\x1a")[0]
    for line in text.replace(b"\r", b"\n").split(b"\n"):
        if line.startswith(b"LA:"):
            return line[3:].decode("latin-1").strip()
    return ""


def build_svg(blocks, title=""):
    """SVG document tree (1 user unit = 1 mm), one group per colour block."""
    points = [p for _, pts in blocks for p in pts] or [(0, 0)]
    min_x = min(p[0] for p in points)
    min_y = min(p[1] for p in points)
    max_x = max(p[0] for p in points)
    max_y = max(p[1] for p in points)
    margin = 50  # 5 mm
    width = (max_x - min_x + 2 * margin) * UNIT_MM
    height = (max_y - min_y + 2 * margin) * UNIT_MM

    document = inkex.InputExtension.get_template(
        width=f"{width:.3f}", height=f"{height:.3f}", unit="mm"
    )
    svg = document.getroot()
    layer = svg.add(inkex.Layer.new(title or "Embroidery"))
    groups = {}
    for color, pts in blocks:
        if color not in groups:
            groups[color] = layer.add(inkex.Group(id=f"thread{color + 1}"))
            groups[color].label = f"Thread {color + 1}"
        coords = " ".join(
            f"{(px - min_x + margin) * UNIT_MM:.2f},{(py - min_y + margin) * UNIT_MM:.2f}"
            for px, py in pts
        )
        elem = groups[color].add(inkex.PathElement())
        elem.set("d", "M " + coords)
        # Same look as the CorelDRAW macro: black hairline, no fill.
        elem.style = {"fill": "none", "stroke": "#000000", "stroke-width": "0.1"}
    return document


class EmbroideryInput(inkex.InputExtension):
    """Open a .dst or .dsb embroidery file."""

    multi_inx = True  # shared by rdlaser_dst_input.inx and rdlaser_dsb_input.inx

    def load(self, stream):
        data = stream.read()
        if len(data) < HEADER_SIZE:
            raise inkex.AbortExtension("File is too short to be an embroidery file.")
        name = str(getattr(stream, "name", "") or self.options.input_file or "")
        decoder = decode_dsb if name.lower().endswith(".dsb") else decode_dst
        blocks = stitch_blocks(decoder(data))
        if not blocks:
            raise inkex.AbortExtension("No stitches found in this file.")
        return build_svg(blocks, header_label(data))


if __name__ == "__main__":
    EmbroideryInput().run()
