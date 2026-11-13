#!/usr/bin/env python3
# coding=utf-8
#
# Send to LaserWorks V5 -- Inkscape port of the CorelDRAW macro RLaserV5.bas
#
# LaserWorks V5 (RDCAM) has no import API. Its CorelDRAW plug-in drops a few
# files into "<LaserWorks folder>/temp/" and then starts LaserWorkV5.exe. The
# hand-off protocol below was confirmed by disassembling LaserWorkV5.exe:
#
#   dest.tmp   int16 kind          0 = rd.ai + bitmaps, nonzero = rd.dxf only
#              int16 n_bitmaps     (only read when kind == 0)
#              n * float64[4]      width, height, centre x, centre y   (mm)
#   1.bmp..n.bmp                   grayscale bitmaps, matched to the list above
#   dest2.tmp  float64[4]          width, height, centre x, centre y of rd.ai
#   rd.ai      Illustrator (AI5-style PostScript) file with the vectors
#
# All coordinates are in millimetres with the Y axis pointing up and the
# origin in the bottom-left corner of the page (CorelDRAW conventions).
# LaserWorks scales and moves the imported rd.ai so that its bounding box
# matches dest2.tmp exactly, loads the bitmaps, then deletes the temp folder.
# If LaserWorks is already running, starting the exe again only broadcasts a
# window message that makes the running instance reload the temp folder.
"""Export the drawing to RDCAM LaserWorks V5 and start it."""

import math
import os
import shlex
import shutil
import struct
import subprocess
import sys
import tempfile

import inkex
from inkex import command
from inkex.units import convert_unit

import rdlaser_geometry as geometry

MM_TO_PT = 72.0 / 25.4
EPSILON = 1e-6
FLATTEN_TOLERANCE = 0.01  # mm, for clipped outlines only

# Files LaserWorks reads from its temp folder (see the header above).
HANDOFF_FILES = ("dest.tmp", "dest2.tmp", "rd.ai", "rd.dxf")

# Elements whose children are never drawn directly.
NON_RENDERED = {
    "defs",
    "clipPath",
    "mask",
    "marker",
    "pattern",
    "symbol",
    "metadata",
    "title",
    "desc",
}

# Environment variables set by the Inkscape app bundles that must not leak into
# Wine or LaserWorks (they point at Inkscape's private libraries).
ENV_BLOCKLIST_PREFIXES = ("DYLD_", "PYTHON", "GTK_", "GDK_", "GI_", "GIO_", "FONTCONFIG_")


# --------------------------------------------------------------------------- #
# Hand-off file writers (pure functions, unit tested)
# --------------------------------------------------------------------------- #


def write_dest(path, bitmaps):
    """Write dest.tmp: kind 0, bitmap count, then (w, h, cx, cy) per bitmap."""
    with open(path, "wb") as stream:
        stream.write(struct.pack("<hh", 0, len(bitmaps)))
        for box in bitmaps:
            stream.write(struct.pack("<4d", *box))


def write_dest2(path, box):
    """Write dest2.tmp: (w, h, cx, cy) of the vector artwork in rd.ai."""
    with open(path, "wb") as stream:
        stream.write(struct.pack("<4d", *box))


def ai_stroke_cmyk(rgb):
    """CMYK operands for the AI ``K`` operator that LaserWorks turns back into rgb.

    LaserWorks decodes ``c m y k K`` as ``trunc(255 * (1 - min(1, c + k)))`` per
    channel (truncation, not rounding), so aim at the middle of each integer.
    """
    return tuple(min(1.0, max(0.0, (254.5 - channel) / 255.0)) for channel in rgb) + (
        0.0,
    )


class AIWriter:
    """Minimal Illustrator writer using the same dialect LaserWorks writes itself.

    The structure mirrors ``_file_SaveAIHeader`` / ``_file_SaveCurveAsAI`` /
    ``_file_SaveAIEnd`` in LaserWorks' ExFileMgr.dll. Coordinates are points.
    """

    def __init__(self, stream):
        self.stream = stream

    def _w(self, text):
        self.stream.write(text.encode("ascii"))

    def header(self, bbox_pt):
        x0, y0, x1, y1 = bbox_pt
        self._w(
            "%!PS-Adobe-3.0\n"
            "%%Creator:Adobe Illustrator (TM) 7.000000 Exported from Inkscape\n"
            "%%Title:\n"
            f"%%BoundingBox:{x0:.0f} {y0:.0f} {x1:.0f} {y1:.0f}\n"
            f"%%HiResBoundingBox:{x0:.4f} {y0:.4f} {x1:.4f} {y1:.4f}\n"
            "%%DocumentProcessColors: Cyan Magenta Yellow Black\n"
            "%AI3_ColorUsage: Color\n"
            "%AI5_FileFormat 2.0\n"
            "%AI5_NumLayers: 1\n"
            "%%EndComments\n"
            "%%BeginProlog\n"
            "%%EndProlog\n"
            "%%BeginSetup\n"
            "%%EndSetup\n"
            "%AI5_BeginLayer\n"
            "1 1 1 1 0 0 0 0 0 0 Lb\n"
            "(Layer 1) Ln\n"
        )

    def path(self, rgb, subpath, closed):
        """Write one stroked subpath.

        ``subpath`` is a list of segments: ("M", (x, y)), ("L", (x, y)) or
        ("C", (x1, y1), (x2, y2), (x, y)), in points.
        """
        cyan, magenta, yellow, black = ai_stroke_cmyk(rgb)
        self._w(
            "1 XR\n0 R\n"
            f"{cyan:.4f} {magenta:.4f} {yellow:.4f} {black:.4f} K\n"
            "0 J\n0 j\n0.2160 w\n[]0 d\n"
        )
        for seg in subpath:
            if seg[0] == "M":
                self._w("{:.4f} {:.4f} m\n".format(*seg[1]))
            elif seg[0] == "L":
                self._w("{:.4f} {:.4f} L\n".format(*seg[1]))
            else:
                self._w(
                    "{:.4f} {:.4f} {:.4f} {:.4f} {:.4f} {:.4f} c\n".format(
                        *seg[1], *seg[2], *seg[3]
                    )
                )
        self._w("s\n" if closed else "S\n")

    def footer(self):
        self._w(
            "LB\n"
            "%AI5_EndLayer--\n"
            "%%PageTrailer\n"
            "gsave annotatepage grestore showpage\n"
            "%%Trailer\n"
            "%%EOF\n"
        )


# --------------------------------------------------------------------------- #
# Geometry helpers
# --------------------------------------------------------------------------- #


def localname(elem):
    tag = elem.tag
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def page_transform(svg):
    """Transform from SVG user units to millimetres, Y up, origin bottom-left."""
    vb_x, vb_y, vb_w, vb_h = svg.get_viewbox()
    width = svg.get("width")
    height = svg.get("height")
    if not vb_w or not vb_h:
        vb_x, vb_y = 0.0, 0.0
        vb_w = svg.to_dimensionless(width)
        vb_h = svg.to_dimensionless(height)
    width_mm = _to_mm(width, vb_w)
    height_mm = _to_mm(height, vb_h)
    sx = width_mm / vb_w
    sy = height_mm / vb_h
    return inkex.Transform(((sx, 0.0, -vb_x * sx), (0.0, -sy, (vb_y + vb_h) * sy)))


def page_origin(svg):
    """Top-left corner of the page in user units."""
    vb_x, vb_y, vb_w, vb_h = svg.get_viewbox()
    return (vb_x, vb_y) if vb_w and vb_h else (0.0, 0.0)


def _to_mm(length, fallback_px):
    if not length or str(length).strip().endswith("%"):
        return convert_unit(fallback_px, "mm")
    return convert_unit(length, "mm")


SOLID, COMPLEX, NONE = "solid", "complex", "none"


def paint(style, key):
    """Classify a fill/stroke: (SOLID, rgb), (COMPLEX, None) or (NONE, None).

    Gradients and patterns are COMPLEX: LaserWorks layers need one solid
    colour per path, so those parts are rendered to a bitmap instead.
    """
    try:
        value = style(key)
    except Exception:  # pylint: disable=broad-except
        value = None
    if value is None:
        raw = str(style.get(key, "none")).strip()
        return (COMPLEX, None) if raw.startswith("url(") else (NONE, None)
    if isinstance(value, inkex.Color):
        rgb = _color_tuple(value)
        return (SOLID, rgb) if rgb is not None else (NONE, None)
    return (COMPLEX, None)


def _color_tuple(color):
    try:
        rgb = color.to_rgb()
    except Exception:  # pylint: disable=broad-except
        return None
    return tuple(int(round(c)) for c in list(rgb)[:3])


def url_target(svg, value):
    """Element referenced by a 'url(#id)' value, or None."""
    value = str(value or "").strip()
    if not value.startswith("url("):
        return None
    ref = value[4:].split(")", 1)[0].strip().strip("'\"")
    return svg.getElementById(ref[1:]) if ref.startswith("#") else None


def own_style(node, key):
    """Value from the element itself (CSS, presentation attribute or style)."""
    return node.cascaded_style().get(key) or node.get(key)


def hidden(node):
    """display:none / visibility:hidden / opacity:0 on the element itself."""
    style = node.cascaded_style()
    if style.get("display", "inline") == "none":
        return True
    if style.get("visibility", "visible") in ("hidden", "collapse"):
        return True
    try:
        return float(style.get("opacity", 1.0)) == 0.0
    except ValueError:
        return False


def clip_shape(svg, node, clip_path, to_mm):
    """ClipShape (mm) for a <clipPath> referenced by ``node``."""
    transform = to_mm @ node.composed_transform()
    if clip_path.get("clipPathUnits") == "objectBoundingBox":
        # Reference box: the unclipped geometry in the node's own user space
        # (bounding_box() would already apply this very clip).
        box = node.shape_box(-node.transform)
        if box is None or not box.width or not box.height or math.isinf(box.width):
            return geometry.ClipShape([])
        transform = transform @ inkex.Transform(
            ((box.width, 0, box.left), (0, box.height, box.top))
        )
    transform = transform @ clip_path.transform
    regions = []
    for child in clip_path:
        if not isinstance(child, inkex.ShapeElement) or hidden(child):
            continue
        try:
            if isinstance(child, inkex.Use):
                path = child.path
                local = child.transform @ inkex.Transform(
                    translate=(
                        child.to_dimensionless(child.get("x", 0)),
                        child.to_dimensionless(child.get("y", 0)),
                    )
                )
            else:
                path = child.path
                local = child.transform
            path = path.to_absolute().transform(transform @ local)
        except (NotImplementedError, AttributeError, ValueError, TypeError):
            inkex.errormsg(f"Clip path '{clip_path.get_id()}': ignored an unsupported part.")
            continue
        rule = child.specified_style().get("clip-rule", "nonzero")
        rings = [
            geometry.flatten_subpath(sub, FLATTEN_TOLERANCE)
            for sub in path.to_superpath()
            if len(sub) >= 2
        ]
        regions.append(geometry.Region(rings, rule))
    return geometry.ClipShape(regions)


def parts_to_segments(parts):
    """Turn clip_outline() parts into AI segment lists, with their mm bbox."""
    out = []
    for part in parts:
        if part[0] == "bezier":
            sub = part[1]
            segments = csp_segments(sub)
            if segments:
                closed = _same(sub[0][1], sub[-1][1])
                box = inkex.CubicSuperPath([sub]).to_path().bounding_box()
                out.append((segments, closed, box))
        else:
            points, closed = part[1], part[2]
            if len(points) < 2:
                continue
            segments = [("M", _pt(points[0]))] + [("L", _pt(p)) for p in points[1:]]
            xs = [p[0] for p in points]
            ys = [p[1] for p in points]
            box = inkex.BoundingBox((min(xs), max(xs)), (min(ys), max(ys)))
            out.append((segments, closed, box))
    return out


def csp_segments(sub):
    """AI segments (points) of one CubicSuperPath subpath in mm."""
    segments = [("M", _pt(sub[0][1]))]
    for prev, cur in zip(sub, sub[1:]):
        p0, c1, c2, p3 = prev[1], prev[2], cur[0], cur[1]
        if _same(p0, c1) and _same(c2, p3):
            if not _same(p0, p3):
                segments.append(("L", _pt(p3)))
        else:
            segments.append(("C", _pt(c1), _pt(c2), _pt(p3)))
    return segments if len(segments) >= 2 else None


def _pt(point):
    return (point[0] * MM_TO_PT, point[1] * MM_TO_PT)


def _same(a, b):
    return abs(a[0] - b[0]) < EPSILON and abs(a[1] - b[1]) < EPSILON


def box_tuple(bbox):
    """(width, height, centre x, centre y) of an inkex BoundingBox."""
    return (bbox.width, bbox.height, bbox.center_x, bbox.center_y)


class VectorJob:  # pylint: disable=too-few-public-methods
    """A path sent to rd.ai."""

    def __init__(self, elem, rgb, filled, fill_rule, clips):
        self.elem, self.rgb, self.filled = elem, rgb, filled
        self.fill_rule, self.clips = fill_rule, clips


class RasterJob:  # pylint: disable=too-few-public-methods
    """An element rendered by Inkscape and sent as a bitmap."""

    def __init__(self, elem, native=False, strip=()):
        self.elem = elem
        self.native = native  # <image>: keep its own resolution
        self.strip = strip  # solid paints removed before rendering (sent as vectors)


# --------------------------------------------------------------------------- #
# The extension
# --------------------------------------------------------------------------- #


class SendToLaserWorks(inkex.EffectExtension):
    """Export the drawing to LaserWorks V5's temp folder and start LaserWorks."""

    def add_arguments(self, pars):
        pars.add_argument("--tab", default="send", choices=("send", "setup", "about"))
        pars.add_argument(
            "--scope", default="auto", choices=("auto", "selection", "all")
        )
        pars.add_argument("--bitmaps", type=inkex.Boolean, default=True)
        pars.add_argument("--bitmap_dpi", type=float, default=0.0)
        pars.add_argument("--raster_dpi", type=float, default=300.0)
        pars.add_argument("--launch", type=inkex.Boolean, default=True)
        pars.add_argument("--exe", default="")
        pars.add_argument("--wine", default="wine")
        pars.add_argument("--wineprefix", default="")

    # -- entry point -------------------------------------------------------- #

    def effect(self):
        exe = self.locate_exe()
        temp_dir = os.path.join(os.path.dirname(exe), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        clear_handoff(temp_dir)

        ids = self.scope_ids()
        if ids is not None and not ids:
            raise inkex.AbortExtension(
                "Nothing is selected. Select objects or set Scope to "
                "'Whole drawing'."
            )

        if self.has_drawable_content():
            self.export(temp_dir, ids)

        if self.options.launch:
            self.launch(exe)
        return False  # the document is not modified: no undo step

    # -- configuration ------------------------------------------------------ #

    def locate_exe(self):
        exe = os.path.expanduser(self.options.exe.strip().strip('"'))
        if not exe:
            raise inkex.AbortExtension(
                "Set the path to LaserWorkV5.exe in the 'Setup' tab first."
            )
        if os.path.isdir(exe):
            exe = os.path.join(exe, "LaserWorkV5.exe")
        if not os.path.isfile(exe):
            raise inkex.AbortExtension(f"LaserWorks executable not found:\n{exe}")
        return os.path.abspath(exe)

    def scope_ids(self):
        """Selected ids to export, or None for the whole drawing."""
        scope = self.options.scope
        selected = [elem.get_id() for elem in self.svg.selection.values()]
        if scope == "all" or (scope == "auto" and not selected):
            return None
        return set(selected)

    def has_drawable_content(self):
        for elem in self.svg.descendants():
            if isinstance(elem, inkex.ShapeElement) and not isinstance(
                elem, inkex.Group
            ):
                return True
        return False

    # -- export ------------------------------------------------------------- #

    def export(self, temp_dir, ids):
        with tempfile.TemporaryDirectory(prefix="rdlaser-") as work:
            svg = self.flatten_document()
            to_mm = page_transform(svg)
            vectors, rasters = self.collect(svg, ids)

            # Vectors first: rendering bitmaps edits the styles of their elements.
            if vectors:
                self.export_vectors(svg, vectors, to_mm, temp_dir)
            bitmap_boxes = []
            if self.options.bitmaps and rasters:
                bitmap_boxes = self.export_bitmaps(svg, rasters, to_mm, work, temp_dir)
            write_dest(os.path.join(temp_dir, "dest.tmp"), bitmap_boxes)

    def flatten_document(self):
        """Let Inkscape turn text, shapes and clones into plain paths.

        This replaces RLaserV5.bas's ``ConvertToCurves`` step; ids are kept.
        Text inside clip paths is converted too.
        """
        absolutize_image_links(self.svg)
        try:
            data = command.inkscape_command(
                self.svg,
                actions="select-all:all;object-unlink-clones;object-to-path;"
                "select-by-selector:clipPath > *;object-to-path",
            )
        except Exception as err:  # pylint: disable=broad-except
            inkex.errormsg(
                "Could not convert objects to paths with Inkscape; text and "
                f"clones may be missing.\n{err}"
            )
            return self.svg
        return inkex.load_svg(data).getroot()

    def collect(self, svg, ids):
        """Walk the drawing in z-order; return (VectorJobs, RasterJobs)."""
        vectors, rasters = [], []

        def walk(node, in_scope, under_mask):
            for elem in node:
                if not isinstance(elem, inkex.ShapeElement):
                    continue
                if localname(elem) in NON_RENDERED or hidden(elem):
                    continue
                selected = in_scope or ids is None or elem.get("id") in ids
                masked = url_target(svg, own_style(elem, "mask")) is not None
                if masked and selected:
                    # Masks are transparency, not geometry: render the whole
                    # masked object as Inkscape shows it.
                    rasters.append(RasterJob(elem))
                    continue
                if isinstance(elem, (inkex.Group, inkex.Anchor)):
                    walk(elem, selected, under_mask or masked)
                    continue
                if not selected or isinstance(elem, (inkex.TextElement, inkex.FlowRoot)):
                    continue  # (text that Inkscape could not convert is skipped)
                if isinstance(elem, inkex.Image) or under_mask:
                    rasters.append(RasterJob(elem, native=isinstance(elem, inkex.Image)))
                    continue
                self.classify_shape(svg, elem, vectors, rasters)

        walk(svg, False, False)
        return vectors, rasters

    def classify_shape(self, svg, elem, vectors, rasters):
        style = elem.specified_style()
        stroke_kind, stroke_rgb = paint(style, "stroke")
        fill_kind, fill_rgb = paint(style, "fill")

        # Gradient/pattern paint -> bitmap of just those parts.
        if COMPLEX in (stroke_kind, fill_kind):
            strip = tuple(
                key
                for key, kind in (("stroke", stroke_kind), ("fill", fill_kind))
                if kind != COMPLEX
            )
            rasters.append(RasterJob(elem, strip=strip))

        # Solid paint -> laser outline: stroke colour, else fill colour (as the
        # CorelDRAW macro did for text).
        rgb = stroke_rgb if stroke_kind == SOLID else fill_rgb if fill_kind == SOLID else None
        if rgb is None:
            return
        clips = []
        for node in [elem] + list(elem.ancestors()):
            if not isinstance(node, inkex.ShapeElement) or node is svg:
                continue
            target = url_target(svg, own_style(node, "clip-path"))
            if target is not None:
                clips.append((node, target))
        vectors.append(
            VectorJob(
                elem, rgb, fill_kind == SOLID, style.get("fill-rule", "nonzero"), clips
            )
        )

    def export_vectors(self, svg, vectors, to_mm, temp_dir):
        items = []
        bbox = None
        clip_cache = {}
        for job in vectors:
            try:
                path = job.elem.path.to_absolute().transform(
                    to_mm @ job.elem.composed_transform()
                )
            except (NotImplementedError, AttributeError, ValueError):
                continue
            subpaths = [sub for sub in path.to_superpath() if len(sub) >= 2]
            clips = []
            for node, target in job.clips:
                key = (id(node), id(target))
                if key not in clip_cache:
                    clip_cache[key] = clip_shape(svg, node, target, to_mm)
                clips.append(clip_cache[key])
            parts = geometry.clip_outline(
                subpaths, clips, job.filled, job.fill_rule, FLATTEN_TOLERANCE
            )
            for segments, closed, box in parts_to_segments(parts):
                if box is None:
                    continue
                bbox = box if bbox is None else bbox + box
                items.append((job.rgb, segments, closed))

        if not items:
            return
        ai_box = (
            bbox.left * MM_TO_PT,
            bbox.top * MM_TO_PT,  # inkex "top" is the minimum y, i.e. bottom here
            bbox.right * MM_TO_PT,
            bbox.bottom * MM_TO_PT,
        )
        with open(os.path.join(temp_dir, "rd.ai"), "wb") as stream:
            writer = AIWriter(stream)
            writer.header(ai_box)
            for rgb, segments, closed in items:
                writer.path(rgb, segments, closed)
            writer.footer()
        write_dest2(os.path.join(temp_dir, "dest2.tmp"), box_tuple(bbox))

    def export_bitmaps(self, svg, rasters, to_mm, work, temp_dir):
        """Render jobs with Inkscape, save as grayscale BMP. Return their boxes."""
        for job in rasters:
            job.elem.get_id()  # make sure every job can be addressed by id
            for key in job.strip:
                job.elem.style[key] = "none"
        source = os.path.join(work, "bitmaps.svg")
        command.write_svg(svg, source)

        # Inkscape's own (visual, clip-aware) bounding boxes: px at 96 dpi from
        # the page's top-left corner. Exports by id cover exactly these boxes.
        try:
            query = command.inkscape(source, query_all=True)
        except Exception as err:  # pylint: disable=broad-except
            inkex.errormsg(f"Bitmap rendering failed, bitmaps skipped.\n{err}")
            return []
        boxes_px = {}
        for line in query.splitlines():
            fields = line.strip().split(",")
            if len(fields) == 5:
                try:
                    boxes_px[fields[0]] = tuple(float(v) for v in fields[1:])
                except ValueError:
                    pass

        page_height = to_mm.apply_to_point(page_origin(svg)).y
        jobs = []
        for job in rasters:
            elem_id = job.elem.get("id")
            if elem_id not in boxes_px:
                continue
            x, y, w, h = (v * 25.4 / 96.0 for v in boxes_px[elem_id])
            if w <= 0 or h <= 0:
                continue
            box = (w, h, x + w / 2.0, page_height - (y + h / 2.0))
            if job.native:
                dpi = self.options.bitmap_dpi or native_dpi(job.elem, to_mm) or 300.0
            else:
                dpi = self.options.raster_dpi or 300.0
            size = (max(1, int(round(w / 25.4 * dpi))), max(1, int(round(h / 25.4 * dpi))))
            png = os.path.join(work, f"img{len(jobs) + 1}.png")
            jobs.append((elem_id, box, size, png))
        if not jobs:
            return []

        actions = [
            "export-type:png",
            "export-png-color-mode:Gray_8",
            "export-background:#ffffff",
            "export-background-opacity:1",
            "export-id-only",
        ]
        for elem_id, _, (width_px, height_px), png in jobs:
            actions += [
                f"export-id:{elem_id}",
                f"export-width:{width_px}",
                f"export-height:{height_px}",
                f"export-filename:{png}",
                "export-do",
            ]
        try:
            command.inkscape(source, actions=";".join(actions))
        except Exception as err:  # pylint: disable=broad-except
            inkex.errormsg(f"Bitmap rendering failed, bitmaps skipped.\n{err}")
            return []

        boxes = []
        for _, box, _, png in jobs:
            if not os.path.isfile(png):
                continue
            try:
                png_to_gray_bmp(png, os.path.join(temp_dir, f"{len(boxes) + 1}.bmp"))
            except (OSError, ValueError) as err:
                inkex.errormsg(f"Skipped a bitmap that could not be converted: {err}")
                continue
            boxes.append(box)
        return boxes


    # -- launching ---------------------------------------------------------- #

    def launch(self, exe):
        if sys.platform == "win32":
            argv = [exe]
        else:
            argv = shlex.split(self.options.wine.strip() or "wine")
            argv[0] = find_wine(argv[0])
            argv.append(exe)

        env = {
            key: value
            for key, value in os.environ.items()
            if not key.startswith(ENV_BLOCKLIST_PREFIXES)
        }
        if self.options.wineprefix.strip():
            env["WINEPREFIX"] = os.path.expanduser(self.options.wineprefix.strip())

        try:
            ensure_r5_dll(appdata_dir(argv[:-1], env))
        except OSError as err:
            raise inkex.AbortExtension(
                f"Could not write %AppData%\\{R5_DLL_NAME}:\n{err}"
            ) from err

        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = (
                subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
            )
        else:
            kwargs["start_new_session"] = True
        try:
            # Detach completely: Inkscape waits for our stdout/stderr to close.
            proc = subprocess.Popen(  # pylint: disable=consider-using-with
                argv,
                cwd=os.path.dirname(exe),
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
                **kwargs,
            )
            # Detached on purpose; tell Popen not to warn that it is still running.
            proc.returncode = 0
        except OSError as err:
            raise inkex.AbortExtension(
                f"Could not start LaserWorks ({' '.join(argv)}):\n{err}"
            ) from err


R5_DLL_NAME = "r5.dll"
R5_DLL_CONTENT = b"aaaaaaaaaa"  # exactly 10 bytes, no line ending


def ensure_r5_dll(appdata):
    """Make sure %AppData%\\r5.dll exists with exactly R5_DLL_CONTENT."""
    os.makedirs(appdata, exist_ok=True)
    target = os.path.join(appdata, R5_DLL_NAME)
    try:
        with open(target, "rb") as stream:
            if stream.read() == R5_DLL_CONTENT:
                return target
    except OSError:
        pass
    with open(target, "wb") as stream:  # binary mode: no newline translation
        stream.write(R5_DLL_CONTENT)
    return target


def appdata_dir(wine_argv, env):
    """Host path of %AppData% as LaserWorks sees it (native or inside Wine)."""
    if sys.platform == "win32":
        appdata = env.get("APPDATA") or os.path.join(
            os.path.expanduser("~"), "AppData", "Roaming"
        )
        return appdata

    prefix = env.get("WINEPREFIX") or os.path.join(os.path.expanduser("~"), ".wine")
    users = os.path.join(prefix, "drive_c", "users")
    user = env.get("USER") or env.get("LOGNAME") or ""
    for candidate in (
        os.path.join(users, user, "AppData", "Roaming"),  # Wine >= 5
        os.path.join(users, user, "Application Data"),  # older Wine
    ):
        if user and os.path.isdir(candidate):
            return candidate

    # Unusual layout or a fresh prefix: ask Wine itself.
    try:
        win_path = subprocess.run(
            wine_argv + ["cmd", "/c", "echo %APPDATA%"],
            env=env,
            capture_output=True,
            text=True,
            timeout=120,
            check=True,
        ).stdout.strip().splitlines()[-1].strip()
        winepath = os.path.join(os.path.dirname(wine_argv[0]), "winepath")
        if not os.access(winepath, os.X_OK):
            winepath = shutil.which("winepath") or "winepath"
        host_path = subprocess.run(
            [winepath, "-u", win_path],
            env=env,
            capture_output=True,
            text=True,
            timeout=120,
            check=True,
        ).stdout.strip()
        if host_path:
            return host_path
    except (OSError, subprocess.SubprocessError, IndexError):
        pass
    return os.path.join(users, user or "user", "AppData", "Roaming")


def clear_handoff(temp_dir):
    """Remove files left by a previous export so LaserWorks never loads stale data."""
    for name in os.listdir(temp_dir):
        stem, ext = os.path.splitext(name)
        if name.lower() in HANDOFF_FILES or (ext.lower() == ".bmp" and stem.isdigit()):
            try:
                os.remove(os.path.join(temp_dir, name))
            except OSError:
                pass


def image_href(elem):
    return elem.get("xlink:href") or elem.get("href") or ""


def absolutize_image_links(svg):
    """Make relative image links absolute: rendering happens in a temp folder."""
    doc_path = os.environ.get("DOCUMENT_PATH") or ""
    base = os.path.dirname(doc_path) or svg.get("sodipodi:docbase") or ""
    for elem in svg.xpath("//svg:image"):
        href = image_href(elem)
        if not href or href.startswith(("data:", "http:", "https:")):
            continue
        path = href[7:] if href.startswith("file://") else href
        if not os.path.isabs(path) and base:
            path = os.path.join(base, path)
        elem.set("xlink:href", os.path.abspath(path))
        elem.attrib.pop("href", None)


def native_dpi(elem, to_mm):
    """Resolution at which an <image> is shown, so it is exported pixel for pixel."""
    href = image_href(elem)
    try:
        if href.startswith("data:"):
            import base64  # pylint: disable=import-outside-toplevel

            head = base64.b64decode(href.split(",", 1)[1])
        else:
            path = href[7:] if href.startswith("file://") else href
            with open(path, "rb") as stream:
                head = stream.read(1 << 16)
        width_px = image_pixel_width(head)
        box = elem.path.to_absolute().transform(to_mm @ elem.composed_transform()).bounding_box()
    except (OSError, ValueError, IndexError, AttributeError):
        return None
    if not width_px or box is None or box.width <= 0:
        return None
    return width_px / (box.width / 25.4)


def image_pixel_width(head):
    """Pixel width from the first bytes of a PNG, JPEG, GIF, BMP or WebP file."""
    if head[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">I", head[16:20])[0]
    if head[:6] in (b"GIF87a", b"GIF89a"):
        return struct.unpack("<H", head[6:8])[0]
    if head[:2] == b"BM":
        return abs(struct.unpack("<i", head[18:22])[0])
    if head[:4] == b"RIFF" and head[8:12] == b"WEBP":
        chunk = head[12:16]
        if chunk == b"VP8 ":
            return struct.unpack("<H", head[26:28])[0] & 0x3FFF
        if chunk == b"VP8L":
            return (struct.unpack("<I", head[21:25])[0] & 0x3FFF) + 1
        if chunk == b"VP8X":
            return int.from_bytes(head[24:27], "little") + 1
    if head[:2] == b"\xff\xd8":
        pos = 2
        while pos + 9 < len(head):
            if head[pos] != 0xFF:
                pos += 1
                continue
            marker = head[pos + 1]
            if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7 or marker == 0xFF:
                pos += 1 if marker == 0xFF else 2
                continue
            length = struct.unpack(">H", head[pos + 2 : pos + 4])[0]
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                return struct.unpack(">H", head[pos + 7 : pos + 9])[0]
            pos += 2 + length
    return None


def png_to_gray_bmp(png_path, bmp_path):
    """Convert an 8-bit PNG (as exported by Inkscape) to an 8-bit grayscale BMP.

    Pure Python, because Pillow is missing or broken in some Inkscape bundles
    (e.g. the macOS app outside its own launcher).
    """
    width, height, rows = read_png_gray(png_path)
    write_gray_bmp(bmp_path, width, height, rows)


def read_png_gray(path):
    """Decode an 8-bit non-interlaced PNG into rows of grayscale bytes."""
    import zlib  # pylint: disable=import-outside-toplevel

    with open(path, "rb") as stream:
        data = stream.read()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG file")
    pos, idat, palette = 8, [], None
    width = height = depth = ctype = interlace = None
    while pos < len(data):
        length, kind = struct.unpack(">I4s", data[pos : pos + 8])
        body = data[pos + 8 : pos + 8 + length]
        pos += 12 + length
        if kind == b"IHDR":
            width, height, depth, ctype, _, _, interlace = struct.unpack(">IIBBBBB", body)
        elif kind == b"PLTE":
            palette = body
        elif kind == b"IDAT":
            idat.append(body)
        elif kind == b"IEND":
            break
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}.get(ctype)
    if depth != 8 or interlace or channels is None:
        raise ValueError(f"unsupported PNG (depth {depth}, type {ctype})")
    raw = zlib.decompress(b"".join(idat))
    stride = width * channels
    rows, prev = [], bytearray(stride)
    for y in range(height):
        start = y * (stride + 1)
        ftype = raw[start]
        line = bytearray(raw[start + 1 : start + 1 + stride])
        _unfilter(ftype, line, prev, channels)
        rows.append(_to_gray(line, ctype, channels, palette))
        prev = line
    return width, height, rows


def _unfilter(ftype, line, prev, bpp):
    if ftype == 0:
        return
    n = len(line)
    if ftype == 1:
        for i in range(bpp, n):
            line[i] = (line[i] + line[i - bpp]) & 0xFF
    elif ftype == 2:
        for i in range(n):
            line[i] = (line[i] + prev[i]) & 0xFF
    elif ftype == 3:
        for i in range(n):
            left = line[i - bpp] if i >= bpp else 0
            line[i] = (line[i] + ((left + prev[i]) >> 1)) & 0xFF
    elif ftype == 4:
        for i in range(n):
            a = line[i - bpp] if i >= bpp else 0
            b = prev[i]
            c = prev[i - bpp] if i >= bpp else 0
            p = a + b - c
            pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
            pred = a if pa <= pb and pa <= pc else (b if pb <= pc else c)
            line[i] = (line[i] + pred) & 0xFF
    else:
        raise ValueError(f"bad PNG filter {ftype}")


def _to_gray(line, ctype, channels, palette):
    if ctype == 0:
        return bytes(line)
    if ctype == 4:
        return bytes(line[0::2])
    if ctype == 3:
        palette = palette or b""
        rgb = [tuple(palette[3 * i : 3 * i + 3]) or (0, 0, 0) for i in range(256)]
        return bytes(_luma(*rgb[i]) if len(rgb[i]) == 3 else 0 for i in line)
    return bytes(
        _luma(line[i], line[i + 1], line[i + 2]) for i in range(0, len(line), channels)
    )


def _luma(red, green, blue):
    return (299 * red + 587 * green + 114 * blue + 500) // 1000


def write_gray_bmp(path, width, height, rows):
    """Write an 8-bit BMP with a grayscale palette (bottom-up rows)."""
    pad = (-width) % 4
    image_size = (width + pad) * height
    offset = 14 + 40 + 256 * 4
    with open(path, "wb") as stream:
        stream.write(struct.pack("<2sIHHI", b"BM", offset + image_size, 0, 0, offset))
        stream.write(
            struct.pack("<IiiHHIIiiII", 40, width, height, 1, 8, 0, image_size, 3780, 3780, 256, 256)
        )
        stream.write(b"".join(bytes((v, v, v, 0)) for v in range(256)))
        for row in reversed(rows):
            stream.write(row + b"\0" * pad)


def find_wine(name):
    """Resolve the Wine binary; GUI apps on macOS do not inherit the shell PATH."""
    if os.path.isabs(name):
        return name
    found = shutil.which(name)
    if found:
        return found
    candidates = [
        f"/opt/homebrew/bin/{name}",
        f"/usr/local/bin/{name}",
        f"/usr/bin/{name}",
        f"/Applications/Wine Stable.app/Contents/Resources/wine/bin/{name}",
        f"/Applications/Wine Devel.app/Contents/Resources/wine/bin/{name}",
        f"/Applications/Wine Staging.app/Contents/Resources/wine/bin/{name}",
    ]
    for candidate in candidates:
        if os.access(candidate, os.X_OK):
            return candidate
    raise inkex.AbortExtension(
        f"'{name}' was not found. Install Wine or set the full path to it in the "
        "'Setup' tab."
    )


if __name__ == "__main__":
    SendToLaserWorks().run()
