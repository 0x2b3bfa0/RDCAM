# coding=utf-8
"""Clipping of laser outlines against SVG clip paths, in pure Python.

Inkscape's own boolean operations (Path > Intersection, Flatten, ...) treat every
path as a filled area: an open cut line comes back as a closed sliver, and a
stroke comes back as a filled band, which doubles every cut. Laser outlines need
the opposite: keep lines as lines and only drop the parts the clip hides.

Approach: curves are flattened to polylines (tolerance in mm), split wherever
they cross a clip boundary, and every piece is kept or dropped by testing a
point on it against the clip regions (honouring clip-rule / fill-rule). Pieces
of a path that never touch a clip edge keep their original Bezier curves.
"""

import math

EDGE_EPS = 1e-4  # mm: points closer than this to a boundary count as "on" it

INSIDE, OUTSIDE, ON = 1, 0, -1


# --------------------------------------------------------------------------- #
# Flattening
# --------------------------------------------------------------------------- #


def flatten_subpath(sub, tolerance):
    """Polyline of one CubicSuperPath subpath ([[c_in, p, c_out], ...])."""
    points = [tuple(sub[0][1])]
    for prev, cur in zip(sub, sub[1:]):
        _flatten_cubic(prev[1], prev[2], cur[0], cur[1], tolerance, points)
    return points


def _flatten_cubic(p0, p1, p2, p3, tolerance, out):
    stack = [(tuple(p0), tuple(p1), tuple(p2), tuple(p3), 0)]
    while stack:
        a, b, c, d, depth = stack.pop()
        if depth >= 18 or (
            _line_distance(b, a, d) <= tolerance and _line_distance(c, a, d) <= tolerance
        ):
            out.append(d)
            continue
        ab, bc, cd = _mid(a, b), _mid(b, c), _mid(c, d)
        abc, bcd = _mid(ab, bc), _mid(bc, cd)
        m = _mid(abc, bcd)
        # second half first so the first half is processed first (LIFO)
        stack.append((m, bcd, cd, d, depth + 1))
        stack.append((a, ab, abc, m, depth + 1))


def _mid(a, b):
    return ((a[0] + b[0]) * 0.5, (a[1] + b[1]) * 0.5)


def _line_distance(p, a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    length = math.hypot(dx, dy)
    if length < 1e-12:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    return abs((p[0] - a[0]) * dy - (p[1] - a[1]) * dx) / length


# --------------------------------------------------------------------------- #
# Regions
# --------------------------------------------------------------------------- #


class Region:
    """Area enclosed by closed rings with a nonzero/evenodd rule."""

    def __init__(self, rings, rule="nonzero"):
        self.evenodd = str(rule).strip() == "evenodd"
        self.rings = [ring for ring in rings if len(ring) >= 3]
        self.edges = []
        for ring in self.rings:
            for a, b in zip(ring, ring[1:] + ring[:1]):
                if a != b:
                    self.edges.append((a[0], a[1], b[0], b[1]))
        if self.edges:
            xs = [e[0] for e in self.edges] + [e[2] for e in self.edges]
            ys = [e[1] for e in self.edges] + [e[3] for e in self.edges]
            self.bbox = (min(xs), min(ys), max(xs), max(ys))
        else:
            self.bbox = None
        self.index = EdgeIndex(self.edges)

    def classify(self, point):
        """INSIDE, OUTSIDE or ON (within EDGE_EPS of the boundary)."""
        if self.bbox is None:
            return OUTSIDE
        x, y = point
        x0, y0, x1, y1 = self.bbox
        if x < x0 - EDGE_EPS or x > x1 + EDGE_EPS or y < y0 - EDGE_EPS or y > y1 + EDGE_EPS:
            return OUTSIDE
        box = (x - EDGE_EPS, y - EDGE_EPS, x + EDGE_EPS, y + EDGE_EPS)
        for edge in self.index.query(box):
            if _segment_distance(point, edge) <= EDGE_EPS:
                return ON
        winding = 0
        for ax, ay, bx, by in self.edges:
            if ay <= y:
                if by > y and _is_left(ax, ay, bx, by, x, y) > 0:
                    winding += 1
            elif by <= y and _is_left(ax, ay, bx, by, x, y) < 0:
                winding -= 1
        inside = (winding % 2 == 1) if self.evenodd else (winding != 0)
        return INSIDE if inside else OUTSIDE


class ClipShape:
    """A <clipPath>: the union of its children's regions."""

    def __init__(self, regions):
        self.regions = [r for r in regions if r.bbox is not None]

    @property
    def bbox(self):
        boxes = [r.bbox for r in self.regions]
        if not boxes:
            return None
        return (
            min(b[0] for b in boxes),
            min(b[1] for b in boxes),
            max(b[2] for b in boxes),
            max(b[3] for b in boxes),
        )

    def classify(self, point):
        result = OUTSIDE
        for region in self.regions:
            state = region.classify(point)
            if state == INSIDE:
                return INSIDE
            if state == ON:
                result = ON
        return result


def _is_left(ax, ay, bx, by, x, y):
    return (bx - ax) * (y - ay) - (x - ax) * (by - ay)


def _segment_distance(p, edge):
    ax, ay, bx, by = edge
    dx, dy = bx - ax, by - ay
    denom = dx * dx + dy * dy
    t = 0.0 if denom == 0 else max(0.0, min(1.0, ((p[0] - ax) * dx + (p[1] - ay) * dy) / denom))
    return math.hypot(p[0] - (ax + t * dx), p[1] - (ay + t * dy))


class EdgeIndex:
    """Uniform grid over edges, to find candidate intersections quickly."""

    def __init__(self, edges):
        self.edges = edges
        self.cells = {}
        if not edges:
            self.size = 1.0
            return
        xs = [e[0] for e in edges] + [e[2] for e in edges]
        ys = [e[1] for e in edges] + [e[3] for e in edges]
        span = max(max(xs) - min(xs), max(ys) - min(ys), 1e-6)
        self.size = span / max(1, int(math.sqrt(len(edges))))
        for i, (ax, ay, bx, by) in enumerate(edges):
            for cell in self._cells((min(ax, bx), min(ay, by), max(ax, bx), max(ay, by))):
                self.cells.setdefault(cell, []).append(i)

    def _cells(self, box):
        s = self.size
        for cx in range(int(math.floor(box[0] / s)), int(math.floor(box[2] / s)) + 1):
            for cy in range(int(math.floor(box[1] / s)), int(math.floor(box[3] / s)) + 1):
                yield (cx, cy)

    def query(self, box):
        seen = set()
        for cell in self._cells(box):
            for i in self.cells.get(cell, ()):
                if i not in seen:
                    seen.add(i)
                    yield self.edges[i]


# --------------------------------------------------------------------------- #
# Cutting
# --------------------------------------------------------------------------- #


def split_polyline(points, closed, regions):
    """Split a polyline at every crossing with the regions' edges.

    Segments that run along a region edge (collinear overlap) become pieces of
    their own, so every piece is uniformly inside, outside or on a boundary and
    one probe point decides its fate. For a closed polyline the closing segment
    is included.
    """
    pts = list(points)
    if closed and pts[0] != pts[-1]:
        pts.append(pts[0])
    pieces = [[pts[0]]]
    for a, b in zip(pts, pts[1:]):
        cuts = set()
        overlap = False
        box = (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
        for region in regions:
            for edge in region.index.query(box):
                ts, collinear = _intersections(a, b, edge)
                cuts.update(ts)
                overlap = overlap or collinear
        if overlap and len(pieces[-1]) > 1:
            pieces.append([a])
        for t in sorted(cuts):
            p = (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
            pieces[-1].append(p)
            pieces.append([p])
        pieces[-1].append(b)
        if overlap:
            pieces.append([b])
    return [piece for piece in pieces if _length(piece) > EDGE_EPS]


def _intersections(a, b, edge):
    """(parameters t in (0, 1) where a->b meets the edge, collinear overlap?)."""
    cx, cy, dx, dy = edge
    rx, ry = b[0] - a[0], b[1] - a[1]
    sx, sy = dx - cx, dy - cy
    denom = rx * sy - ry * sx
    qx, qy = cx - a[0], cy - a[1]
    rr = rx * rx + ry * ry
    if rr == 0:
        return (), False
    if abs(denom) <= 1e-12 * math.sqrt(rr * (sx * sx + sy * sy) or 1.0):
        # parallel: only collinear overlaps matter (split at the edge's ends)
        if abs(qx * ry - qy * rx) / math.sqrt(rr) > EDGE_EPS:
            return (), False
        t_c = ((cx - a[0]) * rx + (cy - a[1]) * ry) / rr
        t_d = ((dx - a[0]) * rx + (dy - a[1]) * ry) / rr
        lo, hi = min(t_c, t_d), max(t_c, t_d)
        if hi <= 1e-9 or lo >= 1 - 1e-9:
            return (), False  # collinear but not overlapping
        return [t for t in (t_c, t_d) if 1e-9 < t < 1 - 1e-9], True
    t = (qx * sy - qy * sx) / denom
    u = (qx * ry - qy * rx) / denom
    if 1e-9 < t < 1 - 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        return (t,), False
    return (), False


def _length(piece):
    return sum(math.hypot(b[0] - a[0], b[1] - a[1]) for a, b in zip(piece, piece[1:]))


def probe_point(piece):
    """A point inside a piece (away from its ends, which lie on boundaries)."""
    half = _length(piece) / 2.0
    for a, b in zip(piece, piece[1:]):
        seg = math.hypot(b[0] - a[0], b[1] - a[1])
        if seg >= half and seg > 0:
            t = half / seg
            return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
        half -= seg
    return piece[-1]


def join_pieces(pieces, closed):
    """Merge consecutive kept pieces (each piece ends where the next starts)."""
    out = []
    for piece in pieces:
        if out and _close(out[-1][-1], piece[0]):
            out[-1].extend(piece[1:])
        else:
            out.append(list(piece))
    if closed and len(out) > 1 and _close(out[-1][-1], out[0][0]):
        out[0] = out.pop()[:-1] + out[0]
    return out


def _close(a, b):
    return abs(a[0] - b[0]) <= 1e-9 and abs(a[1] - b[1]) <= 1e-9


def bbox_overlap(a, b):
    return a is not None and b is not None and not (
        a[2] < b[0] - EDGE_EPS
        or b[2] < a[0] - EDGE_EPS
        or a[3] < b[1] - EDGE_EPS
        or b[3] < a[1] - EDGE_EPS
    )


def clip_outline(subpaths, clips, filled=False, fill_rule="nonzero", tolerance=0.01):
    """Clip a laser outline against a stack of clip paths (all must apply).

    ``subpaths``: CubicSuperPath subpaths in mm. ``clips``: list of ClipShape.
    Returns a list of parts:
      ("bezier", subpath)       untouched subpath, fully visible
      ("poly", points, closed)  visible piece as a polyline
    When ``filled`` is true, the visible edges of the clips over the filled area
    are added too, so the result outlines the visible region of the shape.
    """
    if not clips:
        return [("bezier", sub) for sub in subpaths]
    if any(clip.bbox is None for clip in clips):
        return []  # an empty clip path hides everything

    flats = []
    for sub in subpaths:
        if len(sub) < 2:
            continue
        pts = flatten_subpath(sub, tolerance)
        closed = _close(pts[0], pts[-1]) or (
            abs(pts[0][0] - pts[-1][0]) < EDGE_EPS and abs(pts[0][1] - pts[-1][1]) < EDGE_EPS
        )
        flats.append((sub, pts, closed))

    all_clip_regions = [r for clip in clips for r in clip.regions]
    parts = []

    def inside_clips(point, strict_before=None):
        for j, clip in enumerate(clips):
            state = clip.classify(point)
            if state == OUTSIDE or (state == ON and strict_before is not None and j < strict_before):
                return False
        return True

    # 1. the object's own outline, inside every clip (touching counts)
    for sub, pts, closed in flats:
        pieces = split_polyline(pts, closed, all_clip_regions)
        if len(pieces) == 1 and all(
            not _crosses(pts, closed, clip) for clip in clips
        ):
            if inside_clips(probe_point(pieces[0])):
                parts.append(("bezier", sub))
            continue
        kept = [p for p in pieces if inside_clips(probe_point(p))]
        for piece in join_pieces(kept, closed):
            parts.append(("poly", piece, _close(piece[0], piece[-1]) and len(piece) > 2))

    if not filled:
        return merge_parts(parts)

    # 2. clip edges that now bound the visible part of the filled shape
    shape = Region([pts for _, pts, _ in flats], fill_rule)
    if shape.bbox is None:
        return parts
    for i, clip in enumerate(clips):
        for k, region in enumerate(clip.regions):
            if not bbox_overlap(region.bbox, shape.bbox):
                continue
            siblings = [r for n, r in enumerate(clip.regions) if n != k]
            for ring in region.rings:
                cutters = [shape] + [r for c in clips for r in c.regions if r is not region]
                pieces = split_polyline(ring, True, cutters)
                kept = []
                for piece in pieces:
                    point = probe_point(piece)
                    if shape.classify(point) != INSIDE:
                        continue  # on the shape outline: already drawn in step 1
                    if any(s.classify(point) == INSIDE for s in siblings):
                        continue  # inside the union, not on its edge
                    others_ok = all(
                        c.classify(point) == INSIDE or (c.classify(point) == ON and j > i)
                        for j, c in enumerate(clips)
                        if j != i
                    )
                    if others_ok:
                        kept.append(piece)
                for piece in join_pieces(kept, True):
                    parts.append(("poly", piece, _close(piece[0], piece[-1]) and len(piece) > 2))
    return merge_parts(parts)


def merge_parts(parts, tol=1e-6):
    """Chain open polylines whose ends meet, so the laser cuts them in one go."""
    beziers = [p for p in parts if p[0] == "bezier"]
    closed = [p for p in parts if p[0] == "poly" and p[2]]
    open_lines = [list(p[1]) for p in parts if p[0] == "poly" and not p[2]]

    def meet(a, b):
        return abs(a[0] - b[0]) <= tol and abs(a[1] - b[1]) <= tol

    chains = []
    while open_lines:
        chain = open_lines.pop(0)
        grown = True
        while grown:
            grown = False
            for i, line in enumerate(open_lines):
                if meet(chain[-1], line[0]):
                    chain += line[1:]
                elif meet(chain[-1], line[-1]):
                    chain += line[-2::-1]
                elif meet(chain[0], line[-1]):
                    chain = line[:-1] + chain
                elif meet(chain[0], line[0]):
                    chain = line[:0:-1] + chain
                else:
                    continue
                open_lines.pop(i)
                grown = True
                break
        is_closed = len(chain) > 2 and meet(chain[0], chain[-1])
        if is_closed:
            chain[-1] = chain[0]
        chains.append(("poly", chain, is_closed))
    return beziers + closed + chains


def _crosses(pts, closed, clip):
    """True if the polyline touches any edge of the clip."""
    ring = list(pts) + ([pts[0]] if closed else [])
    for a, b in zip(ring, ring[1:]):
        box = (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
        for region in clip.regions:
            for edge in region.index.query(box):
                ts, collinear = _intersections(a, b, edge)
                if ts or collinear or _segment_distance(a, edge) <= EDGE_EPS:
                    return True
    return False
