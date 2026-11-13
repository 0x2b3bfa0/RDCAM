# LaserWorks V5 extension for Inkscape

Inkscape (1.2 or newer) port of `RLaserV5.bas`, the CorelDRAW macro that ships with
RDCAM LaserWorks V5. It adds:

| CorelDRAW macro button | Inkscape |
| --- | --- |
| **Laser Running** (`LaserWorking`) | **Extensions › Laser (RDCAM) › Send to LaserWorks V5** |
| **Import Dst/Dsb Data** (`ImportDstFile`) | **File › Open / Import** of `.dst` (Tajima) and `.dsb` (Barudan) files |

The CorelDRAW toolbar setup (`UserInit`) has no equivalent because Inkscape builds its menu from the `.inx` files.

## Install

Copy all `*.inx` and `*.py` files (not `tests/`) into your Inkscape user extensions folder, then restart Inkscape.
**Edit › Preferences › System › User extensions** shows the exact path:

- macOS: `~/Library/Application Support/org.inkscape.Inkscape/config/inkscape/extensions/`
- Windows: `%APPDATA%\inkscape\extensions\`
- Linux: `~/.config/inkscape/extensions/`

## Setup

Open **Send to LaserWorks V5**, go to the **Setup** tab and choose `LaserWorkV5.exe`.
On macOS and Linux, LaserWorks runs under Wine. The plugin finds `wine` from Homebrew, `/usr/local`
or the *Wine Stable/Devel/Staging* apps. You can also enter a command or full path, and optionally a `WINEPREFIX`.

## What "Send" does

1. Uses Inkscape to convert text, shapes and clones to paths (what `ConvertToCurves` did in the macro).
2. Takes the selection, or the whole drawing if nothing is selected. Hidden objects and layers are skipped.
3. Writes these files to `<LaserWorks folder>/temp/`:
   - `rd.ai`: every solid-coloured path as a stroked outline. The colour is the stroke colour, or the fill colour when
     there is no stroke (text, for example). LaserWorks creates its layers from these colours.
     Clipped objects are cut down to the part that is visible (see below).
   - `1.bmp`, `2.bmp`, …: 8-bit grayscale BMPs of
     - images, at their own resolution unless you set one;
     - gradient and pattern fills/strokes, and masked objects, rendered by Inkscape at the
       *Gradient/mask resolution* (300 dpi by default). A solid stroke on a gradient-filled shape is still sent as a
       vector, so it can be cut while the fill is engraved.
   - `dest.tmp` / `dest2.tmp`: sizes and positions, so the drawing lands in LaserWorks at the same size and place
     (mm, origin at the page's bottom-left).
4. Makes sure `%AppData%\r5.dll` contains exactly `aaaaaaaaaa` (10 bytes, no line ending). Under Wine, `%AppData%` is the
   prefix's `drive_c/users/<you>/AppData/Roaming`.
5. Starts `LaserWorkV5.exe`. If LaserWorks is already open, it reloads the new drawing.

The Inkscape document is never modified, so no undo step is created.

### Hand-off format (verified against `LaserWorkV5.exe`)

| File | Content |
| --- | --- |
| `dest.tmp` | `int16 kind` (0 = AI + bitmaps), `int16 count`, then `count × float64[4]` (width, height, centre x, centre y) |
| `dest2.tmp` | `float64[4]` (width, height, centre x, centre y) of the vectors |
| `rd.ai` | AI5-style PostScript: `c m y k K` stroke colour, then `m` / `L` / `c` path operators, ending in `S` / `s` |

LaserWorks reads these files at startup, or when another instance is started while it is already open. It then fits
`rd.ai` to the box in `dest2.tmp` and deletes the `temp` folder.

## Clipping

Inkscape's own boolean operations (Path › Intersection, Flatten) treat every path as a filled area. An open cut line
comes back as a closed sliver, and a stroke comes back as a filled band, which doubles every cut. So the clipping is
done in `rdlaser_geometry.py`:

- Clip paths are resolved with inkex: `userSpaceOnUse` and `objectBoundingBox` units, transforms, `clip-rule`, clips
  on parent groups (nested clips intersect), and text inside clip paths (converted to paths by Inkscape first).
- Each outline is split wherever it crosses a clip edge, and only the visible pieces are kept. Lines stay open.
  Parts that never touch a clip edge keep their original curves; cut parts become polylines (0.01 mm tolerance).
- For filled shapes, the clip edges that now bound the visible area are added. The pieces are chained into continuous,
  closed outlines where possible, and edges shared by the shape and the clip are cut only once.
- Bitmaps (images, gradients, masks) are rendered by Inkscape itself, so their clips are exact.

## Notes

- Masks are transparency, not geometry, so masked objects are always sent as bitmaps.
- Images are sent as grayscale (as in the original macro). Inkscape converts colour to gray with Rec.709 weights.

## Tests

```sh
python -m venv .venv && .venv/bin/pip install "inkex>=1.2" pytest pillow
INKSCAPE_COMMAND=/path/to/inkscape .venv/bin/pytest tests
```

End-to-end tests run Inkscape for the path conversion and bitmap rendering; they are skipped if Inkscape is not found.
