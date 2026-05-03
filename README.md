# SharpCap Solar Scanning — MLAStro SHG Fast Burst Scanner

An event-driven Python script that automates fast, repeatable solar disk
scans for a Spectroheliograph (SHG) using SharpCap's scripting API. It
finds the solar limb, centres the disk in declination, sets exposure
from the live histogram, runs a constant-rate scan across the disk, and
loops — without ever leaving SharpCap or touching temp files.

---

## How to use it in SharpCap

The script is a SharpCap **IronPython** script. It runs inside SharpCap's
embedded Python console and uses the `SharpCap` automation object that
the host injects at runtime. It is **not** a standalone Python program —
do not try to run it from a regular CPython interpreter.

### Prerequisites

- **SharpCap Pro** (the scripting console requires the Pro licence).
- A telescope **mount** connected and selected in SharpCap
  (`Mount` panel — `SharpCap.Mounts.SelectedMount` must be non-null).
- A **camera** connected and live-previewing
  (`SharpCap.SelectedCamera` must be non-null and streaming frames).
- The mount's RA axis aligned with the scan direction (the script
  drives axis 0 forward to cross the disk and axis 1 to nudge Dec).
- A roughly-pointed sun — close enough that the script can find the
  limb within `LIMB_MAX_SEARCH` seconds (default 120 s) of slewing.

### Running the script

1. Connect the camera and mount in SharpCap. Start the live view.
2. Make sure the mount is tracking (sidereal) and the sun is somewhere
   inside the slit/field, or just off it.
3. Verify the camera is streaming frames at a stable FPS — the scan
   speed is auto-derived from `cam.CurrentFrameRate`.
4. Open the SharpCap scripting console
   (`Tools` → `Scripting` → `Open Console`).
5. Either paste the contents of `Solarscan_fastpaced.py` into the
   console and run it, or use the `Run Script File…` option and point
   it at `Solarscan_fastpaced.py`.
6. Watch the console output. Pre-flight info is printed first
   (FPS, scan multiplier, crossing time, full-scale ADU, estimated
   total runtime). The cycles run automatically.
7. **Press `Ctrl+C` in the console to abort cleanly** — the script
   stops the mount, stops capture if active, and detaches its
   `FrameCaptured` handlers.

### What you should adjust before first use

Edit the constants near the top of the file:

| Constant         | What it is                                                     | Default |
|------------------|----------------------------------------------------------------|---------|
| `FOCAL_LENGTH`   | Telescope focal length in millimetres                          | 714     |
| `PIXEL_SIZE`     | Camera pixel pitch in micrometres                              | 2       |
| `SOLAR_DIAMETER` | Apparent solar diameter in arcseconds (varies through the year)| 1920    |
| `NUM_CYCLES`     | Number of full scans in this session                           | 1       |
| `ENABLE_CAPTURE` | Set `False` for a dry run (no SER file written)                | `True`  |
| `ENABLE_AUTO_EXP`| Set `False` if you want to control exposure manually           | `True`  |

The remaining constants tune the limb / Dec / exposure algorithms and
are documented inline in the file.

### Capture output

When `ENABLE_CAPTURE = True`, capture is started by the script with
`cam.PrepareToCapture()` / `cam.RunCapture()` and stopped at the end of
each scan leg. SharpCap's normal capture settings apply: the file
format, target folder, naming, gain, and any pre-/post-processing are
whatever you have configured in the SharpCap UI before launching the
script. Set the camera to **MONO16** for the best SHG dynamic range.

---

## What the script does

A single cycle, end-to-end:

1. **Find the solar limb.** Slew RA forward at `LIMB_SEARCH_SPEED`
   while monitoring the live frame's standard deviation. On dark sky
   the slit shows a near-uniform field (low stddev); on the disk it
   shows a spectrum with deep absorption lines (high stddev). When the
   stddev crosses a dynamically computed threshold the mount stops —
   the slit is now on the limb.
2. **Move to disk centre.** Slew forward at `RETURN_MULTIPLIER` for
   exactly half the disk-crossing time.
3. **Centre in declination.** Capture one frame, split it into left
   and right halves with `Frame.CutROI`, and compare the mean
   brightness of each half. Nudge Dec toward the dimmer side until
   `(L − R) / (L + R)` is below `DEC_CENTER_TOLERANCE`.
4. **Auto-expose.** Read the live histogram, compute the pixel value
   at the `EXP_PERCENTILE` cumulative point, and scale exposure so
   that percentile lands at `EXP_TARGET_LEVEL` of full scale. This
   targets the spectrum's continuum peak (which mean-based AE
   misses) so absorption-line detail is preserved without clipping.
5. **Reposition past the leading limb.** Back the mount up by
   `PADDED_DURATION` seconds at the search rate so the scan starts on
   blank sky.
6. **Run the scan.** Start SharpCap capture, slew at the auto-derived
   `scan_multiplier` so the disk drifts across the slit at exactly one
   pixel per frame, and stop both mount and capture once the disk
   plus padding has crossed.
7. **Fast-return** to the start position at `RETURN_MULTIPLIER` and
   begin the next cycle.

Every `DEC_INTERVAL` cycles the script repeats the full alignment
(limb + Dec + exposure). Other cycles do a quick re-find of the limb
only.

After the last cycle the mount is repositioned to the disk centre.

## Why event-driven (v2 design notes)

This is v2 of the scanner. The core change vs. v1 is that **all
brightness and edge measurements happen inside a `FrameCaptured`
handler on the live stream** — the mount never has to stop, switch
modes, save a PNG, or read a file from disk to take a measurement.

Other v1 → v2 changes:

- **Limb signal is stddev, not mean brightness.** Robust whether the
  spectrum is bright or dim, and works with either spectral axis
  orientation.
- **Dec centring uses `Frame.CutROI`** to split the live frame in
  half, so it does not depend on saved files.
- **Auto-exposure targets a histogram percentile** at a fraction of
  full scale, so it is bit-depth agnostic (MONO8 or MONO16) and
  resistant to a few hot pixels.
- **`Ctrl+C` is honoured everywhere.** Wait loops poll an abort flag,
  so aborting always stops the mount and detaches handlers cleanly.

## Polar-alignment drift report

At the end of the session the script prints a drift summary. RA and
Dec of the leading limb are recorded each cycle; total drift is
divided by elapsed time to give arcseconds per minute, and a hint is
printed about which axis (azimuth or altitude) of the polar mount to
adjust if drift exceeds 5"/min.

## Hardware this was developed against

- **Telescope:** SVBONY SV503 102 ED, f/7 (102 mm aperture, 714 mm focal length)
- **Camera:** TOUPTEK G3M867M (mono, MONO16)
- **Mount:** MLAStro SAL-33

The code is hardware-agnostic in principle — anything SharpCap can
drive should work — but the default `DEC_NUDGE_SPEED`, settle times,
and step caps were tuned on the setup above. Expect to retune
`DEC_SETTLE_TIME` and `DEC_NUDGE_SPEED` for a heavier OTA or a
different mount.

## Disclaimer — use at your own risk

This script is provided **as-is**, with no warranty of any kind. The
author accepts **no responsibility** for any damage, loss, or injury
that results from its use, including but not limited to:

- Damage to the mount, telescope, camera, or any other equipment
  caused by unexpected slews, runaway motion, or stalls.
- Damage to the camera sensor, eyepieces, or your eyesight from
  pointing at or near the sun without an appropriate, undamaged
  solar filter.
- Loss of captured data or corruption of SharpCap state.

The script drives a motorised mount at non-trivial speeds while the
telescope is pointed at the sun. It is **entirely your responsibility**
to:

- Ensure a properly fitted, undamaged solar filter is in place at all
  times when the telescope is pointed at or near the sun.
- Verify that the mount's slew limits, cable routing, and counterweight
  balance are safe before running the script.
- Stay within reach of the abort path (`Ctrl+C` in the SharpCap console
  and the mount's emergency stop) for the entire session.
- Test with `ENABLE_CAPTURE = False` and reduced slew rates first when
  trying the script on a new rig.

If you are not comfortable accepting these risks, do not run this
script.

## License

Released under the **MIT License** — see [LICENSE](LICENSE) for the
full text. In short: do whatever you like with the code, keep the
copyright notice, and don't blame the author if it breaks. Issues and
PRs welcome.
