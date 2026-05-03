# MLAStro SHG Fast Burst Scanner — v2
# Event-driven limb/Dec/exposure detection via GetStats on live frames.
# No mode switching, no PNG round-trip, no temp files.
#
# Port notes from v1 -> v2:
#   - Limb signal is stddev (Item2), not mean brightness. On-disk the slit
#     shows a spectrum with dark absorption lines => high stddev. Off-disk
#     the frame is dark-uniform => low stddev. Robust on either axis.
#   - All brightness/edge measurements happen inside a FrameCaptured handler
#     on the live stream. Mount never stops to measure.
#   - Dec centering compares mean of left half vs right half of the frame
#     via Frame.CutROI, nudges toward the dimmer side.
#   - Auto-exposure target is a fraction of full scale (bit-depth agnostic).
#   - Ctrl+C aborts: stops mount, detaches any handler, returns cleanly.
#
# Acknowledgements:
#   The v2 architecture — staying in MONO16 throughout, attaching a
#   FrameCaptured handler to the live stream, and using
#   Frame.GetStats().Item2 (stddev) plus Frame.CutROI(Rectangle(...)) to
#   measure the disk without any temp-file round-trip — is directly
#   inspired by Patrick Hsieh's (FlankerOneTwo) SHGScan project:
#       https://github.com/FlankerOneTwo/SHGScan
#   Huge thanks to Patrick for publishing that code and showing how to do
#   all the measurement work on the live stream — that approach is what
#   led to dropping MONO8 and operating purely in MONO16.
#
#   Thanks also to Fabio Silvi for the SideOfPier-aware Dec centering
#   contributed in feat/sideofpier-dec-correction.

import time
import clr
# SharpCap runs IronPython, which can import .NET types via `clr`. We need
# System.Drawing.Rectangle because that's the type SharpCap's Frame.CutROI
# expects when picking a sub-region of a frame.
clr.AddReference("System.Drawing")
from System.Drawing import Rectangle

# ══════════════════════════════════════════════════════════════════════════════
# PARAMETERS
# ══════════════════════════════════════════════════════════════════════════════

# ── Hardware ──────────────────────────────────────────────────────────────────
# Together these define the image scale: arcsec/pixel = 206265 * (PIXEL_SIZE/1000) / FOCAL_LENGTH.
# SOLAR_DIAMETER varies through the year by ~3% — 1920" is a safe annual mean.
# SIDEREAL is the ASCOM unit conversion: a "1.0x" mount slew = 15.04"/s of axis motion.
FOCAL_LENGTH       = 714     # mm
PIXEL_SIZE         = 2       # um (G3M678M)
SIDEREAL           = 15.04   # arcsec/s per sidereal multiplier unit
SOLAR_DIAMETER     = 1920.0  # arcsec

# ── Scanning ──────────────────────────────────────────────────────────────────
# RETURN_MULTIPLIER and LIMB_SEARCH_SPEED below are sidereal multipliers — the
# value passed to mount.MoveAxis(). 32x sidereal ≈ 481"/s, fast enough to make
# repositioning between cycles negligible without exceeding most mount limits.
ENABLE_CAPTURE     = True
ENABLE_AUTO_EXP    = True
NUM_CYCLES         = 1
PADDED_DURATION    = 3       # seconds of sky captured before/after the disk
RETURN_MULTIPLIER  = 32.0    # fast slew rate (no capture)

# ── Limb detection (stddev-based) ─────────────────────────────────────────────
LIMB_SEARCH_SPEED  = 4.0     # forward search multiplier
LIMB_SAMPLE_INTERVAL = 0.05  # how often to poll the watcher (s)
LIMB_MAX_SEARCH    = 120     # seconds before giving up
# Absolute stddev floor that counts as "on disk". At MONO16, sensor dark
# noise stddev is ~20-100, on-disk stddev is typically 10000+.
LIMB_STDDEV_MIN    = 500
# Stddev above this = clearly on disk (used to decide whether to back off first).
ON_DISK_STDDEV     = 2000
# Dynamic threshold multiplier over observed dark baseline.
LIMB_DARK_MULT     = 5.0

# ── Dec centering (CutROI left/right mean) ────────────────────────────────────
DEC_INTERVAL       = 2       # re-center every N cycles
DEC_CENTER_TOLERANCE = 0.03  # (L-R)/(L+R) within this = centered
DEC_MAX_CORRECTIONS = 8
DEC_NUDGE_SPEED    = 3.0
DEC_SETTLE_TIME    = 0.3

# ── Auto-exposure (histogram percentile) ──────────────────────────────────────
# Drive the specified percentile of the histogram to a target fraction of full
# scale. Prevents continuum-peak clipping that a mean-based target can't see.
# 0.995 = look at the 99.5th percentile (ignores a few hot pixels).
# 0.88  = aim for that percentile at 88% of full scale (leaves headroom).
EXP_PERCENTILE     = 0.995
EXP_TARGET_LEVEL   = 0.88
EXP_TOLERANCE      = 0.04
EXP_MAX_ATTEMPTS   = 6

# ══════════════════════════════════════════════════════════════════════════════

# 206265 = arcseconds per radian. Multiplied by (pixel size / focal length) in
# matching length units, this gives the angular size each pixel subtends.
PIXEL_SCALE = 206265 * (PIXEL_SIZE / 1000.0) / FOCAL_LENGTH

# `SharpCap` is a host-injected global available inside the SharpCap scripting
# console. Both will be None if the user has not connected a mount/camera; the
# script does not guard against that — fail loudly is better than silently.
mount = SharpCap.Mounts.SelectedMount
cam   = SharpCap.SelectedCamera

# IronPython 2.7 has no `nonlocal`, so we use a one-element list as a mutable
# closure cell. Ctrl+C in the SharpCap console raises KeyboardInterrupt at
# the `try:` block far below; we set the flag there and check it in every
# wait loop so handlers and watchers can exit cleanly.
_abort = [False]

def check_abort():
    if _abort[0]:
        raise KeyboardInterrupt()

def ra_diff_arcsec(ra1, ra2):
    """Signed RA difference in arcsec, handling the 0/24h wrap-around.

    `mount.RA` returns hours in [0, 24). A naive subtraction of 23.9h - 0.1h
    would give 23.8h (~360°), but the actual angular difference is 0.2h.
    Anything beyond ±12h crosses the wrap, so add/subtract 24 to bring it
    back into the short-arc range. 1 hour of RA = 15° = 54000".
    """
    diff = ra1 - ra2
    if diff > 12:
        diff -= 24
    elif diff < -12:
        diff += 24
    return diff * 15 * 3600

def full_scale():
    """Detect bit depth from current color space. MONO16 => 65535, MONO8 => 255."""
    try:
        cs = str(cam.Controls.ColourSpace.Value).upper()
        if "16" in cs:
            return 65535.0
        if "8" in cs:
            return 255.0
    except:
        pass
    return 65535.0

# ── FrameCaptured watcher: latches stats from the live stream ─────────────────
# The FrameCaptured + Frame.GetStats() pattern used here is borrowed from
# Patrick Hsieh's SHGScan (https://github.com/FlankerOneTwo/SHGScan).
#
# SharpCap fires the camera's FrameCaptured event for every frame in the live
# preview. We attach a handler, let it run continuously, and latch only the
# most recent (mean, stddev) pair so the rest of the script can read "current
# brightness" at any moment without ever stopping the stream. This is the key
# trick that makes the v2 architecture event-driven.
class FrameWatcher:
    """Attaches a FrameCaptured handler and latches the most recent (mean, stddev)."""
    def __init__(self, camera):
        self.camera = camera
        self.latest = None
        # Guards against a frame arriving after stop() has begun unhooking.
        self._active = False

    def start(self):
        # `+=` on a .NET event subscribes the handler. SharpCap will then call
        # _on_frame on its capture thread for every frame until we unsubscribe.
        self.camera.FrameCaptured += self._on_frame
        self._active = True

    def stop(self):
        if self._active:
            try:
                self.camera.FrameCaptured -= self._on_frame
            except:
                # Some SharpCap builds throw if the handler was already detached
                # or the camera was disconnected — non-fatal, swallow it.
                pass
            self._active = False

    def _on_frame(self, sender, args):
        if not self._active:
            return
        try:
            # Frame.GetStats() returns a .NET Tuple<float, float>:
            #   Item1 = mean pixel value, Item2 = standard deviation.
            # IronPython exposes these as attributes (not [0] / [1] indices).
            s = args.Frame.GetStats()
            self.latest = (s.Item1, s.Item2)
        except:
            # Don't take down the live stream if a single frame's stats fail.
            pass

    def wait_first(self, timeout=3.0):
        """Block until the first frame has populated `latest`, or time out."""
        t0 = time.time()
        while self.latest is None:
            if time.time() - t0 > timeout:
                return False
            check_abort()
            time.sleep(0.05)
        return True

# ── One-shot frame grab: left/right mean via CutROI ───────────────────────────
# The Frame.CutROI(Rectangle(...)) approach for measuring sub-regions of a
# live frame (instead of writing the frame to disk and reading it back) is
# borrowed from Patrick Hsieh's SHGScan.
#
# Unlike FrameWatcher (which keeps running and is read by polling), this
# helper attaches a handler, takes the *next* frame that arrives, computes
# left/right half-frame means from it, and detaches. We do this on a single
# fresh frame (rather than reading FrameWatcher.latest) because Dec centering
# needs a stable measurement after the mount has settled, not whichever frame
# happened to be cached during the previous nudge.
def measure_left_right_means(timeout=3.0):
    """Capture one live frame, split into left and right halves, return means."""
    # One-element list = mutable closure cell (IronPython 2.7 has no nonlocal).
    result = [None]

    def on_frame(sender, args):
        # Only act on the first frame after subscription; later frames are
        # ignored so we don't keep doing work after the result is in.
        if result[0] is not None:
            return
        try:
            f = args.Frame
            # cam.ROI is the camera's currently-set sensor ROI; the live frame
            # has these dimensions. Splitting at the midpoint gives us two
            # halves to compare for Dec centering.
            roi = cam.ROI
            w = roi.Width
            h = roi.Height
            mid = w // 2
            # Frame.CutROI(Rectangle(x, y, w, h)) returns a sub-frame view.
            # GetStats() on each half gives us per-half (mean, stddev).
            left  = f.CutROI(Rectangle(0,   0, mid,     h))
            right = f.CutROI(Rectangle(mid, 0, w - mid, h))
            ls = left.GetStats()
            rs = right.GetStats()
            result[0] = (ls.Item1, rs.Item1)  # means only — caller compares L vs R
        except Exception as e:
            # Smuggle the error back across the closure boundary as a tagged
            # tuple; the caller unpacks and prints it after detaching.
            result[0] = ("error", str(e))

    cam.FrameCaptured += on_frame
    try:
        t0 = time.time()
        while result[0] is None:
            if time.time() - t0 > timeout:
                return None
            check_abort()
            time.sleep(0.05)
    finally:
        # Always detach, even on timeout / abort, so we don't leak handlers.
        try:
            cam.FrameCaptured -= on_frame
        except:
            pass

    if isinstance(result[0], tuple) and result[0][0] == "error":
        print("  CutROI error: {}".format(result[0][1]))
        return None
    return result[0]

# ── Limb detection ────────────────────────────────────────────────────────────
# Using stddev (GetStats().Item2), not mean brightness, as the on-disk vs
# off-disk signal is an idea taken from Patrick Hsieh's SHGScan. The dynamic
# dark-baseline thresholding around it is local to this script.
#
# Why stddev and not mean: the SHG slit projects a spectrum onto the sensor.
# When the slit is on the solar disk that spectrum has dark Fraunhofer lines
# crossing it, so the per-frame stddev is enormous (typ. 10000+ at MONO16).
# Off-disk the frame is near-uniform sensor noise, so stddev is tiny (10s).
# The signal is robust whichever way the spectral axis is oriented and does
# not depend on a particular brightness target.
def find_solar_limb():
    # mount.MoveAxis(0, ...) is the RA axis; positive multiplier slews "east"
    # in the apparent-sky sense (sun appears to drift west across the slit
    # when the mount moves slower than sidereal in this direction).
    watcher = FrameWatcher(cam)
    watcher.start()
    try:
        if not watcher.wait_first():
            print("  No frames from camera")
            return False

        initial_stddev = watcher.latest[1]

        # If we started already on disk, find the leading limb by reversing
        # off the disk first, then resuming the forward search. This keeps the
        # rest of the routine simple — it only has to handle "we're in dark
        # sky and need to find the rising edge".
        if initial_stddev > ON_DISK_STDDEV:
            print("  On disk (stddev={:.0f}) - backing up...".format(initial_stddev), end="")
            mount.MoveAxis(0, -RETURN_MULTIPLIER)
            t0 = time.time()
            off_disk = False
            while (time.time() - t0) < LIMB_MAX_SEARCH:
                check_abort()
                if watcher.latest[1] < LIMB_STDDEV_MIN:
                    mount.Stop()
                    off_disk = True
                    print(" off after {:.1f}s".format(time.time() - t0))
                    time.sleep(0.3)  # let the mount settle before sampling dark
                    break
                time.sleep(LIMB_SAMPLE_INTERVAL)
            if not off_disk:
                mount.Stop()
                print(" FAILED")
                return False

        # Sample the actual dark-sky stddev and build a threshold relative to it,
        # rather than using a fixed absolute number. Hot pixels, gain settings,
        # and exposure time all change the noise floor; a relative threshold
        # adapts. We require all three: at least LIMB_DARK_MULT× the floor, at
        # least floor+300, and at least the absolute LIMB_STDDEV_MIN — whichever
        # is largest wins, so noisy darks don't trigger spurious "limb found".
        time.sleep(0.2)
        dark = watcher.latest[1]
        threshold = max(dark * LIMB_DARK_MULT, dark + 300.0, LIMB_STDDEV_MIN)

        # Slew forward at LIMB_SEARCH_SPEED (4× sidereal) and watch for the
        # stddev jump that marks the leading limb crossing into the slit.
        mount.MoveAxis(0, LIMB_SEARCH_SPEED)
        t0 = time.time()
        while (time.time() - t0) < LIMB_MAX_SEARCH:
            check_abort()
            s = watcher.latest[1]
            if s > threshold:
                mount.Stop()
                print("  Limb found stddev={:.0f} (dark={:.0f}, thr={:.0f}, {:.1f}s)".format(
                      s, dark, threshold, time.time() - t0))
                return True
            time.sleep(LIMB_SAMPLE_INTERVAL)

        mount.Stop()
        print("  Limb not found (dark={:.0f}, thr={:.0f})".format(dark, threshold))
        return False
    finally:
        watcher.stop()

# ── Dec centering ─────────────────────────────────────────────────────────────
def center_dec():
    """Nudge Dec until left-half and right-half means are balanced.

    Strategy: at the disk centre the slit is illuminated symmetrically along
    its length, so the left and right halves of the frame should have equal
    mean brightness. If the sun is offset in Dec, one half is brighter than
    the other; we nudge Dec toward the dimmer side and re-measure.
    """
    for iteration in range(1, DEC_MAX_CORRECTIONS + 1):
        check_abort()
        means = measure_left_right_means()
        if means is None:
            print("  Dec: frame read failed")
            return False
        l, r = means
        denom = l + r
        if denom <= 0:
            # Both halves dark — likely off-disk or very over-exposed black.
            # Nothing useful to do; abort centering rather than guess.
            print("  Dec: no signal")
            return False
        # Normalised imbalance in [-1, +1]. Sign tells us which half is brighter,
        # magnitude how far off-centre we are. Dividing by (L+R) cancels overall
        # brightness, so the value doesn't depend on exposure.
        offset = (l - r) / denom  # + => sun pushed left, - => pushed right

        print("  Dec {}: L={:.0f} R={:.0f} offset={:+.3f}".format(
              iteration, l, r, offset), end="")

        if abs(offset) < DEC_CENTER_TOLERANCE:
            print("  OK")
            return True

        # Proportional control: nudge time scales with how far off-centre we
        # are. Cap the upper end so a misread doesn't trigger a long slew, and
        # the lower end so a tiny nudge still actually moves the mount past
        # static friction / gear backlash.
        nudge_time = min(abs(offset) * 8.0, 1.5)
        nudge_time = max(nudge_time, 0.1)
        # +offset means left half brighter => sun is pushed left in the frame
        # => nudge Dec in the direction that moves the sun right (toward center).
        direction = -1 if offset > 0 else 1

        # MoveAxis(1, ...) is the Dec axis (axis index 1); axis 0 is RA.
        mount.MoveAxis(1, direction * DEC_NUDGE_SPEED)
        time.sleep(nudge_time)
        mount.Stop()
        # Brief settle pause so the next measure_left_right_means() reads a
        # mechanically-still frame rather than one mid-vibration.
        time.sleep(DEC_SETTLE_TIME)
        print("  nudge {:.2f}s {}".format(nudge_time, "+" if direction > 0 else "-"))

    print("  Dec: max corrections")
    return False

# ── Histogram percentile measurement ──────────────────────────────────────────
def measure_percentile_pixel(pct, timeout=3.0):
    """Capture one frame, return pixel value at the given cumulative percentile.

    For pct=0.99, this returns the bin index N such that 99% of all pixels in
    the frame have a value <= N. We use this for auto-exposure: targeting the
    99.5th percentile is robust against a few hot pixels but still tracks the
    bright continuum peak of the spectrum, which is what we don't want to clip.
    """
    result = [None]

    def on_frame(sender, args):
        if result[0] is not None:
            return
        try:
            # CalculateHistogram() returns a Histogram object with a Values
            # property that is a list-of-channel-arrays. For mono cameras
            # there is one channel; Values[0] is its bin array. Each bin
            # holds the pixel count for that pixel value (0..255 for MONO8,
            # 0..65535 for MONO16).
            h = args.Frame.CalculateHistogram()
            bins = h.Values[0]
            nbins = len(bins)
            # Total pixel count = sum of all bins. Walk from bin 0 upward and
            # find the first bin where the running cumulative count crosses
            # `pct` of the total — that bin index is the percentile pixel value.
            total = 0
            for i in range(nbins):
                total += bins[i]
            if total <= 0:
                result[0] = 0
                return
            target_count = total * pct
            cum = 0
            for i in range(nbins):
                cum += bins[i]
                if cum >= target_count:
                    result[0] = i
                    return
            result[0] = nbins - 1
        except Exception as e:
            result[0] = ("error", str(e))

    cam.FrameCaptured += on_frame
    try:
        t0 = time.time()
        while result[0] is None:
            if time.time() - t0 > timeout:
                return None
            check_abort()
            time.sleep(0.05)
    finally:
        try:
            cam.FrameCaptured -= on_frame
        except:
            pass

    if isinstance(result[0], tuple) and result[0][0] == "error":
        print("  Histogram error: {}".format(result[0][1]))
        return None
    return result[0]

# ── Auto-exposure ─────────────────────────────────────────────────────────────
def auto_exposure():
    """Iteratively scale the exposure so the EXP_PERCENTILE pixel sits at
    EXP_TARGET_LEVEL of full scale.

    The signal we care about for SHG reconstruction is the continuum peak of
    the spectrum, which is well above the mean. A mean-based AE would happily
    drive that peak up into the clipping rail. Instead we look at the high
    percentile of the histogram and aim it at ~88% of full scale, leaving
    room above for atmospheric bumps without losing any line detail.
    """
    fs = full_scale()
    target_val = EXP_TARGET_LEVEL * fs  # absolute pixel value we're aiming for
    tol_val    = EXP_TOLERANCE * fs     # ± window that counts as "converged"

    # cam.Controls.FindByName returns the named control object; its .Value
    # property is read/write and is the exposure in milliseconds for most
    # ZWO/QHY/Touptek drivers under SharpCap.
    exp_ctrl = cam.Controls.FindByName("Exposure")
    entry_exp = exp_ctrl.Value  # remember starting value so we can revert on failure

    for attempt in range(EXP_MAX_ATTEMPTS):
        check_abort()
        p = measure_percentile_pixel(EXP_PERCENTILE)
        if p is None:
            print("  Auto-exp: measurement failed, reverting to {:.2f}ms".format(entry_exp))
            exp_ctrl.Value = entry_exp
            return False

        current_exp = exp_ctrl.Value
        print("  Exp {:.2f}ms  p{:.1f}={:.0f}/{:.0f} ({:.0%})".format(
              current_exp, EXP_PERCENTILE * 100, p, fs, p / fs), end="")

        if abs(p - target_val) < tol_val:
            print("  OK")
            return True

        if p > 0:
            # Linear approximation: brightness scales roughly with exposure,
            # so new_exp = current_exp × (target / measured). Clamp the
            # multiplier to [0.5, 2.0] so a single noisy/blocked frame can't
            # send the exposure to an extreme; we'll converge over a few
            # iterations instead.
            ratio = target_val / p
            ratio = max(0.5, min(2.0, ratio))
            new_exp = current_exp * ratio
            exp_ctrl.Value = new_exp
            # Wait long enough for the new exposure to take effect on at
            # least one frame before we measure again.
            time.sleep(0.3)
            print("  -> {:.2f}ms".format(new_exp))
        else:
            # Zero percentile = no signal at all, so the multiplier would
            # be infinite. Bail rather than ramp exposure into oblivion.
            print("  (no signal)")
            exp_ctrl.Value = entry_exp
            return False

    # Fell through without converging — don't leave a possibly-bad value.
    print("  Auto-exp: max attempts, reverting to {:.2f}ms".format(entry_exp))
    exp_ctrl.Value = entry_exp
    return False

# ── Alignment routines ────────────────────────────────────────────────────────
# do_full_alignment: re-find the limb, walk to disk centre, re-centre Dec, set
# exposure, then walk back to the limb. This is the slow path, used on cycle 0
# and every DEC_INTERVAL cycles. Returns the RA of the leading limb.
#
# do_quick_realign: just re-find the leading limb. Used on intermediate cycles
# where Dec/exposure are assumed not to have drifted enough to need re-tuning.
def do_full_alignment():
    print("\n  Full alignment...")
    if not find_solar_limb():
        return False
    limb_ra = mount.RA  # remembered so we can return here later

    # Time to walk half a solar diameter at RETURN_MULTIPLIER × sidereal.
    # This puts us roughly at the centre of the disk for Dec measurement.
    half_t = (SOLAR_DIAMETER / 2.0) / (RETURN_MULTIPLIER * SIDEREAL)
    mount.MoveAxis(0, RETURN_MULTIPLIER)
    time.sleep(half_t)
    mount.Stop()
    time.sleep(0.2)

    print("  Dec centering...")
    center_dec()

    # Let the mount fully damp before sampling brightness. The per-iteration
    # DEC_SETTLE_TIME inside center_dec is too short for the SAL-33 with a
    # 4" OTA — consecutive nudges accumulate mechanical oscillation.
    time.sleep(1.0)

    if ENABLE_AUTO_EXP:
        print("  Auto-exposure...")
        auto_exposure()

    # Walk back to the leading limb so the actual scan starts in the right place.
    mount.MoveAxis(0, -RETURN_MULTIPLIER)
    time.sleep(half_t)
    mount.Stop()
    time.sleep(0.2)

    return limb_ra

def do_quick_realign():
    print("  Quick realign...")
    if not find_solar_limb():
        return False
    return mount.RA

# ── Auto-calculate scan speed from live fps ───────────────────────────────────
# For a 1:1 reconstruction X:Y ratio we need the slit to traverse exactly one
# pixel-width of sky per captured frame. So:
#   scan_speed (arcsec/s) = PIXEL_SCALE (arcsec/px) × fps (frames/s)
# and the corresponding sidereal multiplier the mount needs is that speed
# divided by SIDEREAL (15.04"/s per 1× sidereal). This couples the slew
# directly to the live frame rate, so any FPS change picked up here is
# automatically compensated.
fps = cam.CurrentFrameRate
scan_multiplier = PIXEL_SCALE * fps / SIDEREAL
scan_speed      = scan_multiplier * SIDEREAL
crossing_time   = SOLAR_DIAMETER / scan_speed
# scan_duration includes PADDED_DURATION seconds of dark sky padding on each
# side of the disk, captured at the same scan_speed.
scan_duration   = PADDED_DURATION * 2 + crossing_time

# ── Pre-flight info ───────────────────────────────────────────────────────────
est_first = 15 + PADDED_DURATION + scan_duration + 6
est_fast  = 3 + PADDED_DURATION + scan_duration + 6
est_total = est_first + (NUM_CYCLES - 1) * est_fast

print("=" * 60)
print("MLAStro SHG Fast Burst Scanner v2")
if not ENABLE_CAPTURE or not ENABLE_AUTO_EXP:
    flags = []
    if not ENABLE_CAPTURE:
        flags.append("CAPTURE OFF")
    if not ENABLE_AUTO_EXP:
        flags.append("AUTO-EXP OFF")
    print("*** DRY RUN: {} ***".format(", ".join(flags)))
print("=" * 60)
print("FPS: {:.1f}  Scale: {:.3f}\"/px  Scan: {:.2f}x ({:.1f}\"/s)".format(
      fps, PIXEL_SCALE, scan_multiplier, scan_speed))
print("Crossing: {:.1f}s  Duration: {:.1f}s  X/Y: {:.3f}".format(
      crossing_time, scan_duration,
      PIXEL_SCALE / (scan_speed / fps)))
print("Cycles: {}  Return: {:.0f}x  Dec every: {} cycles".format(
      NUM_CYCLES, RETURN_MULTIPLIER, DEC_INTERVAL))
print("Full scale: {:.0f} ADU  Exp: p{:.1f} -> {:.0%}".format(
      full_scale(), EXP_PERCENTILE * 100, EXP_TARGET_LEVEL))
print("Est. total: {:.0f}s ({:.1f} min)  (Ctrl+C to abort)".format(
      est_total, est_total / 60))
print("=" * 60)

# ── Main cycle loop ───────────────────────────────────────────────────────────
drift_log = []
session_start = time.time()
return_time = 0

try:
    for cycle in range(NUM_CYCLES):
        cycle_start = time.time()
        print("\n-- CYCLE {} of {} --".format(cycle + 1, NUM_CYCLES))

        # Re-do the full alignment (limb + Dec + exposure) on the first cycle
        # and every DEC_INTERVAL cycles thereafter. Intermediate cycles only
        # need to re-find the leading limb, which is much cheaper.
        need_full = (cycle == 0) or (cycle % DEC_INTERVAL == 0)

        if need_full:
            limb_ra = do_full_alignment()
        else:
            limb_ra = do_quick_realign()

        if limb_ra is False:
            print("ABORTING: Could not find solar limb.")
            break

        # Log (cycle, t, RA, Dec) of the leading limb each cycle so we can
        # report drift against cycle 0 at the end of the session.
        limb_dec = mount.Dec
        drift_log.append((cycle + 1, time.time(), limb_ra, limb_dec))

        if len(drift_log) > 1:
            ref = drift_log[0]
            dt  = drift_log[-1][1] - ref[1]
            dra = ra_diff_arcsec(drift_log[-1][2], ref[2])
            # mount.Dec is in degrees; ×3600 → arcsec.
            ddec = (drift_log[-1][3] - ref[3]) * 3600
            if dt > 0:
                print("  Drift: RA {:+.1f}\"  Dec {:+.1f}\"  ({:.0f}s)".format(
                      dra, ddec, dt))

        # Step the mount back into dark sky just before the leading limb so
        # the scan begins on sky and ramps in cleanly. Note the rate here
        # must match the leading-edge padding scan_duration accounts for
        # (PADDED_DURATION × scan_speed). Backing up at LIMB_SEARCH_SPEED
        # would overshoot and cause the scan to stop short of the trailing
        # limb at low FPS / short focal length.
        mount.MoveAxis(0, -LIMB_SEARCH_SPEED)
        time.sleep(PADDED_DURATION)
        mount.Stop()
        time.sleep(0.2)

        # The actual scan: kick off capture, slew at the FPS-matched rate
        # for the full disk-crossing time + 2× padding, then stop.
        if ENABLE_CAPTURE:
            cam.PrepareToCapture()
            cam.RunCapture()
        mount.MoveAxis(0, scan_multiplier)
        time.sleep(scan_duration)
        mount.Stop()
        if ENABLE_CAPTURE:
            cam.StopCapture()

        # How far the mount actually moved during this cycle, so we can
        # reverse it at the fast rate to get back to the start.
        displacement = abs(ra_diff_arcsec(mount.RA, limb_ra))
        return_time  = displacement / (RETURN_MULTIPLIER * SIDEREAL)

        cycle_elapsed = time.time() - cycle_start
        print("  {} {:.0f}\" moved. Cycle: {:.1f}s".format(
              "Scan done." if ENABLE_CAPTURE else "Dry run done.",
              displacement, cycle_elapsed))

        # Fast return to start position; skipped on the last cycle since the
        # post-loop "Repositioning to disk center" block handles that case.
        if cycle < NUM_CYCLES - 1:
            mount.MoveAxis(0, -RETURN_MULTIPLIER)
            time.sleep(return_time)
            mount.Stop()
            time.sleep(0.5)

    # After all cycles: park on disk centre. We were at the trailing-edge end
    # of the last scan; reverse for return_time, then re-find the leading
    # limb and step half a disk-diameter forward to the centre.
    print("\nRepositioning to disk center...")
    mount.MoveAxis(0, -RETURN_MULTIPLIER)
    time.sleep(return_time)
    mount.Stop()
    time.sleep(0.5)

    if find_solar_limb():
        half_t = (SOLAR_DIAMETER / 2.0) / (RETURN_MULTIPLIER * SIDEREAL)
        mount.MoveAxis(0, RETURN_MULTIPLIER)
        time.sleep(half_t)
        mount.Stop()
        print("Centered on solar disk.")

    mount.Stop()

except KeyboardInterrupt:
    # Set the abort flag so any handlers/watchers still spinning will exit
    # their wait loops on the next check_abort(). Then make a best-effort
    # attempt to stop the mount and any in-progress capture — wrapped in
    # try/except because we may be in any state (mount disconnected,
    # capture not actually running, etc.) and we never want the cleanup
    # itself to throw.
    print("\n\n*** ABORTED BY USER ***")
    _abort[0] = True
    try:
        mount.Stop()
    except:
        pass
    try:
        if cam.Capturing:
            cam.StopCapture()
    except:
        pass

# ── Session summary ───────────────────────────────────────────────────────────
total_time = time.time() - session_start

print("")
print("=" * 60)
print("SESSION COMPLETE")
print("=" * 60)
print("Cycles: {}  Total: {:.0f}s ({:.1f} min)  Avg: {:.1f}s/cycle".format(
      len(drift_log), total_time, total_time / 60,
      total_time / len(drift_log) if drift_log else 0))

if len(drift_log) > 1:
    ref  = drift_log[0]
    last = drift_log[-1]
    dt   = last[1] - ref[1]
    dra  = ra_diff_arcsec(last[2], ref[2])
    ddec = (last[3] - ref[3]) * 3600

    print("")
    print("POLAR ALIGNMENT DRIFT")
    print("RA  : {:+.1f}\"  ({:+.1f}\"/min)".format(dra, dra / dt * 60 if dt > 0 else 0))
    print("Dec : {:+.1f}\"  ({:+.1f}\"/min)".format(ddec, ddec / dt * 60 if dt > 0 else 0))

# If drift exceeds 5"/min in either axis, suggest a polar-alignment tweak.
    # The classic drift-alignment heuristic: large RA drift = azimuth error;
    # large Dec drift = altitude error. The east/west and up/down hints are
    # rules of thumb for the northern hemisphere — flip them in the south.
    if dt > 0 and (abs(dra / dt * 60) > 5 or abs(ddec / dt * 60) > 5):
        if abs(dra) > abs(ddec):
            print("-> Adjust azimuth slightly {}".format(
                  "east" if dra > 0 else "west"))
        else:
            print("-> Adjust altitude slightly {}".format(
                  "up" if ddec > 0 else "down"))
    else:
        print("-> Alignment is good for solar work.")

print("=" * 60)
