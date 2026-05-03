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
clr.AddReference("System.Drawing")
from System.Drawing import Rectangle

# ══════════════════════════════════════════════════════════════════════════════
# PARAMETERS
# ══════════════════════════════════════════════════════════════════════════════

# ── Hardware ──────────────────────────────────────────────────────────────────
FOCAL_LENGTH       = 714     # mm
PIXEL_SIZE         = 2       # um (G3M678M)
SIDEREAL           = 15.04   # arcsec/s per sidereal multiplier unit
SOLAR_DIAMETER     = 1920.0  # arcsec

# ── Scanning ──────────────────────────────────────────────────────────────────
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

PIXEL_SCALE = 206265 * (PIXEL_SIZE / 1000.0) / FOCAL_LENGTH

mount = SharpCap.Mounts.SelectedMount
cam   = SharpCap.SelectedCamera

_abort = [False]  # flipped by KeyboardInterrupt; checked in every wait loop

def check_abort():
    if _abort[0]:
        raise KeyboardInterrupt()

def ra_diff_arcsec(ra1, ra2):
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
class FrameWatcher:
    """Attaches a FrameCaptured handler and latches the most recent (mean, stddev)."""
    def __init__(self, camera):
        self.camera = camera
        self.latest = None
        self._active = False

    def start(self):
        self.camera.FrameCaptured += self._on_frame
        self._active = True

    def stop(self):
        if self._active:
            try:
                self.camera.FrameCaptured -= self._on_frame
            except:
                pass
            self._active = False

    def _on_frame(self, sender, args):
        if not self._active:
            return
        try:
            s = args.Frame.GetStats()
            self.latest = (s.Item1, s.Item2)
        except:
            pass

    def wait_first(self, timeout=3.0):
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
def measure_left_right_means(timeout=3.0):
    """Capture one live frame, split into left and right halves, return means."""
    result = [None]

    def on_frame(sender, args):
        if result[0] is not None:
            return
        try:
            f = args.Frame
            roi = cam.ROI
            w = roi.Width
            h = roi.Height
            mid = w // 2
            left  = f.CutROI(Rectangle(0,   0, mid,     h))
            right = f.CutROI(Rectangle(mid, 0, w - mid, h))
            ls = left.GetStats()
            rs = right.GetStats()
            result[0] = (ls.Item1, rs.Item1)
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
        print("  CutROI error: {}".format(result[0][1]))
        return None
    return result[0]

# ── Limb detection ────────────────────────────────────────────────────────────
# Using stddev (GetStats().Item2), not mean brightness, as the on-disk vs
# off-disk signal is an idea taken from Patrick Hsieh's SHGScan. The dynamic
# dark-baseline thresholding around it is local to this script.
def find_solar_limb():
    watcher = FrameWatcher(cam)
    watcher.start()
    try:
        if not watcher.wait_first():
            print("  No frames from camera")
            return False

        initial_stddev = watcher.latest[1]

        # If clearly on disk, back off until we see dark sky first.
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
                    time.sleep(0.3)
                    break
                time.sleep(LIMB_SAMPLE_INTERVAL)
            if not off_disk:
                mount.Stop()
                print(" FAILED")
                return False

        # Sample dark baseline and compute dynamic threshold.
        time.sleep(0.2)
        dark = watcher.latest[1]
        threshold = max(dark * LIMB_DARK_MULT, dark + 300.0, LIMB_STDDEV_MIN)

        # Slew forward looking for the stddev jump at the limb.
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
    """Nudge Dec until left-half and right-half means are balanced."""
    # On a German Equatorial mount the camera orientation flips after a
    # meridian flip, so the Dec nudge direction has to be inverted on the
    # opposite side of the pier. ASCOM SideOfPier: 0 = pierEast, 1 = pierWest.
    try:
        side = int(mount.SideOfPier)
    except:
        side = 0
        print("  Dec: SideOfPier not exposed by driver, assuming pierEast")
    pier_flip = -1 if side == 1 else 1

    for iteration in range(1, DEC_MAX_CORRECTIONS + 1):
        check_abort()
        means = measure_left_right_means()
        if means is None:
            print("  Dec: frame read failed")
            return False
        l, r = means
        denom = l + r
        if denom <= 0:
            print("  Dec: no signal")
            return False
        offset = (l - r) / denom  # + => sun pushed left, - => pushed right

        print("  Dec {} (side={}): L={:.0f} R={:.0f} offset={:+.3f}".format(
              iteration, side, l, r, offset), end="")

        if abs(offset) < DEC_CENTER_TOLERANCE:
            print("  OK")
            return True

        # Scale nudge to the magnitude of the imbalance.
        nudge_time = min(abs(offset) * 8.0, 1.5)
        nudge_time = max(nudge_time, 0.1)
        # +offset means left half brighter => sun is pushed left in the frame
        # => nudge Dec in the direction that moves the sun right (toward center).
        direction = (-1 if offset > 0 else 1) * pier_flip

        mount.MoveAxis(1, direction * DEC_NUDGE_SPEED)
        time.sleep(nudge_time)
        mount.Stop()
        time.sleep(DEC_SETTLE_TIME)
        print("  nudge {:.2f}s {}".format(nudge_time, "+" if direction > 0 else "-"))

    print("  Dec: max corrections")
    return False

# ── Histogram percentile measurement ──────────────────────────────────────────
def measure_percentile_pixel(pct, timeout=3.0):
    """Capture one frame, return pixel value at the given cumulative percentile."""
    result = [None]

    def on_frame(sender, args):
        if result[0] is not None:
            return
        try:
            h = args.Frame.CalculateHistogram()
            bins = h.Values[0]
            nbins = len(bins)
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
    fs = full_scale()
    target_val = EXP_TARGET_LEVEL * fs
    tol_val    = EXP_TOLERANCE * fs

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
            # Cap step size to prevent runaway on noisy frames / mount-settling artefacts.
            ratio = target_val / p
            ratio = max(0.5, min(2.0, ratio))
            new_exp = current_exp * ratio
            exp_ctrl.Value = new_exp
            time.sleep(0.3)
            print("  -> {:.2f}ms".format(new_exp))
        else:
            print("  (no signal)")
            exp_ctrl.Value = entry_exp
            return False

    # Fell through without converging — don't leave a possibly-bad value.
    print("  Auto-exp: max attempts, reverting to {:.2f}ms".format(entry_exp))
    exp_ctrl.Value = entry_exp
    return False

# ── Alignment routines ────────────────────────────────────────────────────────
def do_full_alignment():
    print("\n  Full alignment...")
    if not find_solar_limb():
        return False
    limb_ra = mount.RA

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
fps = cam.CurrentFrameRate
scan_multiplier = PIXEL_SCALE * fps / SIDEREAL
scan_speed      = scan_multiplier * SIDEREAL
crossing_time   = SOLAR_DIAMETER / scan_speed
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

        need_full = (cycle == 0) or (cycle % DEC_INTERVAL == 0)

        if need_full:
            limb_ra = do_full_alignment()
        else:
            limb_ra = do_quick_realign()

        if limb_ra is False:
            print("ABORTING: Could not find solar limb.")
            break

        limb_dec = mount.Dec
        drift_log.append((cycle + 1, time.time(), limb_ra, limb_dec))

        if len(drift_log) > 1:
            ref = drift_log[0]
            dt  = drift_log[-1][1] - ref[1]
            dra = ra_diff_arcsec(drift_log[-1][2], ref[2])
            ddec = (drift_log[-1][3] - ref[3]) * 3600
            if dt > 0:
                print("  Drift: RA {:+.1f}\"  Dec {:+.1f}\"  ({:.0f}s)".format(
                      dra, ddec, dt))

        # Backup past limb
        mount.MoveAxis(0, -LIMB_SEARCH_SPEED)
        time.sleep(PADDED_DURATION)
        mount.Stop()
        time.sleep(0.2)

        # Scan
        if ENABLE_CAPTURE:
            cam.PrepareToCapture()
            cam.RunCapture()
        mount.MoveAxis(0, scan_multiplier)
        time.sleep(scan_duration)
        mount.Stop()
        if ENABLE_CAPTURE:
            cam.StopCapture()

        displacement = abs(ra_diff_arcsec(mount.RA, limb_ra))
        return_time  = displacement / (RETURN_MULTIPLIER * SIDEREAL)

        cycle_elapsed = time.time() - cycle_start
        print("  {} {:.0f}\" moved. Cycle: {:.1f}s".format(
              "Scan done." if ENABLE_CAPTURE else "Dry run done.",
              displacement, cycle_elapsed))

        # Fast return
        if cycle < NUM_CYCLES - 1:
            mount.MoveAxis(0, -RETURN_MULTIPLIER)
            time.sleep(return_time)
            mount.Stop()
            time.sleep(0.5)

    # Return to disk center
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
