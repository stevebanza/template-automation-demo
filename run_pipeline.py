"""Command-line entry point.

    python run_pipeline.py all                 # whole cycle, end to end
    python run_pipeline.py validate            # one step (after the earlier ones)
    python run_pipeline.py all --out runs/test # write somewhere else
"""
import argparse
import errno
import os
import sys

from templatekit.pipeline import Pipeline

STEPS = ["generate", "rollforward", "translate", "prefill", "simulate", "validate", "consolidate"]


def main() -> int:
    """Parse arguments, run the requested step(s) and print a summary."""
    ap = argparse.ArgumentParser(description="Excel template automation demo pipeline")
    ap.add_argument("step", choices=STEPS + ["all"])
    ap.add_argument("--config", default="config.yaml")
    ap.add_argument("--out", default="output")
    a = ap.parse_args()

    pl = Pipeline(a.config, a.out)
    if a.step == "all":
        s = pl.run_all()
        print(f"Release database: {s['db']}  ({s['rows']} rows, {s['revisions']} revisions)")
        print(f"Templates loaded (country + mined): {s['loaded']}, "
              f"returned to country: {s['excluded']}")
        for r in pl.results:
            print(f"  {r.country_code}  {r.status:<22} must fix {r.count('blocking')}, "
                  f"confirm {r.count('to_confirm')}, info {r.count('info')}")
        print(f"Review reports: {pl.out / 'reports' / 'index.html'}")
    else:
        print(getattr(pl, a.step)())
    return 0


def _run() -> int:
    """Run main(), exiting quietly if stdout is closed early (e.g. `... | head`)."""
    # Piped output on Windows uses the ANSI code page: never crash on a character it lacks.
    sys.stdout.reconfigure(errors="backslashreplace")
    try:
        code = main()
        sys.stdout.flush()
    except (BrokenPipeError, OSError) as exc:
        # Windows reports a closed pipe as EINVAL rather than EPIPE.
        if not isinstance(exc, BrokenPipeError) and exc.errno != errno.EINVAL:
            raise
        # Point stdout at devnull so the interpreter's final flush does not fail again.
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        return 1
    return code


if __name__ == "__main__":
    sys.exit(_run())
