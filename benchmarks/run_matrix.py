"""Run every engine across every profile (each in its own process) and print a
markdown table grouped by profile. Drives bench_engines.py.

    python run_matrix.py
"""

import platform
import subprocess
import sys
from pathlib import Path

ENGINES = ["fastxlsx", "openpyxl_writeonly", "xlsxwriter", "xlsxwriter_constant", "pandas"]
PROFILES = ["mixed", "numeric", "strings", "wide"]
PROFILE_DESCRIPTIONS = {
    "mixed": "200k rows x 5 cols, mixed types (str/int/float/bool)",
    "numeric": "200k rows x 5 cols, all numeric",
    "strings": "200k rows x 5 cols, all unique strings (stresses shared strings)",
    "wide": "20k rows x 50 cols, mixed types (wide rows)",
}
BENCH = Path(__file__).with_name("bench_engines.py")


def run_one(engine, profile):
    proc = subprocess.run(
        [sys.executable, str(BENCH), engine, profile],
        capture_output=True, text=True, check=True,
    )
    _, _, _, time_s, rss_mb = proc.stdout.split()
    return float(time_s), float(rss_mb)


def main():
    results = {(e, p): run_one(e, p) for p in PROFILES for e in ENGINES}

    print(f"Machine: {platform.platform()}, Python {platform.python_version()}\n")
    for profile in PROFILES:
        print(f"### {profile} — {PROFILE_DESCRIPTIONS[profile]}\n")
        print("| Engine | Time (s) | Peak RSS (MB) |")
        print("|---|---:|---:|")
        for engine in ENGINES:
            time_s, rss_mb = results[(engine, profile)]
            print(f"| {engine} | {time_s:.2f} | {rss_mb:.0f} |")
        print()


if __name__ == "__main__":
    main()
