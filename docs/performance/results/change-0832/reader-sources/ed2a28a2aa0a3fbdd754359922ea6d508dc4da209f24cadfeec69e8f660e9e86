"""Root-owned stage launcher; preserve terminal failures and driver identity."""
import subprocess
import sys
import time

import driver as d


def main():
    stage = sys.argv[1]
    assert stage in ("prepare", "before", "install", "after", "capture")
    folder = d.P / "stages" / stage
    folder.mkdir(parents=True)
    row = {"argv": ["python3", "-B", str(d.P / "driver.py"), stage],
           "cwd": str(d.ROOT), "driver_sha256": d.sha(d.P / "driver.py"),
           "started_unix": time.time()}
    d.write(folder / "started.json", row)
    with (folder / "output.log").open("xb") as log:
        child = subprocess.run(row["argv"], cwd=d.ROOT, stdout=log,
                               stderr=subprocess.STDOUT)
    row.update(exit_code=child.returncode, finished_unix=time.time(),
               output_sha256=d.sha(folder / "output.log"))
    d.write(folder / "receipt.json", row)
    print(stage, child.returncode, flush=True)
    assert d.sha(d.P / "driver.py") == row["driver_sha256"]
    assert child.returncode == 0


if __name__ == "__main__":
    main()
