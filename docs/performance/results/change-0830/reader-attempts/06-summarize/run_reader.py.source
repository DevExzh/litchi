"""Retain exact reader sources and every attempt, including failed attempts."""
import shutil
import sys

import driver as d


def main():
    script = sys.argv[1]
    assert script in ("heaptrack_preflight.py", "analyze.py", "audit.py", "summarize.py")
    extra = sys.argv[2:]
    assert not extra or (script == "analyze.py" and extra == ["--check"])
    attempts = d.P / "reader-attempts"
    attempts.mkdir(exist_ok=True)
    number = len(list(attempts.iterdir()))
    folder = attempts / f"{number:02}-{script.removesuffix('.py')}"
    folder.mkdir()
    sources = {path.name: d.sha(path) for path in d.P.glob("*.py")}
    for name in sources:
        shutil.copyfile(d.P / name, folder / (name + ".source"))
    d.write(folder / "sources.json", sources)
    try:
        d.run(f"reader-{number:02}-{script.removesuffix('.py')}", ["python3", "-B", d.P / script, *extra])
    finally:
        assert sources == {name: d.sha(d.P / name) for name in sources}, "reader source changed during execution"


if __name__ == "__main__":
    main()
