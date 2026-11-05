"""Handles used from several threads (bug B5, fixed in S8a; ``File``, R9).

Every class is ``frozen`` and thread-safe over the ``Send + Sync``
``litchi::File``, and the document-wide reads release the GIL. These tests
share handles across threads and check that every thread reads the same
values the opening thread did. They run in a subprocess, as the B5 test does,
so a regression that panics cannot abort the whole run (release builds use
``panic = "abort"``).

``FileThreadTests`` does the same for the unified ``File``: shared reads,
a handle opened on a worker, ``close()`` racing reads on other threads, and
an observation that the reads release the GIL: a ticker thread must run
while one long read is in progress.
"""

from __future__ import annotations

import ast
import unittest

from _support import TEST_DATA, run_isolated

SHARED_READS = '''
import sys, threading
import litchi_py as L

root = {root!r} + "/"
handles = {{
    "docx": L.Document.open(root + "ooxml/docx/table-alignment.docx"),
    "doc": L.Document.open(root + "ole/doc/HeaderFooterUnicode.doc"),
    "pages": L.Document.open(root + "iwork/pages/body-sections-unicode.pages"),
    "pptx": L.Presentation.open(root + "ooxml/pptx/shapes.pptx"),
    "ppt": L.Presentation.open(root + "ole/ppt/SampleShow.ppt"),
    "key": L.Presentation.open(root + "iwork/keynote/basic.key"),
    "xlsx": L.Workbook.open(root + "ooxml/xlsx/SimpleNormal.xlsx"),
    "xls": L.Workbook.open(root + "ole/xls/SimpleMultiCell.xls"),
    "numbers": L.Workbook.open(root + "iwork/numbers/cell-value-native.numbers"),
}}
paragraph = handles["docx"].paragraphs()[4]
table = handles["docx"].tables()[0]
slide = handles["pptx"].slides()[1]

def snapshot():
    values = []
    for name in ("docx", "doc", "pages"):
        document = handles[name]
        values.append(document.text())
        values.append([p.text() for p in document.paragraphs()])
        values.append([[c.text() for r in t.rows() for c in r.cells()] for t in document.tables()])
    for name in ("pptx", "ppt", "key"):
        presentation = handles[name]
        values += [presentation.text(), presentation.slide_count(), [s.text() for s in presentation.slides()]]
    for name in ("xlsx", "xls", "numbers"):
        workbook = handles[name]
        values += [workbook.text(), workbook.worksheet_count(), workbook.worksheet_names()]
    values.append([(r.text(), r.bold(), r.italic(), r.underline()) for r in paragraph.runs()])
    values.append([c.text() for r in table.rows() for c in r.cells()])
    values.append((slide.text(), repr(slide), repr(paragraph)))
    return values

expected = snapshot()
matches, failures = [], []

def work():
    try:
        for _ in range({rounds}):
            matches.append(snapshot() == expected)
    except BaseException as error:
        failures.append(repr(error))

threads = [threading.Thread(target=work) for _ in range({threads})]
for thread in threads:
    thread.start()
for thread in threads:
    thread.join()
print(len(matches), all(matches), failures)
'''

OPEN_IN_A_WORKER = '''
import threading
import litchi_py as L

root = {root!r} + "/"
opened = {{}}

def open_all():
    opened["document"] = L.Document.open(root + "ooxml/docx/Hyperlink.docx")
    opened["paragraphs"] = opened["document"].paragraphs()
    opened["workbook"] = L.Workbook.open(root + "iwork/numbers/basic.numbers")
    try:
        L.Presentation.open(root + "ooxml/docx/Hyperlink.docx")
    except L.UnsupportedError as error:
        opened["error"] = error

worker = threading.Thread(target=open_all)
worker.start()
worker.join()
# Handles and the exception made on the worker are used on the main thread.
print(repr(opened["document"].text()))
print([p.text() for p in opened["paragraphs"]])
print(opened["workbook"].worksheet_names())
print(opened["error"].code)
'''


SHARED_FILES = '''
import threading
import litchi_py as L

root = {root!r} + "/"
docx = L.open(root + "ooxml/docx/drawing.docx")
xlsx = L.open(root + "ooxml/xlsx/ExcelTables.xlsx")
signed = L.open(root + "poi/test-data/xmldsign/hello-world-signed.docx")

def snapshot():
    return [
        docx.get("/body/p[10]"),
        docx.get("/body/tbl[1]", 2),
        docx.find("/body/p[999]"),
        docx.query("paragraph", 25),
        docx.view("text"),
        docx.view("stats", json=True),
        docx.capabilities(),
        docx.text(),
        xlsx.get("/", 2),
        xlsx.query("cell"),
        xlsx.view("outline", json=True),
        xlsx.capabilities(),
        xlsx.metadata(),
        signed.get("/body", 2),
    ]

expected = snapshot()
matches, failures = [], []

def work():
    try:
        for _ in range({rounds}):
            matches.append(snapshot() == expected)
    except BaseException as error:
        failures.append(repr(error))

threads = [threading.Thread(target=work) for _ in range({threads})]
for thread in threads:
    thread.start()
for thread in threads:
    thread.join()
print(len(matches), all(matches), failures)
'''

FILE_FROM_A_WORKER = '''
import threading
import litchi_py as L

root = {root!r} + "/"
made = {{}}

def open_all():
    made["file"] = L.open(root + "ooxml/xlsx/SimpleNormal.xlsx")
    made["view"] = made["file"].get("/Sheet1/row[1]")
    try:
        made["file"].get("/Sheet1/[")
    except L.PathError as error:
        made["error"] = error

worker = threading.Thread(target=open_all)
worker.start()
worker.join()
# The file, a view and an exception made on the worker, used here.
print(made["file"].get("/Sheet1/row[1]") == made["view"])
print([cell["path"] for cell in made["view"]["children"]])
print(made["error"].code, made["file"].view("stats", json=True)["totalCells"])
'''

CLOSE_WHILE_READING = '''
import threading, time
import litchi_py as L

file = L.open({root!r} + "/ooxml/docx/drawing.docx")
expected = file.get("/body", 1, max_nodes=10_000)
outcomes, lock = {{}}, threading.Lock()
started = threading.Barrier({threads} + 1)

def read():
    started.wait()
    while True:
        try:
            key = "same" if file.get("/body", 1, max_nodes=10_000) == expected else "different"
        except L.IoError as error:
            key = "closed" if error.code == "file_required" else repr(error)
        except BaseException as error:
            key = repr(error)
        with lock:
            outcomes[key] = outcomes.get(key, 0) + 1
        if key != "same":
            return

threads = [threading.Thread(target=read) for _ in range({threads})]
for thread in threads:
    thread.start()
started.wait()
while not outcomes.get("same"):
    time.sleep(0.001)
file.close()
for thread in threads:
    thread.join()
print(sorted(outcomes), outcomes["closed"], file.closed)
'''

# A read that takes long enough to observe must let a ticker thread run while
# it is in progress. Each attempt opens the file afresh, so the read decodes
# the sheet again. Ticks count only when they fall MARGIN inside the call, and
# a read counts as releasing the GIL from MIN_TICKS ticks on: a C call that
# holds the GIL can still let one tick through (the json.loads control, which
# the test also runs, shows one), while the reads show dozens to thousands.
GIL_RELEASE = '''
import json, threading, time
import litchi_py as L

path = {root!r} + "/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx"
MARGIN, MIN_TICKS = 0.005, 10
ticks, running = [], True

def tick():
    while running:
        ticks.append(time.perf_counter())
        time.sleep(0.0005)

def observe(read):
    seen = "short"
    for _ in range(3):
        file = L.open(path)
        begin = time.perf_counter()
        read(file)
        end = time.perf_counter()
        if end - begin < 4 * MARGIN:
            continue
        inside = sum(1 for moment in list(ticks) if begin + MARGIN < moment < end - MARGIN)
        if inside >= MIN_TICKS:
            return "released"
        seen = "held"
    return seen

control = json.dumps([{{"a": index, "b": str(index)}} for index in range(400_000)])
ticker = threading.Thread(target=tick, daemon=True)
ticker.start()
time.sleep(0.02)
reads = {{
    "text": lambda file: file.text(),
    "view": lambda file: file.view("text"),
    "view_json": lambda file: file.view("text", json=True),
    "get": lambda file: file.get("/sheet[1]/A1:J7577", max_nodes=10**6),
    "query": lambda file: file.query("cell"),
    "zz control: json.loads": lambda file: json.loads(control),
}}
results = {{name: observe(read) for name, read in reads.items()}}
running = False
print(sorted(results.items()))
'''


class FileThreadTests(unittest.TestCase):
    def test_shared_files_read_the_same_values_on_every_thread(self):
        result = run_isolated(SHARED_FILES.format(root=str(TEST_DATA), rounds=3, threads=8), timeout=600)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "24 True []", result.stderr)

    def test_a_file_opened_on_a_worker_is_read_on_the_main_thread(self):
        result = run_isolated(FILE_FROM_A_WORKER.format(root=str(TEST_DATA)))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            result.stdout.splitlines(),
            ["True", str(["/sheet[1]/A1", "/sheet[1]/C1"]), "invalid_path 19"],
        )

    def test_close_while_other_threads_read(self):
        # Reads in progress finish on their own reference; later ones raise
        # IoError (file_required). Nothing else happens.
        result = run_isolated(CLOSE_WHILE_READING.format(root=str(TEST_DATA), threads=4), timeout=600)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "['closed', 'same'] 4 True", result.stderr)

    def test_long_reads_release_the_gil(self):
        result = run_isolated(GIL_RELEASE.format(root=str(TEST_DATA)), timeout=600)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            result.stdout.strip(),
            str(
                [
                    ("get", "released"),
                    ("query", "released"),
                    ("text", "released"),
                    ("view", "released"),
                    ("view_json", "released"),
                    # The method's control: a C call that holds the GIL.
                    ("zz control: json.loads", "held"),
                ]
            ),
            result.stderr,
        )


# Writes from many threads on one shared File: every thread's result holds its
# own change and nothing else, and the shared File never changes.
SHARED_WRITES = '''
import threading
import litchi_py as L

base = L.create("xlsx")
base = base.set("/sheet[1]/A1", {{"value": "base"}})
before = base.to_bytes()
results, failures = {{}}, []

def work(index):
    try:
        mine = base
        for round in range({rounds}):
            mine = mine.set("/sheet[1]/B%d" % (index + 1), {{"value": index * 100 + round}})
            items = [{{"command": "set", "path": "/sheet[1]/C%d" % (index * 50 + row + 1),
                       "props": {{"value": index}}}} for row in range(50)]
            mine = mine.batch(items).file
        def staged(tx):
            tx.set("/sheet[1]/D%d" % (index + 1), {{"value": index}})
        mine = mine.edit(staged).file
        results[index] = mine
    except BaseException as error:
        failures.append(repr(error))

threads = [threading.Thread(target=work, args=(index,)) for index in range({threads})]
for thread in threads:
    thread.start()
for thread in threads:
    thread.join()

ok = []
for index, mine in sorted(results.items()):
    ok.append(mine.get("/sheet[1]/B%d" % (index + 1), 0)["text"] == str(index * 100 + {rounds} - 1))
    ok.append(mine.get("/sheet[1]/D%d" % (index + 1), 0)["text"] == str(index))
    ok.append(len(mine.query("cell[type=Number]")) == 1 + 50 + 1)
print(len(results), all(ok), base.to_bytes() == before, failures)
'''

# One Edit shared by many threads: calls on it run one at a time, and the
# commit holds every staged operation.
SHARED_EDIT = '''
import threading
import litchi_py as L

base = L.create("xlsx")
tx = base.edit()
failures = []

def work(index):
    try:
        for row in range({rows}):
            tx.set("/sheet[1]/%s%d" % ("ABCDEFGH"[index], row + 1), {{"value": row}})
        tx.get("/sheet[1]/A1", 0)
    except BaseException as error:
        failures.append(repr(error))

threads = [threading.Thread(target=work, args=(index,)) for index in range({threads})]
for thread in threads:
    thread.start()
for thread in threads:
    thread.join()
commit = tx.commit()
print(commit.operations, len(commit.file.query("cell")), commit.changed, failures,
      base.find("/sheet[1]/A1", 0)["text"])
'''

# Writes long enough to observe must release the GIL, as the reads do.
GIL_RELEASE_WRITES = '''
import json, os, tempfile, threading, time
import litchi_py as L

path = {root!r} + "/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx"
MARGIN, MIN_TICKS = 0.005, 10
ticks, running = [], True
directory = tempfile.mkdtemp()

def tick():
    while running:
        ticks.append(time.perf_counter())
        time.sleep(0.0005)

items = [{{"command": "set", "path": "/sheet[1]/A%d" % (row + 1), "props": {{"value": row}}}}
         for row in range(9_000)]
# Saving an edited workbook rewrites its sheet: long enough to observe.
edited = L.open(path).batch(items).file

def observe(write):
    seen = "short"
    for attempt in range(3):
        file = L.open(path)
        begin = time.perf_counter()
        write(file, items, attempt)
        end = time.perf_counter()
        if end - begin < 4 * MARGIN:
            continue
        inside = sum(1 for moment in list(ticks) if begin + MARGIN < moment < end - MARGIN)
        if inside >= MIN_TICKS:
            return "released"
        seen = "held"
    return seen

ticker = threading.Thread(target=tick, daemon=True)
ticker.start()
time.sleep(0.02)
writes = {{
    "open": lambda file, items, attempt: L.open(path),
    "batch": lambda file, items, attempt: file.batch(items),
    "save_as": lambda file, items, attempt: edited.save_as(os.path.join(directory, "%d.xlsx" % attempt)),
    "to_bytes": lambda file, items, attempt: edited.to_bytes(),
}}
results = {{name: observe(write) for name, write in writes.items()}}
running = False
print(sorted(results.items()))
'''


class WriteThreadTests(unittest.TestCase):
    def test_writes_on_a_shared_file_are_independent(self):
        result = run_isolated(SHARED_WRITES.format(rounds=3, threads=8), timeout=600)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "8 True True []", result.stderr)

    def test_a_shared_edit_serializes_its_calls(self):
        result = run_isolated(SHARED_EDIT.format(rows=25, threads=8), timeout=600)
        self.assertEqual(result.returncode, 0, result.stderr)
        # 8 × 25 staged sets; the shared base File is unchanged.
        self.assertEqual(result.stdout.strip(), "200 200 True [] (empty)", result.stderr)

    def test_long_writes_release_the_gil(self):
        result = run_isolated(GIL_RELEASE_WRITES.format(root=str(TEST_DATA)), timeout=600)
        self.assertEqual(result.returncode, 0, result.stderr)
        results = dict(ast.literal_eval(result.stdout.strip()))
        # A source-backed open reads only the container index, which is
        # often too short to observe; it must never be seen holding the GIL.
        self.assertIn(results.pop("open"), ("released", "short"))
        self.assertEqual(
            results,
            {"batch": "released", "save_as": "released", "to_bytes": "released"},
            result.stderr,
        )


class CrossThreadTests(unittest.TestCase):
    def test_shared_handles_read_the_same_values_on_every_thread(self):
        result = run_isolated(SHARED_READS.format(root=str(TEST_DATA), rounds=5, threads=8))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "40 True []", result.stderr)

    def test_handles_opened_on_a_worker_are_read_on_the_main_thread(self):
        result = run_isolated(OPEN_IN_A_WORKER.format(root=str(TEST_DATA)))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            result.stdout.splitlines(),
            [
                repr("Following is a hyperlink.\nEricWhite.com"),
                str(["Following is a hyperlink.", "EricWhite.com"]),
                str(["Sheet 1"]),
                "unsupported_format",
            ],
        )


if __name__ == "__main__":
    unittest.main()
