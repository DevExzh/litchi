# xml-substrate survey: scratch evidence summary (synthetic, survey-only)

Host: AMD EPYC 9R45, callgrind 3.26, rustc 1.95, release + debug=1, taskset -c 3.
Probe: xmlprobe/ (path deps on litchi-xlsx, litchi-ooxml-common at HEAD 2fc5fc657).
Fixture: test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx
  (sheet1.xml 209,931 B, 680 rows, 5,440 cells, 11,578 elements, root declares
   xmlns, xmlns:r, xmlns:mc, xmlns:x14ac + mc:Ignorable="x14ac", 681 x14ac:dyDescent).
Control: same package with only x14ac:dyDescent / xmlns:mc / mc:Ignorable / xmlns:x14ac
  attributes removed from sheet1.xml (strip_markers.py; 194,257 B, same rows/cells).
Cell read: H680 (last cell, t="s", s="7").

Whole-process callgrind Ir (includes process setup; single run each):
  eager  real    1,037,845,761  process_markup_compatibility 849,541,700 (81.9%)  Parser::parse 152,616,896 (14.7%)  x14ac::capture 26,566,647 (2.6%)
  eager  control    80,541,821  Parser::parse 59,371,665 (73.7%)  process_markup_compatibility (scan only) 13,849,335 (17.2%)
  source real    1,075,496,653  codec 733,249,619 (68.2%)  selected stream 154,247,846 (14.3%)  Parser::parse 153,218,873 (14.3%)  x14ac 26,467,715 (2.5%)
  source control   221,112,686  selected stream 140,687,894 (63.6%; Processor::start 101,477,345)  eager_cell/store 76,520,815 (34.4%; Parser::parse 59,739,170; styles::parse 9,690,888)
  (source path: worksheet marked NotEligible(Styles) at <cols>, stream still runs to EOF, then store parse)
Files: cg-{eager,source}-{real,control}.{inclusive,self}.txt

MCE codec alone (mode mce, process_ooxml on the extracted sheet1.xml):
  real:    in 209,931 -> out 3,540,261 B (16.9x), owned; 841,817,148 Ir total (4,008 Ir/input byte)
           start 818.3M (97.2%, 11,578 calls); esc 687.4M (81.7%, 58,794 calls);
           BoundedOutput::extend_from_slice 2,833,883 calls; __rust_realloc 2,954,106 calls 327.0M (38.8%)
           output has 46,312 "xmlns" (input 4): every element re-declares all 4 in-scope bindings
  control: in 194,257 -> out 194,257 B, borrowed; 5,003,353 Ir (4,660,836 in the codec = naive windows() scan, 24 Ir/byte)
  DOCX testComment.docx word/document.xml: 3,108 -> 50,987 B (16.4x), 31 root xmlns, 23 elements
  PPTX shapes-litchi.pptx ppt/slides/slide1.xml: 27,582 -> 385,997 B (14.0x), 6 root xmlns, 850 elements

Fixture breadth (first 60 .xlsx under test-data, sheet1.xml): 41 declare the mc namespace,
  39 contain x14ac/dyDescent, 27 contain <cols>, 8 have row s=/customFormat,
  19 pass every raw::worksheet::source_stream_eligible gate.
