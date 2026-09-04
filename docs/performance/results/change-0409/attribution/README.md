# XLSX selected-cell CPU attribution

`attribute_0409.py` parses period-weighted `perf script --no-inline` blocks. Frame zero is the leaf; callers follow toward the root. The exact facade cell ancestor selects a subset of the query timer. Whole-process and selected denominators are explicit; immediate-caller rows additionally have a leaf-specific denominator. Inclusive rows overlap, and sample period sums are not elapsed phase timers. Frequency adaptation across short fresh children may bias whole-process phase shares.

The input hashes and original Git-blob hashes in `attribution-summary.json` bind the captured revision. Reproduce with `python3 attribute_0409.py --script PATH_TO_DECOMPRESSED_SCRIPT --report perf-all-self.stdout --capture ../initial --repo PATH_TO_REPOSITORY_WITH_ORIGINAL_HISTORY --output reproduced.json`. Original default paths and exact postprocessing argv are in commands.txt. Paths in the reproduced JSON will reflect the chosen location; counts and periods should match.

The initial supplementary parser draft used the leaf denominator for a field named share_of_timed_percent in immediate-caller rows; the published parser corrects that field and filters the selected scan/x14ac subsets by the same exact timed ancestor. No workload was rerun for this postprocessing correction.
