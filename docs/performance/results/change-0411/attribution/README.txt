0411 XLS attribution artifacts

The frozen postprocessor is /tmp/litchi-goal-0411-attribute.py.  The exact
argv used for both captures, their report inputs, output paths, and log paths
are in commands.json.  The JSON reports are eager-attribution.json and
source-backed-attribution.json; findings.txt is a short human-readable
extract.

The postprocessor binds source files with `git show <revision>:<path>`.  The
capture metadata must therefore be available as materialized JSON while it is
run.  If an evidence bundle stores a profile as `*.json.zst`, create a
temporary capture metadata directory and decompress the profile next to the
other metadata before invoking the recorded argv:

    mkdir -p /tmp/litchi-goal-0411-capture-json
    cp /path/to/build-identity.json /tmp/litchi-goal-0411-capture-json/
    cp /path/to/commands.json /tmp/litchi-goal-0411-capture-json/
    zstd -dc /path/to/eager-profile.json.zst > /tmp/litchi-goal-0411-capture-json/eager-profile.json
    zstd -dc /path/to/source_backed-profile.json.zst > /tmp/litchi-goal-0411-capture-json/source_backed-profile.json

Use the temporary directory as `--capture`; keep the perf script and normal
`perf report --no-inline` output at the paths recorded in commands.json.  The
profile JSON is used for the selected case and 1000/20 identity checks;
source binding itself comes from build-identity.json's captured revision.

All reported weights are cycles:u sample periods.  They are CPU attribution,
not phase latency.  The `run_xls_source_backed_case` row is an observed parent
stack subset and includes setup/oracle work that shares that parent.
