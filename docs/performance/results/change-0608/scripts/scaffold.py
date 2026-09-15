import sys, pathlib
mode = sys.argv[1]
p = pathlib.Path("/home/zhuhe/code/litchi-worktrees/0608/crates/litchi-xls/src/records.rs")
s = p.read_text()
LOOP = """    let mut entries = Vec::new();
    entries
        .try_reserve_exact(unique_count)
        .map_err(|_| SharedStringScanError::Allocation {
            resource: "SST entry locator",
            requested: unique_count,
        })?;
    for string_index in 0..unique_count {
        let start = cursor.logical_position();
        // The scan retains offsets, never text, so production walks the string
        // rather than decoding it. The walk refuses the same inputs at the same
        // point in the sequence, with the same typed errors.
        walk_one_shared_string::<T>(&mut cursor, string_index)?;
        let end = cursor.logical_position();
        entries.push(SharedStringEntryLocation { start, end });
    }
"""
assert LOOP in s, "loop text drifted"
if mode == "nostore":
    NEW = """    // CHANGE 0608 MEASUREMENT SCAFFOLD -- NOT A CANDIDATE.
    // Walk every shared string exactly as production does, but reserve nothing
    // and record nothing, so that `open(base) - open(nostore)` prices what the
    // "validate eagerly, store lazily" form of item XLS-2 could save.
    let entries = Vec::new();
    for string_index in 0..unique_count {
        walk_one_shared_string::<T>(&mut cursor, string_index)?;
    }
"""
    # the locator type is then never constructed inside the crate
    s = s.replace(
        "#[derive(Debug, Clone, Copy)]\npub(crate) struct SharedStringEntryLocation {",
        "#[derive(Debug, Clone, Copy)]\n#[allow(dead_code)]\npub(crate) struct SharedStringEntryLocation {",
    )
    s = s.replace("    fn logical_position(&self) -> usize {",
                  "    #[allow(dead_code)]\n    fn logical_position(&self) -> usize {")
elif mode == "nowalk":
    NEW = """    // CHANGE 0608 MEASUREMENT SCAFFOLD -- NOT A CANDIDATE.
    // Build the segments and run the SST header checks, then skip the
    // per-string walk entirely, so that `open(base) - open(nowalk)` prices the
    // ceiling of fully deferred indexing at open. `black_box` keeps the loop
    // and its instantiation in the binary.
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(unique_count)
        .map_err(|_| SharedStringScanError::Allocation {
            resource: "SST entry locator",
            requested: unique_count,
        })?;
    let scaffold_count = std::hint::black_box(0usize);
    for string_index in 0..scaffold_count {
        let start = cursor.logical_position();
        walk_one_shared_string::<T>(&mut cursor, string_index)?;
        let end = cursor.logical_position();
        entries.push(SharedStringEntryLocation { start, end });
    }
"""
else:
    raise SystemExit("mode?")
p.write_text(s.replace(LOOP, NEW))
print("scaffold", mode, "applied")
