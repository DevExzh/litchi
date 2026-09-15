import sys, pathlib
src = pathlib.Path(sys.argv[1]).read_text(); mode = sys.argv[3]
if mode == "prefix":          # change 0580: central only, both branches
    src = src.replace(
        "    let payload_end = variable_end\n        .checked_add(neighbour_payload_length(&file_header, entry))",
        "    let payload_end = variable_end\n        .checked_add(entry.compressed_size)")
elif mode == "max-everywhere":  # the intermediate that broke monotonicity
    src = src.replace(
        "        let payload_end = variable_end\n            .checked_add(entry.compressed_size)\n            .ok_or_else(|| Error::from(ErrorKind::Eof))?;",
        "        let payload_end = variable_end\n            .checked_add(neighbour_payload_length(&file_header, entry))\n            .ok_or_else(|| Error::from(ErrorKind::Eof))?;")
elif mode == "local-only":    # take the local value instead of the maximum
    src = src.replace(
        "    entry\n        .compressed_size\n        .max(u64::from(file_header.compressed_size))\n}",
        "    u64::from(file_header.compressed_size)\n}")
elif mode == "sentinel-literal":  # treat 0xFFFFFFFF as a length
    src = src.replace(
        "    if file_header.compressed_size == u32::MAX {\n        return entry.compressed_size;\n    }\n", "")
elif mode == "usize-max":     # extend the maximum to the uncompressed size
    src = src.replace(
        "    entry\n        .compressed_size\n        .max(u64::from(file_header.compressed_size))\n}",
        "    entry\n        .compressed_size\n        .max(u64::from(file_header.compressed_size))\n        .max(u64::from(file_header.uncompressed_size))\n}")
else:
    raise SystemExit("unknown mode")
pathlib.Path(sys.argv[2]).write_text(src)
