#![no_main]

use libfuzzer_sys::fuzz_target;

// Drives raw bytes through litchi-cfb's CFB/OLE2 parser.
// Errors are expected on malformed input; we want to ensure
// the parser does not panic, OOM, or hit UB on arbitrary bytes.
fuzz_target!(|data: &[u8]| {
    // Cheap sniff helper; exercises the public is_ole_file path.
    let _ = litchi_cfb::is_ole_file(data);

    if let Ok(mut ole) = litchi_cfb::OleFile::open(std::io::Cursor::new(data)) {
        // Walk the directory tree to give the fuzzer extra reach.
        let streams = ole.list_streams();
        let _ = ole.list_directory_entries(&[]);
        let _ = ole.get_root_name();

        // Attempt to read each enumerated stream. Cap iterations so a
        // pathological directory tree can't dominate one fuzz iteration.
        for path in streams.into_iter().take(64) {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let _ = ole.exists(&refs);
            let _ = ole.open_stream(&refs);
        }
    }
});
