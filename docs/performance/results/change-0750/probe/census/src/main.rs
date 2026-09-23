//! Change 0750 census probe: audits every XML member it is given with the
//! `xml-minifier` it was built against and prints one verdict line per member.
//!
//! Input on stdin: records of a little-endian `u32` length and that many bytes.
//! Output: `source<TAB>authored<TAB>compact<TAB>reader` per record, where each
//! field is `OK` or the error's `Display` text with tabs and newlines replaced.

use std::io::{self, BufWriter, Cursor, Read, Write};

use xml_minifier::audit::{self, Limits};

fn verdict<T, E: std::fmt::Display>(result: Result<T, E>) -> String {
    match result {
        Ok(_) => "OK".to_owned(),
        Err(error) => error.to_string().replace(['\t', '\n', '\r'], " "),
    }
}

fn main() -> io::Result<()> {
    let mut input = Vec::new();
    io::stdin().read_to_end(&mut input)?;
    let stdout = io::stdout();
    let mut out = BufWriter::new(stdout.lock());
    let limits = Limits::default();
    let mut cursor = 0;
    while cursor + 4 <= input.len() {
        let length = u32::from_le_bytes(input[cursor..cursor + 4].try_into().unwrap()) as usize;
        cursor += 4;
        let member = &input[cursor..cursor + length];
        cursor += length;
        let source = verdict(audit::verify_source(member, limits));
        let authored = verdict(audit::verify_authored(member, limits));
        let compact = verdict(audit::verify(member, limits));
        let reader = verdict(audit::verify_reader(Cursor::new(member), limits));
        writeln!(out, "{source}\t{authored}\t{compact}\t{reader}")?;
    }
    out.flush()
}
