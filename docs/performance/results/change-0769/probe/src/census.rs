//! The reader-agreement lanes of change 0769.
//!
//! `census --input LIST`: every public way of reading every stream of every
//! listed compound file, and whether the answers agree. For each file (one
//! path per line of `LIST`) it records:
//!
//! - `OleFile::open` and `SharedOleFile::open_owned` verdicts;
//! - for every stream `OleFile::list_streams` names, the result of
//!   - `OleFile::open_stream` on one reader that reads every stream in order,
//!   - `OleFile::read_stream_range` of the whole stream on that reader,
//!   - `SharedOleFile::open_stream` on a fresh shared reader per stream (the
//!     bounded direct MiniFAT read when the stream is eligible),
//!   - `SharedOleFile::open_stream` on one shared reader reading every stream
//!     in order (the root mini-stream cache after the first target),
//!   - `SharedOleFile::read_stream_range` of the whole stream,
//!   - a `SharedOleStreamCursor` reading the whole stream;
//! - `agree`: every reader returned the same bytes, or every reader refused;
//! - `open_ok_read_err`: streams the open admitted that some reader refused.
//!
//! `root-sweep --input LIST`: for every listed file, the root entry's
//! stream size (the mini stream's size) is rewritten to every value within
//! 128 bytes of its own, and each rewritten copy gets the census above,
//! without the fresh shared reader per stream (the bounded direct read is
//! the range reader's code, which the sweep runs). One line per copy.
//!
//! Errors are kept as display strings; data as length and SHA-256. The output
//! depends only on the inputs, so two builds' outputs compare line for line.

use std::io::Cursor;
use std::path::Path;
use std::sync::Arc;

use litchi_cfb::{OleFile, SharedOleFile};
use litchi_core::SourceVersion;

use crate::{BoxError, json_string, sha};

fn verdict(result: Result<Vec<u8>, litchi_cfb::OleError>) -> String {
    match result {
        Ok(data) => format!("ok:{}:{}", data.len(), sha(&data)),
        Err(error) => format!("err:{error}"),
    }
}

fn shared_of(bytes: &[u8]) -> Result<SharedOleFile, litchi_cfb::OleError> {
    let source: Arc<[u8]> = Arc::from(bytes.to_vec().into_boxed_slice());
    SharedOleFile::open_owned(source, SourceVersion::new(769, 0))
}

/// One file's agreement: the two open verdicts, one JSON object per stream,
/// whether everything agreed, and how many admitted streams a reader refused.
struct Agreement {
    open: String,
    shared_open: String,
    streams: Vec<String>,
    agree: bool,
    open_ok_read_err: usize,
    reads_ok: usize,
    reads_err: usize,
}

fn agreement(bytes: &[u8], fresh_shared: bool) -> Agreement {
    let open = OleFile::open(Cursor::new(bytes));
    let shared_open = shared_of(bytes);
    let open_text = match &open {
        Ok(_) => "ok".to_string(),
        Err(error) => format!("err:{error}"),
    };
    let shared_text = match &shared_open {
        Ok(_) => "ok".to_string(),
        Err(error) => format!("err:{error}"),
    };
    let mut result = Agreement {
        agree: open_text == shared_text,
        open: open_text,
        shared_open: shared_text,
        streams: Vec::new(),
        open_ok_read_err: 0,
        reads_ok: 0,
        reads_err: 0,
    };
    if let (Ok(mut ole), Ok(shared_all)) = (open, shared_open) {
        for stream in ole.list_streams() {
            let refs: Vec<&str> = stream.iter().map(String::as_str).collect();
            let len = ole.stream_len(&refs).unwrap_or(0);
            let whole = usize::try_from(len).unwrap_or(0);
            let mut reads = vec![
                verdict(ole.open_stream(&refs)),
                verdict({
                    let mut out = vec![0u8; whole];
                    ole.read_stream_range(&refs, 0, &mut out).map(|()| out)
                }),
            ];
            if fresh_shared {
                reads.push(verdict(
                    shared_of(bytes).and_then(|fresh| fresh.open_stream(&refs)),
                ));
            }
            reads.push(verdict(shared_all.open_stream(&refs)));
            reads.push(verdict({
                let mut out = vec![0u8; whole];
                shared_all
                    .read_stream_range(&refs, 0, &mut out)
                    .map(|()| out)
            }));
            reads.push(verdict(shared_all.stream_cursor_at(&refs, 0).and_then(
                |mut cursor| {
                    let mut out = vec![0u8; whole];
                    cursor.read_exact(&mut out).map(|()| out)
                },
            )));
            let agree = reads.iter().all(|read| read == &reads[0])
                || reads.iter().all(|read| read.starts_with("err:"));
            let refused = reads.iter().any(|read| read.starts_with("err:"));
            result.agree &= agree;
            result.open_ok_read_err += usize::from(refused);
            result.reads_ok += reads.iter().filter(|read| read.starts_with("ok:")).count();
            result.reads_err += reads.iter().filter(|read| read.starts_with("err:")).count();
            result.streams.push(format!(
                "{{\"stream\":{},\"len\":{len},\"agree\":{agree},\"reads\":[{}]}}",
                json_string(&stream.join("/")),
                reads
                    .iter()
                    .map(|read| json_string(read))
                    .collect::<Vec<_>>()
                    .join(","),
            ));
        }
    }
    result
}

/// One file's census line.
fn census_line(path: &str) -> Result<String, BoxError> {
    let bytes = std::fs::read(path)?;
    let found = agreement(&bytes, true);
    Ok(format!(
        "{{\"path\":{},\"sha256\":{},\"open\":{},\"shared_open\":{},\"agree\":{},\"open_ok_read_err\":{},\"streams\":[{}]}}",
        json_string(path),
        json_string(&sha(&bytes)),
        json_string(&found.open),
        json_string(&found.shared_open),
        found.agree,
        found.open_ok_read_err,
        found.streams.join(","),
    ))
}

/// Prints one JSON line per listed file.
pub(crate) fn run(list: &Path) -> Result<(), BoxError> {
    let paths = std::fs::read_to_string(list)?;
    for path in paths.lines().filter(|line| !line.is_empty()) {
        println!("{}", census_line(path)?);
    }
    Ok(())
}

/// Where the root entry's stream size field lies, and its value (the low
/// 32 bits only for 512-byte sectors, as the reader masks them), when the
/// header and first directory sector are within the file.
fn root_size_field(bytes: &[u8]) -> Option<(usize, u64)> {
    if bytes.len() < 512 || bytes[..8] != [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        return None;
    }
    let shift = u16::from_le_bytes([bytes[0x1E], bytes[0x1F]]);
    if !matches!(shift, 9 | 12) {
        return None;
    }
    let sector_size = 1usize << shift;
    let first_dir = usize::try_from(u32::from_le_bytes(bytes[0x30..0x34].try_into().ok()?)).ok()?;
    let field = first_dir.checked_add(1)?.checked_mul(sector_size)?.checked_add(0x78)?;
    let raw = u64::from_le_bytes(bytes.get(field..field + 8)?.try_into().ok()?);
    let size = if sector_size == 512 { raw & 0xFFFF_FFFF } else { raw };
    Some((field, size))
}

/// Prints one JSON line per listed file and rewritten root size.
pub(crate) fn run_root_sweep(list: &Path) -> Result<(), BoxError> {
    let paths = std::fs::read_to_string(list)?;
    for (file, path) in paths.lines().filter(|line| !line.is_empty()).enumerate() {
        let clean = std::fs::read(path)?;
        let Some((field, original)) = root_size_field(&clean) else {
            println!(
                "{{\"file\":{file},\"path\":{},\"skipped\":\"no root size field\"}}",
                json_string(path)
            );
            continue;
        };
        let low = original.saturating_sub(128).max(1);
        for root_size in low..=original + 128 {
            let mut bytes = clean.clone();
            bytes[field..field + 8].copy_from_slice(&root_size.to_le_bytes());
            let found = agreement(&bytes, false);
            println!(
                "{{\"file\":{file},\"root\":{root_size},\"delta\":{},\"open\":{},\"shared_open\":{},\"streams\":{},\"reads_ok\":{},\"reads_err\":{},\"agree\":{},\"open_ok_read_err\":{},\"digest\":{}}}",
                i128::from(root_size) - i128::from(original),
                json_string(&found.open),
                json_string(&found.shared_open),
                found.streams.len(),
                found.reads_ok,
                found.reads_err,
                found.agree,
                found.open_ok_read_err,
                json_string(&sha(found.streams.join("\n").as_bytes())),
            );
        }
        println!(
            "{{\"file\":{file},\"path\":{},\"sha256\":{},\"original_root\":{original},\"summary\":true}}",
            json_string(path),
            json_string(&sha(&clean)),
        );
    }
    Ok(())
}
