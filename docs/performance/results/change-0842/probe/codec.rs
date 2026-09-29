//! Packet-only byte qualification for the root-lock compression probe.
//!
//! This module intentionally depends only on the public `soapberry-zip` API.
//! It writes one archive per codec route and a JSON manifest beside them.  The
//! caller can run the same command against a baseline and a candidate build;
//! route files can then be compared byte-for-byte without comparing bytes
//! across routes with different ZIP framing.

use std::fmt::Write as FmtWrite;
use std::fs;
use std::io::{self, Read, Write};
use std::path::Path;

use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, IndexedArchive, StreamingArchiveWriter};
use soapberry_zip::{CompressionMethod, PreservationPlan};

const SCHEMA: &str = "litchi.compression-codec-qualification-0842.v1";
const TOOL: &str = "compression-probe-0842-codecs";
const LARGE_BYTES: usize = 32 * 1024 + 257;
const RANDOM_SEED: u32 = 0x9e37_79b9;
const PATTERN: &[u8] = b"litchi-0842-codec-pattern-0123456789abcdef";

const ROUTES: &[Route] = &[
    Route::Borrowed,
    Route::Owned { chunk: 1 },
    Route::Owned { chunk: 7 },
    Route::Owned { chunk: 32 * 1024 },
    Route::Sized,
    Route::ReaderStream,
    Route::ReaderAlias,
    Route::Reusable,
    Route::PreservedCopy,
];

const MEMBER_NAMES: [&str; 4] = [
    "codec/empty.bin",
    "codec/tiny.bin",
    "codec/pattern.bin",
    "codec/random.bin",
];

#[derive(Clone, Copy, Debug)]
enum Route {
    Borrowed,
    Owned { chunk: usize },
    Sized,
    ReaderStream,
    ReaderAlias,
    Reusable,
    PreservedCopy,
}

impl Route {
    const fn name(self) -> &'static str {
        match self {
            Self::Borrowed => "borrowed",
            Self::Owned { chunk: 1 } => "owned-write-1",
            Self::Owned { chunk: 7 } => "owned-write-7",
            Self::Owned { chunk: 32_768 } => "owned-write-32768",
            Self::Owned { .. } => "owned-write-other",
            Self::Sized => "sized",
            Self::ReaderStream => "reader-stream",
            Self::ReaderAlias => "reader-alias",
            Self::Reusable => "reusable",
            Self::PreservedCopy => "preserved-copy",
        }
    }

    const fn api(self) -> &'static str {
        match self {
            Self::Borrowed => "StreamingArchiveWriter::write_deflated",
            Self::Owned { .. } => "StreamingArchiveWriter::start_entry + Write",
            Self::Sized => "StreamingArchiveWriter::write_deflated_sized",
            Self::ReaderStream => "StreamingArchiveWriter::write_deflated_stream",
            Self::ReaderAlias => "StreamingArchiveWriter::write_deflated_reader",
            Self::Reusable => "StreamingArchiveWriter::write_deflated repeated on one writer",
            Self::PreservedCopy => {
                "IndexedArchive::preservation_index + PreservationPlan::copy_all"
            },
        }
    }
}

#[derive(Debug)]
struct PayloadCase {
    name: &'static str,
    kind: &'static str,
    payload: Vec<u8>,
}

#[derive(Debug, Serialize)]
struct Identity {
    bytes: u64,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct Generator {
    version: &'static str,
    large_bytes: usize,
    pattern: &'static str,
    pattern_rule: &'static str,
    random_algorithm: &'static str,
    random_seed: u32,
    tiny_literal: &'static str,
}

#[derive(Debug, Serialize)]
struct PayloadSummary {
    name: &'static str,
    kind: &'static str,
    bytes: u64,
    payload_sha256: String,
}

#[derive(Debug, Serialize)]
struct MemberSummary {
    name: &'static str,
    bytes: u64,
    expected_payload_sha256: String,
    decoded_payload_sha256: String,
}

#[derive(Debug, Serialize)]
struct RouteSummary {
    route: &'static str,
    api: &'static str,
    archive: Identity,
    logical: Identity,
    expected_logical: Identity,
    members: Vec<MemberSummary>,
}

#[derive(Debug, Serialize)]
struct Summary {
    schema: &'static str,
    tool: &'static str,
    generator: Generator,
    payloads: Vec<PayloadSummary>,
    expected_logical: Identity,
    routes: Vec<RouteSummary>,
}

/// Generate every packet archive and an independently checkable JSON summary.
///
/// `dir` must already exist and must be empty of the output names.  Every ZIP
/// and the summary is opened with `create_new`, so a rerun cannot silently
/// overwrite evidence from a prior leg.
pub fn qualify(dir: &Path) -> Result<(), Box<dyn std::error::Error>> {
    if dir.as_os_str().is_empty() {
        return Err("codec qualification directory must not be empty".into());
    }
    if !dir.is_dir() {
        return Err(format!(
            "codec qualification directory is not a directory: {}",
            dir.display()
        )
        .into());
    }

    let cases = payload_cases();
    let expected_logical = logical_identity(&cases)?;
    let mut route_outputs = Vec::with_capacity(ROUTES.len());
    for route in ROUTES {
        let bytes = route_bytes(*route, &cases)?;
        let (logical, members) = inspect_archive(&bytes, &cases)?;
        if logical.sha256 != expected_logical.sha256 || logical.bytes != expected_logical.bytes {
            return Err(format!("{} route logical payload mismatch", route.name()).into());
        }
        write_new(&dir.join(format!("{}.zip", route.name())), &bytes)?;
        route_outputs.push(RouteSummary {
            route: route.name(),
            api: route.api(),
            archive: identity(&bytes),
            logical,
            expected_logical: Identity {
                bytes: expected_logical.bytes,
                sha256: expected_logical.sha256.clone(),
            },
            members,
        });
    }

    let payloads = cases
        .iter()
        .map(|case| {
            Ok(PayloadSummary {
                name: case.name,
                kind: case.kind,
                bytes: u64::try_from(case.payload.len())?,
                payload_sha256: digest(&case.payload),
            })
        })
        .collect::<Result<Vec<_>, Box<dyn std::error::Error>>>()?;
    let summary = Summary {
        schema: SCHEMA,
        tool: TOOL,
        generator: Generator {
            version: "0842-packet-v1",
            large_bytes: LARGE_BYTES,
            pattern: std::str::from_utf8(PATTERN)?,
            pattern_rule: "PATTERN[i modulo pattern.len()]",
            random_algorithm: "xorshift32; emit state shifted right 24 after each step",
            random_seed: RANDOM_SEED,
            tiny_literal: "soapberry-0842-tiny",
        },
        payloads,
        expected_logical,
        routes: route_outputs,
    };
    let summary_bytes = serde_json::to_vec_pretty(&summary)?;
    write_new(&dir.join("summary.json"), &summary_bytes)
}

fn payload_cases() -> Vec<PayloadCase> {
    let pattern = (0..LARGE_BYTES)
        .map(|index| PATTERN[index % PATTERN.len()])
        .collect();
    let mut state = RANDOM_SEED;
    let random = (0..LARGE_BYTES)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            (state >> 24) as u8
        })
        .collect();
    vec![
        PayloadCase {
            name: MEMBER_NAMES[0],
            kind: "empty",
            payload: Vec::new(),
        },
        PayloadCase {
            name: MEMBER_NAMES[1],
            kind: "tiny",
            payload: b"soapberry-0842-tiny".to_vec(),
        },
        PayloadCase {
            name: MEMBER_NAMES[2],
            kind: "pattern",
            payload: pattern,
        },
        PayloadCase {
            name: MEMBER_NAMES[3],
            kind: "random",
            payload: random,
        },
    ]
}

fn route_bytes(route: Route, cases: &[PayloadCase]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    match route {
        Route::Borrowed | Route::Reusable => write_borrowed(cases),
        Route::Owned { chunk } => write_owned(cases, chunk),
        Route::Sized => write_sized(cases),
        Route::ReaderStream => write_reader(cases, false),
        Route::ReaderAlias => write_reader(cases, true),
        Route::PreservedCopy => {
            let source = write_borrowed(cases)?;
            let source_len = u64::try_from(source.len())?;
            let indexed = IndexedArchive::from_reader(source, source_len)?;
            let mut scratch = vec![0_u8; 16 * 1024];
            let preservation = indexed.preservation_index(&mut scratch)?;
            let plan = PreservationPlan::copy_all(&preservation);
            Ok(preservation.write_to(&plan, Vec::new())?)
        },
    }
}

fn write_borrowed(cases: &[PayloadCase]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut writer = StreamingArchiveWriter::new();
    for case in cases {
        writer.write_deflated(case.name, &case.payload)?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn write_owned(cases: &[PayloadCase], chunk: usize) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    if chunk == 0 {
        return Err("owned entry chunk size must be nonzero".into());
    }
    let mut writer = StreamingArchiveWriter::new();
    for case in cases {
        let mut entry = writer.start_entry(case.name, CompressionMethod::Deflate)?;
        for bytes in case.payload.chunks(chunk) {
            entry.write_all(bytes)?;
        }
        writer = entry.finish().map_err(|failure| failure.into_error())?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn write_sized(cases: &[PayloadCase]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut writer = StreamingArchiveWriter::new();
    for case in cases {
        writer.write_deflated_sized(case.name, &case.payload)?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn write_reader(cases: &[PayloadCase], alias: bool) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut writer = StreamingArchiveWriter::new();
    for case in cases {
        let reader = PayloadReader::new(&case.payload, if alias { 1 } else { 7 });
        if alias {
            writer.write_deflated_reader(case.name, reader)?;
        } else {
            writer.write_deflated_stream(case.name, reader)?;
        }
    }
    Ok(writer.finish_to_bytes()?)
}

struct PayloadReader<'a> {
    data: &'a [u8],
    offset: usize,
    chunk: usize,
}

impl<'a> PayloadReader<'a> {
    fn new(data: &'a [u8], chunk: usize) -> Self {
        Self {
            data,
            offset: 0,
            chunk: chunk.max(1),
        }
    }
}

impl Read for PayloadReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.offset == self.data.len() || output.is_empty() {
            return Ok(0);
        }
        let count = (self.data.len() - self.offset)
            .min(output.len())
            .min(self.chunk);
        output[..count].copy_from_slice(&self.data[self.offset..self.offset + count]);
        self.offset += count;
        Ok(count)
    }
}

fn inspect_archive(
    bytes: &[u8],
    cases: &[PayloadCase],
) -> Result<(Identity, Vec<MemberSummary>), Box<dyn std::error::Error>> {
    let reader = ArchiveReader::new(bytes)?;
    if reader.len() != cases.len() {
        return Err(format!(
            "archive has {} members, expected {}",
            reader.len(),
            cases.len()
        )
        .into());
    }
    let names = reader.file_names().collect::<Vec<_>>();
    if names != cases.iter().map(|case| case.name).collect::<Vec<_>>() {
        return Err("archive member order or names differ from packet".into());
    }
    let mut digest = Sha256::new();
    let mut total = 0_u64;
    let mut members = Vec::with_capacity(cases.len());
    for case in cases {
        let payload = reader.read(case.name)?;
        if payload != case.payload {
            return Err(format!("decoded payload mismatch for {}", case.name).into());
        }
        let name_len = u64::try_from(case.name.len())?;
        let payload_len = u64::try_from(payload.len())?;
        digest.update(name_len.to_le_bytes());
        digest.update(case.name.as_bytes());
        digest.update(payload_len.to_le_bytes());
        digest.update(&payload);
        total = total
            .checked_add(payload_len)
            .ok_or("logical payload byte count overflow")?;
        let payload_sha256 = self::digest(&case.payload);
        members.push(MemberSummary {
            name: case.name,
            bytes: payload_len,
            expected_payload_sha256: payload_sha256.clone(),
            decoded_payload_sha256: payload_sha256,
        });
    }
    Ok((
        Identity {
            bytes: total,
            sha256: hex_digest(digest.finalize()),
        },
        members,
    ))
}

fn logical_identity(cases: &[PayloadCase]) -> Result<Identity, Box<dyn std::error::Error>> {
    let mut digest = Sha256::new();
    let mut total = 0_u64;
    for case in cases {
        let name_len = u64::try_from(case.name.len())?;
        let payload_len = u64::try_from(case.payload.len())?;
        digest.update(name_len.to_le_bytes());
        digest.update(case.name.as_bytes());
        digest.update(payload_len.to_le_bytes());
        digest.update(&case.payload);
        total = total
            .checked_add(payload_len)
            .ok_or("logical payload byte count overflow")?;
    }
    Ok(Identity {
        bytes: total,
        sha256: hex_digest(digest.finalize()),
    })
}

fn identity(bytes: &[u8]) -> Identity {
    Identity {
        bytes: bytes.len() as u64,
        sha256: digest(bytes),
    }
}

fn digest(bytes: &[u8]) -> String {
    hex_digest(Sha256::digest(bytes))
}

fn hex_digest(bytes: impl AsRef<[u8]>) -> String {
    let mut output = String::with_capacity(bytes.as_ref().len() * 2);
    for byte in bytes.as_ref() {
        write!(&mut output, "{byte:02x}").expect("writing to String cannot fail");
    }
    output
}

fn write_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)?;
    file.write_all(bytes)?;
    file.flush()?;
    Ok(())
}
