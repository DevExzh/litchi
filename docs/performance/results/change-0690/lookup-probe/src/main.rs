//! Small, deterministic CFB directory-lookup probe for change 0690.
//!
//! Fixture construction and CFB parsing happen before the timed region.  The
//! timed operation is only repeated `OleFile::stream_len` calls, with the
//! result consumed by a checksum so the lookup cannot be removed.  The probe
//! intentionally uses the public writer and reader APIs; it does not reach
//! into the directory index or duplicate its comparison implementation.
//!
//! Examples:
//!
//! ```text
//! cfb-lookup-probe-0690 --case ascii-mixed-257 --repetitions 1000000
//! cfb-lookup-probe-0690 --case all --format json --repetitions 100000
//! cfb-lookup-probe-0690 --list
//! ```

use litchi_cfb::writer::OleWriter;
use litchi_cfb::{OleError, OleFile};
use sha2::{Digest, Sha256};
use std::env;
use std::hint::black_box;
use std::io::{self, Cursor, Write};
use std::time::Instant;

const DEFAULT_REPETITIONS: u64 = 100_000;
const DEFAULT_WARMUPS: u64 = 1_000;
const MAX_REPETITIONS: u64 = 100_000_000;
const MAX_WARMUPS: u64 = 10_000_000;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OutputFormat {
    Tsv,
    Json,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum QueryKind {
    Exact,
    MixedAscii,
    UnicodeEquivalent,
    UnicodeQuery,
    UnicodeFallback,
    Missing,
    Invalid,
    Nested,
}

impl QueryKind {
    const fn label(self) -> &'static str {
        match self {
            Self::Exact => "exact",
            Self::MixedAscii => "mixed_ascii",
            Self::UnicodeEquivalent => "unicode_equivalent",
            Self::UnicodeQuery => "unicode_query",
            Self::UnicodeFallback => "unicode_fallback",
            Self::Missing => "missing",
            Self::Invalid => "invalid",
            Self::Nested => "nested",
        }
    }

    const fn expects_success(self) -> bool {
        matches!(
            self,
            Self::Exact
                | Self::MixedAscii
                | Self::UnicodeEquivalent
                | Self::UnicodeQuery
                | Self::UnicodeFallback
                | Self::Nested
        )
    }
}

#[derive(Clone, Copy, Debug)]
struct CaseSpec {
    name: &'static str,
    width: usize,
    kind: QueryKind,
}

const CASES: &[CaseSpec] = &[
    CaseSpec {
        name: "ascii-exact-1",
        width: 1,
        kind: QueryKind::Exact,
    },
    CaseSpec {
        name: "ascii-mixed-1",
        width: 1,
        kind: QueryKind::MixedAscii,
    },
    CaseSpec {
        name: "ascii-exact-31",
        width: 31,
        kind: QueryKind::Exact,
    },
    CaseSpec {
        name: "ascii-mixed-31",
        width: 31,
        kind: QueryKind::MixedAscii,
    },
    CaseSpec {
        name: "ascii-exact-257",
        width: 257,
        kind: QueryKind::Exact,
    },
    CaseSpec {
        name: "ascii-mixed-257",
        width: 257,
        kind: QueryKind::MixedAscii,
    },
    CaseSpec {
        name: "unicode-equivalent-31",
        width: 31,
        kind: QueryKind::UnicodeEquivalent,
    },
    CaseSpec {
        name: "unicode-query-31",
        width: 31,
        kind: QueryKind::UnicodeQuery,
    },
    CaseSpec {
        name: "unicode-fallback-257",
        width: 257,
        kind: QueryKind::UnicodeFallback,
    },
    CaseSpec {
        name: "missing-257",
        width: 257,
        kind: QueryKind::Missing,
    },
    CaseSpec {
        name: "invalid-31",
        width: 31,
        kind: QueryKind::Invalid,
    },
    CaseSpec {
        name: "nested-257",
        width: 257,
        kind: QueryKind::Nested,
    },
];

struct Fixture {
    file: OleFile<Cursor<Vec<u8>>>,
    path: Vec<String>,
    query: Vec<String>,
    expected_length: Option<u64>,
    bytes: usize,
    sha256: String,
}

#[derive(Debug)]
struct Row {
    case: &'static str,
    width: usize,
    query_kind: QueryKind,
    repetitions: u64,
    elapsed_ns: u128,
    checksum: u64,
    successes: u64,
    errors: u64,
    expected: bool,
    fixture_bytes: usize,
    fixture_sha256: String,
}

fn usage() -> &'static str {
    "usage: cfb-lookup-probe-0690 [--case NAME|all] [--repetitions N] [--warmups N] [--format tsv|json] [--list]\n\nThe default case is ascii-mixed-257. Repetitions are timed lookup calls; fixture construction, serialization, and parsing are outside the timer."
}

fn parse_u64(value: &str, option: &str, maximum: u64) -> Result<u64, String> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| format!("{option} requires a non-negative integer"))?;
    if parsed == 0 || parsed > maximum {
        return Err(format!(
            "{option} must be between 1 and {maximum}, got {parsed}"
        ));
    }
    Ok(parsed)
}

fn case_by_name(name: &str) -> Option<CaseSpec> {
    CASES.iter().copied().find(|case| case.name == name)
}

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let mut case_name = String::from("ascii-mixed-257");
    let mut repetitions = DEFAULT_REPETITIONS;
    let mut warmups = DEFAULT_WARMUPS;
    let mut format = OutputFormat::Tsv;
    let mut list = false;

    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--case" => {
                case_name = args
                    .next()
                    .ok_or_else(|| "--case requires a value".to_string())?;
            },
            "--repetitions" => {
                repetitions = parse_u64(
                    &args
                        .next()
                        .ok_or_else(|| "--repetitions requires a value".to_string())?,
                    "--repetitions",
                    MAX_REPETITIONS,
                )?;
            },
            "--warmups" => {
                warmups = parse_u64(
                    &args
                        .next()
                        .ok_or_else(|| "--warmups requires a value".to_string())?,
                    "--warmups",
                    MAX_WARMUPS,
                )?;
            },
            "--format" => {
                format = match args
                    .next()
                    .ok_or_else(|| "--format requires tsv or json".to_string())?
                    .as_str()
                {
                    "tsv" => OutputFormat::Tsv,
                    "json" => OutputFormat::Json,
                    other => return Err(format!("unknown output format {other:?}").into()),
                };
            },
            "--list" => list = true,
            "--help" | "-h" => {
                println!("{}", usage());
                return Ok(());
            },
            other => return Err(format!("unknown option {other:?}\n{}", usage()).into()),
        }
    }

    if list {
        for case in CASES {
            println!(
                "{}\twidth={}\tquery={}",
                case.name,
                case.width,
                case.kind.label()
            );
        }
        return Ok(());
    }

    let specs: Vec<CaseSpec> = if case_name == "all" {
        CASES.to_vec()
    } else {
        vec![
            case_by_name(&case_name)
                .ok_or_else(|| format!("unknown case {case_name:?}; use --list"))?,
        ]
    };

    let mut rows = Vec::with_capacity(specs.len());
    for spec in specs {
        rows.push(run_case(spec, warmups, repetitions)?);
    }
    match format {
        OutputFormat::Tsv => write_tsv(&rows)?,
        OutputFormat::Json => write_json(&rows)?,
    }
    Ok(())
}

fn run_case(
    spec: CaseSpec,
    warmups: u64,
    repetitions: u64,
) -> Result<Row, Box<dyn std::error::Error + Send + Sync>> {
    // This includes all writer work, serialization, and reader validation, so
    // none of those costs can enter the lookup-only timer below.
    let fixture = build_fixture(spec)?;
    let path: Vec<&str> = fixture.path.iter().map(String::as_str).collect();
    let query: Vec<&str> = fixture.query.iter().map(String::as_str).collect();

    for _ in 0..warmups {
        let _ = black_box(lookup_once(&fixture.file, &query));
    }

    let mut checksum = 0_u64;
    let mut successes = 0_u64;
    let mut errors = 0_u64;
    let start = Instant::now();
    for _ in 0..repetitions {
        match fixture.file.stream_len(black_box(&query)) {
            Ok(length) => {
                successes += 1;
                checksum = checksum
                    .wrapping_mul(0x9e37_79b9_7f4a_7c15)
                    .wrapping_add(length ^ 0x51ed_270b);
            },
            Err(error) => {
                errors += 1;
                checksum = checksum
                    .wrapping_mul(0x9e37_79b9_7f4a_7c15)
                    .wrapping_add(error_code(&error));
            },
        }
    }
    let elapsed_ns = start.elapsed().as_nanos();
    checksum = black_box(checksum);

    // Keep the path alive through the timed call. This also makes it obvious
    // that nested and root queries use exactly the same lookup API.
    black_box(&path);
    if spec.kind.expects_success() {
        if errors != 0 || successes != repetitions {
            return Err(format!(
                "case {} expected all lookups to succeed: successes={successes} errors={errors}",
                spec.name
            )
            .into());
        }
    } else if successes != 0 || errors != repetitions {
        return Err(format!(
            "case {} expected all lookups to refuse: successes={successes} errors={errors}",
            spec.name
        )
        .into());
    }
    if let Some(expected_length) = fixture.expected_length {
        let observed = fixture.file.stream_len(&query)?;
        if observed != expected_length {
            return Err(format!(
                "case {} returned {observed} instead of {expected_length}",
                spec.name
            )
            .into());
        }
    }

    Ok(Row {
        case: spec.name,
        width: spec.width,
        query_kind: spec.kind,
        repetitions,
        elapsed_ns,
        checksum,
        successes,
        errors,
        expected: spec.kind.expects_success(),
        fixture_bytes: fixture.bytes,
        fixture_sha256: fixture.sha256,
    })
}

fn lookup_once(file: &OleFile<Cursor<Vec<u8>>>, query: &[&str]) -> Result<u64, OleError> {
    file.stream_len(query)
}

fn build_fixture(spec: CaseSpec) -> Result<Fixture, Box<dyn std::error::Error + Send + Sync>> {
    let mut writer = OleWriter::new();
    let target = match spec.kind {
        QueryKind::UnicodeEquivalent | QueryKind::UnicodeFallback => "ſtream",
        QueryKind::UnicodeQuery => "élan",
        QueryKind::Nested => "Payload",
        _ => "Stream",
    };

    let nested = spec.kind == QueryKind::Nested;
    if nested {
        writer.create_storage(&["Nested"])?;
        writer.create_storage(&["Nested", "Deep"])?;
    }
    for index in 0..spec.width {
        // Six UTF-16-unit names keep the generated siblings in one length
        // bucket with "Stream"/"ſtream", making width the meaningful axis.
        let name = if index == spec.width / 2 {
            target.to_owned()
        } else {
            format!("E{index:05}")
        };
        let payload: &[u8] = if name == target { &[0x37; 37] } else { &[] };
        if nested {
            writer.create_stream(&["Nested", "Deep", &name], payload)?;
        } else {
            writer.create_stream(&[&name], payload)?;
        }
    }

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes)?;
    let bytes = bytes.into_inner();
    let digest = Sha256::digest(&bytes);
    let sha256 = digest.iter().map(|byte| format!("{byte:02x}")).collect();
    let file = OleFile::open(Cursor::new(bytes.clone()))?;

    let (path, query, expected_length) = match spec.kind {
        QueryKind::Exact => (vec![target.to_owned()], vec![target.to_owned()], Some(37)),
        QueryKind::MixedAscii => (vec![target.to_owned()], vec!["sTrEaM".to_owned()], Some(37)),
        QueryKind::UnicodeEquivalent => {
            (vec![target.to_owned()], vec!["Stream".to_owned()], Some(37))
        },
        QueryKind::UnicodeQuery => (vec![target.to_owned()], vec!["ÉLAN".to_owned()], Some(37)),
        QueryKind::UnicodeFallback => {
            (vec![target.to_owned()], vec!["ſTrEaM".to_owned()], Some(37))
        },
        QueryKind::Missing => (vec!["Stream".to_owned()], vec!["Missing".to_owned()], None),
        QueryKind::Invalid => (vec!["Stream".to_owned()], vec!["bad/name".to_owned()], None),
        QueryKind::Nested => (
            vec!["Nested".to_owned(), "Deep".to_owned(), target.to_owned()],
            vec!["nested".to_owned(), "DEEP".to_owned(), "pAYLOAD".to_owned()],
            Some(37),
        ),
    };

    // The ordinary stream path is retained as a debug oracle while the timed
    // query is allowed to use a mixed-case spelling of the same path.
    if spec.kind == QueryKind::Nested {
        let path_refs: Vec<&str> = path.iter().map(String::as_str).collect();
        if file.stream_len(&path_refs)? != 37 {
            return Err("nested fixture oracle failed".into());
        }
    }
    Ok(Fixture {
        file,
        path,
        query,
        expected_length,
        bytes: bytes.len(),
        sha256,
    })
}

fn error_code(error: &OleError) -> u64 {
    match error {
        OleError::StreamNotFound => 0x4e4f_5446,
        OleError::InvalidData(_) => 0x494e_5644,
        OleError::InvalidFormat(_) => 0x494e_5646,
        OleError::Io(_) => 0x494f_0000,
        OleError::Allocation { .. } => 0x414c_4c4f,
        OleError::Committed { .. } => 0x434f_4d4d,
        OleError::LimitExceeded { .. } => 0x4c49_4d54,
        OleError::InvalidLimit { .. } => 0x4c49_4d49,
        OleError::NotOleFile => 0x4e4f_4c45,
        OleError::CorruptedFile(_) => 0x434f_5252,
        OleError::SourceChanged { .. } => 0x534f_5552,
    }
}

fn write_tsv(rows: &[Row]) -> io::Result<()> {
    let stdout = io::stdout();
    let mut out = stdout.lock();
    writeln!(
        out,
        "case\twidth\tquery_kind\trepetitions\telapsed_ns\tchecksum\tsuccesses\terrors\texpected\tfixture_bytes\tfixture_sha256"
    )?;
    for row in rows {
        writeln!(
            out,
            "{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
            row.case,
            row.width,
            row.query_kind.label(),
            row.repetitions,
            row.elapsed_ns,
            row.checksum,
            row.successes,
            row.errors,
            row.expected,
            row.fixture_bytes,
            row.fixture_sha256,
        )?;
    }
    Ok(())
}

fn write_json(rows: &[Row]) -> io::Result<()> {
    let stdout = io::stdout();
    let mut out = stdout.lock();
    writeln!(out, "[")?;
    for (index, row) in rows.iter().enumerate() {
        let comma = if index + 1 == rows.len() { "" } else { "," };
        writeln!(
            out,
            "  {{\"case\":\"{}\",\"width\":{},\"query_kind\":\"{}\",\"repetitions\":{},\"elapsed_ns\":{},\"checksum\":{},\"successes\":{},\"errors\":{},\"expected\":{},\"fixture_bytes\":{},\"fixture_sha256\":\"{}\"}}{comma}",
            json_escape(row.case),
            row.width,
            row.query_kind.label(),
            row.repetitions,
            row.elapsed_ns,
            row.checksum,
            row.successes,
            row.errors,
            row.expected,
            row.fixture_bytes,
            json_escape(&row.fixture_sha256),
        )?;
    }
    writeln!(out, "]")
}

fn json_escape(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '"' => escaped.push_str("\\\""),
            '\\' => escaped.push_str("\\\\"),
            '\n' => escaped.push_str("\\n"),
            '\r' => escaped.push_str("\\r"),
            '\t' => escaped.push_str("\\t"),
            character if character.is_control() => {
                use std::fmt::Write as _;
                write!(&mut escaped, "\\u{:04x}", character as u32).expect("String write");
            },
            character => escaped.push(character),
        }
    }
    escaped
}
