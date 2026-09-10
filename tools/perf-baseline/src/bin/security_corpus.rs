//! Opt-in executable rows for the checked-in security corpus.
//!
//! This binary is deliberately separate from the ordinary baseline selector.
//! It keeps the 201-row default matrix unchanged while providing three
//! runnable operations for every security fixture: bounded read, semantic
//! validation, and exact no-op preservation.  The normal invocation executes
//! each selected row once as a correctness smoke.  A timer is available only
//! behind an explicit environment gate and `--measure`, so an audit cannot
//! accidentally turn a correctness run into an unreviewed measurement.

#![forbid(unsafe_code)]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the opt-in fixture runner reports a failing oracle immediately"
)]

use std::{
    collections::BTreeMap,
    error::Error,
    fs::{self, File},
    io::{self, Cursor, Read, Write},
    path::{Path, PathBuf},
    time::Instant,
};

use litchi_cfb::{OleFile, OleFileLimits};
use litchi_crypto::ooxml::{self, IntegrityStatus, Kind, Limits as CryptoLimits, Mode};
use litchi_crypto::spaces;
use litchi_doc::{Error as DocError, Limits as DocLimits, OpenOptions, Package as DocPackage};
use litchi_docx::source_backed::Package as DocxSourcePackage;
use litchi_opc::{OpcError, OpcPackage, ReadLimits, SourceBackedPackage};
use litchi_sign::{Policy, Status};
use litchi_xls::Workbook as XlsWorkbook;
use litchi_xlsx::Workbook as XlsxWorkbook;
use serde::Serialize;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveLimits, ArchiveReader};
use soapberry_zip::{
    PreservationIndex, PreservationPlan, RECOMMENDED_BUFFER_SIZE, ReaderAt, ZipArchive,
};

#[cfg(test)]
use soapberry_zip::office::StreamingArchiveWriter;

const MAX_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const MAX_MEMBERS: usize = 4_096;
const MAX_RELATIONSHIPS: usize = 8_192;
const MAX_MEMBER_NAME_BYTES: u64 = 4 * 1024;
const CFB_DIRECTORY_ENTRY_BYTES: u64 = 128;
const CFB_DIRECTORY_BYTES: usize = (MAX_MEMBERS as u64 * CFB_DIRECTORY_ENTRY_BYTES) as usize;
const MEASURE_ENV: &str = "LITCHI_SECURITY_CORPUS_CAPTURE";
const MEASURE_TOKEN: &str = "approved";
const CATALOG_ID: &str = "litchi-security-corpus-v3";
const PROFILE_ID: &str = "security-corpus-bounded-read-v3";
const CLEAR_CRYPTO_SHA256: &str =
    "6e71a4d554ea90df1ddcc9b6fbbcc677a94c6617cb35f9e5b7d4ccab46f1a79e";
const SOURCE_PATHS: &[&str] = &[
    "tools/perf-baseline/src/bin/security_corpus.rs",
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/Cargo.lock",
    "tools/perf-baseline/src/security_corpus.rs",
    "crates/litchi-crypto/tests/ooxml_interoperability.rs",
    "crates/litchi-crypto/tests/data/ooxml/standard-manifest.json",
    "crates/litchi-crypto/tests/data/ooxml/agile-mixed-manifest.json",
    "crates/litchi-crypto/tests/data/ooxml/native-agile-manifest.json",
    "crates/litchi-crypto/tests/data/ooxml/rust-output-manifest.json",
    "crates/litchi-crypto/src/ooxml/mod.rs",
    "crates/litchi-crypto/src/ooxml/container.rs",
    "crates/litchi-crypto/src/ooxml/standard.rs",
    "crates/litchi-crypto/src/ooxml/agile.rs",
    "crates/litchi-crypto/src/spaces.rs",
    "crates/litchi-cfb/src/file.rs",
    "crates/litchi-cfb/src/allocation_validation_tests.rs",
    "crates/litchi-doc/src/package/codec.rs",
    "crates/litchi-doc/src/package/mod.rs",
    "crates/litchi-doc/src/package/model.rs",
    "crates/litchi-doc/src/document/mod.rs",
    "crates/litchi-doc/src/document/package.rs",
    "crates/litchi-docx/src/source_backed.rs",
    "crates/litchi-opc/src/pkgreader.rs",
    "crates/litchi-opc/src/package.rs",
    "crates/litchi-opc/src/limits.rs",
    "crates/litchi-opc/src/phys_pkg.rs",
    "crates/litchi-opc/src/source_backed.rs",
    "crates/litchi-ooxml-common/src/properties/read.rs",
    "crates/litchi-sign/src/xml.rs",
    "crates/litchi-xls/src/workbook/package.rs",
    "crates/litchi-xls/src/vba.rs",
    "crates/litchi-xls/src/cell_values/mod.rs",
    "crates/litchi-xls/src/cell_values/structural.rs",
    "crates/litchi-xls/src/workbook/source.rs",
    "crates/litchi-ole-common/src/object/mod.rs",
    "crates/litchi-ole-common/src/object/model.rs",
    "crates/litchi-ole-common/src/object/editor.rs",
    "crates/litchi-xlsx/src/workbook/model.rs",
    "crates/litchi-xlsx/tests/reader_tolerance.rs",
    "crates/litchi-xlsx/src/xml_maps/tests.rs",
    "crates/soapberry-zip/src/archive.rs",
    "crates/soapberry-zip/src/office.rs",
    "crates/soapberry-zip/src/preserve.rs",
    "docs/adr/0005-io-memory-and-performance.md",
    "docs/adr/0006-validation-security-and-compatibility.md",
    "docs/performance/results/security-corpus-v3/catalog.json",
    "docs/performance/results/security-corpus-v3/oracles.json",
    "docs/performance/results/security-corpus-v3/profile.json",
    "docs/performance/results/security-corpus-v3/README.md",
];

type Result<T> = std::result::Result<T, Box<dyn Error>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ContainerKind {
    Zip,
    Cfb,
    EncryptedCfb,
}

impl ContainerKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Zip => "OOXML/ZIP",
            Self::Cfb => "CFB/OLE2",
            Self::EncryptedCfb => "OOXML encrypted CFB",
        }
    }

    const fn is_zip(self) -> bool {
        matches!(self, Self::Zip)
    }
}

#[derive(Clone, Copy, Debug)]
struct CryptoProfile {
    mode: Mode,
    password: &'static str,
    strict_data_spaces_refusal: bool,
}

#[derive(Clone, Copy, Debug)]
enum Oracle {
    Signed,
    Protected,
    External,
    Macro,
    EncryptedDoc {
        password: &'static str,
        semantic_sha256: &'static str,
    },
    Malformed,
    CfbBoundary,
    Xxe,
    SharedStrings,
    Crypto(CryptoProfile),
}

impl Oracle {
    const fn name(self) -> &'static str {
        match self {
            Self::Signed => "signed-ooxml-verify-noop-refusal",
            Self::Protected => "protected-docx-noop-refusal",
            Self::External => "external-xlsx-inventory-no-fetch",
            Self::Macro => "macro-xls-inert-preservation",
            Self::EncryptedDoc { .. } => "encrypted-ole-password-readback",
            Self::Malformed => "opc-malformed-refusal",
            Self::CfbBoundary => "cfb-boundary-read",
            Self::Xxe => "xlsx-xxe-schema-inert-read",
            Self::SharedStrings => "xlsx-malformed-shared-string-read",
            Self::Crypto(_) => "ooxml-encryption-interop-readback",
        }
    }

    const fn expected_behavior(self) -> &'static str {
        match self {
            Self::Malformed => "refuse",
            Self::CfbBoundary
            | Self::Signed
            | Self::Protected
            | Self::External
            | Self::Macro
            | Self::EncryptedDoc { .. }
            | Self::Xxe
            | Self::SharedStrings
            | Self::Crypto(_) => "accept",
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct Fixture {
    id: &'static str,
    path: &'static str,
    name: &'static str,
    kind: ContainerKind,
    categories: &'static [&'static str],
    case_name: &'static str,
    role: &'static str,
    target: &'static str,
    source_repository: &'static str,
    oracle: Oracle,
}

const POI_SOURCE: &str = "Apache POI test-data snapshot";
const CRYPTO_SOURCE: &str = "synthetic interoperability fixture";

const FIXTURES: &[Fixture] = &[
    Fixture {
        id: "poi-signed-docx",
        path: "../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx",
        name: "Apache POI signed DOCX fixture",
        kind: ContainerKind::Zip,
        categories: &["fixture", "security", "signed", "ooxml"],
        case_name: "security-signed-ooxml-verify-noop-refusal",
        role: "guard",
        target: "_xmlsignatures/sig1.xml",
        source_repository: POI_SOURCE,
        oracle: Oracle::Signed,
    },
    Fixture {
        id: "poi-signed-xlsx",
        path: "../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx",
        name: "Apache POI signed XLSX fixture",
        kind: ContainerKind::Zip,
        categories: &["fixture", "security", "signed", "ooxml"],
        case_name: "security-signed-ooxml-verify-noop-refusal",
        role: "guard",
        target: "_xmlsignatures/sig1.xml",
        source_repository: POI_SOURCE,
        oracle: Oracle::Signed,
    },
    Fixture {
        id: "poi-signed-pptx",
        path: "../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx",
        name: "Apache POI signed PPTX fixture",
        kind: ContainerKind::Zip,
        categories: &["fixture", "security", "signed", "ooxml"],
        case_name: "security-signed-ooxml-verify-noop-refusal",
        role: "guard",
        target: "_xmlsignatures/sig1.xml",
        source_repository: POI_SOURCE,
        oracle: Oracle::Signed,
    },
    Fixture {
        id: "ooxml-protected-docx",
        path: "../../test-data/ooxml/docx/documentProtection_readonly_no_password.docx",
        name: "OOXML document-protection fixture",
        kind: ContainerKind::Zip,
        categories: &["fixture", "protected", "security", "ooxml"],
        case_name: "security-protected-docx-noop-refusal",
        role: "guard",
        target: "word/settings.xml",
        source_repository: "repository fixture",
        oracle: Oracle::Protected,
    },
    Fixture {
        id: "ooxml-external-startup-xlsx",
        path: "../../test-data/ooxml/xlsx/external-link-path-startup.xlsx",
        name: "OOXML external-link inventory fixture",
        kind: ContainerKind::Zip,
        categories: &["external-links", "fixture", "ooxml", "security"],
        case_name: "security-external-xlsx-inventory-no-fetch",
        role: "inventory-only",
        target: "xl/externalLinks/externalLink1.xml",
        source_repository: "repository fixture",
        oracle: Oracle::External,
    },
    Fixture {
        id: "poi-simple-macro-xls",
        path: "../../test-data/poi/test-data/spreadsheet/SimpleMacro.xls",
        name: "Apache POI VBA-bearing XLS fixture",
        kind: ContainerKind::Cfb,
        categories: &["cfb", "fixture", "macro", "security"],
        case_name: "security-macro-xls-inert-preservation",
        role: "guard",
        target: "__raw_archive__",
        source_repository: POI_SOURCE,
        oracle: Oracle::Macro,
    },
    Fixture {
        id: "poi-cryptoapi-encrypted-doc",
        path: "../../test-data/poi/test-data/document/password_password_cryptoapi.doc",
        name: "Apache POI CryptoAPI encrypted DOC fixture",
        kind: ContainerKind::Cfb,
        categories: &["cfb", "encrypted", "fixture", "security"],
        case_name: "security-encrypted-ole-password-readback",
        role: "guard",
        target: "__raw_archive__",
        source_repository: POI_SOURCE,
        oracle: Oracle::EncryptedDoc {
            password: "password",
            semantic_sha256: "6dd4273bea0a8f70f4b6d8448e0ea1cb22b54713ad78d177394bd8b496e0aea6",
        },
    },
    Fixture {
        id: "poi-binaryrc4-encrypted-doc",
        path: "../../test-data/poi/test-data/document/password_tika_binaryrc4.doc",
        name: "Apache POI binary RC4 encrypted DOC fixture",
        kind: ContainerKind::Cfb,
        categories: &["cfb", "encrypted", "fixture", "security"],
        case_name: "security-encrypted-ole-password-readback",
        role: "guard",
        target: "__raw_archive__",
        source_repository: POI_SOURCE,
        oracle: Oracle::EncryptedDoc {
            password: "tika",
            semantic_sha256: "5c5c945257fcd1569b5161722e15a9d73283daf786aca1daf04d0029d3736b78",
        },
    },
    Fixture {
        id: "poi-opc-multiple-core-properties",
        path: "../../test-data/poi/test-data/openxml4j/OPCCompliance_CoreProperties_OnlyOneCorePropertiesPartFAIL.docx",
        name: "Apache POI OPC malformed multiple-core-properties fixture",
        kind: ContainerKind::Zip,
        categories: &["fixture", "malformed", "ooxml", "security"],
        case_name: "security-opc-malformed-refusal",
        role: "guard",
        target: "[Content_Types].xml",
        source_repository: POI_SOURCE,
        oracle: Oracle::Malformed,
    },
    Fixture {
        id: "poi-opc-derived-part-name",
        path: "../../test-data/poi/test-data/openxml4j/OPCCompliance_DerivedPartNameFAIL.docx",
        name: "Apache POI OPC malformed derived-part-name fixture",
        kind: ContainerKind::Zip,
        categories: &["fixture", "malformed", "ooxml", "security"],
        case_name: "security-opc-malformed-refusal",
        role: "guard",
        target: "word/document.xml",
        source_repository: POI_SOURCE,
        oracle: Oracle::Malformed,
    },
    Fixture {
        id: "cfb-short-final-sector",
        path: "../../test-data/ole/doc/cfb-truncated-final-sector.doc",
        name: "CFB short-final-sector compatibility fixture",
        kind: ContainerKind::Cfb,
        categories: &["adversarial", "cfb", "fixture", "malformed"],
        case_name: "security-cfb-boundary-read",
        role: "guard",
        target: "__raw_archive__",
        source_repository: "repository fixture",
        oracle: Oracle::CfbBoundary,
    },
    Fixture {
        id: "cfb-uninitialized-size-high-word",
        path: "../../test-data/ole/doc/cfb-v3-uninitialized-size-high-word.doc",
        name: "CFB version-3 uninitialized-size-word compatibility fixture",
        kind: ContainerKind::Cfb,
        categories: &["adversarial", "cfb", "fixture", "malformed"],
        case_name: "security-cfb-boundary-read",
        role: "guard",
        target: "__raw_archive__",
        source_repository: "repository fixture",
        oracle: Oracle::CfbBoundary,
    },
    Fixture {
        id: "poi-xxe-schema-inert-xlsx",
        path: "../../test-data/poi/test-data/spreadsheet/xxe_in_schema.xlsx",
        name: "Apache POI XML-maps external-entity-looking fixture",
        kind: ContainerKind::Zip,
        categories: &[
            "adversarial",
            "fixture",
            "ooxml",
            "security",
            "unknown-extension",
        ],
        case_name: "security-xlsx-adversarial-inert-read",
        role: "inventory-only",
        target: "xl/xmlMaps.xml",
        source_repository: POI_SOURCE,
        oracle: Oracle::Xxe,
    },
    Fixture {
        id: "ooxml-malformed-shared-string-hints",
        path: "../../test-data/ooxml/xlsx/shared-strings-malformed-count.xlsx",
        name: "OOXML malformed shared-string hint fixture",
        kind: ContainerKind::Zip,
        categories: &["adversarial", "fixture", "malformed", "ooxml"],
        case_name: "security-xlsx-adversarial-inert-read",
        role: "guard",
        target: "xl/sharedStrings.xml",
        source_repository: "repository fixture",
        oracle: Oracle::SharedStrings,
    },
    Fixture {
        id: "crypto-component-standard-aes192",
        path: "../../crates/litchi-crypto/tests/data/ooxml/component-standard-aes192-sha1.docx",
        name: "Litchi crypto component Standard AES-192/SHA-1 fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::StandardAes192,
            password: "Litchi synthetic crypto fixture 2026",
            strict_data_spaces_refusal: false,
        }),
    },
    Fixture {
        id: "crypto-component-standard-aes256",
        path: "../../crates/litchi-crypto/tests/data/ooxml/component-standard-aes256-sha1.docx",
        name: "Litchi crypto component Standard AES-256/SHA-1 fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::StandardAes256,
            password: "Litchi synthetic crypto fixture 2026",
            strict_data_spaces_refusal: false,
        }),
    },
    Fixture {
        id: "crypto-component-agile-mixed-key-size",
        path: "../../crates/litchi-crypto/tests/data/ooxml/component-agile-aes128-wrap-aes256-data-sha512.docx",
        name: "Litchi crypto component Agile mixed key-size fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::AgileAes256Sha512,
            password: "Litchi synthetic crypto fixture 2026",
            strict_data_spaces_refusal: false,
        }),
    },
    Fixture {
        id: "crypto-msoffcrypto-agile-valid-graph",
        path: "../../crates/litchi-crypto/tests/data/ooxml/msoffcrypto-agile-aes256-sha512-valid-graph.docx",
        name: "msoffcrypto Agile valid-graph interoperability fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::AgileAes256Sha512,
            password: "Litchi synthetic crypto fixture 2026",
            strict_data_spaces_refusal: false,
        }),
    },
    Fixture {
        id: "crypto-msoffcrypto-agile-zero-block-graph",
        path: "../../crates/litchi-crypto/tests/data/ooxml/msoffcrypto-agile-aes256-sha512.docx",
        name: "msoffcrypto Agile zero-DataSpaces-block compatibility fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "advisory-compatibility",
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-advisory-compatibility",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::AgileAes256Sha512,
            password: "Litchi synthetic crypto fixture 2026",
            strict_data_spaces_refusal: true,
        }),
    },
    Fixture {
        id: "crypto-rust-standard-aes192",
        path: "../../crates/litchi-crypto/tests/data/ooxml/rust-standard-aes192.docx",
        name: "Litchi crypto authored Standard AES-192/SHA-1 fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::StandardAes192,
            password: "Litchi rust interop output 2026",
            strict_data_spaces_refusal: false,
        }),
    },
    Fixture {
        id: "crypto-rust-standard-aes256",
        path: "../../crates/litchi-crypto/tests/data/ooxml/rust-standard-aes256.docx",
        name: "Litchi crypto authored Standard AES-256/SHA-1 fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::StandardAes256,
            password: "Litchi rust interop output 2026",
            strict_data_spaces_refusal: false,
        }),
    },
    Fixture {
        id: "crypto-rust-agile-aes256-sha512",
        path: "../../crates/litchi-crypto/tests/data/ooxml/rust-agile-aes256-sha512.docx",
        name: "Litchi crypto authored Agile AES-256/SHA-512 fixture",
        kind: ContainerKind::EncryptedCfb,
        categories: &[
            "crypto",
            "encrypted",
            "fixture",
            "interop",
            "ooxml",
            "synthetic",
        ],
        case_name: "security-ooxml-encryption-interop-readback",
        role: "guard",
        target: "EncryptedPackage",
        source_repository: CRYPTO_SOURCE,
        oracle: Oracle::Crypto(CryptoProfile {
            mode: Mode::AgileAes256Sha512,
            password: "Litchi rust interop output 2026",
            strict_data_spaces_refusal: false,
        }),
    },
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    Read,
    Validate,
    Preserve,
}

impl Operation {
    const ALL: [Self; 3] = [Self::Read, Self::Validate, Self::Preserve];

    const fn name(self) -> &'static str {
        match self {
            Self::Read => "bounded_read",
            Self::Validate => "validate",
            Self::Preserve => "preserve",
        }
    }

    const fn selector(self) -> &'static str {
        match self {
            Self::Read => "security-corpus --operation read",
            Self::Validate => "security-corpus --operation validate",
            Self::Preserve => "security-corpus --operation preserve",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "read" | "bounded_read" => Some(Self::Read),
            "validate" => Some(Self::Validate),
            "preserve" => Some(Self::Preserve),
            _ => None,
        }
    }
}

#[derive(Debug)]
struct OperationResult {
    logical_bytes: usize,
    output_bytes: usize,
    output_sha256: String,
}

#[derive(Debug, Serialize)]
struct Statistics {
    samples: Vec<u64>,
    min_ns: u64,
    max_ns: u64,
    mean_ns: u64,
}

#[derive(Debug, Serialize)]
struct PerfRow {
    fixture_id: &'static str,
    case: &'static str,
    role: &'static str,
    operation: &'static str,
    selector: &'static str,
    format: &'static str,
    executable: bool,
    timed: bool,
    input_bytes: usize,
    input_sha256: String,
    logical_bytes: usize,
    output_bytes: usize,
    output_sha256: String,
    warmup_iterations: usize,
    sample_count: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    elapsed_ns: Option<Statistics>,
    admission: &'static str,
    host_code_execution: bool,
    external_reference_resolution: bool,
}

#[derive(Debug)]
struct Cli {
    operations: Vec<Operation>,
    fixtures: Vec<&'static Fixture>,
    samples: usize,
    warmup: usize,
    measure: bool,
    json: Option<PathBuf>,
    catalog_dir: Option<PathBuf>,
}

fn main() -> Result<()> {
    let cli = parse_cli()?;
    if cli.measure && std::env::var(MEASURE_ENV).ok().as_deref() != Some(MEASURE_TOKEN) {
        return Err(format!(
            "--measure requires {MEASURE_ENV}={MEASURE_TOKEN}; no capture is enabled by default"
        )
        .into());
    }
    if let Some(directory) = &cli.catalog_dir {
        write_catalog(directory)?;
    }

    let mut rows = Vec::new();
    for fixture in &cli.fixtures {
        let bytes = load_fixture(fixture)?;
        for &operation in &cli.operations {
            rows.push(run_row(
                fixture,
                &bytes,
                operation,
                cli.warmup,
                cli.samples,
                cli.measure,
            )?);
        }
    }

    let report = json!({
        "schema": "litchi-security-corpus-perf-v3",
        "catalog_id": CATALOG_ID,
        "profile_id": PROFILE_ID,
        "default_bindings_unchanged": 201,
        "timed_capture": cli.measure,
        "rows": rows,
        "policy": {
            "host_code_execution": false,
            "external_reference_resolution": false,
            "macro_policy": "inventory and inert preservation only; never execute",
            "producer_claim": "repository fixture provenance; document producer remains unknown unless independently recorded",
        },
    });
    write_json(cli.json.as_deref(), &report)
}

fn parse_cli() -> Result<Cli> {
    let mut operations = Operation::ALL.to_vec();
    let mut fixtures = FIXTURES.iter().collect::<Vec<_>>();
    let mut samples = 15usize;
    let mut warmup = 3usize;
    let mut measure = false;
    let mut json_path = None;
    let mut catalog_dir = None;
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--operation" => {
                let value = args
                    .next()
                    .ok_or("--operation requires a comma-separated value")?;
                operations = value
                    .split(',')
                    .map(|value| {
                        Operation::parse(value)
                            .ok_or_else(|| format!("unknown security corpus operation {value:?}"))
                    })
                    .collect::<std::result::Result<Vec<_>, _>>()?;
                if operations.is_empty() {
                    return Err("--operation must select at least one operation".into());
                }
            },
            "--fixture" => {
                let value = args
                    .next()
                    .ok_or("--fixture requires a comma-separated value")?;
                fixtures = value
                    .split(',')
                    .map(|id| {
                        FIXTURES
                            .iter()
                            .find(|fixture| fixture.id == id)
                            .ok_or_else(|| format!("unknown security corpus fixture {id:?}"))
                    })
                    .collect::<std::result::Result<Vec<_>, _>>()?;
                if fixtures.is_empty() {
                    return Err("--fixture must select at least one fixture".into());
                }
            },
            "--samples" => {
                samples = args
                    .next()
                    .ok_or("--samples requires a positive integer")?
                    .parse()?;
                if samples == 0 {
                    return Err("--samples must be greater than zero".into());
                }
            },
            "--warmup" => {
                warmup = args
                    .next()
                    .ok_or("--warmup requires a non-negative integer")?
                    .parse()?;
            },
            "--measure" => measure = true,
            "--json" => {
                let value = args.next().ok_or("--json requires PATH or -")?;
                if value != "-" {
                    json_path = Some(PathBuf::from(value));
                }
            },
            "--catalog-dir" => {
                catalog_dir = Some(PathBuf::from(
                    args.next().ok_or("--catalog-dir requires PATH")?,
                ));
            },
            "--help" | "-h" => {
                println!(
                    "security-corpus [--operation read,validate,preserve] [--fixture ID,...] \\\n                     [--samples N] [--warmup N] [--measure] [--json PATH|-] [--catalog-dir PATH]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unrecognized argument {other:?}; use --help").into()),
        }
    }
    Ok(Cli {
        operations,
        fixtures,
        samples,
        warmup,
        measure,
        json: json_path,
        catalog_dir,
    })
}

fn run_row(
    fixture: &Fixture,
    bytes: &[u8],
    operation: Operation,
    warmup: usize,
    samples: usize,
    measure: bool,
) -> Result<PerfRow> {
    let checked = execute(fixture, bytes, operation)?;
    let mut elapsed = None;
    if measure {
        for _ in 0..warmup {
            let warm = execute(fixture, bytes, operation)?;
            ensure_same_result(&checked, &warm, fixture, operation)?;
        }
        let mut values = Vec::with_capacity(samples);
        for _ in 0..samples {
            let started = Instant::now();
            let result = execute(fixture, bytes, operation)?;
            let elapsed_ns = u64::try_from(started.elapsed().as_nanos())?;
            ensure_same_result(&checked, &result, fixture, operation)?;
            values.push(elapsed_ns);
        }
        elapsed = Some(statistics(values)?);
    }
    Ok(PerfRow {
        fixture_id: fixture.id,
        case: fixture.case_name,
        role: fixture.role,
        operation: operation.name(),
        selector: operation.selector(),
        format: fixture.kind.name(),
        executable: true,
        timed: measure,
        input_bytes: bytes.len(),
        input_sha256: sha256_hex(bytes),
        logical_bytes: checked.logical_bytes,
        output_bytes: checked.output_bytes,
        output_sha256: checked.output_sha256,
        warmup_iterations: if measure { warmup } else { 0 },
        sample_count: if measure { samples } else { 1 },
        elapsed_ns: elapsed,
        admission: "stat length before Vec; bounded Rust ZIP/CFB admission before metadata or payload allocation",
        host_code_execution: false,
        external_reference_resolution: false,
    })
}

fn execute(fixture: &Fixture, bytes: &[u8], operation: Operation) -> Result<OperationResult> {
    match operation {
        Operation::Read => read_fixture(fixture, bytes),
        Operation::Validate => validate_fixture(fixture, bytes),
        Operation::Preserve => preserve_fixture(fixture, bytes),
    }
}

fn read_fixture(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    if matches!(fixture.oracle, Oracle::Malformed) {
        let _ = validate_fixture(fixture, bytes)?;
        return Ok(empty_result(bytes));
    }
    match fixture.kind {
        ContainerKind::Zip => {
            let (archive, _) = bounded_archive(fixture.id, bytes)?;
            let payload = archive.read(fixture.target)?;
            Ok(result_from_bytes(payload))
        },
        ContainerKind::Cfb => {
            let payload = read_cfb_stream(fixture, bytes)?;
            Ok(result_from_bytes(payload))
        },
        ContainerKind::EncryptedCfb => read_crypto(fixture, bytes),
    }
}

fn validate_fixture(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    match fixture.kind {
        ContainerKind::Zip => validate_zip(fixture, bytes),
        ContainerKind::Cfb => validate_cfb(fixture, bytes),
        ContainerKind::EncryptedCfb => validate_crypto(fixture, bytes),
    }
}

fn preserve_fixture(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    if matches!(fixture.oracle, Oracle::Malformed) {
        let _ = validate_fixture(fixture, bytes)?;
        return Ok(empty_result(bytes));
    }
    match fixture.kind {
        ContainerKind::Zip => {
            let mut locate_buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
            let archive = ZipArchive::from_seekable(Cursor::new(bytes), &mut locate_buffer)?;
            admit_zip_archive_entries(fixture.id, &archive)?;
            let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
            let index =
                PreservationIndex::new_with_limits(&archive, &mut buffer, archive_limits())?;
            let output = index.write_to(&PreservationPlan::copy_all(&index), Vec::new())?;
            if output.as_slice() != bytes {
                return Err(format!("{} no-op ZIP preservation changed bytes", fixture.id).into());
            }
            Ok(result_from_bytes(output))
        },
        ContainerKind::Cfb | ContainerKind::EncryptedCfb => {
            let _ = validate_fixture(fixture, bytes)?;
            let mut output = Vec::new();
            output.try_reserve_exact(bytes.len())?;
            output.extend_from_slice(bytes);
            if output.as_slice() != bytes {
                return Err(format!("{} no-op CFB preservation changed bytes", fixture.id).into());
            }
            Ok(result_from_bytes(output))
        },
    }
}

fn validate_zip(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    let package = OpcPackage::from_vec_with_limits(bytes.to_vec(), opc_limits()?);
    if matches!(fixture.oracle, Oracle::Malformed) {
        if package.is_ok() {
            return Err(format!("{} unexpectedly opened as an OPC package", fixture.id).into());
        }
        return Ok(empty_result(bytes));
    }
    let package = package?;
    let (archive, _) = bounded_archive(fixture.id, bytes)?;
    let target = archive.read(fixture.target)?;
    match fixture.oracle {
        Oracle::Signed => {
            if !package.is_signed() {
                return Err(format!("{} lost its signature graph", fixture.id).into());
            }
            let reports = package.signatures_with(&Policy::compatible())?;
            if reports.is_empty()
                || reports.iter().any(|report| {
                    report.integrity() != Status::Valid || report.signature() != Status::Valid
                })
            {
                return Err(format!("{} signature verification failed", fixture.id).into());
            }
            let source = SourceBackedPackage::from_vec_with_limits(bytes.to_vec(), opc_limits()?)?;
            let (target, replacement) = {
                let main = source.main_document_part()?;
                let target = main.partname().clone();
                let mut replacement = main.data()?.as_bytes().to_vec();
                replacement.push(b' ');
                (target, replacement)
            };
            let mut output = Vec::new();
            if !matches!(
                source.write_part_overlay_to_stream(&mut output, &target, replacement),
                Err(OpcError::SignedSourceRequiresExplicitPolicy)
            ) || !output.is_empty()
            {
                return Err(
                    format!("{} signed edit was not refused before output", fixture.id).into(),
                );
            }
        },
        Oracle::Protected => {
            if !target
                .windows(b"documentProtection".len())
                .any(|window| window == b"documentProtection")
            {
                return Err(format!("{} protection marker missing", fixture.id).into());
            }
            let source = SourceBackedPackage::from_vec_with_limits(bytes.to_vec(), opc_limits()?)?;
            let docx = DocxSourcePackage::from_source_backed_package(source)?;
            let mut edit = docx.edit_document_variables()?;
            edit.set_variable("security_matrix", "must-refuse")?;
            let commit = edit.commit()?;
            if !commit.changed() {
                return Err(
                    format!("{} protected edit was unexpectedly unchanged", fixture.id).into(),
                );
            }
            let mut output = Vec::new();
            if !matches!(
                docx.publish_document_variables_commit_to_stream(&mut output, &commit),
                Err(litchi_docx::Error::UnsafeEdit { .. })
            ) || !output.is_empty()
            {
                return Err(format!(
                    "{} protected edit was not refused before output",
                    fixture.id
                )
                .into());
            }
        },
        Oracle::External => {
            let inventory = collect_external_targets(&package);
            let expected = "part=/xl/externalLinks/externalLink1.xml|rId1|http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup|personal.xls|external";
            if inventory != [expected.to_owned()] {
                return Err(format!("{} external inventory changed", fixture.id).into());
            }
            if archive.contains("personal.xls") {
                return Err(format!("{} unexpectedly fetched external target", fixture.id).into());
            }
        },
        Oracle::Xxe => {
            let _ = XlsxWorkbook::from_slice_with_limits(bytes, opc_limits()?)?;
            if !target
                .windows(b"schemaLocation".len())
                .any(|window| window == b"schemaLocation")
                || !target
                    .windows(b"http://localhost".len())
                    .any(|window| window == b"http://localhost")
            {
                return Err(format!("{} XML-maps target missing", fixture.id).into());
            }
        },
        Oracle::SharedStrings => {
            let _ = XlsxWorkbook::from_slice_with_limits(bytes, opc_limits()?)?;
            if !target.windows(b"sst".len()).any(|window| window == b"sst") {
                return Err(format!("{} shared-string target missing", fixture.id).into());
            }
        },
        Oracle::Malformed => unreachable!("handled before package unwrap"),
        Oracle::CfbBoundary | Oracle::Macro | Oracle::EncryptedDoc { .. } | Oracle::Crypto(_) => {
            return Err(format!("{} has a non-ZIP oracle", fixture.id).into());
        },
    }
    Ok(OperationResult {
        logical_bytes: target.len(),
        output_bytes: 0,
        output_sha256: sha256_hex(bytes),
    })
}

fn validate_cfb(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    let mut ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits()?)?;
    match fixture.oracle {
        Oracle::Macro => {
            let mut workbook = XlsWorkbook::from_ole_file(OleFile::open_with_limits(
                Cursor::new(bytes),
                cfb_limits()?,
            )?)?;
            let metadata = workbook.vba_metadata();
            if !metadata.has_project_marker()
                || !metadata.has_project_storage()
                || !metadata.may_contain_executable_code()
            {
                return Err(format!("{} VBA marker is incomplete", fixture.id).into());
            }
            let storage = workbook
                .vba_project_storage()
                .ok_or("macro fixture has no VBA project storage")?;
            if !storage.is_structurally_complete() || !storage.may_contain_macro_code() {
                return Err(format!("{} VBA storage is incomplete", fixture.id).into());
            }
            let project = workbook.vba()?.ok_or("macro fixture has no VBA project")?;
            if !project
                .modules()
                .iter()
                .any(|module| module.source().text().contains("Sub "))
            {
                return Err(
                    format!("{} VBA project has no inert module source", fixture.id).into(),
                );
            }
        },
        Oracle::EncryptedDoc {
            password,
            semantic_sha256,
        } => validate_encrypted_doc(fixture, bytes, password, semantic_sha256)?,
        Oracle::CfbBoundary => {
            let payload = open_named_stream(&mut ole, "WordDocument")?;
            if payload.is_empty() {
                return Err(format!("{} WordDocument stream is empty", fixture.id).into());
            }
        },
        Oracle::Signed
        | Oracle::Protected
        | Oracle::External
        | Oracle::Malformed
        | Oracle::Xxe
        | Oracle::SharedStrings
        | Oracle::Crypto(_) => return Err(format!("{} has a non-CFB oracle", fixture.id).into()),
    }
    let stream_count = ole.list_streams().len();
    Ok(OperationResult {
        logical_bytes: stream_count,
        output_bytes: 0,
        output_sha256: sha256_hex(bytes),
    })
}

fn validate_crypto(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    let Oracle::Crypto(profile) = fixture.oracle else {
        return Err(format!("{} has a non-crypto oracle", fixture.id).into());
    };
    let limits = crypto_limits();
    enforce_strict_data_spaces_oracle(fixture, bytes, profile)?;
    if ooxml::inspect_with(bytes, &limits)? != Kind::Encrypted(profile.mode) {
        return Err(format!("{} encryption mode changed", fixture.id).into());
    }
    let opened = ooxml::open_with(bytes.to_vec(), profile.password, &limits)?;
    if opened.mode() != Some(profile.mode) {
        return Err(format!("{} opened under the wrong encryption mode", fixture.id).into());
    }
    if opened.bytes().is_empty() || sha256_hex(opened.bytes()) != CLEAR_CRYPTO_SHA256 {
        return Err(format!("{} clear-package digest changed", fixture.id).into());
    }
    if profile.mode.is_agile() && opened.integrity() != Some(IntegrityStatus::Authenticated) {
        return Err(format!("{} lost Agile integrity authentication", fixture.id).into());
    }
    Ok(OperationResult {
        logical_bytes: opened.bytes().len(),
        output_bytes: 0,
        output_sha256: sha256_hex(bytes),
    })
}

fn validate_encrypted_doc(
    fixture: &Fixture,
    bytes: &[u8],
    password: &str,
    semantic_sha256: &str,
) -> Result<()> {
    let mut package = open_doc_package(bytes)?;
    if !matches!(package.document(), Err(DocError::PasswordRequired)) {
        return Err(format!("{} did not require a password", fixture.id).into());
    }
    let mut package = open_doc_package(bytes)?;
    if !matches!(
        package
            .document_with_options(OpenOptions::default().with_password("wrong".to_owned().into())),
        Err(DocError::InvalidPassword)
    ) {
        return Err(format!("{} accepted the wrong password", fixture.id).into());
    }
    let mut package = open_doc_package(bytes)?;
    let document = package
        .document_with_options(OpenOptions::default().with_password(password.to_owned().into()))?;
    let text = document.text()?;
    if text.trim().is_empty() || sha256_hex(text.as_bytes()) != semantic_sha256 {
        return Err(format!("{} semantic password readback changed", fixture.id).into());
    }
    Ok(())
}

fn open_doc_package(bytes: &[u8]) -> Result<DocPackage<Cursor<&[u8]>>> {
    let ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits()?)?;
    let limits = DocLimits::try_new(
        MAX_INPUT_BYTES as usize,
        MAX_INPUT_BYTES as usize,
        MAX_INPUT_BYTES as usize,
    )?;
    Ok(DocPackage::from_ole_file_with_limits(ole, limits)?)
}

fn read_crypto(fixture: &Fixture, bytes: &[u8]) -> Result<OperationResult> {
    let Oracle::Crypto(profile) = fixture.oracle else {
        return Err(format!("{} has a non-crypto oracle", fixture.id).into());
    };
    enforce_strict_data_spaces_oracle(fixture, bytes, profile)?;
    let opened = ooxml::open_with(bytes.to_vec(), profile.password, &crypto_limits())?;
    if opened.mode() != Some(profile.mode)
        || sha256_hex(opened.bytes()) != CLEAR_CRYPTO_SHA256
        || (profile.mode.is_agile() && opened.integrity() != Some(IntegrityStatus::Authenticated))
    {
        return Err(format!("{} crypto readback changed", fixture.id).into());
    }
    Ok(result_from_bytes(opened.into_bytes()))
}

fn enforce_strict_data_spaces_oracle(
    fixture: &Fixture,
    bytes: &[u8],
    profile: CryptoProfile,
) -> Result<()> {
    if !profile.strict_data_spaces_refusal {
        return Ok(());
    }
    let mut ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits()?)?;
    if spaces::inspect(&mut ole).is_ok() {
        return Err(format!(
            "{} strict DataSpaces inspection unexpectedly accepted the zero-block graph",
            fixture.id
        )
        .into());
    }
    Ok(())
}

fn read_cfb_stream(fixture: &Fixture, bytes: &[u8]) -> Result<Vec<u8>> {
    let mut ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits()?)?;
    let wanted = match fixture.oracle {
        Oracle::Macro => "_VBA_PROJECT_CUR",
        Oracle::CfbBoundary | Oracle::EncryptedDoc { .. } => "WordDocument",
        _ => return Err(format!("{} has no CFB stream read oracle", fixture.id).into()),
    };
    if wanted == "_VBA_PROJECT_CUR" {
        let path = ole
            .list_streams()
            .into_iter()
            .find(|path| path.iter().any(|component| component == wanted))
            .ok_or("macro project stream is missing")?;
        let references = path.iter().map(String::as_str).collect::<Vec<_>>();
        Ok(ole.open_stream(&references)?)
    } else {
        open_named_stream(&mut ole, wanted)
    }
}

fn open_named_stream<R: Read + io::Seek>(ole: &mut OleFile<R>, name: &str) -> Result<Vec<u8>> {
    let path = ole
        .list_streams()
        .into_iter()
        .find(|path| path.last().is_some_and(|component| component == name))
        .ok_or_else(|| format!("CFB stream {name} is missing"))?;
    let references = path.iter().map(String::as_str).collect::<Vec<_>>();
    Ok(ole.open_stream(&references)?)
}

fn archive_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: MAX_MEMBERS,
        max_member_name_bytes: MAX_MEMBER_NAME_BYTES,
        max_metadata_bytes: MAX_INPUT_BYTES,
        max_compressed_size: MAX_INPUT_BYTES,
        max_entry_size: MAX_INPUT_BYTES,
        max_total_size: MAX_INPUT_BYTES,
    }
}

/// Admit every ZIP central-directory record before constructing the high-level
/// file index. `ArchiveLimits::max_files` intentionally counts only
/// non-directory members, so security callers must perform this separate
/// total-entry pass to bound directory records as well.
fn bounded_archive<'data>(
    fixture_id: &str,
    bytes: &'data [u8],
) -> Result<(ArchiveReader<'data>, usize)> {
    let preflight = ZipArchive::from_slice(bytes)?.into_zip_archive();
    let total_entries = admit_zip_archive_entries(fixture_id, &preflight)?;
    Ok((
        ArchiveReader::new_with_limits(bytes, archive_limits())?,
        total_entries,
    ))
}

fn admit_zip_archive_entries<R: ReaderAt>(
    fixture_id: &str,
    archive: &ZipArchive<R>,
) -> Result<usize> {
    let declared = usize::try_from(archive.entries_hint())
        .map_err(|_| format!("{fixture_id} ZIP entry count does not fit this platform"))?;
    if declared > MAX_MEMBERS {
        return Err(
            format!("{fixture_id} ZIP total entry count {declared} exceeds {MAX_MEMBERS}").into(),
        );
    }

    let mut buffer = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let mut entries = archive.entries_with_metadata_limit(&mut buffer, MAX_INPUT_BYTES);
    let mut actual = 0usize;
    while entries.next_entry()?.is_some() {
        actual = actual
            .checked_add(1)
            .ok_or("ZIP total entry count overflow")?;
        if actual > MAX_MEMBERS {
            return Err(format!("{fixture_id} ZIP total entry count exceeds {MAX_MEMBERS}").into());
        }
    }
    if actual != declared {
        return Err(format!(
            "{fixture_id} ZIP declared entry count {declared} differs from central-directory count {actual}"
        )
        .into());
    }
    Ok(actual)
}

fn opc_limits() -> Result<ReadLimits> {
    Ok(ReadLimits::builder()
        .max_input_bytes(MAX_INPUT_BYTES)?
        .max_archive_members(MAX_MEMBERS)?
        .max_archive_total_entries(MAX_MEMBERS)?
        .max_archive_member_name_bytes(MAX_MEMBER_NAME_BYTES)?
        .max_archive_metadata_bytes(MAX_INPUT_BYTES)?
        .max_archive_compressed_bytes(MAX_INPUT_BYTES)?
        .max_archive_entry_bytes(MAX_INPUT_BYTES)?
        .max_archive_total_bytes(MAX_INPUT_BYTES)?
        .max_parts(MAX_MEMBERS)?
        .max_part_bytes(MAX_INPUT_BYTES)?
        .max_total_part_bytes(MAX_INPUT_BYTES)?
        .max_content_types_bytes(MAX_INPUT_BYTES as usize)?
        .max_content_type_mappings(MAX_MEMBERS)?
        .max_relationship_parts(MAX_MEMBERS)?
        .max_relationship_xml_bytes(MAX_INPUT_BYTES as usize)?
        .max_total_relationship_xml_bytes(MAX_INPUT_BYTES as usize)?
        .max_relationships_per_part(MAX_RELATIONSHIPS)?
        .max_total_relationships(MAX_RELATIONSHIPS)?
        .max_relationship_graph_nodes(MAX_MEMBERS)?
        .max_xml_events(MAX_RELATIONSHIPS)?
        .max_total_relationship_xml_events(MAX_RELATIONSHIPS)?
        .max_xml_depth(256)?
        .max_xml_attribute_bytes(MAX_INPUT_BYTES as usize)?
        .max_relationship_target_bytes(MAX_INPUT_BYTES as usize)?
        .build()?)
}

fn cfb_limits() -> Result<OleFileLimits> {
    Ok(OleFileLimits::new(MAX_INPUT_BYTES)?
        .with_max_directory_bytes(MAX_MEMBERS as u64 * CFB_DIRECTORY_ENTRY_BYTES)?
        .with_max_allocation_table_bytes(MAX_INPUT_BYTES)?)
}

fn crypto_limits() -> CryptoLimits {
    CryptoLimits {
        max_input_bytes: MAX_INPUT_BYTES as usize,
        max_info_bytes: 1024 * 1024,
        max_xml_bytes: 1024 * 1024,
        max_xml_depth: 64,
        max_xml_nodes: 4_096,
        max_xml_attributes: 4_096,
        max_spin_count: 1_000_000,
        max_password_chars: 255,
        max_plaintext_bytes: MAX_INPUT_BYTES as usize,
        max_encrypted_bytes: MAX_INPUT_BYTES as usize,
        max_output_bytes: MAX_INPUT_BYTES as usize,
        max_cfb_directory_bytes: CFB_DIRECTORY_BYTES,
        max_cfb_allocation_table_bytes: MAX_INPUT_BYTES as usize,
        allow_missing_data_spaces: false,
    }
}

fn load_fixture(fixture: &Fixture) -> Result<Vec<u8>> {
    let path = fixture_path(fixture);
    let declared = fs::metadata(&path)?.len();
    let capacity = admit_input_length(declared)?;
    let mut file = File::open(&path)?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(capacity)?;
    let mut buffer = [0u8; 16 * 1024];
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        let next = bytes
            .len()
            .checked_add(count)
            .ok_or("fixture length overflow")?;
        if next > capacity {
            return Err(format!("{} grew beyond its admitted source length", fixture.id).into());
        }
        bytes.extend_from_slice(&buffer[..count]);
    }
    if bytes.len() != capacity {
        return Err(format!("{} changed length while being read", fixture.id).into());
    }
    let actual = sha256_hex(&bytes);
    let expected = expected_hash(fixture.id).ok_or("fixture hash table is incomplete")?;
    if actual != expected {
        return Err(format!(
            "{} content hash changed: {actual} != {expected}",
            fixture.id
        )
        .into());
    }
    Ok(bytes)
}

fn admit_input_length(length: u64) -> Result<usize> {
    if length > MAX_INPUT_BYTES {
        return Err(format!("input length {length} exceeds {MAX_INPUT_BYTES}").into());
    }
    Ok(usize::try_from(length)?)
}

fn fixture_path(fixture: &Fixture) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(fixture.path)
}

fn result_from_bytes(bytes: Vec<u8>) -> OperationResult {
    let logical_bytes = bytes.len();
    let output_sha256 = sha256_hex(&bytes);
    OperationResult {
        logical_bytes,
        output_bytes: bytes.len(),
        output_sha256,
    }
}

fn empty_result(bytes: &[u8]) -> OperationResult {
    OperationResult {
        logical_bytes: bytes.len(),
        output_bytes: 0,
        output_sha256: sha256_hex(bytes),
    }
}

fn ensure_same_result(
    expected: &OperationResult,
    actual: &OperationResult,
    fixture: &Fixture,
    operation: Operation,
) -> Result<()> {
    if expected.logical_bytes != actual.logical_bytes
        || expected.output_bytes != actual.output_bytes
        || expected.output_sha256 != actual.output_sha256
    {
        return Err(format!(
            "{} {} result changed between iterations",
            fixture.id,
            operation.name()
        )
        .into());
    }
    Ok(())
}

fn statistics(samples: Vec<u64>) -> Result<Statistics> {
    let min_ns = samples.iter().copied().min().ok_or("empty timing sample")?;
    let max_ns = samples.iter().copied().max().ok_or("empty timing sample")?;
    let total = samples
        .iter()
        .try_fold(0u64, |total, sample| total.checked_add(*sample))
        .ok_or("timing sample sum overflow")?;
    Ok(Statistics {
        mean_ns: total / u64::try_from(samples.len())?,
        samples,
        min_ns,
        max_ns,
    })
}

fn collect_external_targets(package: &OpcPackage) -> Vec<String> {
    let mut inventory = Vec::new();
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| relationship.is_external())
    {
        inventory.push(format_relationship("package", relationship));
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| relationship.is_external())
        {
            inventory.push(format_relationship(part.partname().as_str(), relationship));
        }
    }
    inventory.sort();
    inventory
}

fn format_relationship(owner: &str, relationship: &litchi_opc::Relationship) -> String {
    let mode = if relationship.is_external() {
        "external"
    } else {
        "internal"
    };
    format!(
        "part={owner}|{}|{}|{}|{mode}",
        relationship.r_id(),
        relationship.reltype(),
        relationship.target_ref()
    )
}

fn write_catalog(directory: &Path) -> Result<()> {
    fs::create_dir_all(directory)?;
    let profile = profile_json();
    write_json(Some(&directory.join("profile.json")), &profile)?;

    let mut corpora = Vec::with_capacity(FIXTURES.len());
    let mut bindings = Vec::with_capacity(FIXTURES.len());
    let mut oracle_rows = Vec::with_capacity(FIXTURES.len());
    for fixture in FIXTURES {
        let bytes = load_fixture(fixture)?;
        let corpus = corpus_json(fixture, &bytes)?;
        let corpus_id = corpus["id"].clone();
        corpora.push(corpus);
        bindings.push(json!({
            "fixture_id": fixture.id,
            "case": fixture.case_name,
            "corpus_id": corpus_id,
            "role": fixture.role,
            "operations": Operation::ALL.iter().map(|operation| json!({
                "name": operation.name(),
                "selector": operation.selector(),
                "timed": true,
                "executable": true,
            })).collect::<Vec<_>>(),
        }));
        oracle_rows.push(oracle_json(fixture, &corpus_id));
    }
    let content_set = json!({"corpora": corpora, "bindings": bindings});
    let content_set_sha256 = sha256_json(&content_set);
    let mut catalog = json!({
        "manifest_version": 3,
        "manifest_kind": "corpus-catalog",
        "catalog_id": CATALOG_ID,
        "canonicalization": {"algorithm": "stable-struct-json-utf8-v1", "hash": "sha256"},
        "catalog_sha256": "",
        "content_set_sha256": content_set_sha256,
        "profile_id": PROFILE_ID,
        "profile_sha256": sha256_json(&profile),
        "default_bindings_unchanged": 201,
        "build": {
            "tool": "litchi-security-corpus-perf",
            "tool_version": 3,
            "git_revision": option_env!("LITCHI_SECURITY_CORPUS_REVISION"),
            "git_worktree_dirty": true,
            "bounded_oracle": "Rust ArchiveReader/PreservationIndex/OleFile/litchi-crypto",
        },
        "operations": Operation::ALL.iter().map(|operation| json!({
            "name": operation.name(),
            "selector": operation.selector(),
            "timed": true,
            "capture": "requires --measure and explicit capture environment gate",
            "host_code_execution": false,
            "external_reference_resolution": false,
        })).collect::<Vec<_>>(),
        "corpora": corpora,
        "case_bindings": bindings,
    });
    let catalog_hash = {
        catalog["catalog_sha256"] = Value::String(String::new());
        sha256_json(&catalog)
    };
    catalog["catalog_sha256"] = Value::String(catalog_hash);
    write_json(Some(&directory.join("catalog.json")), &catalog)?;
    let oracles = json!({
        "schema_version": 3,
        "catalog_id": CATALOG_ID,
        "profile_id": PROFILE_ID,
        "rows": oracle_rows,
        "policy": {
            "default_catalog_bindings": 201,
            "timed_operation_definitions": 3,
            "timed_captures_recorded": false,
            "macros": "inspect-and-preserve-inertly; never execute",
            "external_links": "inventory only; never resolve or fetch",
            "producer_claim": "source repository only; all document producers are unknown or explicitly synthetic",
        },
    });
    write_json(Some(&directory.join("oracles.json")), &oracles)?;
    let source_manifest = source_manifest_json()?;
    write_json(
        Some(&directory.join("source-manifest.json")),
        &source_manifest,
    )?;
    Ok(())
}

fn profile_json() -> Value {
    json!({
        "profile_id": PROFILE_ID,
        "max_input_bytes": MAX_INPUT_BYTES,
        "max_members": MAX_MEMBERS,
        "max_zip_non_directory_entries": MAX_MEMBERS,
        "max_zip_total_entries": MAX_MEMBERS,
        "max_opc_zip_total_entries": MAX_MEMBERS,
        "max_member_bytes": MAX_INPUT_BYTES,
        "max_relationships": MAX_RELATIONSHIPS,
        "max_cfb_directory_bytes": CFB_DIRECTORY_BYTES,
        "max_cfb_allocation_table_bytes": MAX_INPUT_BYTES,
        "max_materialized_bytes": MAX_INPUT_BYTES,
        "max_output_bytes": MAX_INPUT_BYTES,
        "timed_operation_definitions": 3,
        "timed_captures_recorded": false,
        "host_code_execution": false,
        "external_reference_resolution": false,
        "allocation_order": {
            "filesystem": "metadata length is admitted before Vec reservation and bounded chunk reads",
            "zip": "a bounded Rust central-directory preflight counts every entry, including directories, before ArchiveReader or PreservationIndex ownership; ArchiveLimits::max_files remains the non-directory file limit",
            "cfb": "source length, header counts, 512KiB directory, and 64MiB allocation-table ceilings are checked before CFB index allocation",
            "crypto": "the 512KiB CFB directory and 64MiB allocation-table ceilings are applied through the explicit OOXML encryption profile before DataSpaces inspection or decryption",
            "opc": "ReadLimits apply an explicit total central-directory entry preflight, then bound the non-directory member index and part payload reads",
        },
        "notes": "Rows are executable correctness operations. ZIP total-entry admission is explicit because ArchiveLimits::max_files excludes directory records. The zero-DataSpaces-block fixture must refuse strict spaces inspection while compatibility OOXML decryption accepts it. Timing is opt-in and intentionally not captured by this candidate.",
    })
}

fn source_manifest_json() -> Result<Value> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut source_paths = Vec::with_capacity(SOURCE_PATHS.len());
    for relative in SOURCE_PATHS {
        let path = root.join(relative);
        let bytes = fs::read(&path)?;
        source_paths.push(json!({
            "path": relative,
            "bytes": bytes.len(),
            "sha256": sha256_hex(&bytes),
        }));
    }
    Ok(json!({
        "schema": "litchi-security-corpus-source-manifest-v3",
        "artifact": CATALOG_ID,
        "base_revision": option_env!("LITCHI_SECURITY_CORPUS_REVISION"),
        "default_bindings_unchanged": 201,
        "timed_captures_recorded": false,
        "source_paths": source_paths,
        "oracle_bindings": {
            "fixtures": FIXTURES.len(),
            "operations_per_fixture": Operation::ALL.len(),
            "rows": FIXTURES.len() * Operation::ALL.len(),
        },
        "policy": {
            "host_code_execution": false,
            "external_reference_resolution": false,
            "producer_claim": "unknown unless independently recorded; synthetic vectors remain synthetic",
        },
    }))
}

fn corpus_json(fixture: &Fixture, bytes: &[u8]) -> Result<Value> {
    let (members, logical_bytes, member_status, total_entries) = member_inventory(fixture, bytes)?;
    let (target_bytes, target_hash) = if fixture.target == "__raw_archive__" {
        (bytes.len(), sha256_hex(bytes))
    } else {
        let member = members
            .iter()
            .find(|member: &&Value| member["name"] == fixture.target)
            .ok_or_else(|| format!("{} target {} is absent", fixture.id, fixture.target))?;
        (
            member["logical_bytes"]
                .as_u64()
                .and_then(|bytes| usize::try_from(bytes).ok())
                .ok_or("target logical byte count is invalid")?,
            member["sha256"]
                .as_str()
                .ok_or("target digest is invalid")?
                .to_owned(),
        )
    };
    let relationship = relationship_inventory(fixture, bytes)?;
    let producer_kind = if fixture.source_repository == POI_SOURCE {
        "third-party-test-fixture-not-native-office"
    } else if fixture.source_repository == CRYPTO_SOURCE {
        "synthetic-interoperability-not-native-office"
    } else {
        "repository-fixture-producer-unknown"
    };
    let archive_hash = sha256_hex(bytes);
    Ok(json!({
        "id": format!(
            "{}:sha256:{archive_hash}",
            fixture.kind.name().replace(['/', ' '], "-").to_lowercase()
        ),
        "fixture_id": fixture.id,
        "name": fixture.name,
        "format": fixture.kind.name(),
        "categories": fixture.categories,
        "path": fixture.path.strip_prefix("../../").unwrap_or(fixture.path),
        "case": fixture.case_name,
        "role": fixture.role,
        "generator": "checked-in-fixture-content-v3",
        "provenance": {
            "source_kind": "repository-fixture",
            "provenance_kind": producer_kind,
            "source_repository": fixture.source_repository,
            "document_producer": Value::Null,
            "producer_status": "unknown",
            "producer_evidence": "The checked-in source identifies a fixture repository or synthetic vector; it does not establish native Office authoring.",
            "source_sha256": archive_hash,
        },
        "input": {
            "expected_behavior": fixture.oracle.expected_behavior(),
            "within_limits": true,
            "observed_input_bytes": bytes.len(),
        },
        "limits": {
            "profile_id": PROFILE_ID,
            "observed_input_bytes": bytes.len(),
            "observed_members": members.len(),
            "observed_total_entries": total_entries,
            "observed_materialized_bytes": logical_bytes,
            "observed_relationships": relationship["count"].clone(),
        },
        "relationships": relationship,
        "members": {"status": member_status, "items": members},
        "targets": [{"entry": fixture.target, "logical_bytes": target_bytes, "sha256": target_hash}],
        "security_oracle": fixture.oracle.name(),
        "coverage": {
            "operations": Operation::ALL.iter().map(|operation| operation.name()).collect::<Vec<_>>(),
            "timed_operation_definitions": Operation::ALL.len(),
            "timed_captures_recorded": false,
            "host_code_execution": false,
            "external_reference_resolution": false,
        },
    }))
}

fn member_inventory(
    fixture: &Fixture,
    bytes: &[u8],
) -> Result<(Vec<Value>, usize, &'static str, usize)> {
    match fixture.kind {
        ContainerKind::Zip => {
            let (archive, total_entries) = bounded_archive(fixture.id, bytes)?;
            let mut members = Vec::new();
            let mut logical_bytes = 0usize;
            for (ordinal, name) in archive.file_names().enumerate() {
                let metadata = archive.metadata(name)?;
                if metadata.uncompressed_size() > MAX_INPUT_BYTES {
                    return Err(format!("{} member {} exceeds bound", fixture.id, name).into());
                }
                let payload = archive.read(name)?;
                logical_bytes = logical_bytes
                    .checked_add(payload.len())
                    .ok_or("ZIP logical bytes overflow")?;
                if logical_bytes as u64 > MAX_INPUT_BYTES {
                    return Err(
                        format!("{} aggregate member bytes exceed bound", fixture.id).into(),
                    );
                }
                members.push(json!({
                    "ordinal": ordinal,
                    "name": name,
                    "logical_bytes": payload.len(),
                    "compressed_bytes": metadata.compressed_size(),
                    "sha256": sha256_hex(&payload),
                    "role": "package-member",
                }));
            }
            Ok((members, logical_bytes, "complete", total_entries))
        },
        ContainerKind::Cfb | ContainerKind::EncryptedCfb => {
            let mut ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits()?)?;
            let paths = ole.list_streams();
            if paths.len() > MAX_MEMBERS {
                return Err(format!("{} stream count exceeds bound", fixture.id).into());
            }
            let mut members = Vec::new();
            let mut logical_bytes = 0usize;
            for (ordinal, path) in paths.iter().enumerate() {
                let references = path.iter().map(String::as_str).collect::<Vec<_>>();
                let payload = ole.open_stream(&references)?;
                logical_bytes = logical_bytes
                    .checked_add(payload.len())
                    .ok_or("CFB logical bytes overflow")?;
                if logical_bytes as u64 > MAX_INPUT_BYTES {
                    return Err(
                        format!("{} aggregate stream bytes exceed bound", fixture.id).into(),
                    );
                }
                members.push(json!({
                    "ordinal": ordinal,
                    "name": path.join("/"),
                    "logical_bytes": payload.len(),
                    "sha256": sha256_hex(&payload),
                    "role": "cfb-stream",
                }));
            }
            Ok((members, logical_bytes, "complete", paths.len()))
        },
    }
}

fn relationship_inventory(fixture: &Fixture, bytes: &[u8]) -> Result<Value> {
    if !fixture.kind.is_zip() {
        return Ok(json!({
            "status": "not-applicable",
            "count": Value::Null,
            "evidence": "CFB container has no OPC relationship graph",
        }));
    }
    match OpcPackage::from_vec_with_limits(bytes.to_vec(), opc_limits()?) {
        Ok(package) => {
            let count = package.rels().len()
                + package
                    .iter_parts()
                    .map(|part| part.rels().len())
                    .sum::<usize>();
            if count > MAX_RELATIONSHIPS {
                return Err(format!("{} relationship count exceeds bound", fixture.id).into());
            }
            Ok(json!({
                "status": "complete",
                "count": count,
                "evidence": "bounded OPC package relationship graph",
            }))
        },
        Err(_) if matches!(fixture.oracle, Oracle::Malformed) => Ok(json!({
            "status": "refused",
            "count": Value::Null,
            "evidence": "bounded OPC package refusal oracle",
        })),
        Err(error) => Err(error.into()),
    }
}

fn oracle_json(fixture: &Fixture, corpus_id: &Value) -> Value {
    let mut expected = BTreeMap::new();
    match fixture.oracle {
        Oracle::Signed => {
            expected.insert("signature", json!("valid"));
            expected.insert("edit_publication", json!("refuse_without_explicit_policy"));
        },
        Oracle::Protected => {
            expected.insert("document_protection", json!("present"));
            expected.insert("changed_publication", json!("refuse"));
        },
        Oracle::External => {
            expected.insert("external_resolution", json!(false));
            expected.insert("inventory_only", json!(true));
        },
        Oracle::Macro => {
            expected.insert("macro_inventory", json!("present"));
            expected.insert("macro_execution", json!(false));
        },
        Oracle::EncryptedDoc {
            password: _,
            semantic_sha256,
        } => {
            expected.insert("password_required", json!(true));
            expected.insert("wrong_password", json!("invalid"));
            expected.insert("semantic_sha256", json!(semantic_sha256));
        },
        Oracle::Malformed => {
            expected.insert("open", json!("refuse"));
            expected.insert("output_bytes", json!(0));
        },
        Oracle::CfbBoundary => {
            expected.insert("open", json!("accept"));
            expected.insert("stream", json!("WordDocument"));
        },
        Oracle::Xxe => {
            expected.insert("opaque_schema_preserved", json!(true));
            expected.insert("external_resolution", json!(false));
        },
        Oracle::SharedStrings => {
            expected.insert("open", json!("accept"));
            expected.insert("nonempty_target", json!(true));
        },
        Oracle::Crypto(profile) => {
            expected.insert("mode", json!(profile.mode.to_string()));
            expected.insert("clear_package_sha256", json!(CLEAR_CRYPTO_SHA256));
            expected.insert("external_verifier", json!("msoffcrypto-tool"));
            if profile.strict_data_spaces_refusal {
                expected.insert("strict_dataspaces", json!("refuse_zero_block_size"));
                expected.insert("compatibility_reader", json!("accept_and_decrypt"));
                expected.insert("integrity", json!("authenticated"));
            }
        },
    }
    json!({
        "fixture_id": fixture.id,
        "case": fixture.case_name,
        "role": fixture.role,
        "corpus_id": corpus_id,
        "operations": Operation::ALL.iter().map(|operation| json!({
            "name": operation.name(),
            "selector": operation.selector(),
            "executable": true,
            "timed": true,
            "capture_recorded": false,
            "read_only_or_guarded": true,
        })).collect::<Vec<_>>(),
        "expected": expected,
        "host_code_execution": false,
        "external_reference_resolution": false,
    })
}

fn sha256_json(value: &Value) -> String {
    let bytes = serde_json::to_vec(value).expect("JSON value is serializable");
    sha256_hex(&bytes)
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn write_json(path: Option<&Path>, value: &Value) -> Result<()> {
    let bytes = serde_json::to_vec_pretty(value)?;
    if let Some(path) = path {
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }
        let mut file = File::create(path)?;
        file.write_all(&bytes)?;
        file.write_all(b"\n")?;
    } else {
        let mut stdout = io::stdout().lock();
        stdout.write_all(&bytes)?;
        stdout.write_all(b"\n")?;
    }
    Ok(())
}

fn expected_hash(id: &str) -> Option<&'static str> {
    Some(match id {
        "poi-signed-docx" => "bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f",
        "poi-signed-xlsx" => "4cbd8cbe613f036b7a0c779ffaaec7c5838710896c6ef26b3f27410d25d5ce45",
        "poi-signed-pptx" => "4d925d282dcca86e62b6716647a458246f8b9ea0eae0ec6664bbbf5a3f91bce1",
        "ooxml-protected-docx" => {
            "5d4c919f2e06b84fbe35cfaaa4012e8f469b811e1f643deb3e660b798bfe4544"
        },
        "ooxml-external-startup-xlsx" => {
            "e06155747da482bfb7c1ac5f0ab3a80cbe5b510e664926709c356ba6b59e9bc4"
        },
        "poi-simple-macro-xls" => {
            "0e92c9bb018abd8a5f9121d65827c9e3bd280777219cb77a2efd70635143c00a"
        },
        "poi-cryptoapi-encrypted-doc" => {
            "f2d0dc59ad7ec2356695ad5dc550057052a4017d5f1eb46e887297f5089896fb"
        },
        "poi-binaryrc4-encrypted-doc" => {
            "9231e724bb17a2e5f74815728d90b06e15684cf5fb2443a6fa24deebd33be952"
        },
        "poi-opc-multiple-core-properties" => {
            "e47a11c96726d915f8ed51537123bf2de1c11aa4c2bb2f97d20eb5c7c62deda5"
        },
        "poi-opc-derived-part-name" => {
            "a57fe168c5dcd49a877b4c41ec453571181c3a0644b5b305c76c6ff385b39c29"
        },
        "cfb-short-final-sector" => {
            "f0e4f66622c0b5f8ee1af1931d5a216be84768f5bbe53b20f6e9a7deec563609"
        },
        "cfb-uninitialized-size-high-word" => {
            "d1fa9699291b09ede96267581629125417db2048ddf78361b5d5c9641de7aa7a"
        },
        "poi-xxe-schema-inert-xlsx" => {
            "95dc84089f255d5878d89fdc1826edd112a9addd88b25a444f323aadc870c79d"
        },
        "ooxml-malformed-shared-string-hints" => {
            "c3c025bed1736a240e0e91960aaaaebd795d95b54a7b977c2e451406aa1f9462"
        },
        "crypto-component-standard-aes192" => {
            "40931c6c034e1d1a5542f801797a188c81d01fb399089c9a6df99f7d143f72a2"
        },
        "crypto-component-standard-aes256" => {
            "5820482f9b9318c98606e3a2394392be23fcf1c0d722d76654d017f5662a1500"
        },
        "crypto-component-agile-mixed-key-size" => {
            "d7f53bf7112b416c489afe24d0d0653fb563bf5ac644afe8bce3af90e332fe9c"
        },
        "crypto-msoffcrypto-agile-valid-graph" => {
            "65eb57739477b8a89f2be7c1188afd28caf88c3841b03797ae6035c37a1dabc4"
        },
        "crypto-msoffcrypto-agile-zero-block-graph" => {
            "4202f94dfc3d11ac64d5115e74d35e97cb9f96e09ef01d68b31ae04d10e61ac3"
        },
        "crypto-rust-standard-aes192" => {
            "83bb67744ce5037349c45d7f2ee6ab1cee05b285c600aba92bb3ef4f85861aa8"
        },
        "crypto-rust-standard-aes256" => {
            "b2c8105fe28c91c1955a20f9b3818fa680dfcce89b0dfe9841288e12040189f3"
        },
        "crypto-rust-agile-aes256-sha512" => {
            "7b56a5c882e85420b85fb6fc892ad54a2122319f7fa7b51d3a5e5b8e661f2a01"
        },
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn catalog_has_all_twenty_two_unique_fixture_bindings() {
        assert_eq!(FIXTURES.len(), 22);
        let mut ids = FIXTURES
            .iter()
            .map(|fixture| fixture.id)
            .collect::<Vec<_>>();
        ids.sort_unstable();
        ids.dedup();
        assert_eq!(ids.len(), FIXTURES.len());
        assert!(
            FIXTURES
                .iter()
                .all(|fixture| expected_hash(fixture.id).is_some())
        );
    }

    #[test]
    fn input_bound_is_checked_at_the_admission_boundary() {
        assert_eq!(
            admit_input_length(MAX_INPUT_BYTES).unwrap(),
            MAX_INPUT_BYTES as usize
        );
        assert!(admit_input_length(MAX_INPUT_BYTES + 1).is_err());
    }

    #[test]
    fn all_rows_are_guarded_against_execution_and_resolution() {
        assert!(FIXTURES.iter().all(|fixture| {
            !fixture.categories.is_empty()
                && fixture.oracle.expected_behavior() != "execute"
                && fixture.oracle.expected_behavior() != "resolve"
        }));
    }

    #[test]
    fn declared_zip_member_size_refuses_before_payload_read() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("payload.bin", b"small").unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let central = bytes
            .windows(4)
            .position(|window| window == b"PK\x01\x02")
            .expect("central directory");
        let oversized = (MAX_INPUT_BYTES + 1) as u32;
        bytes[central + 24..central + 28].copy_from_slice(&oversized.to_le_bytes());
        assert!(ArchiveReader::new_with_limits(&bytes, archive_limits()).is_err());
    }

    #[test]
    fn total_zip_entry_admission_includes_directory_records() {
        let mut writer = StreamingArchiveWriter::new();
        for index in 0..=MAX_MEMBERS {
            let name = format!("directory-{index}/");
            writer.write_stored(&name, &[]).unwrap();
        }
        let bytes = writer.finish_to_bytes().unwrap();
        assert!(bounded_archive("directory-entry-admission", &bytes).is_err());
    }

    #[test]
    fn zero_block_crypto_oracle_matches_strict_and_compatibility_paths() {
        let fixture = FIXTURES
            .iter()
            .find(|fixture| fixture.id == "crypto-msoffcrypto-agile-zero-block-graph")
            .unwrap();
        let bytes = load_fixture(fixture).unwrap();
        assert!(spaces::inspect_bytes(&bytes).is_err());
        validate_crypto(fixture, &bytes).unwrap();
    }

    #[test]
    fn malformed_rows_refuse_before_preservation_output() {
        for id in [
            "poi-opc-multiple-core-properties",
            "poi-opc-derived-part-name",
        ] {
            let fixture = FIXTURES.iter().find(|fixture| fixture.id == id).unwrap();
            let bytes = load_fixture(fixture).unwrap();
            for operation in [Operation::Read, Operation::Preserve] {
                let result = execute(fixture, &bytes, operation).unwrap();
                assert_eq!(result.output_bytes, 0, "{id} {operation:?}");
            }
        }
    }
}
