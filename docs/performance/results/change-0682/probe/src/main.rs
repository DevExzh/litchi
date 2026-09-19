//! Change 0682 DOCX paragraph-index before/after probe.
//!
//! The process opens one package before the measured loop. `fresh-*` modes
//! create a new document view for every iteration; `same-*` modes build one
//! view and repeat the query on it; `first-*` measures only the first query on
//! that one view. The package and corpus hashes, source version, elapsed time,
//! allocation delta, and a semantic result digest are emitted as one JSON
//! object per process. No production instrumentation is required.
//!
//! Usage:
//!
//! ```text
//! probe0682 <eager|source|managed> <mode> <repetitions> <generated:200|generated:10000|path>
//! ```

use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::io::Cursor;
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::Path;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource};
use litchi_docx::{Package, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceCacheLimits};
use serde::Serialize;
use sha2::{Digest, Sha256};

static ALLOCATIONS: AtomicU64 = AtomicU64::new(0);
static ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);

struct CountingAllocator;

// SAFETY: every operation delegates to the system allocator. Relaxed counters
// are observational only and never affect returned pointers.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MANAGED_MEMORY: u64 = 512 * 1024 * 1024;
const MANAGED_INPUT: u64 = 2 * 1024 * 1024 * 1024;
const MANAGED_OUTPUT: u64 = 2 * 1024 * 1024 * 1024;
const MANAGED_OBJECTS: u64 = 100_000_000;
const MANAGED_DEPTH: u64 = 4096;
const MANAGED_WORK: u64 = 1 << 50;

type BoxError = Box<dyn std::error::Error + Send + Sync>;

#[derive(Clone, Copy)]
enum Route {
    Eager,
    Source,
    Managed,
}

impl Route {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "eager" => Ok(Self::Eager),
            "source" => Ok(Self::Source),
            "managed" => Ok(Self::Managed),
            other => Err(format!("unknown route {other}").into()),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Eager => "eager",
            Self::Source => "source",
            Self::Managed => "managed",
        }
    }
}

#[derive(Clone, Copy)]
enum Mode {
    FreshDocument,
    FreshCount,
    SameCount,
    FirstCount,
    FreshParagraph,
    SameParagraph,
    FreshParagraphs,
    SameParagraphs,
    FreshText,
    SameText,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "fresh-document" => Ok(Self::FreshDocument),
            "fresh-count" => Ok(Self::FreshCount),
            "same-count" => Ok(Self::SameCount),
            "first-count" => Ok(Self::FirstCount),
            "fresh-paragraph" => Ok(Self::FreshParagraph),
            "same-paragraph" => Ok(Self::SameParagraph),
            "fresh-paragraphs" => Ok(Self::FreshParagraphs),
            "same-paragraphs" => Ok(Self::SameParagraphs),
            "fresh-text" => Ok(Self::FreshText),
            "same-text" => Ok(Self::SameText),
            other => Err(format!("unknown mode {other}").into()),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::FreshDocument => "fresh-document",
            Self::FreshCount => "fresh-count",
            Self::SameCount => "same-count",
            Self::FirstCount => "first-count",
            Self::FreshParagraph => "fresh-paragraph",
            Self::SameParagraph => "same-paragraph",
            Self::FreshParagraphs => "fresh-paragraphs",
            Self::SameParagraphs => "same-paragraphs",
            Self::FreshText => "fresh-text",
            Self::SameText => "same-text",
        }
    }

    fn measured_iterations(self, repetitions: usize) -> usize {
        match self {
            Self::FirstCount => 1,
            _ => repetitions,
        }
    }
}

struct Corpus {
    label: String,
    bytes: Vec<u8>,
    expected_paragraphs: Option<usize>,
    xml_bytes: Option<usize>,
    input_sha256: String,
}

#[derive(Serialize)]
struct SourceDiagnostics {
    cold_loads: u64,
    hits: u64,
    successful_loads: u64,
    retained_entries: usize,
    retained_bytes: usize,
    budget_managed: bool,
    budget_memory_used: u64,
    budget_cache_reserved_bytes: u64,
}

#[derive(Serialize)]
struct Sample {
    schema: &'static str,
    route: &'static str,
    mode: &'static str,
    corpus: String,
    input_sha256: String,
    archive_bytes: usize,
    expected_paragraphs: usize,
    xml_bytes: Option<usize>,
    repetitions: usize,
    measured_iterations: usize,
    elapsed_ns: u64,
    allocations: u64,
    allocated_bytes: u64,
    result_digest: u64,
    source_version_id: Option<u64>,
    source_revision: Option<u64>,
    source_diagnostics: Option<SourceDiagnostics>,
}

fn main() -> Result<(), BoxError> {
    let mut args = std::env::args().skip(1);
    let route = Route::parse(&args.next().ok_or("missing route")?)?;
    let mode = Mode::parse(&args.next().ok_or("missing mode")?)?;
    let repetitions: usize = args
        .next()
        .ok_or("missing repetitions")?
        .parse()
        .map_err(|error| format!("invalid repetitions: {error}"))?;
    if repetitions == 0 {
        return Err("repetitions must be positive".into());
    }
    let corpus = load_corpus(&args.next().ok_or("missing corpus")?)?;
    if args.next().is_some() {
        return Err("unexpected trailing argument".into());
    }

    let expected_hint = corpus.expected_paragraphs.unwrap_or(0);
    let sample = match route {
        Route::Eager => run_eager(&corpus, expected_hint, mode, repetitions)?,
        Route::Source => run_source(&corpus, expected_hint, mode, repetitions)?,
        Route::Managed => run_managed(&corpus, expected_hint, mode, repetitions)?,
    };
    println!("{}", serde_json::to_string(&sample)?);
    Ok(())
}

fn load_corpus(spec: &str) -> Result<Corpus, BoxError> {
    if let Some(paragraphs) = spec.strip_prefix("generated:") {
        let paragraphs: usize = paragraphs
            .parse()
            .map_err(|error| format!("invalid generated paragraph count: {error}"))?;
        let (bytes, xml_bytes) = package_bytes(paragraphs);
        return Ok(Corpus {
            label: format!("generated-{paragraphs}"),
            input_sha256: sha256_hex(&bytes),
            bytes,
            expected_paragraphs: Some(paragraphs),
            xml_bytes: Some(xml_bytes),
        });
    }

    let path = Path::new(spec);
    let bytes = std::fs::read(path)?;
    Ok(Corpus {
        label: path.to_string_lossy().into_owned(),
        input_sha256: sha256_hex(&bytes),
        bytes,
        expected_paragraphs: None,
        xml_bytes: None,
    })
}

fn package_bytes(paragraphs: usize) -> (Vec<u8>, usize) {
    let mut xml = format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body>"#);
    for index in 0..paragraphs {
        xml.push_str(&format!(
            "<w:p><w:r><w:t>litchi-0682-{index:05}</w:t></w:r></w:p>"
        ));
    }
    xml.push_str("</w:body></w:document>");
    let xml_bytes = xml.len();
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").expect("document URI"),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        xml.into_bytes(),
    )));
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    (
        PackageWriter::to_bytes(&package).expect("package bytes"),
        xml_bytes,
    )
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn counters() -> (u64, u64) {
    (
        ALLOCATIONS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
    )
}

fn mix_u64(digest: &mut u64, value: u64) {
    *digest ^= value.wrapping_add(0x9e37_79b9_7f4a_7c15);
    *digest = digest.rotate_left(27).wrapping_mul(0x94d0_49bb_1331_11eb);
}

fn mix_bytes(digest: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        mix_u64(digest, u64::from(*byte));
    }
}

fn measure<F>(
    corpus: &Corpus,
    expected: usize,
    route: Route,
    mode: Mode,
    repetitions: usize,
    source_version: (Option<u64>, Option<u64>),
    source_diagnostics: Option<SourceDiagnostics>,
    mut operation: F,
) -> Result<Sample, BoxError>
where
    F: FnMut(&mut u64) -> Result<(), BoxError>,
{
    let before = counters();
    let started = Instant::now();
    let mut result_digest = 0_u64;
    operation(&mut result_digest)?;
    let elapsed_ns = started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
    let after = counters();
    Ok(Sample {
        schema: "change-0682-sample-v1",
        route: route.as_str(),
        mode: mode.as_str(),
        corpus: corpus.label.clone(),
        input_sha256: corpus.input_sha256.clone(),
        archive_bytes: corpus.bytes.len(),
        expected_paragraphs: expected,
        xml_bytes: corpus.xml_bytes,
        repetitions,
        measured_iterations: mode.measured_iterations(repetitions),
        elapsed_ns,
        allocations: after.0.saturating_sub(before.0),
        allocated_bytes: after.1.saturating_sub(before.1),
        result_digest,
        source_version_id: source_version.0,
        source_revision: source_version.1,
        source_diagnostics,
    })
}

fn expected_eager(bytes: &[u8]) -> Result<usize, BoxError> {
    let package = Package::from_reader(Cursor::new(bytes.to_vec()))?;
    Ok(package.document()?.paragraph_count()?)
}

fn expected_source(bytes: &[u8]) -> Result<usize, BoxError> {
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))?;
    Ok(package.document()?.paragraph_count()?)
}

fn managed_context() -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-paragraph-index-probe",
        Limits::new(
            MANAGED_MEMORY,
            MANAGED_INPUT,
            MANAGED_OUTPUT,
            MANAGED_OBJECTS,
            MANAGED_DEPTH,
            MANAGED_WORK,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(MANAGED_MEMORY).expect("non-zero memory limit"),
        0,
    )
    .expect("valid managed execution limits");
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, limits),
    )
}

fn managed_package(
    bytes: &[u8],
) -> Result<(Budget, CancellationSource, source_backed::Package), BoxError> {
    let (budget, cancellation_source, context) = managed_context();
    let package = source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
        Arc::new(OwnedSource::new(bytes.to_vec())),
        ReadLimits::default(),
        SourceCacheLimits::new(1 << 20, 8)?,
        context,
    )?;
    Ok((budget, cancellation_source, package))
}

fn source_diagnostics(package: &source_backed::Package) -> SourceDiagnostics {
    let diagnostics = package.cache_diagnostics();
    SourceDiagnostics {
        cold_loads: diagnostics.cold_loads,
        hits: diagnostics.hits,
        successful_loads: diagnostics.successful_loads,
        retained_entries: diagnostics.retained_entries,
        retained_bytes: diagnostics.retained_bytes,
        budget_managed: diagnostics.budget_managed,
        budget_memory_used: diagnostics.budget_memory_used,
        budget_cache_reserved_bytes: diagnostics.budget_cache_reserved_bytes,
    }
}

fn run_eager(
    corpus: &Corpus,
    expected_hint: usize,
    mode: Mode,
    repetitions: usize,
) -> Result<Sample, BoxError> {
    let expected = if expected_hint == 0 {
        expected_eager(&corpus.bytes)?
    } else {
        expected_hint
    };
    let package = Package::from_reader(Cursor::new(corpus.bytes.clone()))?;
    let source_version = (None, None);
    match mode {
        Mode::FreshDocument => measure(
            corpus,
            expected,
            Route::Eager,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    black_box(document);
                    mix_u64(digest, 1);
                }
                Ok(())
            },
        ),
        Mode::FreshCount => measure(
            corpus,
            expected,
            Route::Eager,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let count = document.paragraph_count()?;
                    assert_eq!(count, expected);
                    mix_u64(digest, count as u64);
                }
                Ok(())
            },
        ),
        Mode::SameCount | Mode::FirstCount => {
            let document = package.document()?;
            let iterations = mode.measured_iterations(repetitions);
            measure(
                corpus,
                expected,
                Route::Eager,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..iterations {
                        let count = document.paragraph_count()?;
                        assert_eq!(count, expected);
                        mix_u64(digest, count as u64);
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshParagraph => measure(
            corpus,
            expected,
            Route::Eager,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                let index = expected.saturating_sub(1);
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let paragraph = document.paragraph(index)?;
                    assert!(paragraph.is_some());
                    mix_u64(digest, u64::from(paragraph.is_some()));
                }
                Ok(())
            },
        ),
        Mode::SameParagraph => {
            let document = package.document()?;
            let index = expected.saturating_sub(1);
            measure(
                corpus,
                expected,
                Route::Eager,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..repetitions {
                        let paragraph = document.paragraph(index)?;
                        assert!(paragraph.is_some());
                        mix_u64(digest, u64::from(paragraph.is_some()));
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshParagraphs => measure(
            corpus,
            expected,
            Route::Eager,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let paragraphs = document.paragraphs()?;
                    assert_eq!(paragraphs.len(), expected);
                    mix_u64(digest, paragraphs.len() as u64);
                }
                Ok(())
            },
        ),
        Mode::SameParagraphs => {
            let document = package.document()?;
            measure(
                corpus,
                expected,
                Route::Eager,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..repetitions {
                        let paragraphs = document.paragraphs()?;
                        assert_eq!(paragraphs.len(), expected);
                        mix_u64(digest, paragraphs.len() as u64);
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshText | Mode::SameText => Err("text query is source-backed only".into()),
    }
}

fn run_source(
    corpus: &Corpus,
    expected_hint: usize,
    mode: Mode,
    repetitions: usize,
) -> Result<Sample, BoxError> {
    let expected = if expected_hint == 0 {
        expected_source(&corpus.bytes)?
    } else {
        expected_hint
    };
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        corpus.bytes.clone(),
    )))?;
    let version = package.source_version()?;
    let source_version = (Some(version.id()), Some(version.revision()));
    run_source_with_package(corpus, expected, mode, repetitions, package, source_version)
}

fn run_source_with_package(
    corpus: &Corpus,
    expected: usize,
    mode: Mode,
    repetitions: usize,
    package: source_backed::Package,
    source_version: (Option<u64>, Option<u64>),
) -> Result<Sample, BoxError> {
    let package = package;
    let sample = match mode {
        Mode::FreshDocument => measure(
            corpus,
            expected,
            Route::Source,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    black_box(document);
                    mix_u64(digest, 1);
                }
                Ok(())
            },
        ),
        Mode::FreshCount => measure(
            corpus,
            expected,
            Route::Source,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let count = document.paragraph_count()?;
                    assert_eq!(count, expected);
                    mix_u64(digest, count as u64);
                }
                Ok(())
            },
        ),
        Mode::SameCount | Mode::FirstCount => {
            let document = package.document()?;
            let iterations = mode.measured_iterations(repetitions);
            measure(
                corpus,
                expected,
                Route::Source,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..iterations {
                        let count = document.paragraph_count()?;
                        assert_eq!(count, expected);
                        mix_u64(digest, count as u64);
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshParagraph => measure(
            corpus,
            expected,
            Route::Source,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                let index = expected.saturating_sub(1);
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let paragraph = document.paragraph(index)?;
                    assert!(paragraph.is_some());
                    mix_u64(digest, u64::from(paragraph.is_some()));
                }
                Ok(())
            },
        ),
        Mode::SameParagraph => {
            let document = package.document()?;
            let index = expected.saturating_sub(1);
            measure(
                corpus,
                expected,
                Route::Source,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..repetitions {
                        let paragraph = document.paragraph(index)?;
                        assert!(paragraph.is_some());
                        mix_u64(digest, u64::from(paragraph.is_some()));
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshParagraphs => measure(
            corpus,
            expected,
            Route::Source,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let paragraphs = document.paragraphs()?;
                    assert_eq!(paragraphs.len(), expected);
                    mix_u64(digest, paragraphs.len() as u64);
                }
                Ok(())
            },
        ),
        Mode::SameParagraphs => {
            let document = package.document()?;
            measure(
                corpus,
                expected,
                Route::Source,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..repetitions {
                        let paragraphs = document.paragraphs()?;
                        assert_eq!(paragraphs.len(), expected);
                        mix_u64(digest, paragraphs.len() as u64);
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshText => measure(
            corpus,
            expected,
            Route::Source,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                let index = expected.saturating_sub(1);
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let text = document.paragraph_text(index)?.ok_or("missing paragraph")?;
                    mix_bytes(digest, text.as_bytes());
                }
                Ok(())
            },
        ),
        Mode::SameText => {
            let document = package.document()?;
            let index = expected.saturating_sub(1);
            measure(
                corpus,
                expected,
                Route::Source,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..repetitions {
                        let text = document.paragraph_text(index)?.ok_or("missing paragraph")?;
                        mix_bytes(digest, text.as_bytes());
                    }
                    Ok(())
                },
            )
        }
    }?;
    let mut sample = sample;
    sample.source_diagnostics = Some(source_diagnostics(&package));
    Ok(sample)
}

fn run_managed(
    corpus: &Corpus,
    expected_hint: usize,
    mode: Mode,
    repetitions: usize,
) -> Result<Sample, BoxError> {
    let expected = if expected_hint == 0 {
        let (_budget, _cancellation, package) = managed_package(&corpus.bytes)?;
        package.document()?.paragraph_count()?
    } else {
        expected_hint
    };
    let (_budget, _cancellation, package) = managed_package(&corpus.bytes)?;
    let version = package.source_version()?;
    let source_version = (Some(version.id()), Some(version.revision()));
    let package = package;
    let sample = match mode {
        Mode::FreshDocument => measure(
            corpus,
            expected,
            Route::Managed,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    black_box(document);
                    mix_u64(digest, 1);
                }
                Ok(())
            },
        ),
        Mode::FreshCount => measure(
            corpus,
            expected,
            Route::Managed,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let count = document.paragraph_count()?;
                    assert_eq!(count, expected);
                    mix_u64(digest, count as u64);
                }
                Ok(())
            },
        ),
        Mode::SameCount | Mode::FirstCount => {
            let document = package.document()?;
            let iterations = mode.measured_iterations(repetitions);
            measure(
                corpus,
                expected,
                Route::Managed,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..iterations {
                        let count = document.paragraph_count()?;
                        assert_eq!(count, expected);
                        mix_u64(digest, count as u64);
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshText => measure(
            corpus,
            expected,
            Route::Managed,
            mode,
            repetitions,
            source_version,
            None,
            |digest| {
                let index = expected.saturating_sub(1);
                for _ in 0..repetitions {
                    let document = package.document()?;
                    let text = document.paragraph_text(index)?.ok_or("missing paragraph")?;
                    mix_bytes(digest, text.as_bytes());
                }
                Ok(())
            },
        ),
        Mode::SameText => {
            let document = package.document()?;
            let index = expected.saturating_sub(1);
            measure(
                corpus,
                expected,
                Route::Managed,
                mode,
                repetitions,
                source_version,
                None,
                |digest| {
                    for _ in 0..iterations_for_same_text(mode, repetitions) {
                        let text = document.paragraph_text(index)?.ok_or("missing paragraph")?;
                        mix_bytes(digest, text.as_bytes());
                    }
                    Ok(())
                },
            )
        }
        Mode::FreshParagraph | Mode::SameParagraph | Mode::FreshParagraphs | Mode::SameParagraphs => {
            Err("managed selective paragraph views are intentionally refused".into())
        }
    }?;
    let mut sample = sample;
    sample.source_diagnostics = Some(source_diagnostics(&package));
    Ok(sample)
}

fn iterations_for_same_text(_mode: Mode, repetitions: usize) -> usize {
    repetitions
}
