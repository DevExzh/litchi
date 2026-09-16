//! Opt-in facade selectors for the OLE2 document and presentation routes
//! (change 0638).
//!
//! Change [0587](../../../docs/performance/0587-remaining-opportunity-survey.md)
//! recorded, as the first half of its evidence gap 5, that **no selector opens
//! a `.doc` or a `.ppt` through the `litchi` facade**: "no harness selector
//! opens `.doc` or `.ppt` through the facade, so CORE-1 cannot be A/B-measured
//! today". Every DOC and PPT figure the programme holds — 0584's, 0587's,
//! 0589's, 0596's and 0609's — comes from a throwaway driver built outside the
//! workspace. Change [0630](../../../docs/performance/0630-queue-refresh-after-the-first-wave.md)
//! confirmed the gap is still open after the first wave, and its refreshed
//! queue row 10 depends on it: *"until then the facade's `.doc` route stays
//! eager by measurement"*.
//!
//! This module adds six opt-in selectors, none of them in `Case::DEFAULT`:
//!
//! ```text
//! doc_facade_file_open            ppt_facade_file_open
//! doc_facade_file_full_text       ppt_facade_file_full_text
//! doc_facade_file_one_paragraph   ppt_facade_file_one_slide_text
//! ```
//!
//! They call the documented facade entry points and nothing else:
//! [`litchi::Document::open`], [`litchi::Document::text`],
//! [`litchi::Document::paragraph_text`], [`litchi::Presentation::open`],
//! [`litchi::Presentation::text`] and [`litchi::Presentation::slide`] followed
//! by [`litchi::presentation::Slide::text`]. The facade takes a **path**, not
//! bytes, so unlike every other selector in this harness these run against a
//! real file on the filesystem: the route's defining property, established by
//! change [0609](../../../docs/performance/0609-facade-doc-source-route-design.md),
//! is that `Document::open` reads the artifact once through
//! `FileSource`/`detect_document_source_path_with_limits` and then hands the
//! bytes to the eager reader.
//!
//! The corpus is the caller-named file supplied with `--ole2-file PATH`, the
//! flag change [0627](../../../docs/performance/0627-ole2-range-source-selectors.md)
//! introduced, classified by its CFB stream inventory rather than by its
//! extension. Like 0627's and 0601's real-file families, the file is bounded,
//! kept out of the default matrix, and carried in the corpus identity by path,
//! size and SHA-256 together with its CFB stream count, sector size and target
//! stream.
//!
//! Every case is a complete fresh-open lifecycle, for the same reason 0627's
//! are: the facade owns no cache across calls, and a caller of
//! `Document::open` pays for reaching the text, not for a marginal query on an
//! already-open owner.
//!
//! These selectors take no timing, allocation, physical-I/O, cold-cache or
//! speedup claim. They are a descriptive baseline: the first measurement of
//! either facade route by a registered selector.

use std::{
    error::Error,
    path::{Path, PathBuf},
    time::{Duration, Instant},
};

use serde::Serialize;

use crate::{
    Case, CaseResult, Corpus, SourceSummary, boxed_source, iteration_count,
    ole2_range_source::{cfb_inventory, provenance_of, read_bounded},
    producer_shape::RealFileProvenance,
    record_elapsed, sha256_hex, statistics,
};

/// Generator identity of the DOC facade family. Nothing in it is generated, so
/// the identity says so.
pub(crate) const DOC_FACADE_GENERATOR: &str = "litchi-doc-facade-real-file-v1";
/// Generator identity of the PPT facade family.
pub(crate) const PPT_FACADE_GENERATOR: &str = "litchi-ppt-facade-real-file-v1";

/// Which facade route a corpus and a scenario belong to.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Route {
    /// `litchi::Document` over a `.doc`.
    Document,
    /// `litchi::Presentation` over a `.ppt`.
    Presentation,
}

impl Route {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Document => "litchi::Document",
            Self::Presentation => "litchi::Presentation",
        }
    }

    const fn format(self) -> &'static str {
        match self {
            Self::Document => "DOC",
            Self::Presentation => "PPT",
        }
    }

    const fn package_format(self) -> &'static str {
        match self {
            Self::Document => "DOC/CFB/OLE2",
            Self::Presentation => "PPT/CFB/OLE2",
        }
    }

    const fn generator(self) -> &'static str {
        match self {
            Self::Document => DOC_FACADE_GENERATOR,
            Self::Presentation => PPT_FACADE_GENERATOR,
        }
    }

    const fn corpus_name(self) -> &'static str {
        match self {
            Self::Document => "doc-facade-real-file",
            Self::Presentation => "ppt-facade-real-file",
        }
    }
}

/// One measured phase. Every variant includes a fresh facade open.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Scenario {
    /// `Document::open` alone. The projection is derived after the clock
    /// stops, so the timed region holds nothing but the documented open.
    DocOpen,
    /// Open, then `Document::text()`.
    DocFullText,
    /// Open, then `Document::paragraph_text(index)` at the derived target.
    DocOneParagraph,
    /// `Presentation::open` alone, projection derived after the clock stops.
    PptOpen,
    /// Open, then `Presentation::text()`.
    PptFullText,
    /// Open, then `Presentation::slide(index)` and `Slide::text()` at the
    /// derived target. The facade exposes no per-shape accessor on either
    /// legacy route, so one slide's text is the nearest documented analogue of
    /// change 0627's `open+one-shape-text`.
    PptOneSlideText,
}

impl Scenario {
    const fn as_str(self) -> &'static str {
        match self {
            Self::DocOpen | Self::PptOpen => "facade-open",
            Self::DocFullText | Self::PptFullText => "facade-open+full-text",
            Self::DocOneParagraph => "facade-open+one-paragraph",
            Self::PptOneSlideText => "facade-open+one-slide-text",
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::DocOpen => {
                "litchi::Document::open only; the paragraph count that names the sample is read \
                 after the clock stops"
            },
            Self::DocFullText => "fresh litchi::Document::open plus Document::text",
            Self::DocOneParagraph => {
                "fresh litchi::Document::open plus Document::paragraph_text at the derived target"
            },
            Self::PptOpen => {
                "litchi::Presentation::open only; the slide count that names the sample is read \
                 after the clock stops"
            },
            Self::PptFullText => "fresh litchi::Presentation::open plus Presentation::text",
            Self::PptOneSlideText => {
                "fresh litchi::Presentation::open plus Presentation::slide and Slide::text at the \
                 derived target"
            },
        }
    }

    pub(crate) const fn route(self) -> Route {
        match self {
            Self::DocOpen | Self::DocFullText | Self::DocOneParagraph => Route::Document,
            Self::PptOpen | Self::PptFullText | Self::PptOneSlideText => Route::Presentation,
        }
    }
}

/// Everything a facade corpus states about itself beyond the ordinary
/// [`crate::CorpusManifest`]: where the bytes came from, what the CFB holds,
/// and the scenario targets derived from it.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct FacadeEvidence {
    pub(crate) generator: &'static str,
    pub(crate) format: &'static str,
    pub(crate) facade_type: &'static str,
    pub(crate) real_file: RealFileProvenance,
    pub(crate) cfb_stream_count: usize,
    pub(crate) cfb_sector_size: usize,
    pub(crate) target_stream: String,
    pub(crate) target_stream_bytes: usize,
    pub(crate) target_stream_sha256: String,
    /// DOC: the paragraph count the facade reports, and the paragraph the
    /// one-paragraph scenario selects.
    pub(crate) paragraph_count: Option<usize>,
    pub(crate) selected_paragraph: Option<usize>,
    /// PPT: the slide count the facade reports, and the slide the
    /// one-slide-text scenario selects.
    pub(crate) slide_count: Option<usize>,
    pub(crate) selected_slide: Option<usize>,
    /// Frozen oracle projection per scenario, derived once from the file.
    pub(crate) scenario_oracles: Vec<ScenarioOracle>,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct ScenarioOracle {
    pub(crate) scenario: &'static str,
    pub(crate) observation: String,
}

/// A facade corpus: an ordinary [`Corpus`] plus its evidence block and the
/// path the facade is asked to open.
#[derive(Debug)]
pub(crate) struct FacadeCorpus {
    pub(crate) corpus: Corpus,
    pub(crate) evidence: FacadeEvidence,
    route: Route,
    path: PathBuf,
    selected_paragraph: usize,
    selected_slide: usize,
}

/// Per-case evidence written into `source.facade_ole2`.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct FacadeOle2Summary {
    pub(crate) format: &'static str,
    pub(crate) facade_type: &'static str,
    pub(crate) scenario: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) route_scope: &'static str,
    /// The corpus identity and the derived scenario targets, restated per case
    /// so one result is self-describing.
    pub(crate) corpus: FacadeEvidence,
    /// The frozen oracle, derived once from the file before any sample runs.
    pub(crate) observation: String,
    /// SHA-256 of each retained sample's own projection, and whether all of
    /// them reproduced the oracle.
    pub(crate) observation_sha256: Vec<String>,
    pub(crate) observations_identical: bool,
}

const ROUTE_SCOPE: &str = "the documented path-taking facade entry point; the facade owns no cache \
                           across calls, so every sample is a complete fresh open";

// ---------------------------------------------------------------------------
// Corpus construction
// ---------------------------------------------------------------------------

fn text_outcome(text: &Result<String, impl std::fmt::Display>) -> String {
    match text {
        Ok(text) => format!("text:{}:{}", text.len(), sha256_hex(text.as_bytes())),
        Err(error) => format!("refused:{error}"),
    }
}

fn optional_text_outcome(
    text: &Result<Option<String>, impl std::fmt::Display>,
    absent: &str,
) -> String {
    match text {
        Ok(Some(text)) => format!("text:{}:{}", text.len(), sha256_hex(text.as_bytes())),
        Ok(None) => absent.to_owned(),
        Err(error) => format!("refused:{error}"),
    }
}

fn manifest_of(
    route: Route,
    archive: &[u8],
    streams: &[Vec<String>],
    sector_size: usize,
    target_entry: &str,
    target_payload: &[u8],
    entry_count: usize,
) -> crate::CorpusManifest {
    crate::CorpusManifest {
        name: route.corpus_name().to_owned(),
        generator: route.generator(),
        package_format: route.package_format(),
        shape: "real-file",
        payload_kind: "real-producer-ole2",
        compression: "none",
        entry_count,
        archive_member_count: streams.len(),
        entry_bytes: sector_size,
        uncompressed_payload_bytes: target_payload.len(),
        archive_bytes: archive.len(),
        archive_sha256: sha256_hex(archive),
        target_entry: target_entry.to_owned(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256: sha256_hex(target_payload),
        rtf_variant: None,
        xlsx: None,
    }
}

/// Builds the DOC facade corpus from a caller-named fixture.
///
/// Nothing about the file is assumed. The selected paragraph is the median
/// position among the paragraphs the facade itself reports with non-empty
/// text, so the target is derived from the bytes and is stable for a given
/// fixture. A facade refusal is frozen as the scenario's outcome rather than
/// dropped, exactly as change 0627 freezes the XLS full-text refusal: a caller
/// pays for what the route reads before refusing.
pub(crate) fn build_doc_corpus(path: &Path) -> Result<FacadeCorpus, Box<dyn Error>> {
    let archive = read_bounded(path)?;
    let provenance = provenance_of(path, &archive)?;
    let (streams, sector_size, target_path, target_payload) =
        cfb_inventory(&archive, &[&["WordDocument"]], "WordDocument")?;
    let target_entry = target_path.join("/");

    let opened = litchi::Document::open(path);
    let (paragraph_count, selected_paragraph, open_outcome, full_text, one_paragraph) =
        match &opened {
            Ok(document) => {
                let count = document.paragraph_count()?;
                let mut selected = None;
                for index in 0..count {
                    if document
                        .paragraph_text(index)?
                        .is_some_and(|text| !text.trim().is_empty())
                    {
                        selected = Some(index);
                        break;
                    }
                }
                // Prefer the median non-empty paragraph when the document has
                // several, so the target is not systematically the title.
                let mut populated = Vec::new();
                for index in 0..count {
                    if document
                        .paragraph_text(index)?
                        .is_some_and(|text| !text.trim().is_empty())
                    {
                        populated.push(index);
                    }
                }
                let selected = if populated.is_empty() {
                    selected.unwrap_or(0)
                } else {
                    populated[populated.len() / 2]
                };
                (
                    Some(count),
                    Some(selected),
                    format!("paragraphs:{count}"),
                    text_outcome(&document.text()),
                    optional_text_outcome(&document.paragraph_text(selected), "absent"),
                )
            },
            Err(error) => {
                let refusal = format!("refused:{error}");
                (
                    None,
                    None,
                    refusal.clone(),
                    refusal.clone(),
                    refusal.clone(),
                )
            },
        };
    drop(opened);

    let scenario_oracles = vec![
        ScenarioOracle {
            scenario: Scenario::DocOpen.as_str(),
            observation: open_outcome,
        },
        ScenarioOracle {
            scenario: Scenario::DocFullText.as_str(),
            observation: full_text,
        },
        ScenarioOracle {
            scenario: Scenario::DocOneParagraph.as_str(),
            observation: one_paragraph,
        },
    ];

    let evidence = FacadeEvidence {
        generator: DOC_FACADE_GENERATOR,
        format: Route::Document.format(),
        facade_type: Route::Document.as_str(),
        real_file: provenance,
        cfb_stream_count: streams.len(),
        cfb_sector_size: sector_size,
        target_stream: target_entry.clone(),
        target_stream_bytes: target_payload.len(),
        target_stream_sha256: sha256_hex(&target_payload),
        paragraph_count,
        selected_paragraph,
        slide_count: None,
        selected_slide: None,
        scenario_oracles,
    };
    let manifest = manifest_of(
        Route::Document,
        &archive,
        &streams,
        sector_size,
        &target_entry,
        &target_payload,
        paragraph_count.unwrap_or(0),
    );
    Ok(FacadeCorpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: target_entry,
            target_payload,
            xlsx: None,
        },
        evidence,
        route: Route::Document,
        path: path.to_path_buf(),
        selected_paragraph: selected_paragraph.unwrap_or(0),
        selected_slide: 0,
    })
}

/// Builds the PPT facade corpus from a caller-named fixture.
///
/// The selected slide is the first position, in the facade's own order, whose
/// text the facade reports as non-empty. A refusal is frozen rather than
/// dropped, for the same reason as the DOC route.
pub(crate) fn build_ppt_corpus(path: &Path) -> Result<FacadeCorpus, Box<dyn Error>> {
    let archive = read_bounded(path)?;
    let provenance = provenance_of(path, &archive)?;
    let (streams, sector_size, target_path, target_payload) = cfb_inventory(
        &archive,
        &[
            &["PowerPoint Document"],
            &["PP97_DUALSTORAGE", "PowerPoint Document"],
        ],
        "PowerPoint Document",
    )?;
    let target_entry = target_path.join("/");

    let opened = litchi::Presentation::open(path);
    let (slide_count, selected_slide, open_outcome, full_text, one_slide) = match &opened {
        Ok(presentation) => {
            let count = presentation.slide_count()?;
            let mut selected = 0;
            for index in 0..count {
                let populated = presentation
                    .slide(index)?
                    .map(|slide| slide.text())
                    .transpose()?
                    .is_some_and(|text| !text.trim().is_empty());
                if populated {
                    selected = index;
                    break;
                }
            }
            let one_slide = match presentation.slide(selected)? {
                Some(slide) => text_outcome(&slide.text()),
                None => "absent".to_owned(),
            };
            (
                Some(count),
                Some(selected),
                format!("slides:{count}"),
                text_outcome(&presentation.text()),
                one_slide,
            )
        },
        Err(error) => {
            let refusal = format!("refused:{error}");
            (
                None,
                None,
                refusal.clone(),
                refusal.clone(),
                refusal.clone(),
            )
        },
    };
    drop(opened);

    let scenario_oracles = vec![
        ScenarioOracle {
            scenario: Scenario::PptOpen.as_str(),
            observation: open_outcome,
        },
        ScenarioOracle {
            scenario: Scenario::PptFullText.as_str(),
            observation: full_text,
        },
        ScenarioOracle {
            scenario: Scenario::PptOneSlideText.as_str(),
            observation: one_slide,
        },
    ];

    let evidence = FacadeEvidence {
        generator: PPT_FACADE_GENERATOR,
        format: Route::Presentation.format(),
        facade_type: Route::Presentation.as_str(),
        real_file: provenance,
        cfb_stream_count: streams.len(),
        cfb_sector_size: sector_size,
        target_stream: target_entry.clone(),
        target_stream_bytes: target_payload.len(),
        target_stream_sha256: sha256_hex(&target_payload),
        paragraph_count: None,
        selected_paragraph: None,
        slide_count,
        selected_slide,
        scenario_oracles,
    };
    let manifest = manifest_of(
        Route::Presentation,
        &archive,
        &streams,
        sector_size,
        &target_entry,
        &target_payload,
        slide_count.unwrap_or(0),
    );
    Ok(FacadeCorpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: target_entry,
            target_payload,
            xlsx: None,
        },
        evidence,
        route: Route::Presentation,
        path: path.to_path_buf(),
        selected_paragraph: 0,
        selected_slide: selected_slide.unwrap_or(0),
    })
}

// ---------------------------------------------------------------------------
// Measurement
// ---------------------------------------------------------------------------

fn oracle_for(corpus: &FacadeCorpus, scenario: Scenario) -> Result<&str, Box<dyn Error>> {
    corpus
        .evidence
        .scenario_oracles
        .iter()
        .find(|oracle| oracle.scenario == scenario.as_str())
        .map(|oracle| oracle.observation.as_str())
        .ok_or_else(|| format!("{} scenario has no frozen oracle", scenario.as_str()).into())
}

pub(crate) fn run_case(
    case: Case,
    scenario: Scenario,
    corpus: &FacadeCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if scenario.route() != corpus.route {
        return Err(format!(
            "{} scenario cannot run on a {} corpus",
            scenario.as_str(),
            corpus.route.format()
        )
        .into());
    }
    let oracle = oracle_for(corpus, scenario)?.to_owned();

    let mut elapsed = Vec::with_capacity(samples);
    let mut observation_sha256 = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (duration, projection) = measure(corpus, scenario)?;
        if iteration >= warmup_iterations {
            observation_sha256.push(sha256_hex(projection.as_bytes()));
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let oracle_sha256 = sha256_hex(oracle.as_bytes());
    let observations_identical = observation_sha256
        .iter()
        .all(|entry| *entry == oracle_sha256);
    if !observations_identical {
        return Err(format!(
            "{} {} produced a sample that differs from the frozen oracle",
            corpus.route.format(),
            scenario.as_str()
        )
        .into());
    }

    let summary = SourceSummary {
        facade_ole2: Some(Box::new(FacadeOle2Summary {
            format: corpus.route.format(),
            facade_type: corpus.route.as_str(),
            scenario: scenario.as_str(),
            timing_scope: scenario.timing_scope(),
            route_scope: ROUTE_SCOPE,
            corpus: corpus.evidence.clone(),
            observation: oracle,
            observation_sha256,
            observations_identical,
        })),
        ..SourceSummary::default()
    };

    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: boxed_source(summary),
        execution: None,
        output_sha256: None,
        operation_metrics: None,
    })
}

/// Runs one measured phase against the caller-named path and returns the
/// scenario's projection with the elapsed time of the timed region.
///
/// A facade refusal is not an error here: the frozen oracle already records it,
/// and the timed region legitimately covers the work the route performed before
/// refusing.
fn measure(
    corpus: &FacadeCorpus,
    scenario: Scenario,
) -> Result<(Duration, String), Box<dyn Error>> {
    match scenario {
        Scenario::DocOpen => {
            let started = Instant::now();
            let document = litchi::Document::open(&corpus.path);
            let duration = started.elapsed();
            let projection = match &document {
                Ok(document) => format!("paragraphs:{}", document.paragraph_count()?),
                Err(error) => format!("refused:{error}"),
            };
            drop(document);
            Ok((duration, projection))
        },
        Scenario::DocFullText => {
            let started = Instant::now();
            let projection = match litchi::Document::open(&corpus.path) {
                Ok(document) => text_outcome(&std::hint::black_box(document.text())),
                Err(error) => format!("refused:{error}"),
            };
            Ok((started.elapsed(), projection))
        },
        Scenario::DocOneParagraph => {
            let index = corpus.selected_paragraph;
            let started = Instant::now();
            let projection = match litchi::Document::open(&corpus.path) {
                Ok(document) => optional_text_outcome(
                    &std::hint::black_box(document.paragraph_text(index)),
                    "absent",
                ),
                Err(error) => format!("refused:{error}"),
            };
            Ok((started.elapsed(), projection))
        },
        Scenario::PptOpen => {
            let started = Instant::now();
            let presentation = litchi::Presentation::open(&corpus.path);
            let duration = started.elapsed();
            let projection = match &presentation {
                Ok(presentation) => format!("slides:{}", presentation.slide_count()?),
                Err(error) => format!("refused:{error}"),
            };
            drop(presentation);
            Ok((duration, projection))
        },
        Scenario::PptFullText => {
            let started = Instant::now();
            let projection = match litchi::Presentation::open(&corpus.path) {
                Ok(presentation) => text_outcome(&std::hint::black_box(presentation.text())),
                Err(error) => format!("refused:{error}"),
            };
            Ok((started.elapsed(), projection))
        },
        Scenario::PptOneSlideText => {
            let index = corpus.selected_slide;
            let started = Instant::now();
            let projection = match litchi::Presentation::open(&corpus.path) {
                Ok(presentation) => match std::hint::black_box(presentation.slide(index)) {
                    Ok(Some(slide)) => text_outcome(&slide.text()),
                    Ok(None) => "absent".to_owned(),
                    Err(error) => format!("refused:{error}"),
                },
                Err(error) => format!("refused:{error}"),
            };
            Ok((started.elapsed(), projection))
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(relative: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative)
    }

    #[test]
    fn routes_and_scenarios_describe_themselves() {
        assert_eq!(Scenario::DocOpen.route(), Route::Document);
        assert_eq!(Scenario::PptOneSlideText.route(), Route::Presentation);
        assert_eq!(
            Scenario::DocOneParagraph.as_str(),
            "facade-open+one-paragraph"
        );
        assert_eq!(Route::Document.as_str(), "litchi::Document");
        assert_eq!(Route::Presentation.as_str(), "litchi::Presentation");
    }

    #[test]
    fn every_facade_case_is_opt_in_and_round_trips() {
        let mut seen = 0;
        for case in [
            Case::DocFacadeFileOpen,
            Case::DocFacadeFileFullText,
            Case::DocFacadeFileOneParagraph,
            Case::PptFacadeFileOpen,
            Case::PptFacadeFileFullText,
            Case::PptFacadeFileOneSlideText,
        ] {
            assert!(case.is_facade_ole2());
            assert!(!Case::DEFAULT.contains(&case));
            assert_eq!(crate::parse_case(case.name()), Some(case));
            seen += 1;
        }
        assert_eq!(seen, 6);
    }

    #[test]
    fn doc_corpus_derives_its_target_from_the_file() {
        let path = fixture("test-data/ole/doc/documentProperties.doc");
        let corpus = build_doc_corpus(&path).unwrap();
        assert_eq!(corpus.route, Route::Document);
        assert_eq!(corpus.corpus.manifest.generator, DOC_FACADE_GENERATOR);
        assert_eq!(corpus.evidence.target_stream, "WordDocument");
        assert!(corpus.evidence.target_stream_bytes > 0);
        assert_eq!(corpus.evidence.scenario_oracles.len(), 3);
        assert_eq!(
            corpus.evidence.real_file.sha256,
            corpus.corpus.manifest.archive_sha256
        );
    }

    #[test]
    fn doc_facade_phases_reproduce_their_frozen_oracles() {
        let path = fixture("test-data/ole/doc/documentProperties.doc");
        let corpus = build_doc_corpus(&path).unwrap();
        for (case, scenario) in [
            (Case::DocFacadeFileOpen, Scenario::DocOpen),
            (Case::DocFacadeFileFullText, Scenario::DocFullText),
            (Case::DocFacadeFileOneParagraph, Scenario::DocOneParagraph),
        ] {
            let result = run_case(case, scenario, &corpus, 0, 2).unwrap();
            let source = result.source.unwrap().facade_ole2.unwrap();
            assert!(source.observations_identical);
            assert_eq!(source.observation_sha256.len(), 2);
            assert_eq!(source.format, "DOC");
        }
    }

    #[test]
    fn ppt_corpus_selects_a_slide_and_freezes_its_outcome() {
        let path = fixture("test-data/ole/ppt/SampleShow.ppt");
        let corpus = build_ppt_corpus(&path).unwrap();
        assert_eq!(corpus.route, Route::Presentation);
        assert_eq!(corpus.corpus.manifest.generator, PPT_FACADE_GENERATOR);
        assert_eq!(corpus.evidence.target_stream, "PowerPoint Document");
        assert_eq!(corpus.evidence.scenario_oracles.len(), 3);
        let result = run_case(
            Case::PptFacadeFileOneSlideText,
            Scenario::PptOneSlideText,
            &corpus,
            0,
            2,
        )
        .unwrap();
        let source = result.source.unwrap().facade_ole2.unwrap();
        assert!(source.observations_identical);
        assert!(
            source.observation.starts_with("text:")
                || source.observation.starts_with("refused:")
                || source.observation == "absent",
            "unexpected PPT facade one-slide-text outcome {}",
            source.observation
        );
    }
}
