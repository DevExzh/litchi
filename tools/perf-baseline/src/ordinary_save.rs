//! Opt-in selectors for the ordinary documented OOXML save path (change 0638).
//!
//! Change [0587](../../../docs/performance/0587-remaining-opportunity-survey.md)
//! recorded, as the second half of its evidence gap 5, that the programme has
//! **no selector for "the ordinary OOXML save"**: every timed publication in
//! this harness goes through a source-backed preservation route
//! (`publish_multi_commit_to_stream`, `write_to_stream`, `PackageWriter`) or an
//! in-memory `to_bytes`. The documented entry points a caller actually reaches
//! for — [`litchi_docx::Package::save`], [`litchi_xlsx::Workbook::save`] and
//! [`litchi_pptx::Package::save`], "Path A" of change
//! [0593](../../../docs/performance/0593-opc-publication-pristine-members.md) —
//! have never had one. 0593 and [0607](../../../docs/performance/0607-pptx-authored-slide-regeneration-design.md)
//! measured them with throwaway probes, and 0593 named the price of fixing
//! that: a selector "changes `tools/perf-baseline`'s checked catalog SHA-256,
//! its selector registry and its coverage-index minimum, and belongs in its own
//! record". This is that record's module; the checked catalog SHA-256 does not
//! in fact move, because none of these selectors is in `Case::DEFAULT`.
//!
//! Twenty-four opt-in selectors: three formats × two corpus origins × four
//! phases.
//!
//! ```text
//! docx_ordinary_save_lifecycle          docx_real_file_ordinary_save_lifecycle
//! docx_ordinary_save_edit               docx_real_file_ordinary_save_edit
//! docx_ordinary_save_atomic_publish     docx_real_file_ordinary_save_atomic_publish
//! docx_ordinary_save_counting_publish   docx_real_file_ordinary_save_counting_publish
//! ... and the same eight names for xlsx_ and pptx_
//! ```
//!
//! **The phases.** `lifecycle` times the whole documented route — open the path,
//! make one semantic edit, save to a path. `edit` times only the semantic edit,
//! with the open outside the clock. `atomic_publish` times only the save, with
//! the open and the edit outside the clock, so the publication is reported
//! separately from the edit exactly as change
//! [0497](../../../docs/performance/results/change-0497/README.md) reports its
//! atomic arm: the interval is one `litchi_opc::atomic::replace_with` and
//! therefore covers sibling creation in the destination's own directory, the
//! publication write, permission preservation, `sync_all` on the temporary,
//! the `rename` that replaces the destination, and the parent-directory sync.
//! Like 0497's atomic arm the destination directory is prepared before the
//! clock starts and the readback, the digest and the cleanup happen after it
//! stops. `counting_publish` is the sink variant: the same edited owner is
//! serialized through the format's documented sequential entry point into a
//! bounded counting sink, and the result reports the byte split 0593's
//! copy-versus-regenerate accounting implies — how many output payload bytes
//! this save deflated, how many it stored, how many are byte-identical to the
//! source archive's own member payload, and the total output.
//!
//! **What the byte split is and is not.** It is derived from the two archives,
//! not from production instrumentation: `litchi-opc`'s `OpcOperationAccounting`
//! deliberately excludes `PartWriter` and the topology publishers, which is
//! precisely the path a documented `save` takes. So
//! `payload_bytes_identical_to_source` is an *upper bound on what a
//! copy-through publisher could have avoided re-deflating*, measured by
//! comparing each output member's stored compressed payload with the
//! same-named source member's, and not an observation that this writer took a
//! copy path. Change 0607 established the matching fact on the authored side:
//! an authored package has no source archive, so every member is regenerated.
//!
//! These selectors take no timing, allocation, physical-I/O, cold-cache or
//! speedup claim. They are a descriptive baseline.

use std::{
    error::Error,
    fs,
    io::{self, Write},
    path::{Path, PathBuf},
    time::Instant,
};

use serde::Serialize;
use soapberry_zip::{CompressionMethod, ZipArchive};

use crate::{
    Case, CaseResult, Corpus, CorpusManifest, SemanticShape, SinkSummary, SourceSummary,
    XlsxCellCrudShape, allocation_metrics, boxed_source, elapsed_ns, iteration_count,
    operation_metrics, producer_shape::RealFileProvenance, record_elapsed, sha256_hex, statistics,
};

/// Largest `--ooxml-file` input this harness will read. A caller-named file is
/// still a bounded resource; the limit matches 0601's `--real-file` and 0627's
/// `--ole2-file` bounds.
const MAX_OOXML_FILE_BYTES: u64 = 32 * 1024 * 1024;

/// Text this family writes when it makes its one semantic edit. It is a fixed
/// marker so that the edit is a change on every corpus and so that the output
/// is a deterministic function of the input.
const EDIT_MARKER: &str = "litchi-perf-0638-ordinary-save";

/// Generator identity of the DOCX, XLSX and PPTX real-file families. Nothing
/// in them is generated, so the identity says so.
pub(crate) const DOCX_REAL_FILE_GENERATOR: &str = "litchi-docx-real-file-v1";
pub(crate) const XLSX_REAL_FILE_SAVE_GENERATOR: &str = "litchi-xlsx-real-file-save-v1";
pub(crate) const PPTX_REAL_FILE_GENERATOR: &str = "litchi-pptx-real-file-v1";

/// Which OOXML format a corpus and a selector belong to.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Format {
    Docx,
    Xlsx,
    Pptx,
}

impl Format {
    pub(crate) const ALL: [Self; 3] = [Self::Docx, Self::Xlsx, Self::Pptx];

    const fn as_str(self) -> &'static str {
        match self {
            Self::Docx => "DOCX",
            Self::Xlsx => "XLSX",
            Self::Pptx => "PPTX",
        }
    }

    const fn extension(self) -> &'static str {
        match self {
            Self::Docx => "docx",
            Self::Xlsx => "xlsx",
            Self::Pptx => "pptx",
        }
    }

    const fn package_format(self) -> &'static str {
        match self {
            Self::Docx => "DOCX/OPC/ZIP",
            Self::Xlsx => "XLSX/OPC/ZIP",
            Self::Pptx => "PPTX/OPC/ZIP",
        }
    }

    const fn save_entry_point(self) -> &'static str {
        match self {
            Self::Docx => "litchi_docx::Package::save",
            Self::Xlsx => "litchi_xlsx::Workbook::save",
            Self::Pptx => "litchi_pptx::Package::save",
        }
    }

    const fn sink_entry_point(self) -> &'static str {
        match self {
            Self::Docx => "litchi_docx::Package::to_stream",
            Self::Xlsx => "litchi_xlsx::Workbook::write_to",
            // PPTX has no sequential-sink entry point; the counting phase
            // times `to_bytes` and accepts the produced buffer into the sink.
            Self::Pptx => "litchi_pptx::Package::to_bytes",
        }
    }

    const fn real_file_generator(self) -> &'static str {
        match self {
            Self::Docx => DOCX_REAL_FILE_GENERATOR,
            Self::Xlsx => XLSX_REAL_FILE_SAVE_GENERATOR,
            Self::Pptx => PPTX_REAL_FILE_GENERATOR,
        }
    }

    /// The OPC part whose presence decides which format a caller-named file
    /// belongs to. Ordering matters only in that no OOXML package carries two.
    const fn main_part(self) -> &'static str {
        match self {
            Self::Docx => "word/document.xml",
            Self::Xlsx => "xl/workbook.xml",
            Self::Pptx => "ppt/presentation.xml",
        }
    }
}

/// Where the corpus bytes come from.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Origin {
    /// An existing deterministic harness corpus.
    Generated,
    /// A caller-named real file supplied with `--ooxml-file PATH`.
    RealFile,
    /// Change 0664's marker-bearing corpus: the same production writer's
    /// package with the producer's root namespace declarations merged into
    /// every part kind the real fixture marks, so the markup-compatibility
    /// codec takes its rewriting branch on this route.
    MarkerShape,
    /// Change 0664's marker-stripped control: byte-for-byte the same members,
    /// lengths, element counts and attribute counts as `MarkerShape`, with the
    /// namespace URI replaced by an inert URI of the same length, so the codec
    /// takes its borrowing branch instead.
    MarkerControl,
}

impl Origin {
    pub(crate) const ALL: [Self; 4] = [
        Self::Generated,
        Self::RealFile,
        Self::MarkerShape,
        Self::MarkerControl,
    ];

    const fn as_str(self) -> &'static str {
        match self {
            Self::Generated => "generated-harness-corpus",
            Self::RealFile => "caller-named-real-file",
            Self::MarkerShape => "marker-bearing-producer-shape",
            Self::MarkerControl => "marker-stripped-control",
        }
    }

    /// The marker family this origin needs, if any.
    const fn marker_variant(self) -> Option<crate::marker_shape::Variant> {
        match self {
            Self::MarkerShape => Some(crate::marker_shape::Variant::Marker),
            Self::MarkerControl => Some(crate::marker_shape::Variant::Control),
            Self::Generated | Self::RealFile => None,
        }
    }
}

/// One measured phase of the documented open/edit/save route.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Phase {
    /// Open the path, make one semantic edit, save to a path.
    Lifecycle,
    /// The semantic edit alone; the open is outside the clock.
    Edit,
    /// The save-to-path alone; the open and the edit are outside the clock.
    AtomicPublish,
    /// The documented sequential serialization into a bounded counting sink;
    /// the open and the edit are outside the clock.
    CountingPublish,
}

impl Phase {
    #[cfg(test)]
    pub(crate) const ALL: [Self; 4] = [
        Self::Lifecycle,
        Self::Edit,
        Self::AtomicPublish,
        Self::CountingPublish,
    ];

    const fn as_str(self) -> &'static str {
        match self {
            Self::Lifecycle => "open+edit+save",
            Self::Edit => "edit",
            Self::AtomicPublish => "save-to-path",
            Self::CountingPublish => "serialize-to-counting-sink",
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Lifecycle => {
                "open the path through the documented reader, make one semantic edit, save to a \
                 path; destination preparation, readback, digest and cleanup are outside the clock"
            },
            Self::Edit => {
                "the semantic edit and its commit only; the documented open and every verification \
                 are outside the clock"
            },
            Self::AtomicPublish => {
                "one documented save-to-path only: sibling creation in the destination directory, \
                 the publication write, permission preservation, the temporary's data sync, the \
                 rename that replaces the destination, and the parent-directory sync; the open, \
                 the edit, the readback and the cleanup are outside the clock"
            },
            Self::CountingPublish => {
                "the documented sequential serialization into a bounded counting sink only; the \
                 open, the edit and the byte accounting are outside the clock"
            },
        }
    }
}

/// The byte split the counting phase reports, derived from the output archive
/// and the source archive rather than from production instrumentation.
#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub(crate) struct ByteSplit {
    pub(crate) accounting_scope: &'static str,
    /// Length of the whole published archive.
    pub(crate) output_total_bytes: u64,
    /// Sum of the stored compressed payload of Deflate members.
    pub(crate) payload_bytes_deflated: u64,
    /// Sum of the stored payload of Store members.
    pub(crate) payload_bytes_stored: u64,
    /// Sum of the stored compressed payload of members whose bytes are
    /// byte-identical to the same-named source member's. This is the upper
    /// bound on what a copy-through publisher could have avoided.
    pub(crate) payload_bytes_identical_to_source: u64,
    /// Sum of the stored compressed payload of members that are new or whose
    /// bytes differ from the source's.
    pub(crate) payload_bytes_regenerated: u64,
    /// Uncompressed payload every Deflate member declares. A member whose
    /// compressed payload is byte-identical to the source's may have been
    /// preserved rather than recompressed, so this is an upper bound on what
    /// the compressor consumed.
    pub(crate) uncompressed_payload_bytes_compressed: u64,
    /// Uncompressed payload of the Deflate members whose compressed bytes
    /// differ from the source's. This is the part the save certainly fed to
    /// the compressor.
    pub(crate) uncompressed_payload_bytes_regenerated: u64,
    /// `output_total_bytes` minus every member's stored payload: local
    /// headers, the central directory and the end record.
    pub(crate) framing_bytes: u64,
    pub(crate) output_member_count: usize,
    pub(crate) deflate_member_count: usize,
    pub(crate) stored_member_count: usize,
    pub(crate) members_identical_to_source: usize,
    pub(crate) members_regenerated: usize,
}

const BYTE_SPLIT_SCOPE: &str = "derived from the published archive and the source archive, not from production counters: \
     litchi-opc's OpcOperationAccounting excludes PartWriter and the topology publishers, which is \
     the path a documented save takes. `payload_bytes_identical_to_source` is an upper bound on \
     what a copy-through publisher could have avoided re-deflating, not an observation that this \
     writer copied anything.";

/// Everything an ordinary-save corpus states about itself beyond the ordinary
/// [`CorpusManifest`].
#[derive(Clone, Debug, Serialize)]
pub(crate) struct SaveEvidence {
    pub(crate) generator: &'static str,
    pub(crate) format: &'static str,
    pub(crate) origin: &'static str,
    pub(crate) save_entry_point: &'static str,
    pub(crate) sink_entry_point: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) real_file: Option<RealFileProvenance>,
    pub(crate) source_archive_bytes: u64,
    pub(crate) source_archive_sha256: String,
    pub(crate) source_member_count: usize,
    /// What the one semantic edit does, and whether the format's documented
    /// editor admitted it on this corpus.
    pub(crate) edit_description: String,
    pub(crate) edit_admitted: bool,
    /// `"admitted"`, or the editor's typed refusal verbatim. Every retained
    /// sample must reproduce it.
    pub(crate) edit_outcome: String,
    /// The published archive of one untimed reference run: its length and
    /// digest, and the byte split it implies.
    pub(crate) published_bytes: u64,
    pub(crate) published_sha256: String,
    pub(crate) byte_split: ByteSplit,
    /// The 0625/0631 determinism invariant, proved once per corpus before any
    /// sample runs: two fresh open/edit/save cycles agree, and saving the same
    /// edited owner twice agrees.
    pub(crate) repeated_cycle_sha256: String,
    pub(crate) repeated_cycles_identical: bool,
    pub(crate) repeated_save_sha256: String,
    pub(crate) repeated_saves_identical: bool,
}

/// A private directory owned by one corpus. The source artifact is written
/// into it once, outside every timed region, because the documented route
/// takes a path.
#[derive(Debug)]
pub(crate) struct Workspace {
    root: PathBuf,
    source: PathBuf,
    destination: PathBuf,
    alternate: PathBuf,
}

impl Workspace {
    fn prepare(
        requested_root: Option<&Path>,
        format: Format,
        archive: &[u8],
    ) -> Result<Self, Box<dyn Error>> {
        let root = crate::filesystem::scratch_root(requested_root, "ordinary-save")?;
        let extension = format.extension();
        let source = root.join(format!("source.{extension}"));
        let destination = root.join(format!("published.{extension}"));
        let alternate = root.join(format!("published-repeat.{extension}"));
        fs::write(&source, archive)?;
        Ok(Self {
            root,
            source,
            destination,
            alternate,
        })
    }

    pub(crate) fn source(&self) -> &Path {
        &self.source
    }
}

impl Drop for Workspace {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.alternate);
        let _ = fs::remove_file(&self.destination);
        let _ = fs::remove_file(&self.source);
        let _ = fs::remove_dir(&self.root);
    }
}

/// An ordinary-save corpus: a [`Corpus`], its evidence block, and the private
/// workspace the documented path route needs.
#[derive(Debug)]
pub(crate) struct SaveCorpus {
    pub(crate) corpus: Corpus,
    pub(crate) evidence: SaveEvidence,
    format: Format,
    origin: Origin,
    workspace: Workspace,
    xlsx_sheet: String,
    xlsx_address: String,
    pptx_target: Option<(usize, usize)>,
}

/// Per-case evidence written into `source.ordinary_save`.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct OrdinarySaveSummary {
    pub(crate) format: &'static str,
    pub(crate) origin: &'static str,
    pub(crate) phase: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) atomic_publication_steps: &'static str,
    pub(crate) corpus: SaveEvidence,
    /// SHA-256 of every retained sample's published artifact, and whether all
    /// of them agree with the corpus's reference publication.
    pub(crate) published_sha256: Vec<String>,
    pub(crate) publications_identical: bool,
    /// Every retained sample's own edit outcome, and whether all of them
    /// reproduced the frozen one.
    pub(crate) edit_outcome_sha256: Vec<String>,
    pub(crate) edit_outcomes_identical: bool,
    /// The byte split each retained sample's own output implies, restated only
    /// for the counting phase (the others do not build an in-memory archive).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sample_byte_split: Option<ByteSplit>,
}

const ATOMIC_STEPS: &str = "litchi_opc::atomic::replace_with: destination permission probe, \
                            sibling temporary creation in the destination's own directory, the \
                            publication write, permission preservation, sync_all on the temporary, \
                            persist (rename) over the destination, parent-directory sync";

// ---------------------------------------------------------------------------
// Byte accounting
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
struct MemberFacts {
    method: CompressionMethod,
    compressed: Vec<u8>,
    uncompressed_size: u64,
}

fn member_facts(bytes: &[u8]) -> Result<Vec<(String, MemberFacts)>, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut members = Vec::new();
    for header in archive.entries() {
        let header = header?;
        if header.is_dir() {
            continue;
        }
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let method = header.compression_method();
        let uncompressed_size = header.uncompressed_size_hint();
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        let start = usize::try_from(start)?;
        let end = usize::try_from(end)?;
        let compressed = bytes
            .get(start..end)
            .ok_or("published archive member payload range is out of bounds")?
            .to_vec();
        members.push((
            name,
            MemberFacts {
                method,
                compressed,
                uncompressed_size,
            },
        ));
    }
    Ok(members)
}

fn byte_split(source: &[u8], published: &[u8]) -> Result<ByteSplit, Box<dyn Error>> {
    let source_members = member_facts(source)?
        .into_iter()
        .collect::<std::collections::BTreeMap<_, _>>();
    let published_members = member_facts(published)?;
    let mut split = ByteSplit {
        accounting_scope: BYTE_SPLIT_SCOPE,
        output_total_bytes: u64::try_from(published.len())?,
        payload_bytes_deflated: 0,
        payload_bytes_stored: 0,
        payload_bytes_identical_to_source: 0,
        payload_bytes_regenerated: 0,
        uncompressed_payload_bytes_compressed: 0,
        uncompressed_payload_bytes_regenerated: 0,
        framing_bytes: 0,
        output_member_count: published_members.len(),
        deflate_member_count: 0,
        stored_member_count: 0,
        members_identical_to_source: 0,
        members_regenerated: 0,
    };
    let mut payload_total = 0_u64;
    for (name, facts) in &published_members {
        let payload = u64::try_from(facts.compressed.len())?;
        payload_total = payload_total
            .checked_add(payload)
            .ok_or("published payload byte total overflows u64")?;
        match facts.method {
            CompressionMethod::Deflate => {
                split.deflate_member_count += 1;
                split.payload_bytes_deflated = split
                    .payload_bytes_deflated
                    .checked_add(payload)
                    .ok_or("deflated payload byte total overflows u64")?;
                split.uncompressed_payload_bytes_compressed = split
                    .uncompressed_payload_bytes_compressed
                    .checked_add(facts.uncompressed_size)
                    .ok_or("compressed input byte total overflows u64")?;
            },
            _ => {
                split.stored_member_count += 1;
                split.payload_bytes_stored = split
                    .payload_bytes_stored
                    .checked_add(payload)
                    .ok_or("stored payload byte total overflows u64")?;
            },
        }
        let identical = source_members.get(name).is_some_and(|source| {
            source.method == facts.method && source.compressed == facts.compressed
        });
        if !identical && facts.method == CompressionMethod::Deflate {
            split.uncompressed_payload_bytes_regenerated = split
                .uncompressed_payload_bytes_regenerated
                .checked_add(facts.uncompressed_size)
                .ok_or("regenerated compressor input byte total overflows u64")?;
        }
        if identical {
            split.members_identical_to_source += 1;
            split.payload_bytes_identical_to_source = split
                .payload_bytes_identical_to_source
                .checked_add(payload)
                .ok_or("source-identical payload byte total overflows u64")?;
        } else {
            split.members_regenerated += 1;
            split.payload_bytes_regenerated = split
                .payload_bytes_regenerated
                .checked_add(payload)
                .ok_or("regenerated payload byte total overflows u64")?;
        }
    }
    split.framing_bytes = split
        .output_total_bytes
        .checked_sub(payload_total)
        .ok_or("published archive is smaller than its own member payloads")?;
    Ok(split)
}

// ---------------------------------------------------------------------------
// Owners: open, edit, publish
// ---------------------------------------------------------------------------

/// What the documented editor did with this corpus.
#[derive(Clone, Debug, PartialEq, Eq)]
enum EditOutcome {
    Admitted,
    Refused(String),
}

impl EditOutcome {
    fn as_str(&self) -> &str {
        match self {
            Self::Admitted => "admitted",
            Self::Refused(reason) => reason.as_str(),
        }
    }
}

/// One opened and edited owner of the documented kind.
enum Owner {
    Docx(Box<litchi_docx::Package>),
    Xlsx(Box<litchi_xlsx::Workbook>),
    Pptx(Box<litchi_pptx::Package>),
}

impl Owner {
    fn open(format: Format, path: &Path) -> Result<Self, Box<dyn Error>> {
        Ok(match format {
            Format::Docx => Self::Docx(Box::new(litchi_docx::Package::open(path)?)),
            Format::Xlsx => Self::Xlsx(Box::new(litchi_xlsx::Workbook::open(path)?)),
            Format::Pptx => Self::Pptx(Box::new(litchi_pptx::Package::open(path)?)),
        })
    }

    /// Applies the corpus's one semantic edit through the documented editor.
    ///
    /// A typed refusal is returned verbatim rather than propagated, for the
    /// reason change 0627 froze the XLS full-text and PPT one-shape refusals:
    /// a caller pays for the attempt either way, and a refusal that a real
    /// producer's file provokes is the measurement, not a harness failure. An
    /// owner that refuses is left exactly as the open produced it, so the
    /// publish phases still measure the documented save of an unedited
    /// package — change 0593's `noop` scenario.
    fn edit(&mut self, corpus: &SaveCorpus) -> Result<EditOutcome, Box<dyn Error>> {
        match self {
            Self::Docx(package) => match package.document_mut() {
                Ok(document) => {
                    document.add_paragraph_with_text(EDIT_MARKER);
                    Ok(EditOutcome::Admitted)
                },
                Err(error) => Ok(EditOutcome::Refused(format!("refused:{error}"))),
            },
            Self::Xlsx(workbook) => {
                let outcome = (|| -> Result<litchi_xlsx::Workbook, String> {
                    let mut edit = workbook.edit().map_err(|error| error.to_string())?;
                    {
                        let mut sheet = edit
                            .sheet(corpus.xlsx_sheet.as_str())
                            .map_err(|error| error.to_string())?
                            .ok_or_else(|| {
                                "the selected worksheet is absent from this workbook".to_owned()
                            })?;
                        sheet
                            .set(corpus.xlsx_address.as_str(), EDIT_MARKER)
                            .map_err(|error| error.to_string())?;
                    }
                    let commit = edit.commit().map_err(|error| error.to_string())?;
                    if commit.patch().is_empty() {
                        return Err("the edit produced an empty patch".to_owned());
                    }
                    Ok(commit.into_workbook())
                })();
                match outcome {
                    Ok(edited) => {
                        **workbook = edited;
                        Ok(EditOutcome::Admitted)
                    },
                    Err(error) => Ok(EditOutcome::Refused(format!("refused:{error}"))),
                }
            },
            Self::Pptx(package) => {
                let Some((slide, shape)) = corpus.pptx_target else {
                    return Ok(EditOutcome::Refused(corpus.evidence.edit_outcome.clone()));
                };
                let outcome = (|| -> Result<litchi_pptx::opened::Commit, String> {
                    let mut edit = package
                        .opened_presentation_transaction()
                        .map_err(|error| error.to_string())?;
                    if !edit
                        .set_shape_text(slide, shape, EDIT_MARKER)
                        .map_err(|error| error.to_string())?
                    {
                        return Err("the derived shape position reported no change".to_owned());
                    }
                    let commit = edit.commit().map_err(|error| error.to_string())?;
                    if !commit.is_changed() {
                        return Err("the commit reports no change".to_owned());
                    }
                    Ok(commit)
                })();
                match outcome {
                    Ok(commit) => {
                        package.apply_opened_presentation_commit(commit)?;
                        Ok(EditOutcome::Admitted)
                    },
                    Err(error) => Ok(EditOutcome::Refused(format!("refused:{error}"))),
                }
            },
        }
    }

    /// The documented save-to-path entry point.
    fn save(&mut self, path: &Path) -> Result<(), Box<dyn Error>> {
        match self {
            Self::Docx(package) => package.save(path)?,
            Self::Xlsx(workbook) => workbook.save(path)?,
            Self::Pptx(package) => package.save(path)?,
        }
        Ok(())
    }

    /// The documented sequential serialization entry point.
    fn write_to(&mut self, sink: &mut BoundedSink) -> Result<(), Box<dyn Error>> {
        match self {
            Self::Docx(package) => package.to_stream(&mut *sink)?,
            Self::Xlsx(workbook) => workbook.write_to(&mut *sink)?,
            Self::Pptx(package) => {
                let bytes = package.to_bytes()?;
                sink.write_all(&bytes)?;
            },
        }
        Ok(())
    }
}

/// A bounded counting sink that retains the bytes so the byte split can be
/// derived from them. It refuses to grow past its budget, so a runaway
/// serialization fails instead of exhausting memory.
#[derive(Debug)]
pub(crate) struct BoundedSink {
    bytes: Vec<u8>,
    max_bytes: usize,
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
}

impl BoundedSink {
    fn bounded(max_bytes: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_bytes,
            accepted_bytes: 0,
            write_calls: 0,
            largest_write: 0,
        }
    }

    fn reserve_budget(&mut self) -> io::Result<()> {
        self.bytes
            .try_reserve_exact(self.max_bytes)
            .map_err(|error| io::Error::other(format!("cannot reserve sink byte budget: {error}")))
    }

    fn summary(&self) -> SinkSummary {
        SinkSummary {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            largest_write: self.largest_write,
            write_size_buckets: crate::WriteSizeBuckets::empty(),
            retained_output_bytes: None,
            retained_authoring_window_bytes: None,
            rows: None,
            cells: None,
            paragraphs: None,
            runs: None,
            input_bytes: None,
            authored_part_bytes: None,
            rtf_tail_append: None,
        }
    }
}

impl Write for BoundedSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.bytes.len().saturating_add(buffer.len()) > self.max_bytes {
            return Err(io::Error::other(
                "ordinary-save counting sink exceeded its byte budget",
            ));
        }
        self.bytes.extend_from_slice(buffer);
        self.accepted_bytes = self
            .accepted_bytes
            .checked_add(buffer.len() as u64)
            .ok_or_else(|| io::Error::other("sink accepted byte total overflows u64"))?;
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("sink write call count overflows u64"))?;
        self.largest_write = self.largest_write.max(buffer.len() as u64);
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Corpus construction
// ---------------------------------------------------------------------------

fn read_bounded(path: &Path) -> Result<Vec<u8>, Box<dyn Error>> {
    let metadata = fs::metadata(path)
        .map_err(|source| format!("--ooxml-file {} is unreadable: {source}", path.display()))?;
    if !metadata.is_file() {
        return Err(format!("--ooxml-file {} is not a regular file", path.display()).into());
    }
    if metadata.len() > MAX_OOXML_FILE_BYTES {
        return Err(format!(
            "--ooxml-file {} is {} bytes, above the {MAX_OOXML_FILE_BYTES}-byte bound",
            path.display(),
            metadata.len()
        )
        .into());
    }
    fs::read(path).map_err(|source| {
        format!(
            "--ooxml-file {} could not be read: {source}",
            path.display()
        )
        .into()
    })
}

/// Decides which OOXML reader a caller-named file belongs to by its own member
/// inventory, so `--ooxml-file` needs no format flag and no extension
/// heuristic. A file that offers several main parts, or none, is refused.
pub(crate) fn classify(path: &Path) -> Result<Format, Box<dyn Error>> {
    let archive = read_bounded(path)?;
    let names = member_facts(&archive)?
        .into_iter()
        .map(|(name, _)| name)
        .collect::<Vec<_>>();
    let matched = Format::ALL
        .into_iter()
        .filter(|format| names.iter().any(|name| name == format.main_part()))
        .collect::<Vec<_>>();
    match matched.as_slice() {
        [format] => Ok(*format),
        [] => Err(format!(
            "--ooxml-file {} holds no word/document.xml, xl/workbook.xml or ppt/presentation.xml",
            path.display()
        )
        .into()),
        _ => Err(format!(
            "--ooxml-file {} holds more than one OOXML main part",
            path.display()
        )
        .into()),
    }
}

/// The caller-named OOXML fixtures of one run, at most one per format.
#[derive(Debug, Default)]
pub(crate) struct OoxmlInputs {
    docx: Option<PathBuf>,
    xlsx: Option<PathBuf>,
    pptx: Option<PathBuf>,
}

impl OoxmlInputs {
    pub(crate) fn get(&self, format: Format) -> Option<&Path> {
        match format {
            Format::Docx => self.docx.as_deref(),
            Format::Xlsx => self.xlsx.as_deref(),
            Format::Pptx => self.pptx.as_deref(),
        }
    }
}

pub(crate) fn classify_inputs(paths: &[PathBuf]) -> Result<OoxmlInputs, Box<dyn Error>> {
    let mut inputs = OoxmlInputs::default();
    for path in paths {
        let slot = match classify(path)? {
            Format::Docx => &mut inputs.docx,
            Format::Xlsx => &mut inputs.xlsx,
            Format::Pptx => &mut inputs.pptx,
        };
        if slot.is_some() {
            return Err(format!(
                "--ooxml-file accepts at most one file per format; {} is the second",
                path.display()
            )
            .into());
        }
        *slot = Some(path.clone());
    }
    Ok(inputs)
}

/// Builds the generated corpus for one format from an existing harness corpus.
fn generated_archive(format: Format) -> Result<(Vec<u8>, CorpusManifest), Box<dyn Error>> {
    let corpus = match format {
        Format::Docx => crate::build_semantic_docx_corpus(SemanticShape::Medium)?,
        Format::Xlsx => crate::build_xlsx_cell_crud_corpus(XlsxCellCrudShape::Medium)?,
        Format::Pptx => crate::build_semantic_pptx_corpus(SemanticShape::Medium)?,
    };
    Ok((corpus.archive, corpus.manifest))
}

fn real_file_manifest(
    format: Format,
    archive: &[u8],
    member_count: usize,
    main_part: &str,
    main_payload_bytes: usize,
    main_payload_sha256: String,
) -> Result<CorpusManifest, Box<dyn Error>> {
    Ok(CorpusManifest {
        name: format!("{}-real-file-save", format.extension()),
        generator: format.real_file_generator(),
        package_format: format.package_format(),
        shape: "real-file",
        payload_kind: "real-producer-ooxml",
        compression: "deflate",
        entry_count: member_count,
        archive_member_count: member_count,
        entry_bytes: 0,
        uncompressed_payload_bytes: main_payload_bytes,
        archive_bytes: archive.len(),
        archive_sha256: sha256_hex(archive),
        target_entry: main_part.to_owned(),
        target_payload_bytes: main_payload_bytes,
        target_payload_sha256: main_payload_sha256,
        rtf_variant: None,
        xlsx: None,
    })
}

/// Derives the XLSX edit target from the workbook itself: the first worksheet
/// the reader reports, and a fixed address inside it.
fn xlsx_edit_target(path: &Path) -> Result<(String, String), Box<dyn Error>> {
    let workbook = litchi_xlsx::Workbook::open(path)?;
    let sheet = workbook
        .sheets()
        .next()
        .ok_or("ordinary-save XLSX corpus declares no worksheet")?;
    let name = sheet.name().to_owned();
    drop(workbook);
    Ok((name, "A1".to_owned()))
}

/// Derives the PPTX edit target from the deck itself: the first (slide, shape)
/// position the documented opened-presentation transaction admits. `None`
/// means the deck refuses every position, which is frozen as the corpus's edit
/// outcome rather than treated as an error.
fn pptx_edit_target(
    path: &Path,
) -> Result<(Option<(usize, usize)>, Option<String>), Box<dyn Error>> {
    const MAX_SLIDES: usize = 64;
    const MAX_SHAPES: usize = 32;
    let package = litchi_pptx::Package::open(path)?;
    let slide_count = match package.opened_presentation() {
        Ok(snapshot) => snapshot.slides().len().min(MAX_SLIDES),
        Err(error) => return Ok((None, Some(format!("refused:{error}")))),
    };
    let mut refusal = None;
    for slide in 0..slide_count {
        for shape in 0..MAX_SHAPES {
            // Each transaction is detached from the package, so one open is
            // enough to probe every position.
            let Ok(mut edit) = package.opened_presentation_transaction() else {
                continue;
            };
            match edit.set_shape_text(slide, shape, EDIT_MARKER) {
                Ok(true) => return Ok((Some((slide, shape)), None)),
                Ok(false) => continue,
                Err(error) => {
                    if refusal.is_none() {
                        refusal = Some(format!("refused:{error}"));
                    }
                },
            }
        }
    }
    Ok((
        None,
        Some(refusal.unwrap_or_else(|| "refused:no admitted shape position".to_owned())),
    ))
}

/// Builds one ordinary-save corpus, proving the 0625/0631 determinism
/// invariant and deriving the reference publication and its byte split before
/// any sample runs. Everything here is untimed and happens once.
pub(crate) fn build_corpus(
    format: Format,
    origin: Origin,
    real_file: Option<&Path>,
    filesystem_root: Option<&Path>,
) -> Result<SaveCorpus, Box<dyn Error>> {
    let (archive, manifest, provenance) = match origin {
        Origin::Generated => {
            let (archive, manifest) = generated_archive(format)?;
            (archive, manifest, None)
        },
        Origin::MarkerShape | Origin::MarkerControl => {
            let variant = origin
                .marker_variant()
                .ok_or("a marker ordinary-save origin has no marker variant")?;
            let family = match format {
                Format::Docx => crate::marker_shape::Family::Docx,
                Format::Pptx => crate::marker_shape::Family::Pptx,
                Format::Xlsx => {
                    return Err(
                        "the marker-shape ordinary-save selectors cover DOCX and PPTX only".into(),
                    );
                },
            };
            let corpus = crate::marker_shape::build(family, variant)?;
            (corpus.corpus.archive, corpus.corpus.manifest, None)
        },
        Origin::RealFile => {
            let path = real_file.ok_or_else(|| {
                format!(
                    "the {} real-file ordinary-save selectors require an --ooxml-file {} fixture",
                    format.extension(),
                    format.as_str()
                )
            })?;
            let archive = read_bounded(path)?;
            let members = member_facts(&archive)?;
            let main = members
                .iter()
                .find(|(name, _)| name == format.main_part())
                .ok_or("ordinary-save real file lost its main part between classify and build")?;
            let main_payload_sha256 = sha256_hex(&main.1.compressed);
            let manifest = real_file_manifest(
                format,
                &archive,
                members.len(),
                format.main_part(),
                main.1.compressed.len(),
                main_payload_sha256.clone(),
            )?;
            let provenance = RealFileProvenance {
                path: path.display().to_string(),
                bytes: u64::try_from(archive.len())?,
                sha256: sha256_hex(&archive),
            };
            (archive, manifest, Some(provenance))
        },
    };

    let source_member_count = member_facts(&archive)?.len();
    let source_sha256 = sha256_hex(&archive);
    let workspace = Workspace::prepare(filesystem_root, format, &archive)?;

    let (xlsx_sheet, xlsx_address) = if format == Format::Xlsx {
        xlsx_edit_target(workspace.source())?
    } else {
        (String::new(), String::new())
    };
    let (pptx_target, pptx_refusal) = if format == Format::Pptx {
        pptx_edit_target(workspace.source())?
    } else {
        (None, None)
    };
    // Provisional: the PPTX branch of `Owner::edit` consults it when no shape
    // position was admitted. The probe below replaces it with the outcome the
    // documented editor actually produced.
    let provisional_outcome = pptx_refusal
        .clone()
        .unwrap_or_else(|| "admitted".to_owned());
    let edit_description = match format {
        Format::Docx => {
            format!("Package::document_mut().add_paragraph_with_text({EDIT_MARKER:?})")
        },
        Format::Xlsx => format!(
            "Workbook::edit().sheet({xlsx_sheet:?}).set({xlsx_address:?}, {EDIT_MARKER:?}) and commit"
        ),
        Format::Pptx => match pptx_target {
            Some((slide, shape)) => format!(
                "Package::opened_presentation_transaction().set_shape_text({slide}, {shape}, \
                 {EDIT_MARKER:?}) and apply"
            ),
            None => "Package::opened_presentation_transaction() admits no shape position on this \
                     deck; the save phases publish the unedited opened package"
                .to_owned(),
        },
    };

    // A partially built corpus is enough for `Owner::edit` to consult.
    let mut corpus = SaveCorpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: format.main_part().to_owned(),
            target_payload: Vec::new(),
            xlsx: None,
        },
        evidence: SaveEvidence {
            generator: match origin {
                Origin::Generated => "litchi-perf-existing-corpus",
                Origin::RealFile => format.real_file_generator(),
                Origin::MarkerShape | Origin::MarkerControl => match format {
                    Format::Docx => crate::marker_shape::DOCX_MARKER_SHAPE_GENERATOR,
                    Format::Pptx => crate::marker_shape::PPTX_MARKER_SHAPE_GENERATOR,
                    Format::Xlsx => "litchi-perf-existing-corpus",
                },
            },
            format: format.as_str(),
            origin: origin.as_str(),
            save_entry_point: format.save_entry_point(),
            sink_entry_point: format.sink_entry_point(),
            real_file: provenance,
            source_archive_bytes: 0,
            source_archive_sha256: source_sha256.clone(),
            source_member_count,
            edit_description,
            edit_admitted: false,
            edit_outcome: provisional_outcome,
            published_bytes: 0,
            published_sha256: String::new(),
            byte_split: ByteSplit {
                accounting_scope: BYTE_SPLIT_SCOPE,
                output_total_bytes: 0,
                payload_bytes_deflated: 0,
                payload_bytes_stored: 0,
                payload_bytes_identical_to_source: 0,
                payload_bytes_regenerated: 0,
                uncompressed_payload_bytes_compressed: 0,
                uncompressed_payload_bytes_regenerated: 0,
                framing_bytes: 0,
                output_member_count: 0,
                deflate_member_count: 0,
                stored_member_count: 0,
                members_identical_to_source: 0,
                members_regenerated: 0,
            },
            repeated_cycle_sha256: String::new(),
            repeated_cycles_identical: false,
            repeated_save_sha256: String::new(),
            repeated_saves_identical: false,
        },
        format,
        origin,
        workspace,
        xlsx_sheet,
        xlsx_address,
        pptx_target,
    };
    corpus.evidence.source_archive_bytes = u64::try_from(corpus.corpus.archive.len())?;

    // Probe the documented edit once, untimed, so its outcome is frozen before
    // any sample runs. A typed refusal is an outcome, not a failure.
    let probe = {
        let mut owner = Owner::open(format, corpus.workspace.source())?;
        owner.edit(&corpus)?
    };
    corpus.evidence.edit_admitted = probe == EditOutcome::Admitted;
    corpus.evidence.edit_outcome = probe.as_str().to_owned();

    // Reference publication, and the two determinism proofs the 0625 and 0631
    // invariants ask for: two fresh open/edit/save cycles must agree, and
    // saving one edited owner twice must agree.
    let first = publish_reference(&corpus, false)?;
    let second = publish_reference(&corpus, false)?;
    let repeated_save = publish_reference(&corpus, true)?;
    let published = fs::read(&corpus.workspace.destination)?;
    corpus.evidence.published_bytes = u64::try_from(published.len())?;
    corpus.evidence.published_sha256 = first.clone();
    corpus.evidence.byte_split = byte_split(&corpus.corpus.archive, &published)?;
    corpus.evidence.repeated_cycle_sha256 = second.clone();
    corpus.evidence.repeated_cycles_identical = first == second;
    corpus.evidence.repeated_save_sha256 = repeated_save.clone();
    corpus.evidence.repeated_saves_identical = first == repeated_save;
    if !corpus.evidence.repeated_cycles_identical {
        return Err(format!(
            "{} ordinary save is not a function of the document: two fresh open/edit/save cycles \
             produced {first} and {second}",
            format.as_str()
        )
        .into());
    }
    if !corpus.evidence.repeated_saves_identical {
        return Err(format!(
            "{} repeated save of one edited owner produced {first} then {repeated_save}",
            format.as_str()
        )
        .into());
    }
    corpus.corpus.target_payload = published;
    let _ = fs::remove_file(&corpus.workspace.destination);
    let _ = fs::remove_file(&corpus.workspace.alternate);
    Ok(corpus)
}

/// Runs one untimed open/edit/save cycle and returns the published digest.
/// With `repeat_save`, the same edited owner is saved twice — to the alternate
/// destination and then to the ordinary one — which is the 0625/0631 invariant
/// that a repeated save digests identically.
fn publish_reference(corpus: &SaveCorpus, repeat_save: bool) -> Result<String, Box<dyn Error>> {
    let mut owner = Owner::open(corpus.format, corpus.workspace.source())?;
    let outcome = owner.edit(corpus)?;
    if !corpus.evidence.edit_outcome.is_empty()
        && corpus.evidence.edit_outcome != "admitted"
        && outcome.as_str() != corpus.evidence.edit_outcome
    {
        return Err(format!(
            "{} edit outcome is not a function of the document: {} then {}",
            corpus.format.as_str(),
            corpus.evidence.edit_outcome,
            outcome.as_str()
        )
        .into());
    }
    if repeat_save {
        owner.save(&corpus.workspace.alternate)?;
        let first = sha256_hex(&fs::read(&corpus.workspace.alternate)?);
        owner.save(&corpus.workspace.destination)?;
        let second = sha256_hex(&fs::read(&corpus.workspace.destination)?);
        if first != second {
            return Err(format!(
                "{} repeated save of one owner produced {first} then {second}",
                corpus.format.as_str()
            )
            .into());
        }
        return Ok(second);
    }
    owner.save(&corpus.workspace.destination)?;
    Ok(sha256_hex(&fs::read(&corpus.workspace.destination)?))
}

// ---------------------------------------------------------------------------
// Measurement
// ---------------------------------------------------------------------------

pub(crate) fn run_case(
    case: Case,
    phase: Phase,
    corpus: &SaveCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let expected = corpus.evidence.published_sha256.clone();
    let expected_edit = sha256_hex(corpus.evidence.edit_outcome.as_bytes());
    let budget = usize::try_from(corpus.evidence.published_bytes)?
        .saturating_mul(4)
        .max(64 * 1024);
    let mut elapsed = Vec::with_capacity(samples);
    let mut published_sha256 = Vec::with_capacity(samples);
    let mut edit_outcome_sha256 = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    let mut sample_split = None;

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // The allocation region opens immediately before the clock and closes
        // immediately after it, so it covers exactly the interval this phase
        // reports. Change 0649 found this family emitted no allocation metrics
        // at all because it opened no region; it does now. The owner is still
        // alive when the region closes (`drop(owner)` is outside the timer in
        // every phase, as it has always been), so `live_bytes_after` is
        // retained memory rather than a leak, exactly as the ODP and XLSX
        // source-backed families already report it.
        let (duration, digest, sink, outcome, allocation) = match phase {
            Phase::Lifecycle => {
                let region = allocation_metrics::begin();
                let started = Instant::now();
                let mut owner = Owner::open(corpus.format, corpus.workspace.source())?;
                let outcome = owner.edit(corpus)?;
                owner.save(&corpus.workspace.destination)?;
                let duration = started.elapsed();
                let allocation = region.finish();
                drop(owner);
                let digest = sha256_hex(&fs::read(&corpus.workspace.destination)?);
                fs::remove_file(&corpus.workspace.destination)?;
                (duration, Some(digest), None, outcome, allocation)
            },
            Phase::Edit => {
                let mut owner = Owner::open(corpus.format, corpus.workspace.source())?;
                let region = allocation_metrics::begin();
                let started = Instant::now();
                let outcome = owner.edit(corpus)?;
                let duration = started.elapsed();
                let allocation = region.finish();
                std::hint::black_box(&outcome);
                drop(owner);
                (duration, None, None, outcome, allocation)
            },
            Phase::AtomicPublish => {
                let mut owner = Owner::open(corpus.format, corpus.workspace.source())?;
                let outcome = owner.edit(corpus)?;
                let region = allocation_metrics::begin();
                let started = Instant::now();
                owner.save(&corpus.workspace.destination)?;
                let duration = started.elapsed();
                let allocation = region.finish();
                drop(owner);
                let digest = sha256_hex(&fs::read(&corpus.workspace.destination)?);
                fs::remove_file(&corpus.workspace.destination)?;
                (duration, Some(digest), None, outcome, allocation)
            },
            Phase::CountingPublish => {
                let mut owner = Owner::open(corpus.format, corpus.workspace.source())?;
                let outcome = owner.edit(corpus)?;
                let mut sink = BoundedSink::bounded(budget);
                sink.reserve_budget()?;
                let region = allocation_metrics::begin();
                let started = Instant::now();
                owner.write_to(&mut sink)?;
                let duration = started.elapsed();
                let allocation = region.finish();
                drop(owner);
                let digest = sha256_hex(&sink.bytes);
                if sample_split.is_none() {
                    sample_split = Some(byte_split(&corpus.corpus.archive, &sink.bytes)?);
                }
                let summary = sink.summary();
                std::hint::black_box(&sink.bytes);
                (duration, Some(digest), Some(summary), outcome, allocation)
            },
        };
        if iteration >= warmup_iterations {
            observations.push(operation_metrics::InProcessObservation {
                elapsed_ns: elapsed_ns(duration)?,
                process_metrics: None,
                allocation_metrics: Some(
                    allocation.unwrap_or_else(allocation_metrics::unavailable_sample),
                ),
            });
            if let Some(digest) = digest {
                published_sha256.push(digest);
            }
            if let Some(summary) = sink {
                sink_summaries.push(summary);
            }
            edit_outcome_sha256.push(sha256_hex(outcome.as_str().as_bytes()));
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let publications_identical = published_sha256.iter().all(|digest| *digest == expected)
        && (phase == Phase::Edit || !published_sha256.is_empty());
    if !publications_identical {
        return Err(format!(
            "{} {} published an artifact that differs from the corpus reference",
            corpus.format.as_str(),
            phase.as_str()
        )
        .into());
    }
    let edit_outcomes_identical = edit_outcome_sha256
        .iter()
        .all(|digest| *digest == expected_edit);
    if !edit_outcomes_identical {
        return Err(format!(
            "{} {} produced an edit outcome that differs from the frozen one",
            corpus.format.as_str(),
            phase.as_str()
        )
        .into());
    }

    let sink = if sink_summaries.is_empty() {
        None
    } else {
        Some(crate::deterministic_sink_summary(
            &sink_summaries,
            "ordinary save counting publication",
        )?)
    };
    let output_sha256 = (phase != Phase::Edit).then(|| expected.clone());

    let summary = SourceSummary {
        ordinary_save: Some(Box::new(OrdinarySaveSummary {
            format: corpus.format.as_str(),
            origin: corpus.origin.as_str(),
            phase: phase.as_str(),
            timing_scope: phase.timing_scope(),
            atomic_publication_steps: ATOMIC_STEPS,
            corpus: corpus.evidence.clone(),
            published_sha256,
            publications_identical,
            edit_outcome_sha256,
            edit_outcomes_identical,
            sample_byte_split: sample_split,
        })),
        ..SourceSummary::default()
    };

    let metrics = operation_metrics::from_in_process_observations_without_sink(&observations)?;

    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink,
        source: boxed_source(summary),
        execution: None,
        output_sha256,
        operation_metrics: Some(metrics),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn formats_origins_and_phases_describe_themselves() {
        assert_eq!(
            Format::Docx.save_entry_point(),
            "litchi_docx::Package::save"
        );
        assert_eq!(
            Format::Xlsx.save_entry_point(),
            "litchi_xlsx::Workbook::save"
        );
        assert_eq!(Format::Pptx.main_part(), "ppt/presentation.xml");
        assert_eq!(Origin::RealFile.as_str(), "caller-named-real-file");
        assert_eq!(
            Origin::MarkerShape.as_str(),
            "marker-bearing-producer-shape"
        );
        assert_eq!(Origin::MarkerControl.as_str(), "marker-stripped-control");
        assert_eq!(Phase::AtomicPublish.as_str(), "save-to-path");
        assert_eq!(Phase::ALL.len(), 4);
        assert_eq!(Format::ALL.len(), 3);
        assert_eq!(Origin::ALL.len(), 4);
        assert_eq!(Origin::Generated.marker_variant(), None);
        assert_eq!(Origin::RealFile.marker_variant(), None);
    }

    #[test]
    fn every_ordinary_save_case_is_opt_in_and_round_trips() {
        let mut seen = 0;
        for format in Format::ALL {
            for origin in Origin::ALL {
                for phase in Phase::ALL {
                    // Change 0664's two marker origins cover DOCX and PPTX
                    // only: change 0601's XLSX producer family already owns
                    // the marker-bearing spreadsheet shapes, so no XLSX
                    // marker selector exists and `build_corpus` refuses one
                    // with a typed error.
                    if format == Format::Xlsx && origin.marker_variant().is_some() {
                        assert_eq!(Case::ordinary_save_case(format, origin, phase), None);
                        continue;
                    }
                    let case = Case::ordinary_save_case(format, origin, phase)
                        .expect("every ordinary-save triple has a selector");
                    assert!(case.is_ordinary_save());
                    assert!(!Case::DEFAULT.contains(&case));
                    assert_eq!(crate::parse_case(case.name()), Some(case));
                    assert_eq!(case.ordinary_save_plan(), Some((format, origin, phase)));
                    seen += 1;
                }
            }
        }
        assert_eq!(seen, 40);
        assert!(
            build_corpus(Format::Xlsx, Origin::MarkerShape, None, None)
                .expect_err("an XLSX marker corpus is refused")
                .to_string()
                .contains("DOCX and PPTX only")
        );
    }

    #[test]
    fn generated_docx_save_is_deterministic_and_reports_its_byte_split() {
        let corpus = build_corpus(Format::Docx, Origin::Generated, None, None).unwrap();
        assert!(corpus.evidence.repeated_cycles_identical);
        assert!(corpus.evidence.repeated_saves_identical);
        assert!(corpus.evidence.edit_admitted);
        assert_eq!(corpus.evidence.edit_outcome, "admitted");
        let split = &corpus.evidence.byte_split;
        assert!(split.output_total_bytes > 0);
        assert_eq!(
            split.payload_bytes_identical_to_source + split.payload_bytes_regenerated,
            split.payload_bytes_deflated + split.payload_bytes_stored
        );
        assert_eq!(
            split.members_identical_to_source + split.members_regenerated,
            split.output_member_count
        );
        assert!(split.framing_bytes > 0);

        for (phase, case) in [
            (Phase::Lifecycle, Case::DocxOrdinarySaveLifecycle),
            (Phase::Edit, Case::DocxOrdinarySaveEdit),
            (Phase::AtomicPublish, Case::DocxOrdinarySaveAtomicPublish),
            (
                Phase::CountingPublish,
                Case::DocxOrdinarySaveCountingPublish,
            ),
        ] {
            let result = run_case(case, phase, &corpus, 0, 2).unwrap();
            let source = result.source.unwrap().ordinary_save.unwrap();
            assert!(source.publications_identical);
            assert_eq!(source.format, "DOCX");
            if phase == Phase::CountingPublish {
                assert!(source.sample_byte_split.is_some());
                assert!(result.sink.is_some());
            }
        }
    }

    #[test]
    fn generated_xlsx_and_pptx_saves_are_deterministic() {
        for format in [Format::Xlsx, Format::Pptx] {
            let corpus = build_corpus(format, Origin::Generated, None, None).unwrap();
            assert!(corpus.evidence.repeated_cycles_identical);
            assert!(corpus.evidence.repeated_saves_identical);
            assert!(corpus.evidence.published_bytes > 0);
            let case = Case::ordinary_save_case(format, Origin::Generated, Phase::Lifecycle)
                .expect("lifecycle selector exists");
            let result = run_case(case, Phase::Lifecycle, &corpus, 0, 2).unwrap();
            assert!(
                result
                    .source
                    .unwrap()
                    .ordinary_save
                    .unwrap()
                    .publications_identical
            );
        }
    }
}
