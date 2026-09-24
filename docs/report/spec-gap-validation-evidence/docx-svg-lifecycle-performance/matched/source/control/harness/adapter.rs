//! Synthetic DOCX fixtures and public-API operations for the SVG lifecycle
//! profile.  This module intentionally depends only on committed public APIs;
//! it does not reach into the DOCX implementation or test helpers.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in evidence harness owns synthetic fixtures and JSON receipts"
)]

use std::error::Error as StdError;
use std::fmt::Write as FmtWrite;
use std::io::{self, Cursor};
use std::ops::Range;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, SourceVersion,
};
use litchi_docx::drawing::DrawingPlacement;
use litchi_docx::source_backed::{self, StorySelector};
use litchi_docx::{Error as DocxError, Package as NormalPackage, ReadLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use sha2::{Digest as _, Sha256};


pub type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

pub const LANES: &[&str] = &[
    "native_svg_capture",
    "native_floating_capture",
    "lazy_inventory_1",
    "lazy_inventory_64",
    "single_attach_1",
    "single_attach_16",
    "single_attach_64",
    "single_detach_1",
    "single_detach_16",
    "single_detach_64",
    "batch_attach_1",
    "batch_attach_16",
    "batch_attach_64",
    "batch_detach_1",
    "batch_detach_16",
    "batch_detach_64",
    "shared_svg_cleanup",
    "exact_inverse_single_1",
    "exact_inverse_batch_64",
    "large_unchanged_media_managed_cap",
    "noop_detach_64",
];

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ASVG: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const PICTURE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const OPAQUE_MARKER: &str = "docx-svg-profile-opaque-v1";
const NATIVE_INLINE: &[u8] = include_bytes!("../fixtures/svg.docx");
const NATIVE_FLOATING: &[u8] = include_bytes!("../fixtures/floating.docx");
const SMALL_SVG_BYTES: usize = 512;
const LARGE_SVG_BYTES: usize = 128 * 1024;
const SMALL_RASTER_BYTES: usize = 1 * 1024;
const LARGE_UNCHANGED_BYTES: usize = 8 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FixtureKind {
    Synthetic {
        owners: usize,
        attached: bool,
        shared_svg: bool,
        root_svg_edge: bool,
        large_unchanged: bool,
    },
    Native {
        floating: bool,
    },
}

#[derive(Clone)]
pub struct Fixture {
    pub package: Arc<[u8]>,
    pub owners: usize,
    pub input_bytes: u64,
    pub input_sha256: String,
    pub input_hash_fnv1a64: u64,
    pub native: bool,
    pub svg_payload: Arc<[u8]>,
    pub raster_payload: Arc<[u8]>,
    pub opaque_payload: Arc<[u8]>,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct PhaseTimes {
    pub capture_ns: u64,
    pub stage_ns: u64,
    pub commit_ns: u64,
    pub publish_ns: u64,
    pub reopen_ns: u64,
    pub inverse_reopen_ns: u64,
    pub inverse_ns: u64,
    pub payload_ns: u64,
    pub validation_ns: u64,
    pub readback_ns: u64,
}

#[derive(Clone, Debug)]
pub struct ErrorReceipt {
    pub class: String,
    pub message: String,
    pub typed_match: bool,
}

#[derive(Clone, Debug)]
pub struct RunResult {
    pub actual_success: bool,
    pub semantic_ok: bool,
    pub opaque_ok: bool,
    pub exact_inverse_ok: bool,
    pub lazy_media_cold_ok: bool,
    pub source_readback_physical_ok: Option<bool>,
    pub source_readback_metadata_ok: Option<bool>,
    pub output_bytes: u64,
    pub phases: PhaseTimes,
    pub error: Option<ErrorReceipt>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    LazyInventory,
    SingleAttach,
    SingleDetach,
    BatchAttach,
    BatchDetach,
    SharedCleanup,
    ExactInverse,
    ManagedLarge,
    NoopDetach,
    NativeCapture,
}

#[derive(Clone, Copy, Debug)]
struct LaneSpec {
    operation: Operation,
    owners: usize,
}

fn parse_count(lane: &str, prefix: &str) -> Option<usize> {
    lane.strip_prefix(prefix)?.parse().ok()
}

fn spec_for_lane(lane: &str) -> Option<LaneSpec> {
    if lane == "native_svg_capture" {
        return Some(LaneSpec {
            operation: Operation::NativeCapture,
            owners: 0,
        });
    }
    if lane == "native_floating_capture" {
        return Some(LaneSpec {
            operation: Operation::NativeCapture,
            owners: 0,
        });
    }
    if let Some(owners) = parse_count(lane, "lazy_inventory_") {
        return Some(LaneSpec {
            operation: Operation::LazyInventory,
            owners,
        });
    }
    if let Some(owners) = parse_count(lane, "single_attach_") {
        return Some(LaneSpec {
            operation: Operation::SingleAttach,
            owners,
        });
    }
    if let Some(owners) = parse_count(lane, "single_detach_") {
        return Some(LaneSpec {
            operation: Operation::SingleDetach,
            owners,
        });
    }
    if let Some(owners) = parse_count(lane, "batch_attach_") {
        return Some(LaneSpec {
            operation: Operation::BatchAttach,
            owners,
        });
    }
    if let Some(owners) = parse_count(lane, "batch_detach_") {
        return Some(LaneSpec {
            operation: Operation::BatchDetach,
            owners,
        });
    }
    match lane {
        "shared_svg_cleanup" => Some(LaneSpec {
            operation: Operation::SharedCleanup,
            owners: 64,
        }),
        "exact_inverse_single_1" => Some(LaneSpec {
            operation: Operation::ExactInverse,
            owners: 1,
        }),
        "exact_inverse_batch_64" => Some(LaneSpec {
            operation: Operation::ExactInverse,
            owners: 64,
        }),
        "large_unchanged_media_managed_cap" => Some(LaneSpec {
            operation: Operation::ManagedLarge,
            owners: 1,
        }),
        "noop_detach_64" => Some(LaneSpec {
            operation: Operation::NoopDetach,
            owners: 64,
        }),
        _ => None,
    }
}

#[must_use]
pub fn is_known_lane(lane: &str) -> bool {
    LANES.contains(&lane)
}

#[must_use]
pub fn expected_success(lane: &str) -> bool {
    spec_for_lane(lane).is_some() && !expected_refusal(lane)
}

fn expected_refusal(lane: &str) -> bool {
    matches!(lane, "batch_attach_64" | "exact_inverse_batch_64")
}

pub fn fixture_for_lane(lane: &str) -> Result<Fixture> {
    let spec = spec_for_lane(lane).ok_or_else(|| format!("unknown DOCX SVG lifecycle lane: {lane}"))?;
    let kind = match spec.operation {
        Operation::NativeCapture => FixtureKind::Native {
            floating: lane == "native_floating_capture",
        },
        Operation::SingleDetach | Operation::BatchDetach | Operation::SharedCleanup => {
            FixtureKind::Synthetic {
                owners: spec.owners,
                attached: true,
                shared_svg: spec.operation == Operation::SharedCleanup
                    || (spec.operation == Operation::BatchDetach && spec.owners == 64),
                root_svg_edge: false,
                large_unchanged: false,
            }
        },
        Operation::ManagedLarge => FixtureKind::Synthetic {
            owners: 1,
            attached: false,
            shared_svg: false,
            root_svg_edge: false,
            large_unchanged: true,
        },
        _ => FixtureKind::Synthetic {
            owners: spec.owners,
            attached: false,
            shared_svg: false,
            root_svg_edge: false,
            large_unchanged: false,
        },
    };
    build_fixture(kind)
}

fn build_fixture(kind: FixtureKind) -> Result<Fixture> {
    let (package, owners, native, svg, raster, opaque) =
        match kind {
            FixtureKind::Native { floating } => {
                let bytes = if floating { NATIVE_FLOATING } else { NATIVE_INLINE };
                (
                    Arc::<[u8]>::from(bytes),
                    0,
                    true,
                    Arc::<[u8]>::from([]),
                    Arc::<[u8]>::from([]),
                    Arc::<[u8]>::from([]),
                )
            }
            FixtureKind::Synthetic {
                owners,
                attached,
                shared_svg,
                root_svg_edge,
                large_unchanged,
            } => {
                let svg = Arc::<[u8]>::from(svg_payload(
                    if large_unchanged {
                        LARGE_SVG_BYTES
                    } else {
                        SMALL_SVG_BYTES
                    },
                ));
                let raster = Arc::<[u8]>::from(raster_payload(if large_unchanged {
                    SMALL_RASTER_BYTES
                } else {
                    SMALL_RASTER_BYTES
                }));
                let opaque = Arc::<[u8]>::from(if large_unchanged {
                    opaque_payload(LARGE_UNCHANGED_BYTES)
                } else {
                    opaque_payload(2048)
                });
                let bytes = synthetic_package(
                    owners,
                    attached,
                    shared_svg,
                    root_svg_edge,
                    large_unchanged,
                    &svg,
                    &raster,
                    &opaque,
                )?;
                (
                    Arc::<[u8]>::from(bytes),
                    owners,
                    false,
                    svg,
                    raster,
                    opaque,
                )
            }
        };
    let input_sha256 = sha256_hex(&package);
    let input_hash_fnv1a64 = fnv1a64(&package);
    Ok(Fixture {
        input_bytes: package.len() as u64,
        package,
        owners,
        input_sha256,
        input_hash_fnv1a64,
        native,
        svg_payload: svg,
        raster_payload: raster,
        opaque_payload: opaque,
    })
}

fn svg_payload(size: usize) -> Vec<u8> {
    let prefix = br#"<svg xmlns="http://www.w3.org/2000/svg"><path id="profile" d=""/>"#;
    let suffix = b"</svg>";
    let mut output = Vec::with_capacity(size.max(prefix.len() + suffix.len()));
    output.extend_from_slice(prefix);
    while output.len() + suffix.len() < size.max(prefix.len() + suffix.len()) {
        output.extend_from_slice(b"<!--opaque-svg-->");
    }
    output.extend_from_slice(suffix);
    output
}

fn raster_payload(size: usize) -> Vec<u8> {
    let mut output = vec![0_u8; size.max(64)];
    output[..8].copy_from_slice(&[137, 80, 78, 71, 13, 10, 26, 10]);
    output
}

fn opaque_payload(size: usize) -> Vec<u8> {
    let mut output = vec![0_u8; size];
    let marker = b"unchanged-media-opaque-profile-v1";
    output[..marker.len()].copy_from_slice(marker);
    let mut state = 0x9e37_79b9_u32;
    for byte in &mut output[marker.len()..] {
        // A deterministic noncompressible stream keeps the unchanged media
        // physically large while remaining reproducible and license-free.
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        *byte = state as u8;
    }
    output
}

fn picture(index: usize, raster_id: &str, svg_id: Option<&str>) -> String {
    let id = index + 1;
    let extension = svg_id.map_or_else(String::new, |svg_id| {
        format!(
            r#"<a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="{svg_id}"/></a:ext></a:extLst>"#
        )
    });
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="2895600" cy="2762250"/><wp:docPr id="{id}" name="picture {id}"/><a:graphic><a:graphicData uri="{PICTURE_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="picture {id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{raster_id}">{extension}</a:blip></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn document_xml(owners: usize, attached: bool, shared_svg: bool) -> Vec<u8> {
    let mut pictures = String::new();
    for index in 0..owners {
        let raster_id = if owners == 1 || shared_svg {
            "rIdRaster".to_owned()
        } else {
            format!("rIdRaster{index}")
        };
        let svg_id = if attached {
            Some(if shared_svg {
                "rIdSvg".to_owned()
            } else {
                format!("rIdSvg{index}")
            })
        } else {
            None
        };
        pictures.push_str(&picture(index, &raster_id, svg_id.as_deref()));
    }
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}" xmlns:asvg="{ASVG}"><w:body><!--{OPAQUE_MARKER}--><w:p><w:r>{pictures}</w:r></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn synthetic_package(
    owners: usize,
    attached: bool,
    shared_svg: bool,
    root_svg_edge: bool,
    large_unchanged: bool,
    svg: &[u8],
    raster: &[u8],
    opaque: &[u8],
) -> Result<Vec<u8>> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml")?,
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document_xml(owners, attached, shared_svg),
    );
    for index in 0..owners {
        let raster_id = if owners == 1 || shared_svg {
            "rIdRaster".to_owned()
        } else {
            format!("rIdRaster{index}")
        };
        let target = if owners == 1 || shared_svg {
            "media/image1.png".to_owned()
        } else {
            format!("media/image{}.png", index + 1)
        };
        if !shared_svg || index == 0 {
            main.rels_mut().try_add_relationship(
                rt::IMAGE.to_owned(),
                target,
                raster_id,
                TargetMode::Internal,
            )?;
        }
        if attached && (!shared_svg || index == 0) {
            let svg_id = if shared_svg {
                "rIdSvg".to_owned()
            } else {
                format!("rIdSvg{index}")
            };
            let svg_target = if shared_svg {
                "media/image2.svg".to_owned()
            } else {
                format!("media/image{}.svg", index + 2)
            };
            main.rels_mut().try_add_relationship(
                rt::IMAGE.to_owned(),
                svg_target,
                svg_id,
                TargetMode::Internal,
            )?;
        }
    }
    package.try_add_part(Box::new(main))?;
    for index in 0..owners {
        if owners == 1 || shared_svg {
            if index > 0 {
                continue;
            }
        }
        let name = if owners == 1 || shared_svg {
            "/word/media/image1.png".to_owned()
        } else {
            format!("/word/media/image{}.png", index + 1)
        };
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(name)?,
            ct::PNG.to_owned(),
            raster.to_vec(),
        )))?;
    }
    if attached {
        for index in 0..owners {
            if shared_svg && index > 0 {
                continue;
            }
            let name = if shared_svg {
                "/word/media/image2.svg".to_owned()
            } else {
                format!("/word/media/image{}.svg", index + 2)
            };
            package.try_add_part(Box::new(BlobPart::new(
                PackURI::new(name)?,
                "image/svg+xml".to_owned(),
                svg.to_vec(),
            )))?;
        }
    }
    if large_unchanged || !opaque.is_empty() {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/media/unchanged.bin")?,
            "application/octet-stream".to_owned(),
            opaque.to_vec(),
        )))?;
    }
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    if root_svg_edge && attached {
        package.relate_to("word/media/image2.svg", rt::IMAGE);
    }
    Ok(PackageWriter::to_bytes(&package)?)
}

#[derive(Clone)]
struct BytesSource {
    bytes: Arc<[u8]>,
}

impl ReadAt for BytesSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0xD0C5, 0))
    }
}

#[derive(Clone)]
struct ObservedSource {
    bytes: Arc<[u8]>,
    ranges: Arc<[(Range<usize>, Arc<AtomicUsize>)]>,
}

impl ObservedSource {
    fn new(bytes: Arc<[u8]>, ranges: Vec<(Range<usize>, Arc<AtomicUsize>)>) -> Self {
        Self {
            bytes,
            ranges: ranges.into(),
        }
    }

    fn reset(&self) {
        for (_, counter) in self.ranges.iter() {
            counter.store(0, Ordering::Release);
        }
    }

    fn reads(&self, index: usize) -> usize {
        self.ranges[index].1.load(Ordering::Acquire)
    }
}

impl ReadAt for ObservedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset
            .checked_add(output.len())
            .unwrap_or(self.bytes.len())
            .min(self.bytes.len());
        for (range, counter) in self.ranges.iter() {
            if offset < range.end && range.start < end {
                counter.fetch_add(1, Ordering::Relaxed);
            }
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0xD0C5, 0))
    }
}

pub fn run_once(lane: &str, fixture: &Fixture) -> Result<RunResult> {
    let spec = spec_for_lane(lane).ok_or_else(|| format!("unknown lane: {lane}"))?;
    if expected_refusal(lane) {
        return run_expected_topology_refusal(fixture, spec.owners);
    }
    match spec.operation {
        Operation::NativeCapture => run_native_capture(lane, fixture),
        Operation::LazyInventory => run_lazy_inventory(fixture),
        Operation::SingleAttach => run_single_attach(fixture, spec.owners),
        Operation::SingleDetach => run_single_detach(fixture, spec.owners),
        Operation::BatchAttach => run_batch_attach(fixture, spec.owners),
        Operation::BatchDetach => run_batch_detach(fixture, spec.owners),
        Operation::SharedCleanup => run_shared_cleanup(fixture),
        Operation::ExactInverse => run_exact_inverse(fixture),
        Operation::ManagedLarge => run_managed_large(fixture),
        Operation::NoopDetach => run_noop_detach(fixture),
    }
}

fn success(
    phases: PhaseTimes,
    output_bytes: u64,
    semantic_ok: bool,
    opaque_ok: bool,
    exact_inverse_ok: bool,
    lazy_media_cold_ok: bool,
) -> RunResult {
    let actual_success = semantic_ok && opaque_ok && exact_inverse_ok && lazy_media_cold_ok;
    RunResult {
        actual_success,
        semantic_ok,
        opaque_ok,
        exact_inverse_ok,
        lazy_media_cold_ok,
        source_readback_physical_ok: None,
        source_readback_metadata_ok: None,
        output_bytes,
        phases,
        error: None,
    }
}

fn elapsed(start: Instant) -> u64 {
    start.elapsed().as_nanos().try_into().unwrap_or(u64::MAX)
}

fn batch_snapshot_metadata_matches(
    expected: &source_backed::SourceBackedSvgAttachmentBatchSnapshot,
    actual: &source_backed::SourceBackedSvgAttachmentBatchSnapshot,
) -> bool {
    expected.len() == actual.len()
        && expected.story_xml() == actual.story_xml()
        && selectors(expected.len()).iter().all(|selector| {
            expected
                .picture(*selector)
                .zip(actual.picture(*selector))
                .is_some_and(|(left, right)| snapshot_metadata_matches(&left, &right))
        })
}

fn snapshot_metadata_matches(
    expected: &source_backed::SourceBackedSvgAttachmentSnapshot,
    actual: &source_backed::SourceBackedSvgAttachmentSnapshot,
) -> bool {
    let raster_matches = expected.selector() == actual.selector()
        && expected.placement() == actual.placement()
        && expected.owner_state() == actual.owner_state()
        && expected.raster_relationship_id() == actual.raster_relationship_id()
        && expected.raster_relationship_type() == actual.raster_relationship_type()
        && expected.raster_part_uri().as_str() == actual.raster_part_uri().as_str()
        && expected.raster().content_type() == actual.raster().content_type()
        && expected.raster().bytes() == actual.raster().bytes();
    let svg_matches = match (expected.svg(), actual.svg()) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            left.relationship_id() == right.relationship_id()
                && left.relationship_type() == right.relationship_type()
                && left.part_uri().as_str() == right.part_uri().as_str()
                && left.bytes() == right.bytes()
        },
        _ => false,
    };
    raster_matches && svg_matches
}

fn views_match_snapshot(
    views: &[source_backed::SourceSvgPictureView],
    expected: &source_backed::SourceBackedSvgAttachmentBatchSnapshot,
) -> bool {
    views.len() == expected.len()
        && selectors(expected.len()).iter().all(|selector| {
            let Some(snapshot) = expected.picture(*selector) else {
                return false;
            };
            views
                .iter()
                .find(|view| view.selector() == *selector)
                .is_some_and(|view| {
                    let raster_matches = view.placement() == snapshot.placement()
                        && view.owner_state() == snapshot.owner_state()
                        && view.raster_relationship_id() == snapshot.raster_relationship_id()
                        && view.raster_part_uri().as_str() == snapshot.raster_part_uri().as_str()
                        && view.raster_content_type() == snapshot.raster().content_type()
                        && view.raster().relationship_type()
                            == snapshot.raster_relationship_type()
                        && view.raster().bytes() == snapshot.raster().bytes();
                    let svg_matches = match (view.svg(), snapshot.svg()) {
                        (None, None) => true,
                        (Some(left), Some(right)) => {
                            left.relationship_id() == right.relationship_id()
                                && left.relationship_type() == right.relationship_type()
                                && left.part_uri().as_str() == right.part_uri().as_str()
                                && left.bytes() == right.bytes()
                        },
                        _ => false,
                    };
                    raster_matches && svg_matches
                })
        })
}

fn run_expected_topology_refusal(fixture: &Fixture, owners: usize) -> Result<RunResult> {
    let capture = Instant::now();
    let package = package_from_fixture(fixture)?;
    let mut edit = package.edit_svg_attachments(selectors(owners))?;
    let capture_ns = elapsed(capture);
    let source_snapshot = edit.source().clone();
    let stage = Instant::now();
    let mut stage_error: Option<DocxError> = None;
    for selector in selectors(owners) {
        if let Err(error) = edit.attach_svg(selector, fixture.svg_payload.as_ref()) {
            stage_error = Some(error);
            break;
        }
    }
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit_result: std::result::Result<(), DocxError> = stage_error
        .map(Err)
        .unwrap_or_else(|| edit.commit().map(|_| ()));
    let commit_ns = elapsed(commit_start);
    let (typed_refusal, class, message, unexpected_success) = match commit_result {
        Ok(()) => (
            false,
            "unexpected_success".to_owned(),
            "64-owner batch closure unexpectedly committed".to_owned(),
            true,
        ),
        Err(error) => {
            let typed_overlay = matches!(
                &error,
                DocxError::Opc(OpcError::SourceBackedOverlayUnavailable { .. })
            );
            let class = if typed_overlay {
                "topology_part_bound"
            } else {
                "unexpected_error"
            };
            (typed_overlay, class.to_owned(), error.to_string(), false)
        }
    };
    let readback = Instant::now();
    let post_failure = package.edit_svg_attachments(selectors(owners))?;
    let metadata_unchanged = batch_snapshot_metadata_matches(&source_snapshot, post_failure.source());
    let no_op_commit = post_failure.commit()?;
    let mut readback_bytes = Vec::new();
    package.publish_svg_attachment_batch_commit_to_stream(&mut readback_bytes, &no_op_commit)?;
    let physical_unchanged = readback_bytes.as_slice() == fixture.package.as_ref();
    let reopened = source_backed::Package::from_reader(Cursor::new(readback_bytes.as_slice()))?;
    let readback_views = reopened.svg_pictures(StorySelector::Main)?;
    let metadata_reopened = views_match_snapshot(&readback_views, &source_snapshot);
    let readback_ns = elapsed(readback);
    let validation = Instant::now();
    let opaque_ok = check_opaque(&readback_bytes, fixture);
    let validation_ns = elapsed(validation);
    let source_unchanged = physical_unchanged && metadata_unchanged && metadata_reopened;
    Ok(RunResult {
        actual_success: false,
        semantic_ok: source_unchanged && typed_refusal && !unexpected_success,
        opaque_ok,
        exact_inverse_ok: source_unchanged && typed_refusal && !unexpected_success,
        lazy_media_cold_ok: true,
        source_readback_physical_ok: Some(physical_unchanged),
        source_readback_metadata_ok: Some(metadata_unchanged && metadata_reopened),
        output_bytes: 0,
        phases: PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            validation_ns,
            readback_ns,
            ..PhaseTimes::default()
        },
        error: Some(ErrorReceipt {
            class,
            message,
            typed_match: typed_refusal,
        }),
    })
}

fn run_native_capture(lane: &str, fixture: &Fixture) -> Result<RunResult> {
    let capture = Instant::now();
    let package = source_backed::Package::from_reader(Cursor::new(fixture.package.as_ref()))?;
    let views = package.svg_picture_sources(StorySelector::Main)?;
    let capture_ns = elapsed(capture);
    let validation = Instant::now();
    let semantic_ok = !views.is_empty()
        && if lane == "native_svg_capture" {
            views.len() == 1
                && matches!(views[0].placement(), DrawingPlacement::Inline)
                && views[0].svg().is_some()
        } else {
            views.len() == 1
                && matches!(views[0].placement(), DrawingPlacement::Floating)
        };
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        fixture.input_bytes,
        semantic_ok,
        true,
        true,
        true,
    ))
}

fn source_media_ranges(bytes: Arc<[u8]>, names: &[&str]) -> Result<Vec<(Range<usize>, Arc<AtomicUsize>)>> {
    names
        .iter()
        .map(|name| {
            Ok((
                payload_range(&bytes, name).ok_or_else(|| format!("ZIP member not found: {name}"))?,
                Arc::new(AtomicUsize::new(0)),
            ))
        })
        .collect()
}

fn run_lazy_inventory(fixture: &Fixture) -> Result<RunResult> {
    let ranges = source_media_ranges(
        Arc::clone(&fixture.package),
        &["word/media/image1.png", "word/media/unchanged.bin"],
    )?;
    let source = ObservedSource::new(Arc::clone(&fixture.package), ranges);
    let package = source_backed::Package::from_read_at(Arc::new(source.clone()))?;
    source.reset();
    let capture = Instant::now();
    let views = package.svg_picture_sources(StorySelector::Main)?;
    let capture_ns = elapsed(capture);
    let cold_before_data = source.reads(0) == 0 && source.reads(1) == 0;
    let payload = Instant::now();
    let raster = views
        .first()
        .ok_or("lazy inventory returned no pictures")?
        .raster()
        .data()?;
    let payload_ns = elapsed(payload);
    let selected_read = source.reads(0) > 0;
    let unrelated_cold = source.reads(1) == 0;
    let validation = Instant::now();
    let semantic_ok = views.len() == fixture.owners && raster.as_bytes() == fixture.raster_payload.as_ref();
    let opaque_ok = check_opaque(&fixture.package, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            payload_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        fixture.input_bytes,
        semantic_ok,
        opaque_ok,
        true,
        cold_before_data && selected_read && unrelated_cold,
    ))
}

fn package_from_fixture(fixture: &Fixture) -> Result<source_backed::Package> {
    Ok(source_backed::Package::from_read_at(Arc::new(BytesSource {
        bytes: Arc::clone(&fixture.package),
    }))?)
}

fn selectors(count: usize) -> Vec<source_backed::PictureSelector> {
    (0..count)
        .map(|index| source_backed::PictureSelector::new(0, index))
        .collect()
}

fn run_single_attach(fixture: &Fixture, owners: usize) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachment(source_backed::PictureSelector::new(0, owners - 1))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    let changed = edit.attach_svg(fixture.svg_payload.as_ref())?;
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    package.publish_svg_attachment_commit_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let selected = views.get(owners - 1).ok_or("selected picture missing")?;
    let validation = Instant::now();
    let semantic_ok = changed
        && views.len() == owners
        && selected.svg().is_some_and(|svg| svg.bytes() == fixture.svg_payload.as_ref())
        && views
            .iter()
            .all(|view| view.raster().bytes() == fixture.raster_payload.as_ref());
    let opaque_ok = check_opaque(&output, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        true,
        true,
    ))
}

fn run_single_detach(fixture: &Fixture, owners: usize) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachment(source_backed::PictureSelector::new(0, owners - 1))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    let changed = edit.detach_svg()?;
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    package.publish_svg_attachment_commit_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let validation = Instant::now();
    let semantic_ok = changed
        && views.len() == owners
        && views[owners - 1].svg().is_none()
        && views[..owners - 1].iter().all(|view| view.svg().is_some())
        && views
            .iter()
            .all(|view| view.raster().bytes() == fixture.raster_payload.as_ref());
    let opaque_ok = check_opaque(&output, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        true,
        true,
    ))
}

fn run_batch_attach(fixture: &Fixture, owners: usize) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachments(selectors(owners))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    for selector in selectors(owners) {
        edit.attach_svg(selector, fixture.svg_payload.as_ref())?;
    }
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    package.publish_svg_attachment_batch_commit_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let validation = Instant::now();
    let semantic_ok = views.len() == owners
        && views.iter().all(|view| {
            view.svg().is_some_and(|svg| svg.bytes() == fixture.svg_payload.as_ref())
                && view.raster().bytes() == fixture.raster_payload.as_ref()
        });
    let opaque_ok = check_opaque(&output, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        true,
        true,
    ))
}

fn run_batch_detach(fixture: &Fixture, owners: usize) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachments(selectors(owners))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    for selector in selectors(owners) {
        edit.detach_svg(selector)?;
    }
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    package.publish_svg_attachment_batch_commit_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let validation = Instant::now();
    let semantic_ok = views.len() == owners
        && views.iter().all(|view| {
            view.svg().is_none() && view.raster().bytes() == fixture.raster_payload.as_ref()
        });
    let opaque_ok = check_opaque(&output, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        true,
        true,
    ))
}

fn run_shared_cleanup(fixture: &Fixture) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachments(selectors(fixture.owners))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    for selector in selectors(fixture.owners) {
        edit.detach_svg(selector)?;
    }
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    package.publish_svg_attachment_batch_commit_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let validation = Instant::now();
    let semantic_ok = views.iter().all(|view| view.svg().is_none())
        && !has_part(&output, "/word/media/image2.svg")
        && views
            .iter()
            .all(|view| view.raster().bytes() == fixture.raster_payload.as_ref());
    let opaque_ok = check_opaque(&output, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        true,
        true,
    ))
}

fn run_exact_inverse(fixture: &Fixture) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachments(selectors(fixture.owners))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    for selector in selectors(fixture.owners) {
        edit.attach_svg(selector, fixture.svg_payload.as_ref())?;
    }
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    let publication = package
        .publish_svg_attachment_batch_commit_with_publication_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let validation_pre = Instant::now();
    let semantic_ok = views.len() == fixture.owners
        && views.iter().all(|view| view.svg().is_some());
    let validation_pre_ns = elapsed(validation_pre);
    let mut restored = Vec::new();
    let inverse = Instant::now();
    let inverse_snapshot = reopened
        .publish_svg_attachment_batch_inverse_to_stream(&mut restored, &publication)?;
    let inverse_ns = elapsed(inverse);
    let validation_mid = Instant::now();
    let inverse_semantic_ok = selectors(fixture.owners).iter().all(|selector| {
        inverse_snapshot
            .picture(*selector)
            .is_some_and(|picture| picture.svg().is_none())
    });
    let physical_inverse_ok = restored.as_slice() == fixture.package.as_ref();
    let validation_mid_ns = elapsed(validation_mid);
    let inverse_reopen = Instant::now();
    let restored_reopened = source_backed::Package::from_reader(Cursor::new(restored.as_slice()))?;
    let restored_views = restored_reopened.svg_pictures(StorySelector::Main)?;
    let inverse_reopen_ns = elapsed(inverse_reopen);
    let validation_post = Instant::now();
    let restored_semantic_ok = restored_views.iter().all(|view| view.svg().is_none());
    let exact_inverse_ok = physical_inverse_ok
        && inverse_semantic_ok
        && restored_semantic_ok;
    let output_opaque_ok = check_opaque(&output, fixture);
    let restored_opaque_ok = check_opaque(&restored, fixture);
    let validation_post_ns = elapsed(validation_post);
    let opaque_ok = output_opaque_ok && restored_opaque_ok;
    let validation_ns = validation_pre_ns + validation_mid_ns + validation_post_ns;
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            inverse_ns,
            inverse_reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        exact_inverse_ok,
        true,
    ))
}

fn run_managed_large(fixture: &Fixture) -> Result<RunResult> {
    let ranges = source_media_ranges(
        Arc::clone(&fixture.package),
        &["word/media/image1.png", "word/media/unchanged.bin"],
    )?;
    let source = ObservedSource::new(Arc::clone(&fixture.package), ranges);
    let budget = Budget::root(
        "docx-svg-lifecycle-profile",
        Limits::new(2 * 1024 * 1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(2 * 1024 * 1024).ok_or("invalid managed memory cap")?,
        0,
    )?;
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(source.clone()),
        ReadLimits::default(),
        ExecutionContext::new(budget, cancellation, execution_limits),
    )?;
    source.reset();
    let capture = Instant::now();
    let views = package.svg_picture_sources(StorySelector::Main)?;
    let capture_ns = elapsed(capture);
    let validation = Instant::now();
    let semantic_ok = views.len() == 1 && views[0].raster().part_uri().as_str() == "/word/media/image1.png";
    let lazy_media_cold_ok = source.reads(1) == 0;
    let opaque_ok = has_member_declared_size(
        &fixture.package,
        "word/media/unchanged.bin",
        fixture.opaque_payload.len(),
    );
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        fixture.input_bytes,
        semantic_ok,
        opaque_ok,
        true,
        lazy_media_cold_ok,
    ))
}

fn run_noop_detach(fixture: &Fixture) -> Result<RunResult> {
    let package = package_from_fixture(fixture)?;
    let capture = Instant::now();
    let mut edit = package.edit_svg_attachment(source_backed::PictureSelector::new(0, fixture.owners - 1))?;
    let capture_ns = elapsed(capture);
    let stage = Instant::now();
    let changed = edit.detach_svg()?;
    let stage_ns = elapsed(stage);
    let commit_start = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed(commit_start);
    let mut output = Vec::new();
    let publish = Instant::now();
    package.publish_svg_attachment_commit_to_stream(&mut output, &commit)?;
    let publish_ns = elapsed(publish);
    let reopen = Instant::now();
    let reopened = source_backed::Package::from_reader(Cursor::new(output.as_slice()))?;
    let views = reopened.svg_pictures(StorySelector::Main)?;
    let reopen_ns = elapsed(reopen);
    let validation = Instant::now();
    let semantic_ok = !changed
        && views.len() == fixture.owners
        && views.iter().all(|view| view.svg().is_none());
    let exact_inverse_ok = output.as_slice() == fixture.package.as_ref();
    let opaque_ok = check_opaque(&output, fixture);
    let validation_ns = elapsed(validation);
    Ok(success(
        PhaseTimes {
            capture_ns,
            stage_ns,
            commit_ns,
            publish_ns,
            reopen_ns,
            validation_ns,
            ..PhaseTimes::default()
        },
        output.len() as u64,
        semantic_ok,
        opaque_ok,
        exact_inverse_ok,
        true,
    ))
}

fn check_opaque(bytes: &[u8], fixture: &Fixture) -> bool {
    if fixture.native {
        return true;
    }
    let Ok(package) = NormalPackage::from_reader(Cursor::new(bytes)) else {
        return false;
    };
    let Ok(document) = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").expect("static URI"))
    else {
        return false;
    };
    let marker_ok = document
        .blob()
        .windows(OPAQUE_MARKER.len())
        .any(|window| window == OPAQUE_MARKER.as_bytes());
    let Ok(opaque) = package
        .opc_package()
        .get_part(&PackURI::new("/word/media/unchanged.bin").expect("static URI"))
    else {
        return false;
    };
    marker_ok && opaque.blob() == fixture.opaque_payload.as_ref()
}

fn has_part(bytes: &[u8], name: &str) -> bool {
    let Ok(package) = NormalPackage::from_reader(Cursor::new(bytes)) else {
        return false;
    };
    let Ok(uri) = PackURI::new(name) else {
        return false;
    };
    package.opc_package().get_part(&uri).is_ok()
}

fn has_member_declared_size(bytes: &[u8], name: &str, expected: usize) -> bool {
    let name = name.as_bytes();
    for (offset, signature) in bytes.windows(4).enumerate() {
        if signature != b"PK\x01\x02" || offset + 46 > bytes.len() {
            continue;
        }
        let size = u32::from_le_bytes(bytes[offset + 24..offset + 28].try_into().unwrap_or([0; 4]));
        let name_len = u16::from_le_bytes(bytes[offset + 28..offset + 30].try_into().unwrap_or([0; 2])) as usize;
        if offset + 46 + name_len <= bytes.len()
            && &bytes[offset + 46..offset + 46 + name_len] == name
        {
            return size as usize == expected;
        }
    }
    false
}

fn payload_range(zip: &[u8], name: &str) -> Option<Range<usize>> {
    let name = name.as_bytes();
    for (offset, signature) in zip.windows(4).enumerate() {
        if signature != b"PK\x01\x02" || offset + 46 > zip.len() {
            continue;
        }
        let compressed =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().ok()?) as usize;
        let name_len =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().ok()?) as usize;
        if offset + 46 + name_len > zip.len()
            || &zip[offset + 46..offset + 46 + name_len] != name
        {
            continue;
        }
        let local =
            u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().ok()?) as usize;
        if local + 30 > zip.len() {
            return None;
        }
        let local_name =
            u16::from_le_bytes(zip[local + 26..local + 28].try_into().ok()?) as usize;
        let local_extra =
            u16::from_le_bytes(zip[local + 28..local + 30].try_into().ok()?) as usize;
        let start = local.checked_add(30 + local_name + local_extra)?;
        return Some(start..start.checked_add(compressed)?);
    }
    None
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest.finalize() {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn fnv1a64(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}
