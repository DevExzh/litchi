//! Process-isolated allocator and runtime evidence for the source-backed PPTX
//! SVG picture lifecycle.  This binary is retained under the evidence tree and
//! is never a production dependency.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in profile owns a process-local allocator observer and emits JSON"
)]

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, ReadLimits, TargetMode};
use litchi_pptx::{Error, SourceBackedPresentation, SourceBackedPresentationEditor};
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error as StdError;
use std::fmt::Write as FmtWrite;
use std::hint::black_box;
use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const RASTER: &str = "/ppt/media/raster.png";
const SVG: &str = "/ppt/media/vector.svg";
const NAMESPACE_HEAVY_BINDINGS: usize = 252;
const NAMESPACE_HEAVY_DESCENDANTS: usize = 1_024;
const NAMESPACE_LIMIT_LEVELS: usize = 64;
const NAMESPACE_LIMIT_BINDINGS_PER_LEVEL: usize = 256;
const NAMESPACE_ACTIVE_LIMIT: usize = 16 * 1024;

const LANES: &[&str] = &[
    "capture_raster_small",
    "capture_raster_large",
    "capture_attached_small",
    "capture_attached_large",
    "capture_namespace_heavy",
    "inventory_many_raster_256",
    "inventory_many_raster_1024",
    "inventory_distinct_local_namespace_256",
    "inventory_distinct_local_namespace_1024",
    "attach_end_to_end_small",
    "attach_end_to_end_large",
    "detach_end_to_end_small",
    "detach_end_to_end_large",
    "noop_detach_end_to_end_small",
    "noop_detach_end_to_end_large",
    "clone_raster_small",
    "clone_raster_large",
    "clone_attached_small",
    "clone_attached_large",
    "limit_small",
    "limit_large",
    "malformed_small",
    "malformed_large",
    "namespace_limit_refusal",
];

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: each method forwards the valid allocator contract to `System`; the
// atomics only observe successful operations and never alter pointer ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair belongs to the allocator caller.
        unsafe { System.dealloc(pointer, layout) };
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            REALLOC_OLD.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            REALLOC_NEW.fetch_add(as_u64(new_size), Ordering::Relaxed);
            if new_size >= layout.size() {
                observe_growth(new_size - layout.size());
            } else {
                subtract_live(layout.size() - new_size);
            }
        }
        result
    }
}

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn observe_alloc(size: usize) {
    ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
    DIRECT_BYTES.fetch_add(as_u64(size), Ordering::Relaxed);
    observe_growth(size);
}

fn observe_growth(size: usize) {
    let size = as_u64(size);
    let live = LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    let mut old = PEAK_BYTES.load(Ordering::Relaxed);
    while live > old {
        match PEAK_BYTES.compare_exchange_weak(old, live, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => break,
            Err(observed) => old = observed,
        }
    }
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        INVALID.store(true, Ordering::Release);
    }
}

#[derive(Clone, Copy)]
struct AllocSnapshot {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live: u64,
    peak: u64,
    failed: u64,
    invalid: bool,
}

impl AllocSnapshot {
    fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: REALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
            direct: DIRECT_BYTES.load(Ordering::Acquire),
            realloc_old: REALLOC_OLD.load(Ordering::Acquire),
            realloc_new: REALLOC_NEW.load(Ordering::Acquire),
            deallocated: DEALLOC_BYTES.load(Ordering::Acquire),
            live: LIVE_BYTES.load(Ordering::Acquire),
            peak: PEAK_BYTES.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: INVALID.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> AllocDelta {
        AllocDelta {
            calls: after.calls.saturating_sub(self.calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            direct: after.direct.saturating_sub(self.direct),
            realloc_old: after.realloc_old.saturating_sub(self.realloc_old),
            realloc_new: after.realloc_new.saturating_sub(self.realloc_new),
            deallocated: after.deallocated.saturating_sub(self.deallocated),
            live_before: self.live,
            live_after: after.live,
            peak_delta: after.peak.saturating_sub(self.peak),
            failed: after.failed.saturating_sub(self.failed),
            invalid: self.invalid || after.invalid,
        }
    }
}

#[derive(Clone, Copy)]
struct AllocDelta {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live_before: u64,
    live_after: u64,
    peak_delta: u64,
    failed: u64,
    invalid: bool,
}

impl AllocDelta {
    fn requested(self) -> u64 {
        self.direct.saturating_add(self.realloc_new)
    }

    fn balanced(self) -> bool {
        self.live_before
            .checked_add(self.direct)
            .and_then(|value| value.checked_add(self.realloc_new))
            .and_then(|value| value.checked_sub(self.realloc_old))
            .and_then(|value| value.checked_sub(self.deallocated))
            == Some(self.live_after)
    }
}

fn reset_counters() {
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
    ALLOC_CALLS.store(0, Ordering::Release);
    REALLOC_CALLS.store(0, Ordering::Release);
    DEALLOC_CALLS.store(0, Ordering::Release);
    DIRECT_BYTES.store(0, Ordering::Release);
    REALLOC_OLD.store(0, Ordering::Release);
    REALLOC_NEW.store(0, Ordering::Release);
    DEALLOC_BYTES.store(0, Ordering::Release);
    ALLOC_FAILED.store(0, Ordering::Release);
    INVALID.store(false, Ordering::Release);
}

const FNV_OFFSET: u64 = 0xcbf29ce484222325;
const FNV_PRIME: u64 = 0x100000001b3;

struct BytesSource {
    bytes: Arc<[u8]>,
}

impl BytesSource {
    fn new(bytes: Arc<[u8]>) -> Self {
        Self { bytes }
    }
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
        Ok(SourceVersion::new(1, 0))
    }
}

struct Fixtures {
    small_svg: Vec<u8>,
    large_svg: Vec<u8>,
    small_raster: Vec<u8>,
    large_raster: Vec<u8>,
    raster_small: Arc<[u8]>,
    raster_large: Arc<[u8]>,
    attached_small: Arc<[u8]>,
    attached_large: Arc<[u8]>,
    namespace_heavy: Arc<[u8]>,
    namespace_limit: Arc<[u8]>,
    many_raster_256: Arc<[u8]>,
    many_raster_1024: Arc<[u8]>,
    distinct_local_namespace_256: Arc<[u8]>,
    distinct_local_namespace_1024: Arc<[u8]>,
}

impl Fixtures {
    fn load() -> Result<Self> {
        let small_svg = svg_payload(4 * 1024);
        let large_svg = svg_payload(512 * 1024);
        let small_raster = raster_payload(8 * 1024);
        let large_raster = raster_payload(512 * 1024);
        let raster_small = Arc::from(source_package(&small_svg, &small_raster, false)?);
        let raster_large = Arc::from(source_package(&large_svg, &large_raster, false)?);
        let attached_small = Arc::from(source_package(&small_svg, &small_raster, true)?);
        let attached_large = Arc::from(source_package(&large_svg, &large_raster, true)?);
        let namespace_heavy = Arc::from(source_package_with_slide(
            namespace_heavy_slide_xml(),
            &small_svg,
            &small_raster,
            false,
        )?);
        let namespace_limit = Arc::from(source_package_with_slide(
            namespace_limit_slide_xml(),
            &small_svg,
            &small_raster,
            false,
        )?);
        let many_raster_256 = Arc::from(source_package_with_slide(
            many_picture_slide_xml(256),
            &small_svg,
            &small_raster,
            false,
        )?);
        let many_raster_1024 = Arc::from(source_package_with_slide(
            many_picture_slide_xml(1024),
            &small_svg,
            &small_raster,
            false,
        )?);
        let distinct_local_namespace_256 = Arc::from(source_package_with_slide(
            distinct_local_namespace_slide_xml(256),
            &small_svg,
            &small_raster,
            false,
        )?);
        let distinct_local_namespace_1024 = Arc::from(source_package_with_slide(
            distinct_local_namespace_slide_xml(1024),
            &small_svg,
            &small_raster,
            false,
        )?);
        Ok(Self {
            small_svg,
            large_svg,
            small_raster,
            large_raster,
            raster_small,
            raster_large,
            attached_small,
            attached_large,
            namespace_heavy,
            namespace_limit,
            many_raster_256,
            many_raster_1024,
            distinct_local_namespace_256,
            distinct_local_namespace_1024,
        })
    }
}

fn svg_payload(size: usize) -> Vec<u8> {
    let prefix = br#"<svg xmlns="http://www.w3.org/2000/svg"><!--"#;
    let suffix = b"--><path d=\"M0 0 L1 1\"/></svg>";
    let body = size.saturating_sub(prefix.len() + suffix.len());
    let mut output = Vec::with_capacity(prefix.len() + body + suffix.len());
    output.extend_from_slice(prefix);
    output.extend(std::iter::repeat_n(b'x', body));
    output.extend_from_slice(suffix);
    output
}

fn raster_payload(size: usize) -> Vec<u8> {
    (0..size)
        .map(|index| (index as u8).wrapping_mul(31))
        .collect()
}

fn source_package(svg: &[u8], raster: &[u8], attached: bool) -> Result<Vec<u8>> {
    source_package_with_slide(slide_xml(attached), svg, raster, attached)
}

fn source_package_with_slide(
    slide_xml: Vec<u8>,
    svg: &[u8],
    raster: &[u8],
    attached: bool,
) -> Result<Vec<u8>> {
    let mut package = OpcPackage::new();
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/presentation.xml")?,
        ct::PML_PRESENTATION_MAIN.to_owned(),
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(SLIDE)?,
        ct::PML_SLIDE.to_owned(),
        slide_xml,
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(RASTER)?,
        "image/png".to_owned(),
        raster.to_vec(),
    )))?;
    if attached {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(SVG)?,
            "image/svg+xml".to_owned(),
            svg.to_vec(),
        )))?;
    }
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/opaque.bin")?,
        "application/octet-stream".to_owned(),
        b"opaque synthetic member".to_vec(),
    )))?;

    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml")?)?
        .rels_mut()
        .try_add_relationship(
            rt::SLIDE.to_owned(),
            "slides/slide1.xml".to_owned(),
            "rIdSlide".to_owned(),
            TargetMode::Internal,
        )?;
    let slide = package.get_part_mut(&PackURI::new(SLIDE)?)?;
    slide.rels_mut().try_add_relationship(
        rt::IMAGE.to_owned(),
        "../media/raster.png".to_owned(),
        "rIdRaster".to_owned(),
        TargetMode::Internal,
    )?;
    if attached {
        slide.rels_mut().try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/vector.svg".to_owned(),
            "rIdSvg".to_owned(),
            TargetMode::Internal,
        )?;
    }
    slide.rels_mut().try_add_relationship(
        "urn:litchi:synthetic-opaque".to_owned(),
        "../media/opaque.bin".to_owned(),
        "rIdOpaque".to_owned(),
        TargetMode::Internal,
    )?;
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn namespace_heavy_slide_xml() -> Vec<u8> {
    let mut declarations = format!(
        r#"xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:n000="urn:litchi:namespace:000""#
    );
    for index in 1..NAMESPACE_HEAVY_BINDINGS {
        let _ = write!(
            declarations,
            r#" xmlns:n{index:03}="urn:litchi:namespace:{index:03}""#
        );
    }
    let mut opaque = String::new();
    for index in 0..NAMESPACE_HEAVY_DESCENDANTS {
        let _ = write!(opaque, r#"<n000:item n000:marker="{index}"/>"#);
    }
    format!(
        r#"<p:sld {declarations}><p:cSld><p:spTree><n000:scope>{opaque}</n000:scope><p:nvGrpSpPr/><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="42" name="namespace-heavy SVG picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdRaster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
    .into_bytes()
}

fn namespace_limit_slide_xml() -> Vec<u8> {
    let mut output = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:n="urn:litchi:namespace-limit"><p:cSld><p:spTree>"#
    );
    for _ in 0..NAMESPACE_LIMIT_LEVELS {
        output.push_str("<n:layer");
        for index in 0..NAMESPACE_LIMIT_BINDINGS_PER_LEVEL {
            let _ = write!(
                output,
                r#" xmlns:l{index:03}="urn:litchi:namespace-limit:{index:03}""#
            );
        }
        output.push('>');
    }
    for _ in 0..NAMESPACE_LIMIT_LEVELS {
        output.push_str("</n:layer>");
    }
    output.push_str(
        r#"<p:nvGrpSpPr/><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="42" name="namespace-limit picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdRaster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#,
    );
    output.into_bytes()
}

fn many_picture_slide_xml(count: usize) -> Vec<u8> {
    let mut output = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>"#
    );
    for index in 0..count {
        // Keep each picture a complete direct p:pic with a distinct non-visual
        // id.  All raster payloads share rIdRaster so this lane measures the
        // picture inventory/owner lookup loop rather than media duplication.
        let id = 100u32.saturating_add(u32::try_from(index).unwrap_or(u32::MAX));
        let _ = write!(
            output,
            r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="raster picture {index}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdRaster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{index}" y="{index}"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
        );
    }
    output.push_str(r#"</p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#);
    output.into_bytes()
}

fn distinct_local_namespace_slide_xml(count: usize) -> Vec<u8> {
    let mut output = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>"#
    );
    for index in 0..count {
        // Every picture owns a distinct local binding.  The declaration is
        // intentionally unused by schema-known children: this isolates
        // namespace-context capture/indexing from relationship or payload
        // variation while still forcing the source scanner to retain it.
        let id = 100u32.saturating_add(u32::try_from(index).unwrap_or(u32::MAX));
        let _ = write!(
            output,
            r#"<p:pic xmlns:local{index:04}="urn:litchi:local:{index:04}"><p:nvPicPr><p:cNvPr id="{id}" name="local namespace raster {index}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdRaster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{index}" y="{index}"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
        );
    }
    output.push_str(r#"</p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#);
    output.into_bytes()
}

fn slide_xml(attached: bool) -> Vec<u8> {
    let svg_extension = if attached {
        format!(
            r#"<a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext><a:ext uri="{{future-extension}}"><future:keep/></a:ext></a:extLst>"#
        )
    } else {
        String::new()
    };
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:asvg="{SVG_NS}" xmlns:future="urn:litchi:future"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="42" name="profile SVG picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdRaster">{svg_extension}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
    .into_bytes()
}

fn open_editor(source: &Arc<[u8]>) -> Result<SourceBackedPresentationEditor> {
    Ok(SourceBackedPresentationEditor::from_read_at(Arc::new(
        BytesSource::new(Arc::clone(source)),
    ))?)
}

fn open_editor_with_limits(
    source: &Arc<[u8]>,
    limits: ReadLimits,
) -> Result<SourceBackedPresentationEditor> {
    Ok(SourceBackedPresentationEditor::from_read_at_with_limits(
        Arc::new(BytesSource::new(Arc::clone(source))),
        limits,
    )?)
}

fn open_presentation(source: Arc<[u8]>) -> Result<SourceBackedPresentation> {
    Ok(SourceBackedPresentation::from_read_at(Arc::new(
        BytesSource::new(source),
    ))?)
}

fn assert_reopened(
    output: &[u8],
    expected_svg: Option<&[u8]>,
    expected_raster: &[u8],
) -> Result<()> {
    let presentation = open_presentation(Arc::from(output.to_vec()))?;
    let slide = presentation
        .slide(0)
        .ok_or("reopened profile package has no slide")?;
    let image = slide.image(0)?;
    if image.target().content_type() != Some("image/png") {
        return Err("reopened raster content type changed".into());
    }
    if slide.read_image(0)?.bytes() != expected_raster {
        return Err("reopened raster payload changed".into());
    }
    match expected_svg {
        Some(expected) => {
            if slide.read_svg_image(0)?.bytes() != expected {
                return Err("reopened SVG payload changed".into());
            }
        },
        None => {
            if slide.read_svg_image(0).is_ok() {
                return Err("reopened raster-only picture unexpectedly has SVG".into());
            }
        },
    }
    let package = OpcPackage::from_bytes(output)?;
    if package
        .get_part(&PackURI::new("/ppt/media/opaque.bin")?)?
        .blob()
        != b"opaque synthetic member"
    {
        return Err("opaque media member was not preserved".into());
    }
    Ok(())
}

fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut hash = FNV_OFFSET;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(FNV_PRIME);
    }
    hash
}

fn json_string(value: &str) -> String {
    let mut output = String::with_capacity(value.len() + 2);
    output.push('"');
    for character in value.chars() {
        match character {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            character if character.is_control() => {
                let _ = write!(output, "\\u{:04x}", character as u32);
            },
            character => output.push(character),
        }
    }
    output.push('"');
    output
}

#[derive(Clone)]
struct Sample {
    elapsed_ns: u128,
    alloc: AllocDelta,
    semantic_ok: bool,
    output_exact: bool,
    expected_success: bool,
    actual_success: bool,
    sink_bytes: u64,
    sink_hash: u64,
    sink_writes: u64,
    error: Option<String>,
}

fn sample_json(
    lane: &str,
    input: &[u8],
    warmup: usize,
    samples: &[Sample],
    expected_success: bool,
) -> String {
    let mut output = String::new();
    output.push_str("{\n");
    output.push_str("  \"schema\": \"pptx-svg-lifecycle-profile-v1\",\n");
    output.push_str("  \"lane\": ");
    output.push_str(&json_string(lane));
    output.push_str(",\n");
    output.push_str(&format!("  \"input_bytes\": {},\n  \"input_hash_fnv1a64\": {},\n  \"warmup\": {warmup},\n  \"sample_count\": {},\n  \"expected_success\": {expected_success},\n  \"samples\": [\n", input.len(), hash_bytes(input), samples.len()));
    for (index, sample) in samples.iter().enumerate() {
        let comma = if index + 1 == samples.len() { "" } else { "," };
        let error = sample
            .error
            .as_deref()
            .map(json_string)
            .unwrap_or_else(|| "null".to_owned());
        output.push_str(&format!(
            "    {{\"elapsed_ns\":{},\"direct_allocated_bytes\":{},\"realloc_new_bytes\":{},\"realloc_old_bytes\":{},\"deallocated_bytes\":{},\"requested_alloc_bytes\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"alloc_calls\":{},\"realloc_calls\":{},\"dealloc_calls\":{},\"alloc_failed\":{},\"alloc_invalid\":{},\"alloc_balance_ok\":{},\"semantic_ok\":{},\"output_exact\":{},\"expected_success\":{},\"actual_success\":{},\"sink_bytes\":{},\"sink_hash\":{},\"sink_writes\":{},\"error\":{}}}{}\n",
            sample.elapsed_ns,
            sample.alloc.direct,
            sample.alloc.realloc_new,
            sample.alloc.realloc_old,
            sample.alloc.deallocated,
            sample.alloc.requested(),
            sample.alloc.live_before,
            sample.alloc.live_after,
            sample.alloc.peak_delta,
            sample.alloc.calls,
            sample.alloc.realloc_calls,
            sample.alloc.dealloc_calls,
            sample.alloc.failed,
            sample.alloc.invalid,
            sample.alloc.balanced(),
            sample.semantic_ok,
            sample.output_exact,
            sample.expected_success,
            sample.actual_success,
            sample.sink_bytes,
            sample.sink_hash,
            sample.sink_writes,
            error,
            comma,
        ));
    }
    output.push_str("  ]\n}\n");
    output
}

fn run_lane<F>(
    lane: &str,
    input: &[u8],
    warmup: usize,
    sample_count: usize,
    expected_success: bool,
    mut operation: F,
) -> Result<()>
where
    F: FnMut() -> Result<(bool, bool, u64, u64, u64)>,
{
    allocator_counter_self_test()?;
    for _ in 0..warmup {
        let warmup_result = operation();
        if expected_success {
            warmup_result?;
        } else if warmup_result.is_ok() {
            return Err(format!("lane {lane} unexpectedly accepted refusal fixture").into());
        }
    }
    let mut samples = Vec::with_capacity(sample_count);
    for _ in 0..sample_count {
        reset_counters();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let result = operation();
        let elapsed_ns = started.elapsed().as_nanos();
        let after = AllocSnapshot::now();
        let alloc = before.delta(after);
        let (actual_success, semantic_ok, output_exact, sink_bytes, sink_hash, sink_writes, error) =
            match result {
                Ok((semantic_ok, output_exact, sink_bytes, sink_hash, sink_writes)) => (
                    true,
                    semantic_ok,
                    output_exact,
                    sink_bytes,
                    sink_hash,
                    sink_writes,
                    None,
                ),
                Err(error) => (false, true, true, 0, 0, 0, Some(error.to_string())),
            };
        samples.push(Sample {
            elapsed_ns,
            alloc,
            semantic_ok,
            output_exact,
            expected_success,
            actual_success,
            sink_bytes,
            sink_hash,
            sink_writes,
            error,
        });
    }
    print!(
        "{}",
        sample_json(lane, input, warmup, &samples, expected_success)
    );
    if samples
        .iter()
        .any(|sample| sample.actual_success != expected_success)
    {
        return Err(format!("lane {lane} did not match expected status").into());
    }
    Ok(())
}

fn allocator_counter_self_test() -> Result<()> {
    reset_counters();
    let before = AllocSnapshot::now();
    let layout = Layout::from_size_align(8, std::mem::align_of::<usize>())?;
    // SAFETY: the profile deliberately exercises valid allocator operations.
    let pointer = unsafe { std::alloc::alloc(layout) };
    if pointer.is_null() {
        return Err("allocator self-test allocation failed".into());
    }
    // SAFETY: `pointer` was allocated with `layout`, and the new size is valid.
    let resized = unsafe { std::alloc::realloc(pointer, layout, 32) };
    if resized.is_null() {
        // SAFETY: failed realloc leaves the original allocation valid.
        unsafe { std::alloc::dealloc(pointer, layout) };
        return Err("allocator self-test reallocation failed".into());
    }
    let resized_layout = Layout::from_size_align(32, layout.align())?;
    // SAFETY: `resized` uses the original alignment and requested new size.
    unsafe { std::alloc::dealloc(resized, resized_layout) };
    let delta = before.delta(AllocSnapshot::now());
    if delta.calls != 1
        || delta.realloc_calls != 1
        || delta.dealloc_calls != 1
        || delta.direct != 8
        || delta.realloc_old != 8
        || delta.realloc_new != 32
        || delta.deallocated != 32
        || delta.requested() != 40
        || !delta.balanced()
        || delta.invalid
        || delta.failed != 0
    {
        return Err("allocator self-test counters did not balance".into());
    }
    reset_counters();
    Ok(())
}

fn execute_capture(source: Arc<[u8]>, attached: bool) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = open_editor(&source)?;
    let edit = editor.edit_svg_attachment(0, 0)?;
    let ok = edit.source().svg().is_some() == attached;
    black_box(edit);
    Ok((ok, true, 0, 0, 0))
}

fn execute_inventory(
    source: Arc<[u8]>,
    expected_count: usize,
) -> Result<(bool, bool, u64, u64, u64)> {
    let presentation = open_presentation(source)?;
    let slide = presentation
        .slide(0)
        .ok_or("many-picture fixture has no slide")?;
    let images = slide.images()?;
    let raster_only = images.iter().all(|image| {
        image.svg().is_none()
            && image.target().content_type() == Some("image/png")
            && image.relationship_id() == "rIdRaster"
    });
    let ok = images.len() == expected_count && raster_only;
    black_box(images);
    if !ok {
        return Err(format!(
            "many-picture inventory mismatch: expected {expected_count} raster descriptors"
        )
        .into());
    }
    Ok((true, true, expected_count as u64, expected_count as u64, 1))
}

fn execute_clone(source: Arc<[u8]>, attached: bool) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = open_editor(&source)?;
    let edit = editor.edit_svg_attachment(0, 0)?;
    let snapshot = edit.source().clone();
    let copy = snapshot.clone();
    let ok =
        (copy.svg().is_some() == attached) && copy.raster_part_uri() == snapshot.raster_part_uri();
    black_box(copy);
    Ok((ok, true, 0, 0, 0))
}

fn execute_attach_end_to_end(
    source: Arc<[u8]>,
    svg: &[u8],
    raster: &[u8],
) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = open_editor(&source)?;
    let mut edit = editor.edit_svg_attachment(0, 0)?;
    edit.attach_svg(&svg)?;
    let commit = edit.commit()?;
    let mut output = Vec::new();
    let target = editor.publish_svg_attachment_commit_to_stream(&mut output, &commit)?;
    let semantic = target
        .svg()
        .is_some_and(|attachment| attachment.bytes() == svg);
    assert_reopened(&output, Some(svg), raster)?;
    let sink_hash = hash_bytes(&output);
    Ok((
        semantic,
        !output.is_empty(),
        output.len() as u64,
        sink_hash,
        1,
    ))
}

fn execute_detach_end_to_end(
    source: Arc<[u8]>,
    svg: &[u8],
    raster: &[u8],
) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = open_editor(&source)?;
    let mut edit = editor.edit_svg_attachment(0, 0)?;
    if edit.source().svg().is_none() {
        return Err("detach fixture did not begin attached".into());
    }
    edit.detach()?;
    let commit = edit.commit()?;
    let mut output = Vec::new();
    let target = editor.publish_svg_attachment_commit_to_stream(&mut output, &commit)?;
    let semantic = target.svg().is_none() && !svg.is_empty();
    assert_reopened(&output, None, raster)?;
    let sink_hash = hash_bytes(&output);
    Ok((
        semantic,
        !output.is_empty(),
        output.len() as u64,
        sink_hash,
        1,
    ))
}

fn execute_noop_detach(source: Arc<[u8]>, raster: &[u8]) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = open_editor(&source)?;
    let mut edit = editor.edit_svg_attachment(0, 0)?;
    if edit.detach()? {
        return Err("raster-only no-op detach unexpectedly changed state".into());
    }
    let commit = edit.commit()?;
    let mut output = Vec::new();
    let target = editor.publish_svg_attachment_commit_to_stream(&mut output, &commit)?;
    assert_reopened(&output, None, raster)?;
    Ok((
        target.svg().is_none(),
        output.as_slice() == source.as_ref(),
        output.len() as u64,
        hash_bytes(&output),
        1,
    ))
}

fn execute_malformed(source: Arc<[u8]>, raster: &[u8]) -> Result<(bool, bool, u64, u64, u64)> {
    let source_part_limit = u64::try_from(raster.len().saturating_add(64 * 1024))?;
    let limits = ReadLimits::builder()
        .max_part_bytes(source_part_limit)
        .map_err(|error| format!("profile limit builder: {error}"))?
        .build()
        .map_err(|error| format!("profile limit build: {error}"))?;
    let editor = open_editor_with_limits(&source, limits)?;
    let mut edit = editor.edit_svg_attachment(0, 0)?;
    let oversized = vec![b'x'; usize::try_from(source_part_limit)?.saturating_add(1)];
    match edit.attach_svg(&oversized) {
        Err(Error::Limit { .. }) => Err("refused:limit".into()),
        Err(error) => Err(format!("unexpected:{error}").into()),
        Ok(_) => Err("accepted:oversized-svg".into()),
    }
}

fn execute_empty_payload(source: Arc<[u8]>) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = open_editor(&source)?;
    let mut edit = editor.edit_svg_attachment(0, 0)?;
    match edit.attach_svg(&[]) {
        Err(Error::Invalid(_)) => Err("refused:invalid".into()),
        Err(error) => Err(format!("unexpected:{error}").into()),
        Ok(_) => Err("accepted:empty-svg".into()),
    }
}

fn execute_namespace_limit_refusal(source: Arc<[u8]>) -> Result<(bool, bool, u64, u64, u64)> {
    let editor = match open_editor(&source) {
        Ok(editor) => editor,
        Err(error) => return Err(format!("unexpected-open:{error}").into()),
    };
    match editor.edit_svg_attachment(0, 0) {
        Err(Error::Limit { resource, limit })
            if resource == "active namespace bindings" && limit == NAMESPACE_ACTIVE_LIMIT =>
        {
            Err("refused:namespace-active-limit".into())
        },
        Err(error) => Err(format!("unexpected:{error}").into()),
        Ok(_) => Err("accepted:namespace-active-limit".into()),
    }
}

fn parse_args() -> Result<(String, usize, usize)> {
    let mut lane = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
            "--warmup" => {
                warmup = arguments
                    .next()
                    .ok_or("--warmup requires a value")?
                    .parse()?
            },
            "--samples" => {
                samples = arguments
                    .next()
                    .ok_or("--samples requires a value")?
                    .parse()?
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    let lane = lane.ok_or("--lane is required")?;
    if !LANES.contains(&lane.as_str()) {
        return Err(format!("unknown lane {lane}").into());
    }
    if samples == 0 {
        return Err("--samples must be nonzero".into());
    }
    Ok((lane, warmup, samples))
}

fn main() -> Result<()> {
    let (lane, warmup, samples) = parse_args()?;
    let fixtures = Fixtures::load()?;
    match lane.as_str() {
        "capture_raster_small" => {
            run_lane(&lane, &fixtures.raster_small, warmup, samples, true, || {
                execute_capture(Arc::clone(&fixtures.raster_small), false)
            })
        },
        "capture_raster_large" => {
            run_lane(&lane, &fixtures.raster_large, warmup, samples, true, || {
                execute_capture(Arc::clone(&fixtures.raster_large), false)
            })
        },
        "capture_attached_small" => run_lane(
            &lane,
            &fixtures.attached_small,
            warmup,
            samples,
            true,
            || execute_capture(Arc::clone(&fixtures.attached_small), true),
        ),
        "capture_attached_large" => run_lane(
            &lane,
            &fixtures.attached_large,
            warmup,
            samples,
            true,
            || execute_capture(Arc::clone(&fixtures.attached_large), true),
        ),
        "capture_namespace_heavy" => run_lane(
            &lane,
            &fixtures.namespace_heavy,
            warmup,
            samples,
            true,
            || execute_capture(Arc::clone(&fixtures.namespace_heavy), false),
        ),
        "inventory_many_raster_256" => run_lane(
            &lane,
            &fixtures.many_raster_256,
            warmup,
            samples,
            true,
            || execute_inventory(Arc::clone(&fixtures.many_raster_256), 256),
        ),
        "inventory_many_raster_1024" => run_lane(
            &lane,
            &fixtures.many_raster_1024,
            warmup,
            samples,
            true,
            || execute_inventory(Arc::clone(&fixtures.many_raster_1024), 1024),
        ),
        "inventory_distinct_local_namespace_256" => run_lane(
            &lane,
            &fixtures.distinct_local_namespace_256,
            warmup,
            samples,
            true,
            || execute_inventory(Arc::clone(&fixtures.distinct_local_namespace_256), 256),
        ),
        "inventory_distinct_local_namespace_1024" => run_lane(
            &lane,
            &fixtures.distinct_local_namespace_1024,
            warmup,
            samples,
            true,
            || execute_inventory(Arc::clone(&fixtures.distinct_local_namespace_1024), 1024),
        ),
        "attach_end_to_end_small" => {
            run_lane(&lane, &fixtures.raster_small, warmup, samples, true, || {
                execute_attach_end_to_end(
                    Arc::clone(&fixtures.raster_small),
                    &fixtures.small_svg,
                    &fixtures.small_raster,
                )
            })
        },
        "attach_end_to_end_large" => {
            run_lane(&lane, &fixtures.raster_large, warmup, samples, true, || {
                execute_attach_end_to_end(
                    Arc::clone(&fixtures.raster_large),
                    &fixtures.large_svg,
                    &fixtures.large_raster,
                )
            })
        },
        "detach_end_to_end_small" => run_lane(
            &lane,
            &fixtures.attached_small,
            warmup,
            samples,
            true,
            || {
                execute_detach_end_to_end(
                    Arc::clone(&fixtures.attached_small),
                    &fixtures.small_svg,
                    &fixtures.small_raster,
                )
            },
        ),
        "detach_end_to_end_large" => run_lane(
            &lane,
            &fixtures.attached_large,
            warmup,
            samples,
            true,
            || {
                execute_detach_end_to_end(
                    Arc::clone(&fixtures.attached_large),
                    &fixtures.large_svg,
                    &fixtures.large_raster,
                )
            },
        ),
        "noop_detach_end_to_end_small" => {
            run_lane(&lane, &fixtures.raster_small, warmup, samples, true, || {
                execute_noop_detach(Arc::clone(&fixtures.raster_small), &fixtures.small_raster)
            })
        },
        "noop_detach_end_to_end_large" => {
            run_lane(&lane, &fixtures.raster_large, warmup, samples, true, || {
                execute_noop_detach(Arc::clone(&fixtures.raster_large), &fixtures.large_raster)
            })
        },
        "clone_raster_small" => {
            run_lane(&lane, &fixtures.raster_small, warmup, samples, true, || {
                execute_clone(Arc::clone(&fixtures.raster_small), false)
            })
        },
        "clone_raster_large" => {
            run_lane(&lane, &fixtures.raster_large, warmup, samples, true, || {
                execute_clone(Arc::clone(&fixtures.raster_large), false)
            })
        },
        "clone_attached_small" => run_lane(
            &lane,
            &fixtures.attached_small,
            warmup,
            samples,
            true,
            || execute_clone(Arc::clone(&fixtures.attached_small), true),
        ),
        "clone_attached_large" => run_lane(
            &lane,
            &fixtures.attached_large,
            warmup,
            samples,
            true,
            || execute_clone(Arc::clone(&fixtures.attached_large), true),
        ),
        "limit_small" => run_lane(
            &lane,
            &fixtures.raster_small,
            warmup,
            samples,
            false,
            || execute_malformed(Arc::clone(&fixtures.raster_small), &fixtures.small_raster),
        ),
        "limit_large" => run_lane(
            &lane,
            &fixtures.raster_large,
            warmup,
            samples,
            false,
            || execute_malformed(Arc::clone(&fixtures.raster_large), &fixtures.large_raster),
        ),
        "malformed_small" => run_lane(
            &lane,
            &fixtures.raster_small,
            warmup,
            samples,
            false,
            || execute_empty_payload(Arc::clone(&fixtures.raster_small)),
        ),
        "malformed_large" => run_lane(
            &lane,
            &fixtures.raster_large,
            warmup,
            samples,
            false,
            || execute_empty_payload(Arc::clone(&fixtures.raster_large)),
        ),
        "namespace_limit_refusal" => run_lane(
            &lane,
            &fixtures.namespace_limit,
            warmup,
            samples,
            false,
            || execute_namespace_limit_refusal(Arc::clone(&fixtures.namespace_limit)),
        ),
        _ => unreachable!("lane validated above"),
    }
}
