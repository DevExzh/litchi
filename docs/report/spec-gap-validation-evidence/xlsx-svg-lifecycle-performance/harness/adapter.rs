//! XLSX profile fixture generation, public-API calls, and receipt encoding.
//!
//! The fixture builders are adapted from
//! `crates/litchi-xlsx/tests/drawing_svg_lifecycle.rs`. They remain synthetic,
//! bounded inputs and are kept here so the profile does not depend on test
//! implementation details. The timed paths call only the public workbook,
//! drawing scanner, selector, and worksheet transaction APIs.

use std::collections::{HashMap, HashSet};
use std::error::Error as StdError;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::hint::black_box;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, ReadLimits, TargetMode};
use litchi_xlsx::drawing::{
    DrawingAnchor, PictureSelector, PictureSource, SourceDrawing, SvgInput, SvgOwnerState,
};
use litchi_xlsx::{Error as XlsxError, Package, Workbook};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::support::{self, AllocDelta, AllocSnapshot};

pub type BoxError = Box<dyn StdError + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

pub const LANES: &[&str] = &[
    "capture_native_fixture",
    "capture_raster_two_cell_small",
    "capture_raster_two_cell_large",
    "capture_raster_one_cell_small",
    "capture_raster_one_cell_large",
    "capture_raster_absolute_small",
    "capture_raster_absolute_large",
    "capture_attached_two_cell_small",
    "capture_attached_two_cell_large",
    "capture_attached_one_cell_small",
    "capture_attached_one_cell_large",
    "capture_attached_absolute_small",
    "capture_attached_absolute_large",
    "clone_raster_small",
    "clone_raster_large",
    "clone_attached_small",
    "clone_attached_large",
    "clone_captured_owner_small",
    "clone_captured_owner_large",
    "inventory_shared_256",
    "inventory_shared_1024",
    "inventory_distinct_256",
    "inventory_distinct_1024",
    "namespace_heavy",
    "namespace_limit_refusal",
    "attach_end_to_end_two_cell_small",
    "attach_end_to_end_two_cell_large",
    "attach_end_to_end_one_cell_small",
    "attach_end_to_end_one_cell_large",
    "attach_end_to_end_absolute_small",
    "attach_end_to_end_absolute_large",
    "strict_attach_end_to_end_two_cell_small",
    "inverse_attach_detach_two_cell_small",
    "inverse_attach_detach_two_cell_large",
    "inverse_attach_detach_one_cell_small",
    "inverse_attach_detach_one_cell_large",
    "inverse_attach_detach_absolute_small",
    "inverse_attach_detach_absolute_large",
    "detach_end_to_end_shared_first_two_cell",
    "detach_end_to_end_shared_first_one_cell",
    "detach_end_to_end_shared_first_absolute",
    "detach_end_to_end_shared_final_two_cell",
    "detach_end_to_end_shared_final_one_cell",
    "detach_end_to_end_shared_final_absolute",
    "strict_detach_end_to_end_shared_final_two_cell",
    "incoming_edge_shared_final_two_cell",
    "detach_end_to_end_distinct_two_cell_small",
    "detach_end_to_end_distinct_two_cell_large",
    "detach_end_to_end_distinct_one_cell_small",
    "detach_end_to_end_distinct_one_cell_large",
    "detach_end_to_end_distinct_absolute_small",
    "detach_end_to_end_distinct_absolute_large",
    "same_picture_attach_detach_two_cell",
    "same_picture_attach_detach_one_cell",
    "same_picture_attach_detach_absolute",
    "multisheet_attach_detach",
    "noop_detach_two_cell",
    "noop_detach_one_cell",
    "noop_detach_absolute",
    "limit_small",
    "limit_large",
    "mixed_caps_rejection",
    "malformed_duplicate_owner",
    "malformed_mce_owner",
    "malformed_linked_owner",
    "malformed_unknown_uri",
    "multi_picture_same_drawing_16",
    "multi_picture_same_drawing_64",
    "multi_picture_same_drawing_256",
];

/// Bounded optimization probes kept outside the freeze-gated acceptance set.
/// Their source and binary manifests live in a separate exploratory bundle.
pub const EXPLORATORY_LANES: &[&str] = &[
    "multi_picture_same_drawing_detach_shared_16",
    "multi_picture_same_drawing_detach_shared_64",
    "multi_picture_same_drawing_detach_shared_256",
    "multi_picture_same_drawing_detach_distinct_16",
    "multi_picture_same_drawing_detach_distinct_64",
    "multi_picture_same_drawing_detach_distinct_256",
    "inventory_shared_root_namespace_32",
];

pub fn is_known_lane(lane: &str) -> bool {
    LANES.contains(&lane) || EXPLORATORY_LANES.contains(&lane)
}

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SHEET2: &str = "/xl/worksheets/sheet2.xml";
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const DRAWING2: &str = "/xl/drawings/drawing2.xml";
const RASTER: &str = "/xl/media/image1.png";
const RASTER2: &str = "/xl/media/image3.png";
const OTHER: &str = "/xl/opaque-owner.xml";
const SMALL_PAYLOAD: usize = 512;
const LARGE_PAYLOAD: usize = 65_536;
const LARGE_RASTER: usize = 65_536;
const SVG_INPUT_LIMIT: usize = 32 * 1024 * 1024;
// drawing_xml adds x, a, r, asvg, mc, future, and n at the root before a
// namespace-pressure fragment contributes its generated declarations.
const FIXED_ROOT_NAMESPACE_BINDINGS: usize = 7;
const NATIVE_FIXTURE_SHA256: &str =
    "0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SvgVariant {
    None,
    Embedded,
    Duplicate,
    Mce,
    Linked,
    Unknown,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TargetTopology {
    None,
    Shared,
    Distinct,
}

#[derive(Clone, Copy, Debug)]
struct FixtureSpec {
    strict: bool,
    picture_count: usize,
    svg: SvgVariant,
    targets: TargetTopology,
    svg_picture: Option<usize>,
    incoming_svg_edge: bool,
    namespace_heavy: bool,
    namespace_root_heavy: bool,
    namespace_limit: bool,
    namespace_bindings: usize,
    opaque_extension: bool,
    payload_size: usize,
    raster_size: usize,
}

impl Default for FixtureSpec {
    fn default() -> Self {
        Self {
            strict: false,
            picture_count: 3,
            svg: SvgVariant::None,
            targets: TargetTopology::None,
            svg_picture: None,
            incoming_svg_edge: false,
            namespace_heavy: false,
            namespace_root_heavy: false,
            namespace_limit: false,
            namespace_bindings: 0,
            opaque_extension: true,
            payload_size: SMALL_PAYLOAD,
            raster_size: 1_024,
        }
    }
}

pub struct Fixture {
    pub package: Arc<[u8]>,
    pub payload: Arc<[u8]>,
    pub picture_count: usize,
    pub drawing_count: usize,
    pub input_bytes: u64,
    pub input_hash: u64,
    pub input_sha256: String,
    pub expected_success: bool,
    pub picture: usize,
    pub strict: bool,
    pub incoming_svg_edge: bool,
    pub opaque_extension: bool,
    pub opaque_descendants: bool,
    pub namespace_generated_bindings: Option<usize>,
    pub namespace_active_bindings: Option<usize>,
    pub namespace_active_limit: Option<usize>,
    pub multisheet: bool,
    pub caller_limit_profile: Option<&'static str>,
    pub caller_limit_ceilings: Option<Vec<CallerLimitCeiling>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CallerLimitCeiling {
    pub name: &'static str,
    pub value: u64,
}

#[derive(Clone, Debug)]
struct Execution {
    output: Option<Vec<u8>>,
    semantic_ok: bool,
    output_exact: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RefusalKind {
    NamespaceLimit,
    CallerLimit,
    DuplicateOwner,
    MceAncestry,
    LinkedOwner,
    MixedLimit,
}

impl RefusalKind {
    const fn as_str(self) -> &'static str {
        match self {
            Self::NamespaceLimit => "namespace_limit",
            Self::CallerLimit => "caller_limit",
            Self::DuplicateOwner => "duplicate_owner",
            Self::MceAncestry => "mce_ancestry",
            Self::LinkedOwner => "linked_owner",
            Self::MixedLimit => "mixed_limit",
        }
    }
}

#[derive(Clone, Debug)]
struct ExpectedRefusal {
    kind: RefusalKind,
    message: String,
}

#[derive(Clone, Debug)]
enum LaneExecution {
    Success(Execution),
    ExpectedRefusal(ExpectedRefusal),
}

#[derive(Clone, Debug)]
struct ErrorInfo {
    class: String,
    message: String,
}

enum Outcome {
    Success {
        semantic_ok: bool,
        output_exact: bool,
    },
    ExpectedRefusal(ExpectedRefusal),
    Failure(BoxError),
}

pub fn fixture_for_lane(lane: &str) -> Result<Fixture> {
    if lane == "capture_native_fixture" {
        return native_fixture();
    }
    let mut spec = FixtureSpec::default();
    let picture = picture_index(lane);
    let multisheet = lane == "multisheet_attach_detach";
    if lane.starts_with("strict_") {
        spec.strict = true;
    }
    let attached = lane.contains("attached")
        || lane.starts_with("detach_end_to_end")
        || lane.starts_with("inventory_");
    spec.svg = if attached {
        SvgVariant::Embedded
    } else {
        SvgVariant::None
    };
    spec.targets = if lane.contains("distinct") {
        TargetTopology::Distinct
    } else if attached {
        TargetTopology::Shared
    } else {
        TargetTopology::None
    };
    if lane.contains("large") {
        spec.payload_size = LARGE_PAYLOAD;
        spec.raster_size = LARGE_RASTER;
    }
    if lane.starts_with("inventory_shared_") || lane.starts_with("inventory_distinct_") {
        spec.picture_count = lane
            .rsplit('_')
            .next()
            .ok_or("inventory lane has no picture count")?
            .parse()?;
    }
    if lane == "namespace_heavy" {
        spec.namespace_heavy = true;
        spec.picture_count = 1;
    }
    if lane == "inventory_shared_root_namespace_32" {
        spec.picture_count = 32;
        spec.svg = SvgVariant::Embedded;
        // The root namespace context is shared by all pictures, while each
        // SVG owner deliberately keeps a distinct relationship id.
        spec.targets = TargetTopology::Distinct;
        spec.namespace_root_heavy = true;
        spec.opaque_extension = false;
    }
    if lane == "namespace_limit_refusal" {
        spec.namespace_limit = true;
        spec.picture_count = 1;
    }
    if lane.starts_with("clone_attached") {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
    }
    if lane.starts_with("clone_captured_owner") {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
    }
    if lane.starts_with("attach_end_to_end")
        || lane.starts_with("inverse_attach_detach")
        || lane.starts_with("noop_detach")
    {
        spec.svg = SvgVariant::None;
        spec.targets = TargetTopology::None;
    }
    if lane == "mixed_caps_rejection" {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
    }
    if lane.starts_with("detach_end_to_end_shared") {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
    }
    if lane.starts_with("detach_end_to_end_shared_final") {
        spec.svg_picture = Some(picture);
    }
    if lane == "strict_detach_end_to_end_shared_final_two_cell" {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
        spec.svg_picture = Some(picture);
    }
    if lane == "incoming_edge_shared_final_two_cell" {
        spec.picture_count = 1;
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
        spec.svg_picture = Some(0);
        spec.incoming_svg_edge = true;
    }
    if lane.starts_with("detach_end_to_end_distinct") {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Distinct;
    }
    if lane == "malformed_duplicate_owner" {
        spec.picture_count = 1;
        spec.svg = SvgVariant::Duplicate;
        spec.targets = TargetTopology::Shared;
    } else if lane == "malformed_mce_owner" {
        spec.picture_count = 1;
        spec.svg = SvgVariant::Mce;
        spec.targets = TargetTopology::Shared;
    } else if lane == "malformed_linked_owner" {
        spec.picture_count = 1;
        spec.svg = SvgVariant::Linked;
        spec.targets = TargetTopology::Shared;
    } else if lane == "malformed_unknown_uri" {
        spec.picture_count = 1;
        spec.svg = SvgVariant::Unknown;
        spec.targets = TargetTopology::None;
    } else if lane.starts_with("multi_picture_same_drawing_detach_") {
        spec.picture_count = lane
            .rsplit('_')
            .next()
            .ok_or("multi-picture detach lane has no picture count")?
            .parse()?;
        spec.svg = SvgVariant::Embedded;
        spec.targets = if lane.contains("_shared_") {
            TargetTopology::Shared
        } else {
            TargetTopology::Distinct
        };
    } else if lane.starts_with("multi_picture_same_drawing_") {
        spec.picture_count = lane
            .rsplit('_')
            .next()
            .ok_or("multi-picture lane has no picture count")?
            .parse()?;
        spec.svg = SvgVariant::None;
        spec.targets = TargetTopology::None;
    }
    if lane.starts_with("limit_") {
        spec.payload_size = SVG_INPUT_LIMIT.saturating_add(1);
        spec.svg = SvgVariant::None;
        spec.targets = TargetTopology::None;
    }

    let payload = payload_bytes(spec.payload_size, lane.as_bytes());
    let (package, namespace_generated_bindings, namespace_active_bindings, namespace_active_limit) =
        if lane == "namespace_limit_refusal" {
            let (package, generated, active, limit) = derive_namespace_limit_fixture(&payload)?;
            (package, Some(generated), Some(active), Some(limit))
        } else if multisheet {
            (package_bytes_multisheet(spec, &payload)?, None, None, None)
        } else {
            (package_bytes(spec, &payload)?, None, None, None)
        };
    let mut identity = package.clone();
    if lane.starts_with("limit_") {
        identity.extend_from_slice(&payload);
    }
    let expected_success = !matches!(
        lane,
        "namespace_limit_refusal"
            | "limit_small"
            | "limit_large"
            | "malformed_duplicate_owner"
            | "malformed_mce_owner"
            | "malformed_linked_owner"
            | "mixed_caps_rejection"
    );
    let caller_limit_profile = if lane.starts_with("limit_") {
        Some("svg_input_bytes")
    } else if lane == "mixed_caps_rejection" {
        Some("composite_read_limits")
    } else {
        None
    };
    let caller_limit_ceilings = if lane == "mixed_caps_rejection" {
        Some(composite_caller_limit_ceilings(&package, payload.len())?)
    } else {
        None
    };
    Ok(Fixture {
        input_bytes: u64::try_from(identity.len())?,
        input_hash: support::fnv1a64(&identity),
        input_sha256: support::sha256_hex(&identity),
        package: Arc::from(package.into_boxed_slice()),
        payload: Arc::from(payload.into_boxed_slice()),
        picture_count: spec.picture_count,
        drawing_count: if multisheet { 2 } else { 1 },
        expected_success,
        picture,
        strict: spec.strict,
        incoming_svg_edge: spec.incoming_svg_edge,
        opaque_extension: spec.svg == SvgVariant::None && spec.opaque_extension,
        opaque_descendants: spec.namespace_heavy,
        namespace_generated_bindings,
        namespace_active_bindings,
        namespace_active_limit,
        multisheet,
        caller_limit_profile,
        caller_limit_ceilings,
    })
}

fn picture_index(lane: &str) -> usize {
    if lane.contains("one_cell") {
        1
    } else if lane.contains("absolute") {
        2
    } else {
        0
    }
}

fn native_fixture() -> Result<Fixture> {
    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../fixtures/tdf169496_hidden_graphic.xlsx");
    let bytes = fs::read(path)?;
    let input_bytes = u64::try_from(bytes.len())?;
    let input_hash = support::fnv1a64(&bytes);
    let input_sha256 = support::sha256_hex(&bytes);
    if input_sha256 != NATIVE_FIXTURE_SHA256 {
        return Err(format!(
            "native producer fixture SHA-256 changed: expected {NATIVE_FIXTURE_SHA256}, got {input_sha256}"
        )
        .into());
    }
    Ok(Fixture {
        input_bytes,
        input_hash,
        input_sha256,
        package: Arc::from(bytes.into_boxed_slice()),
        payload: Arc::from(Vec::<u8>::new().into_boxed_slice()),
        picture_count: 2,
        drawing_count: 1,
        expected_success: true,
        picture: 0,
        strict: false,
        incoming_svg_edge: false,
        opaque_extension: false,
        opaque_descendants: false,
        namespace_generated_bindings: None,
        namespace_active_bindings: None,
        namespace_active_limit: None,
        multisheet: false,
        caller_limit_profile: None,
        caller_limit_ceilings: None,
    })
}

fn payload_bytes(size: usize, seed: &[u8]) -> Vec<u8> {
    let mut payload = Vec::with_capacity(size);
    let prefix = b"<svg xmlns=\"http://www.w3.org/2000/svg\"><path d=\"M0 0 L4 4\"/><!--";
    let suffix = b"--></svg>";
    payload.extend_from_slice(prefix);
    let mut index = 0usize;
    while payload.len().saturating_add(suffix.len()) < size {
        // The package writer validates image/svg+xml members as UTF-8 XML.
        // Keep the generated filler ASCII while leaving its content opaque to
        // the lifecycle operation itself.
        let value = seed[index % seed.len().max(1)] ^ (index as u8).wrapping_mul(17);
        payload.push(b'a'.saturating_add(value % 26));
        index = index.saturating_add(1);
    }
    payload.extend_from_slice(suffix);
    payload.truncate(size);
    payload
}

fn raster_bytes(size: usize) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(size);
    for index in 0..size {
        bytes.push((index as u8).wrapping_mul(31).wrapping_add(7));
    }
    bytes
}

fn q(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

fn marker(prefix: &str, name: &str, col: usize) -> String {
    let marker = q(prefix, name);
    let col_name = q(prefix, "col");
    let col_off_name = q(prefix, "colOff");
    let row_name = q(prefix, "row");
    let row_off_name = q(prefix, "rowOff");
    format!(
        "<{marker}><{col_name}>{col}</{col_name}><{col_off_name}>2</{col_off_name}><{row_name}>{col}</{row_name}><{row_off_name}>4</{row_off_name}></{marker}>"
    )
}

fn svg_relationship_id(spec: FixtureSpec, index: usize) -> String {
    if spec.targets == TargetTopology::Distinct {
        format!("rIdSvg{index}")
    } else {
        String::from("rIdSvg")
    }
}

fn svg_target_name(spec: FixtureSpec, index: usize) -> String {
    if spec.targets == TargetTopology::Distinct {
        format!("/xl/media/image{}.svg", index.saturating_add(2))
    } else {
        String::from("/xl/media/image2.svg")
    }
}

fn owner_xml(spec: FixtureSpec, index: usize) -> String {
    if let Some(selected) = spec.svg_picture {
        if selected != index {
            return String::new();
        }
    }
    let relation_id = svg_relationship_id(spec, index);
    let relation_prefix = if spec.strict { "trans" } else { "r" };
    let admitted = format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip {relation_prefix}:embed="{relation_id}"/></a:ext>"#
    );
    let linked = format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip {relation_prefix}:link="{relation_id}"/></a:ext>"#
    );
    let duplicate = format!("{admitted}{admitted}");
    match spec.svg {
        SvgVariant::Embedded => format!("<a:extLst>{admitted}</a:extLst>"),
        SvgVariant::Duplicate => format!("<a:extLst>{duplicate}</a:extLst>"),
        SvgVariant::Mce => format!(
            r#"<mc:AlternateContent><mc:Choice Requires="asvg"><a:extLst>{admitted}</a:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent>"#
        ),
        SvgVariant::Linked => format!("<a:extLst>{linked}</a:extLst>"),
        SvgVariant::Unknown => {
            let suffix = index % 7;
            format!(
                r#"<a:extLst><a:ext uri="urn:litchi:future:{suffix}"><future:payload future:index="{index}"/></a:ext></a:extLst>"#
            )
        },
        SvgVariant::None if spec.opaque_extension => String::from(
            r#"<a:extLst><a:ext uri="urn:litchi:future-extension"><!--future-comment--><future:payload future:keep="yes"/></a:ext></a:extLst>"#,
        ),
        SvgVariant::None => String::new(),
    }
}

fn namespace_pressure(binding_count: usize) -> String {
    let mut output = String::new();
    output.push_str("<n:scope>");
    let levels = binding_count.saturating_add(255) / 256;
    let mut remaining = binding_count;
    for level in (0..levels).rev() {
        write!(output, "<n:l{level}").expect("String cannot fail");
        let local = remaining.min(256);
        for binding in 0..local {
            write!(
                output,
                " xmlns:p{level}_{binding:03}=\"urn:litchi:limit:{level}:{binding}\""
            )
            .expect("String cannot fail");
        }
        remaining = remaining.saturating_sub(local);
        output.push('>');
    }
    output.push_str("<n:leaf/>");
    for level in 0..levels {
        write!(output, "</n:l{level}>").expect("String cannot fail");
    }
    output.push_str("</n:scope>");
    output
}

fn namespace_descendants() -> String {
    let mut output = String::new();
    for index in 0..1_024 {
        write!(output, "<future:opaque index=\"{index}\"/>").expect("String cannot fail");
    }
    output
}

fn picture_xml(spec: FixtureSpec, index: usize) -> String {
    let anchor_index = index % 3;
    let picture = q("x", "pic");
    let nv = q("x", "nvPicPr");
    let c_nv_pr = q("x", "cNvPr");
    let c_nv_pic_pr = q("x", "cNvPicPr");
    let blip_fill = q("x", "blipFill");
    let blip = q("a", "blip");
    let sp_pr = q("x", "spPr");
    let mut body = String::new();
    write!(
        body,
        r#"<{picture}><{nv}><{c_nv_pr} id="{}" name="picture-{}"/><{c_nv_pic_pr}/></{nv}>"#,
        index.saturating_add(1),
        index.saturating_add(1),
    )
    .expect("String cannot fail");
    if spec.namespace_limit {
        body.push_str(&namespace_pressure(spec.namespace_bindings));
    } else if spec.namespace_heavy {
        body.push_str(&namespace_descendants());
    }
    write!(
        body,
        r#"<{blip_fill}><{blip} r:embed="rIdRaster">{owner}</{blip}></{blip_fill}><{sp_pr}/></{picture}>"#,
        owner = owner_xml(spec, index),
    )
    .expect("String cannot fail");
    let client_data = q("x", "clientData");
    match anchor_index {
        0 => format!(
            r#"<x:twoCellAnchor>{from}{to}{body}<{client_data}/></x:twoCellAnchor>"#,
            from = marker("x", "from", index.saturating_add(1)),
            to = marker("x", "to", index.saturating_add(5)),
        ),
        1 => format!(
            r#"<x:oneCellAnchor>{from}<x:ext cx="123456" cy="654321"/>{body}<{client_data}/></x:oneCellAnchor>"#,
            from = marker("x", "from", index.saturating_add(9)),
        ),
        _ => format!(
            r#"<x:absoluteAnchor><x:pos x="-900" y="456"/><x:ext cx="777888" cy="999000"/>{body}<{client_data}/></x:absoluteAnchor>"#
        ),
    }
}

fn drawing_xml(spec: FixtureSpec) -> Vec<u8> {
    let (xdr, drawing, rel, trans) = if spec.strict {
        (
            STRICT_XDR,
            STRICT_A,
            STRICT_REL,
            format!(r#" xmlns:trans="{REL}""#),
        )
    } else {
        (XDR, A, REL, String::new())
    };
    let mut root = format!(
        r#"<x:wsDr xmlns:x="{xdr}" xmlns:a="{drawing}" xmlns:r="{rel}"{trans} xmlns:asvg="{SVG_NS}" xmlns:mc="{MCE}" xmlns:future="urn:litchi:future" xmlns:n="urn:litchi:namespace:scope""#
    );
    if spec.namespace_root_heavy {
        let payload = "x".repeat(1_000);
        for index in 0..128 {
            write!(
                root,
                " xmlns:p{index:03}=\"urn:litchi:root:{index:03}:{payload}\""
            )
            .expect("String cannot fail");
        }
    }
    if spec.namespace_heavy {
        // Keep the complete declaration count at 252, including the seven
        // fixed bindings above, so this lane stays below the per-element
        // declaration bound while exercising inherited lookup.
        for index in 0..245 {
            write!(
                root,
                " xmlns:n{index:03}=\"urn:litchi:namespace:{index:03}\""
            )
            .expect("String cannot fail");
        }
    }
    root.push('>');
    for index in 0..spec.picture_count {
        root.push_str(&picture_xml(spec, index));
    }
    root.push_str("</x:wsDr>");
    root.into_bytes()
}

fn worksheet_xml(strict: bool) -> Vec<u8> {
    let (main, rel) = if strict {
        ("http://purl.oclc.org/ooxml/spreadsheetml/main", STRICT_REL)
    } else {
        (SML, REL)
    };
    format!(
        r#"<worksheet xmlns="{main}" xmlns:r="{rel}"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#
    )
    .into_bytes()
}

fn package_bytes(spec: FixtureSpec, payload: &[u8]) -> Result<Vec<u8>> {
    let mut package = Package::create()?.into_plain_opc();
    package
        .get_part_mut(&PackURI::new(SHEET)?)?
        .set_blob(worksheet_xml(spec.strict));
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(DRAWING)?,
        ct::OFC_DRAWING.to_owned(),
        drawing_xml(spec),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(RASTER)?,
        ct::PNG.to_owned(),
        raster_bytes(spec.raster_size),
    )))?;
    if spec.svg == SvgVariant::Embedded
        || spec.svg == SvgVariant::Duplicate
        || spec.svg == SvgVariant::Mce
    {
        let count = if spec.targets == TargetTopology::Distinct {
            spec.picture_count
        } else {
            1
        };
        for index in 0..count {
            package.try_add_part(Box::new(BlobPart::new(
                PackURI::new(svg_target_name(spec, index))?,
                "image/svg+xml".to_owned(),
                payload.to_vec(),
            )))?;
        }
    }
    if spec.incoming_svg_edge {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(OTHER)?,
            "application/xml".to_owned(),
            b"<opaque-owner/>".to_vec(),
        )))?;
    }
    let sheet = package.get_part_mut(&PackURI::new(SHEET)?)?;
    sheet.rels_mut().try_add_relationship(
        if spec.strict {
            rt::STRICT_DRAWING
        } else {
            rt::DRAWING
        }
        .to_owned(),
        "../drawings/drawing1.xml".to_owned(),
        "rIdDrawing".to_owned(),
        TargetMode::Internal,
    )?;
    let drawing = package.get_part_mut(&PackURI::new(DRAWING)?)?;
    drawing.rels_mut().try_add_relationship(
        if spec.strict {
            rt::STRICT_IMAGE
        } else {
            rt::IMAGE
        }
        .to_owned(),
        "../media/image1.png".to_owned(),
        "rIdRaster".to_owned(),
        TargetMode::Internal,
    )?;
    if spec.svg != SvgVariant::None && spec.svg != SvgVariant::Unknown {
        let count = if spec.targets == TargetTopology::Distinct {
            spec.picture_count
        } else {
            1
        };
        for index in 0..count {
            drawing.rels_mut().try_add_relationship(
                if spec.strict {
                    rt::STRICT_IMAGE
                } else {
                    rt::IMAGE
                }
                .to_owned(),
                if spec.svg == SvgVariant::Linked {
                    "https://example.invalid/vector.svg".to_owned()
                } else if spec.targets == TargetTopology::Distinct {
                    format!("../media/image{}.svg", index.saturating_add(2))
                } else {
                    String::from("../media/image2.svg")
                },
                svg_relationship_id(spec, index),
                if spec.svg == SvgVariant::Linked {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
    }
    if spec.incoming_svg_edge {
        package
            .get_part_mut(&PackURI::new(OTHER)?)?
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "media/image2.svg".to_owned(),
                "rIdIncomingSvg".to_owned(),
                TargetMode::Internal,
            )?;
    }
    Ok(PackageWriter::to_bytes(&package)?)
}

/// Build the smallest package that exercises public worksheet selection across
/// two independent drawings.  Each drawing keeps the same three-anchor,
/// three-picture shape as the single-sheet fixtures; the lifecycle operation
/// selects different semantic pictures on each worksheet so relationship and
/// media ownership cannot be accidentally treated as one global drawing.
fn package_bytes_multisheet(spec: FixtureSpec, payload: &[u8]) -> Result<Vec<u8>> {
    let mut package = OpcPackage::from_bytes(&package_bytes(spec, payload)?)?;
    let sheet2_xml =
        String::from_utf8(worksheet_xml(spec.strict))?.replace("rIdDrawing", "rIdDrawing2");
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(SHEET2)?,
        ct::SML_WORKSHEET.to_owned(),
        sheet2_xml.as_bytes().to_vec(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(DRAWING2)?,
        ct::OFC_DRAWING.to_owned(),
        drawing_xml(spec),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(RASTER2)?,
        ct::PNG.to_owned(),
        raster_bytes(spec.raster_size),
    )))?;

    package
        .get_part_mut(&PackURI::new(SHEET2)?)?
        .rels_mut()
        .try_add_relationship(
            if spec.strict {
                rt::STRICT_DRAWING
            } else {
                rt::DRAWING
            }
            .to_owned(),
            "../drawings/drawing2.xml".to_owned(),
            "rIdDrawing2".to_owned(),
            TargetMode::Internal,
        )?;
    package
        .get_part_mut(&PackURI::new(DRAWING2)?)?
        .rels_mut()
        .try_add_relationship(
            if spec.strict {
                rt::STRICT_IMAGE
            } else {
                rt::IMAGE
            }
            .to_owned(),
            "../media/image3.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )?;

    let workbook_uri = PackURI::new("/xl/workbook.xml")?;
    let workbook = package.get_part_mut(&workbook_uri)?;
    let workbook_xml = String::from_utf8(workbook.blob().to_vec())?;
    workbook.set_blob(
        workbook_xml
            .replace(
                "</sheets>",
                r#"<sheet name="Sheet2" sheetId="2" r:id="rId3"/></sheets>"#,
            )
            .into_bytes(),
    );
    workbook.rels_mut().try_add_relationship(
        rt::WORKSHEET.to_owned(),
        "worksheets/sheet2.xml".to_owned(),
        "rId3".to_owned(),
        TargetMode::Internal,
    )?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn namespace_limit_accepts(binding_count: usize, payload: &[u8]) -> Result<bool> {
    let spec = FixtureSpec {
        namespace_limit: true,
        namespace_bindings: binding_count,
        picture_count: 1,
        ..FixtureSpec::default()
    };
    let package = package_bytes(spec, payload)?;
    let drawing = drawing_part(&package)?;
    Ok(SourceDrawing::scan(&drawing).is_ok())
}

fn namespace_active_limit(package: &[u8]) -> Result<usize> {
    let drawing = drawing_part(package)?;
    let error = SourceDrawing::scan(&drawing)
        .expect_err("namespace refusal boundary unexpectedly became accepted");
    if !refusal_error_matches(RefusalKind::NamespaceLimit, &error) {
        return Err(
            format!("namespace boundary returned an unexpected source error: {error}").into(),
        );
    }
    let limit = error
        .to_string()
        .rsplit_once("limit of ")
        .and_then(|(_, value)| value.trim().parse::<usize>().ok())
        .ok_or("namespace refusal did not expose a positive active-binding limit")?;
    if limit == 0 {
        return Err("namespace refusal exposed a zero active-binding limit".into());
    }
    Ok(limit)
}

fn derive_namespace_limit_fixture(payload: &[u8]) -> Result<(Vec<u8>, usize, usize, usize)> {
    let mut accepted = 0usize;
    let mut rejected = 1usize;
    while namespace_limit_accepts(rejected, payload)? {
        accepted = rejected;
        rejected = rejected
            .checked_mul(2)
            .ok_or("namespace policy search overflowed")?;
        if rejected > 1_048_576 {
            return Err("namespace policy refusal boundary was not found".into());
        }
    }
    while rejected.saturating_sub(accepted) > 1 {
        let middle = accepted + (rejected - accepted) / 2;
        if namespace_limit_accepts(middle, payload)? {
            accepted = middle;
        } else {
            rejected = middle;
        }
    }
    let spec = FixtureSpec {
        namespace_limit: true,
        namespace_bindings: rejected,
        picture_count: 1,
        ..FixtureSpec::default()
    };
    if !namespace_limit_accepts(accepted, payload)? {
        return Err("namespace policy search lost its last accepted generated count".into());
    }
    let package = package_bytes(spec, payload)?;
    let active = rejected
        .checked_add(FIXED_ROOT_NAMESPACE_BINDINGS)
        .ok_or("namespace active-binding count overflowed")?;
    let limit = namespace_active_limit(&package)?;
    if active != limit.saturating_add(1) {
        return Err(format!(
            "namespace boundary count mismatch: generated={rejected}, active={active}, limit={limit}"
        )
        .into());
    }
    let accepted_active = accepted
        .checked_add(FIXED_ROOT_NAMESPACE_BINDINGS)
        .ok_or("accepted namespace active-binding count overflowed")?;
    if accepted_active > limit {
        return Err(format!(
            "namespace previous accepted count exceeded limit: active={accepted_active}, limit={limit}"
        )
        .into());
    }
    Ok((package, rejected, active, limit))
}

fn refusal_error_matches(kind: RefusalKind, error: &XlsxError) -> bool {
    let message = error.to_string();
    match kind {
        RefusalKind::NamespaceLimit => {
            matches!(error, XlsxError::Invalid(_))
                && message.contains("active namespace")
                && message.contains("limit")
        },
        RefusalKind::CallerLimit => {
            matches!(error, XlsxError::Invalid(_))
                && message.contains("SVG attachment input exceeds")
        },
        RefusalKind::DuplicateOwner => {
            matches!(error, XlsxError::Invalid(_))
                && (message.contains("ambiguous") || message.contains("duplicate SVG"))
        },
        RefusalKind::MceAncestry => {
            matches!(error, XlsxError::Invalid(_))
                && (message.contains("refused SVG owner") || message.contains("ambiguous"))
        },
        RefusalKind::LinkedOwner => {
            matches!(error, XlsxError::Invalid(_)) && message.contains("linked SVG owners cannot")
        },
        RefusalKind::MixedLimit => match error {
            XlsxError::Package(package) => {
                matches!(package, OpcError::ReadLimit { .. })
            },
            XlsxError::Invalid(_) => {
                message.contains("SVG") && (message.contains("exceed") || message.contains("limit"))
            },
            _ => false,
        },
    }
}

fn checked_refusal(
    kind: RefusalKind,
    operation: std::result::Result<(), XlsxError>,
    before: &[u8],
    after: &[u8],
) -> Result<LaneExecution> {
    let error = match operation {
        Ok(()) => {
            return Err(format!(
                "expected {} refusal but the API operation was accepted",
                kind.as_str()
            )
            .into());
        },
        Err(error) => error,
    };
    if before != after {
        return Err(format!(
            "{} refusal changed source bytes after the API returned an error",
            kind.as_str()
        )
        .into());
    }
    if !refusal_error_matches(kind, &error) {
        return Err(format!(
            "{} refusal returned an unexpected API error: {}",
            kind.as_str(),
            error
        )
        .into());
    }
    Ok(LaneExecution::ExpectedRefusal(ExpectedRefusal {
        kind,
        message: error.to_string(),
    }))
}

fn validate_malformed_owner_state(fixture: &Fixture, kind: RefusalKind) -> Result<()> {
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let owner_state = source.picture(fixture.picture)?.svg_owner();
    let valid = match kind {
        RefusalKind::DuplicateOwner => matches!(owner_state, SvgOwnerState::Ambiguous),
        RefusalKind::MceAncestry => matches!(owner_state, SvgOwnerState::Refused),
        RefusalKind::LinkedOwner => matches!(owner_state, SvgOwnerState::Linked(_)),
        _ => false,
    };
    if valid {
        Ok(())
    } else {
        Err(format!(
            "fixture owner state did not match the expected {} refusal",
            kind.as_str()
        )
        .into())
    }
}

fn execute_namespace_limit_refusal(fixture: &Fixture) -> Result<LaneExecution> {
    let (Some(generated), Some(active), Some(limit)) = (
        fixture.namespace_generated_bindings,
        fixture.namespace_active_bindings,
        fixture.namespace_active_limit,
    ) else {
        return Err("namespace limit lane did not derive its binding fields".into());
    };
    if generated == 0
        || active <= generated
        || active != generated + FIXED_ROOT_NAMESPACE_BINDINGS
        || active != limit.saturating_add(1)
    {
        return Err(format!(
            "namespace limit fields are inconsistent: generated={generated}, active={active}, limit={limit}"
        )
        .into());
    }
    let drawing = drawing_part(&fixture.package)?;
    let operation = SourceDrawing::scan(&drawing).map(|_| ());
    checked_refusal(RefusalKind::NamespaceLimit, operation, &drawing, &drawing)
}

fn execute_caller_limit_refusal(fixture: &Fixture) -> Result<LaneExecution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let before = workbook.to_plain_bytes()?;
    let operation = {
        let mut edit = workbook.edit()?;
        let mut sheet = edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
        sheet
            .attach_svg(
                PictureSelector::new(0, fixture.picture),
                SvgInput::borrowed(fixture.payload.as_ref()),
            )
            .map(|_| ())
    };
    let after = workbook.to_plain_bytes()?;
    checked_refusal(RefusalKind::CallerLimit, operation, &before, &after)
}

pub fn run(lane: &str, warmup: usize, samples: usize) -> Result<String> {
    let fixture = fixture_for_lane(lane)?;
    if lane.starts_with("clone_captured_owner") {
        return run_captured_owner_clone(lane, warmup, samples, &fixture);
    }
    for _ in 0..warmup {
        match execute_lane(lane, &fixture) {
            Ok(LaneExecution::Success(execution))
                if fixture.expected_success && execution.semantic_ok && execution.output_exact => {
            },
            Ok(LaneExecution::ExpectedRefusal(_)) if !fixture.expected_success => {},
            Ok(LaneExecution::Success(_)) if !fixture.expected_success => {
                return Err("warm-up unexpectedly accepted a refusal lane".into());
            },
            Ok(LaneExecution::ExpectedRefusal(_)) => {
                return Err("warm-up unexpectedly refused a success lane".into());
            },
            Ok(LaneExecution::Success(_)) => {
                return Err("warm-up semantic or output validation failed".into());
            },
            Err(error) => return Err(format!("warm-up failed: {error}").into()),
        }
    }
    let mut receipts = Vec::with_capacity(samples);
    for _ in 0..samples {
        support::reset_counters();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let result = execute_lane(lane, &fixture);
        // Drop candidate output before the allocator snapshot. Fixture and
        // receipt construction stay outside the measured operation, while a
        // changed package's temporary output must not make live_after look
        // permanently larger than live_before.
        let outcome = match result {
            Ok(LaneExecution::Success(execution)) => {
                let outcome = Outcome::Success {
                    semantic_ok: execution.semantic_ok,
                    output_exact: execution.output_exact,
                };
                drop(execution.output);
                outcome
            },
            Ok(LaneExecution::ExpectedRefusal(refusal)) => Outcome::ExpectedRefusal(refusal),
            Err(error) => Outcome::Failure(error),
        };
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let allocation = before.delta(after);
        receipts.push(sample_json(
            lane,
            fixture.expected_success,
            elapsed_ns,
            allocation,
            outcome,
        ));
    }
    Ok(receipt_json(lane, &fixture, warmup, &receipts))
}

fn run_captured_owner_clone(
    lane: &str,
    warmup: usize,
    samples: usize,
    fixture: &Fixture,
) -> Result<String> {
    // Capture and validate the source-backed owner before any warm-up or
    // measured sample.  The measured operation below is only the clone of the
    // already captured public owner; it must not include drawing extraction,
    // XML scanning, source validation, or OPC graph inspection.
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let all_anchors = source.pictures().len() == 3
        && source
            .pictures()
            .iter()
            .map(|picture| anchor_kind(picture.anchor()))
            .collect::<HashSet<_>>()
            == HashSet::from(["two_cell", "one_cell", "absolute"]);
    if !all_anchors {
        return Err("captured-owner fixture did not expose all three anchor forms".into());
    }
    let captured = source.picture(fixture.picture)?.clone();
    drop(source);
    if !captured.is_direct_embedded_svg() {
        return Err("captured-owner fixture did not expose an embedded SVG owner".into());
    }
    captured.picture_bytes(&drawing)?;
    if validate_graph_closure(&fixture.package, 1).is_err() {
        return Err("captured-owner fixture graph validation failed".into());
    }
    if validate_opaque_fragments(&fixture.package, fixture).is_err() {
        return Err("captured-owner fixture opaque-fragment validation failed".into());
    }

    for _ in 0..warmup {
        let cloned = captured.clone();
        let execution = validate_captured_owner_clone(&drawing, &captured, &cloned)?;
        drop(cloned);
        if !execution.semantic_ok || !execution.output_exact {
            return Err("captured-owner clone warm-up validation failed".into());
        }
    }

    let mut receipts = Vec::with_capacity(samples);
    for _ in 0..samples {
        support::reset_counters();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let cloned = black_box(captured.clone());
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let allocation = before.delta(after);
        let result = validate_captured_owner_clone(&drawing, &captured, &cloned);
        let outcome = match result {
            Ok(execution) => Outcome::Success {
                semantic_ok: execution.semantic_ok,
                output_exact: execution.output_exact,
            },
            Err(error) => Outcome::Failure(error),
        };
        drop(cloned);
        receipts.push(sample_json(
            lane,
            fixture.expected_success,
            elapsed_ns,
            allocation,
            outcome,
        ));
    }
    Ok(receipt_json(lane, fixture, warmup, &receipts))
}

fn execute_lane(lane: &str, fixture: &Fixture) -> Result<LaneExecution> {
    if lane == "namespace_limit_refusal" {
        return execute_namespace_limit_refusal(fixture);
    }
    if lane.starts_with("limit_") {
        return execute_caller_limit_refusal(fixture);
    }
    if lane.starts_with("capture_") {
        return execute_capture(lane, fixture).map(LaneExecution::Success);
    }
    if lane == "namespace_heavy" {
        return execute_capture(lane, fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("clone_") {
        if lane.starts_with("clone_captured_owner") {
            return Err("captured-owner clone must use its pre-captured execution path".into());
        }
        return execute_clone(lane, fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("inventory_") {
        if lane == "inventory_shared_root_namespace_32" {
            return execute_inventory_root_namespace(fixture).map(LaneExecution::Success);
        }
        return execute_inventory(lane, fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("attach_end_to_end") {
        return execute_attach(fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("strict_attach_end_to_end") {
        return execute_attach(fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("inverse_attach_detach") {
        return execute_inverse(fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("detach_end_to_end") {
        return execute_detach(lane, fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("strict_detach_end_to_end") {
        return execute_detach(lane, fixture).map(LaneExecution::Success);
    }
    if lane == "incoming_edge_shared_final_two_cell" {
        return execute_detach(lane, fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("noop_detach") {
        return execute_detach(lane, fixture).map(LaneExecution::Success);
    }
    if lane == "malformed_duplicate_owner"
        || lane == "malformed_mce_owner"
        || lane == "malformed_linked_owner"
    {
        let kind = match lane {
            "malformed_duplicate_owner" => RefusalKind::DuplicateOwner,
            "malformed_mce_owner" => RefusalKind::MceAncestry,
            "malformed_linked_owner" => RefusalKind::LinkedOwner,
            _ => unreachable!("lane was matched above"),
        };
        return execute_expected_refusal(fixture, kind);
    }
    if lane == "malformed_unknown_uri" {
        return execute_detach(lane, fixture).map(LaneExecution::Success);
    }
    if lane.starts_with("same_picture_attach_detach") {
        return execute_same_picture_attach_detach(fixture).map(LaneExecution::Success);
    }
    if lane == "multisheet_attach_detach" {
        return execute_multisheet_attach_detach(fixture).map(LaneExecution::Success);
    }
    if lane == "mixed_caps_rejection" {
        return execute_mixed_caps(fixture);
    }
    if lane.starts_with("multi_picture_same_drawing_") {
        if lane.starts_with("multi_picture_same_drawing_detach_") {
            return execute_multi_picture_detach(fixture).map(LaneExecution::Success);
        }
        return execute_multi_picture_attach(fixture).map(LaneExecution::Success);
    }
    Err(format!("no adapter operation for lane {lane}").into())
}

fn execute_capture(lane: &str, fixture: &Fixture) -> Result<Execution> {
    let _workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let picture = source.picture(fixture.picture)?;
    let expected_attached = lane == "capture_native_fixture" || lane.contains("attached");
    let actual_attached = picture.is_direct_embedded_svg();
    let expected_anchor = lane_anchor_kind(lane);
    let anchor_ok =
        expected_anchor.is_none_or(|expected| anchor_kind(picture.anchor()) == expected);
    let native_ok = if lane == "capture_native_fixture" {
        let package = OpcPackage::from_bytes(fixture.package.as_ref())?;
        let raster = package.get_part(&PackURI::new("/xl/media/image1.png")?)?;
        let pictures = source.pictures();
        pictures.len() == 2
            && pictures.iter().all(|picture| {
                picture.raster_relationship_id() == "rId1"
                    && picture
                        .svg_owner()
                        .owner()
                        .and_then(|owner| owner.embedded_relationship_id())
                        == Some("rId2")
            })
            && raster.content_type() == "image/png"
            && !raster.blob().is_empty()
    } else {
        true
    };
    let graph_ok =
        validate_graph_closure(&fixture.package, if expected_attached { 1 } else { 0 }).is_ok();
    let native_anchor_ok = if lane == "capture_native_fixture" {
        source
            .pictures()
            .iter()
            .all(|picture| anchor_kind(picture.anchor()) == "two_cell")
    } else {
        true
    };
    let opaque_ok = validate_opaque_fragments(&fixture.package, fixture).is_ok();
    Ok(Execution {
        output: None,
        semantic_ok: actual_attached == expected_attached
            && anchor_ok
            && native_ok
            && native_anchor_ok
            && graph_ok
            && opaque_ok,
        output_exact: true,
    })
}

fn execute_clone(lane: &str, fixture: &Fixture) -> Result<Execution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let all_anchors = source.pictures().len() == 3
        && source
            .pictures()
            .iter()
            .map(|picture| anchor_kind(picture.anchor()))
            .collect::<HashSet<_>>()
            == HashSet::from(["two_cell", "one_cell", "absolute"]);
    let _owner = source.picture(fixture.picture)?;
    let first = workbook.clone();
    let second = first.clone();
    let snapshot_bytes = second.to_plain_bytes()?;
    black_box(second);
    let opaque_ok = validate_opaque_fragments(&fixture.package, fixture).is_ok();
    Ok(Execution {
        output: None,
        semantic_ok: all_anchors
            && snapshot_bytes == fixture.package.as_ref()
            && validate_graph_closure(
                &fixture.package,
                if lane.starts_with("clone_attached") {
                    1
                } else {
                    0
                },
            )
            .is_ok()
            && opaque_ok,
        output_exact: snapshot_bytes == fixture.package.as_ref(),
    })
}

fn validate_captured_owner_clone(
    drawing: &[u8],
    captured: &PictureSource<'_>,
    cloned: &PictureSource<'_>,
) -> Result<Execution> {
    let bytes_same = captured.picture_bytes(drawing)? == cloned.picture_bytes(drawing)?;
    let owner_same = captured.svg_owner() == cloned.svg_owner();
    let relationship_refs_same =
        captured.relationship_references() == cloned.relationship_references();
    let semantic_ok =
        bytes_same && owner_same && relationship_refs_same && cloned.is_direct_embedded_svg();
    Ok(Execution {
        output: None,
        semantic_ok,
        output_exact: bytes_same,
    })
}

fn execute_inventory(lane: &str, fixture: &Fixture) -> Result<Execution> {
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let expected = lane
        .rsplit('_')
        .next()
        .ok_or("inventory lane has no count")?
        .parse::<usize>()?;
    let all_present = source.pictures().len() == expected
        && source
            .pictures()
            .iter()
            .enumerate()
            .all(|(index, picture)| {
                picture.picture_ordinal() == index
                    && picture.raster_relationship_id() == "rIdRaster"
                    && picture.anchor_range().start < picture.anchor_range().end
            });
    let relationship_ids = source
        .pictures()
        .iter()
        .filter_map(|picture| picture.svg_owner().owner()?.embedded_relationship_id())
        .collect::<Vec<_>>();
    let topology_ok = relationship_ids.len() == expected
        && if lane.starts_with("inventory_shared") {
            relationship_ids.iter().all(|id| *id == relationship_ids[0])
        } else {
            relationship_ids
                .windows(2)
                .all(|window| window[0] != window[1])
        };
    let graph_ok = validate_graph_closure(
        &fixture.package,
        if lane.starts_with("inventory_shared") {
            1
        } else {
            expected
        },
    )
    .is_ok();
    Ok(Execution {
        output: None,
        semantic_ok: all_present && topology_ok && graph_ok,
        output_exact: true,
    })
}

fn execute_inventory_root_namespace(fixture: &Fixture) -> Result<Execution> {
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let relationship_ids = source
        .pictures()
        .iter()
        .filter_map(|picture| picture.svg_owner().owner()?.embedded_relationship_id())
        .collect::<Vec<_>>();
    let retained_svg_source_bytes = source
        .pictures()
        .iter()
        .filter_map(|picture| match picture.svg_owner() {
            SvgOwnerState::Embedded(owner) => owner.value().raw_source(),
            _ => None,
        })
        .map(|source| source.len())
        .sum::<usize>();
    let raw_source_count = source
        .pictures()
        .iter()
        .filter_map(|picture| match picture.svg_owner() {
            SvgOwnerState::Embedded(owner) => owner.value().raw_source(),
            _ => None,
        })
        .count();
    let contexts = source
        .pictures()
        .iter()
        .filter_map(|picture| match picture.svg_owner() {
            SvgOwnerState::Embedded(owner) => owner.value().namespace_context(),
            _ => None,
        })
        .collect::<Vec<_>>();
    let shared_context = contexts
        .first()
        .is_some_and(|first| contexts.iter().all(|context| first.shares_storage(context)));
    let context_bound = contexts
        .iter()
        .all(|context| context.binding_count() >= 128);
    let semantic_ok = source.pictures().len() == fixture.picture_count
        && source.pictures().iter().all(|picture| {
            picture.is_direct_embedded_svg() && picture.raster_relationship_id() == "rIdRaster"
        })
        && relationship_ids.len() == fixture.picture_count
        && raw_source_count == fixture.picture_count
        && contexts.len() == fixture.picture_count
        && relationship_ids.iter().enumerate().all(|(index, id)| {
            relationship_ids[..index]
                .iter()
                .all(|previous| previous != id)
        })
        && retained_svg_source_bytes > fixture.picture_count
        && shared_context
        && context_bound
        && validate_graph_closure(&fixture.package, fixture.picture_count).is_ok();
    Ok(Execution {
        output: None,
        semantic_ok,
        output_exact: true,
    })
}

fn execute_attach(fixture: &Fixture) -> Result<Execution> {
    let before_drawing = drawing_part(&fixture.package)?;
    let before_source = SourceDrawing::scan(&before_drawing)?;
    let before_anchor = before_source.picture(fixture.picture)?.anchor().clone();
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or("Sheet1 is missing")?
        .attach_svg(
            PictureSelector::new(0, fixture.picture),
            SvgInput::borrowed(fixture.payload.as_ref()),
        )?;
    let commit = edit.commit()?;
    let output = commit.workbook().to_plain_bytes()?;
    let reopened = Workbook::from_bytes(output.clone())?;
    let reopened_bytes = reopened.to_plain_bytes()?;
    let after_drawing = drawing_part(&output)?;
    let after_source = SourceDrawing::scan(&after_drawing)?;
    let after_picture = after_source.picture(fixture.picture)?;
    let svg_payload_ok = svg_payload_matches(&output, &fixture.payload);
    let graph_ok = validate_graph_closure(&output, 1).is_ok()
        && validate_picture_closure(&output, fixture.picture, true).is_ok();
    let strict_ok = !fixture.strict || validate_strict_host(&output, 1).is_ok();
    let opaque_ok = validate_opaque_fragments(&output, fixture).is_ok();
    let semantic_ok = output != fixture.package.as_ref()
        && reopened_bytes == output
        && after_picture.anchor() == &before_anchor
        && after_picture.is_direct_embedded_svg()
        && svg_payload_ok
        && graph_ok;
    Ok(Execution {
        output: Some(output),
        semantic_ok: semantic_ok && strict_ok && opaque_ok,
        // The candidate is checked against the complete expected closure and
        // reopened bytes. "Exact" here means the adapter's deterministic
        // output assertions passed; attach necessarily changes source bytes.
        output_exact: semantic_ok,
    })
}

fn execute_inverse(fixture: &Fixture) -> Result<Execution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let original = workbook.to_plain_bytes()?;
    if original != fixture.package.as_ref() {
        return Err("inverse fixture did not preserve its source snapshot".into());
    }
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or("Sheet1 is missing")?
        .attach_svg(
            PictureSelector::new(0, fixture.picture),
            SvgInput::borrowed(fixture.payload.as_ref()),
        )?;
    let commit = edit.commit()?;
    let published = commit.workbook().to_plain_bytes()?;
    if published == original {
        return Err("inverse forward publication was unexpectedly a no-op".into());
    }
    let reopened = Workbook::from_bytes(published.clone())?;
    let reopened_bytes = reopened.to_plain_bytes()?;
    let inverse = commit.patch().inverse();
    let restored = commit.workbook().apply(&inverse)?.into_workbook();
    let restored_bytes = restored.to_plain_bytes()?;
    let replay_refused = commit.workbook().apply(commit.patch()).is_err();
    let semantic_ok = reopened_bytes == published
        && restored_bytes == original
        && replay_refused
        && validate_graph_closure(&published, 1).is_ok()
        && validate_picture_closure(&published, fixture.picture, true).is_ok()
        && validate_graph_closure(&restored_bytes, 0).is_ok()
        && validate_picture_closure(&restored_bytes, fixture.picture, false).is_ok()
        && validate_opaque_fragments(&published, fixture).is_ok()
        && validate_opaque_fragments(&restored_bytes, fixture).is_ok();
    Ok(Execution {
        output: Some(published),
        semantic_ok,
        output_exact: restored_bytes == original,
    })
}

fn execute_detach(lane: &str, fixture: &Fixture) -> Result<Execution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or("Sheet1 is missing")?
        .detach_svg(PictureSelector::new(0, fixture.picture))?;
    let commit = edit.commit()?;
    let output = commit.workbook().to_plain_bytes()?;
    let reopened = Workbook::from_bytes(output.clone())?;
    let reopened_bytes = reopened.to_plain_bytes()?;
    let drawing = drawing_part(&output)?;
    let source = SourceDrawing::scan(&drawing)?;
    let selected_absent = !source.picture(fixture.picture)?.is_direct_embedded_svg();
    let expected_svg_parts = if lane.starts_with("detach_end_to_end_shared_first") {
        1
    } else if fixture.incoming_svg_edge {
        1
    } else if lane.starts_with("detach_end_to_end_shared_final")
        || lane.starts_with("detach_end_to_end_distinct")
    {
        if lane.starts_with("detach_end_to_end_shared_final") {
            0
        } else {
            fixture.picture_count.saturating_sub(1)
        }
    } else {
        0
    };
    let source_exact = (lane.starts_with("noop_detach") || lane == "malformed_unknown_uri")
        .then_some(output.as_slice() == fixture.package.as_ref())
        .unwrap_or(true);
    let strict_ok = !fixture.strict || validate_strict_host(&output, expected_svg_parts).is_ok();
    let incoming_ok = !fixture.incoming_svg_edge || validate_incoming_svg_edge(&output).is_ok();
    let selected_edge_ok =
        !fixture.incoming_svg_edge || validate_selected_drawing_svg_edge_removed(&output).is_ok();
    let semantic_ok = reopened_bytes == output
        && selected_absent
        && svg_parts(&output).len() == expected_svg_parts
        && source_exact
        && validate_graph_closure(&output, expected_svg_parts).is_ok()
        && validate_picture_closure(&output, fixture.picture, false).is_ok()
        && strict_ok
        && incoming_ok
        && selected_edge_ok
        && validate_opaque_fragments(&output, fixture).is_ok();
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        // Changed detach outputs are validated by the selected-owner,
        // reachability, and reopen checks above; source-byte equality is an
        // additional no-op assertion when this lane has no owner.
        output_exact: semantic_ok,
    })
}

/// Exercise the public attach and detach verbs on the same semantic picture.
/// The operations use separate committed transactions so the lane observes
/// both publication boundaries and verifies that the second publication
/// restores the accepted source members and graph closure.
fn execute_same_picture_attach_detach(fixture: &Fixture) -> Result<Execution> {
    let original_drawing = drawing_part(&fixture.package)?;
    let original_content_types = content_types_member(&fixture.package)?;
    let original_source = SourceDrawing::scan(&original_drawing)?;
    let original_anchor = original_source.picture(fixture.picture)?.anchor().clone();

    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut attach_edit = workbook.edit()?;
    attach_edit
        .sheet("Sheet1")?
        .ok_or("Sheet1 is missing")?
        .attach_svg(
            PictureSelector::new(0, fixture.picture),
            SvgInput::borrowed(fixture.payload.as_ref()),
        )?;
    let attached = attach_edit.commit()?.into_workbook().to_plain_bytes()?;
    let attached_drawing = drawing_part(&attached)?;
    let attached_source = SourceDrawing::scan(&attached_drawing)?;
    let attached_picture = attached_source.picture(fixture.picture)?;
    let attached_reopened = Workbook::from_bytes(attached.clone())?.to_plain_bytes()?;
    if attached_reopened != attached
        || !attached_picture.is_direct_embedded_svg()
        || attached_picture.anchor() != &original_anchor
        || !source_anchors_match(&original_source, &attached_source)
        || validate_graph_closure(&attached, 1).is_err()
        || validate_picture_closure(&attached, fixture.picture, true).is_err()
        || validate_opaque_fragments(&attached, fixture).is_err()
        || !svg_payload_matches(&attached, &fixture.payload)
    {
        return Err("same-picture attach publication failed semantic checks".into());
    }

    let attached_workbook = Workbook::from_bytes(attached)?;
    let mut detach_edit = attached_workbook.edit()?;
    detach_edit
        .sheet("Sheet1")?
        .ok_or("Sheet1 is missing after attach")?
        .detach_svg(PictureSelector::new(0, fixture.picture))?;
    let restored = detach_edit.commit()?.into_workbook().to_plain_bytes()?;
    let restored_drawing = drawing_part(&restored)?;
    let restored_source = SourceDrawing::scan(&restored_drawing)?;
    let restored_picture = restored_source.picture(fixture.picture)?;
    let restored_reopened = Workbook::from_bytes(restored.clone())?.to_plain_bytes()?;
    let semantic_ok = !restored_picture.is_direct_embedded_svg()
        && restored_reopened == restored
        && restored_picture.anchor() == &original_anchor
        && source_anchors_match(&original_source, &restored_source)
        && drawing_part(&restored)? == original_drawing
        && content_types_member(&restored)? == original_content_types
        && svg_parts(&restored).is_empty()
        && validate_graph_closure(&restored, 0).is_ok()
        && validate_picture_closure(&restored, fixture.picture, false).is_ok()
        && validate_opaque_fragments(&restored, fixture).is_ok();
    Ok(Execution {
        output: Some(restored),
        semantic_ok,
        output_exact: semantic_ok,
    })
}

/// Attach on two worksheets in one public transaction, then detach the same
/// semantic owners in a second transaction.  The selected picture ordinals
/// deliberately differ so the lane exercises worksheet selection and drawing
/// selection independently: Sheet1/two-cell and Sheet2/one-cell.
fn execute_multisheet_attach_detach(fixture: &Fixture) -> Result<Execution> {
    if !fixture.multisheet || fixture.drawing_count != 2 {
        return Err("multisheet fixture did not contain two drawing owners".into());
    }
    let original_drawing1 = drawing_part(&fixture.package)?;
    let original_drawing2 = drawing_part_at(&fixture.package, DRAWING2)?;
    let original_content_types = content_types_member(&fixture.package)?;
    let source1 = SourceDrawing::scan(&original_drawing1)?;
    let source2 = SourceDrawing::scan(&original_drawing2)?;
    if source1.pictures().len() != fixture.picture_count
        || source2.pictures().len() != fixture.picture_count
        || anchor_kind(source1.picture(0)?.anchor()) != "two_cell"
        || anchor_kind(source2.picture(1)?.anchor()) != "one_cell"
    {
        return Err("multisheet fixture did not expose the expected anchor projections".into());
    }

    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut attach_edit = workbook.edit()?;
    attach_edit
        .sheet("Sheet1")?
        .ok_or("Sheet1 is missing")?
        .attach_svg(
            PictureSelector::new(0, 0),
            SvgInput::borrowed(fixture.payload.as_ref()),
        )?;
    attach_edit
        .sheet("Sheet2")?
        .ok_or("Sheet2 is missing")?
        .attach_svg(
            PictureSelector::new(0, 1),
            SvgInput::borrowed(fixture.payload.as_ref()),
        )?;
    let attached = attach_edit.commit()?.into_workbook().to_plain_bytes()?;
    let attached_drawing1 = drawing_part(&attached)?;
    let attached_drawing2 = drawing_part_at(&attached, DRAWING2)?;
    let attached_source1 = SourceDrawing::scan(&attached_drawing1)?;
    let attached_source2 = SourceDrawing::scan(&attached_drawing2)?;
    let attached_reopened = Workbook::from_bytes(attached.clone())?.to_plain_bytes()?;
    let attached_ok = attached_source1.picture(0)?.is_direct_embedded_svg()
        && !attached_source1.picture(1)?.is_direct_embedded_svg()
        && attached_source2.picture(1)?.is_direct_embedded_svg()
        && !attached_source2.picture(0)?.is_direct_embedded_svg()
        && attached_source1.picture(0)?.anchor() == source1.picture(0)?.anchor()
        && attached_source2.picture(1)?.anchor() == source2.picture(1)?.anchor()
        && source_anchors_match(&source1, &attached_source1)
        && source_anchors_match(&source2, &attached_source2)
        && attached_reopened == attached
        && svg_parts(&attached).len() == 2
        && svg_parts(&attached)
            .iter()
            .all(|part| part.as_slice() == fixture.payload.as_ref())
        && validate_graph_closure(&attached, 2).is_ok()
        && validate_picture_closure_at(&attached, DRAWING, 0, true).is_ok()
        && validate_picture_closure_at(&attached, DRAWING2, 1, true).is_ok()
        && validate_picture_closure_at(&attached, DRAWING, 1, false).is_ok()
        && validate_picture_closure_at(&attached, DRAWING2, 0, false).is_ok()
        && validate_opaque_fragments(&attached, fixture).is_ok();
    if !attached_ok {
        return Err("multisheet attach publication failed semantic checks".into());
    }

    let attached_workbook = Workbook::from_bytes(attached)?;
    let mut detach_edit = attached_workbook.edit()?;
    detach_edit
        .sheet("Sheet1")?
        .ok_or("Sheet1 is missing after attach")?
        .detach_svg(PictureSelector::new(0, 0))?;
    detach_edit
        .sheet("Sheet2")?
        .ok_or("Sheet2 is missing after attach")?
        .detach_svg(PictureSelector::new(0, 1))?;
    let restored = detach_edit.commit()?.into_workbook().to_plain_bytes()?;
    let restored_reopened = Workbook::from_bytes(restored.clone())?.to_plain_bytes()?;
    let restored_ok = drawing_part(&restored)? == original_drawing1
        && drawing_part_at(&restored, DRAWING2)? == original_drawing2
        && restored_reopened == restored
        && content_types_member(&restored)? == original_content_types
        && svg_parts(&restored).is_empty()
        && validate_graph_closure(&restored, 0).is_ok()
        && validate_picture_closure_at(&restored, DRAWING, 0, false).is_ok()
        && validate_picture_closure_at(&restored, DRAWING2, 1, false).is_ok()
        && validate_opaque_fragments(&restored, fixture).is_ok();
    Ok(Execution {
        output: Some(restored),
        semantic_ok: restored_ok,
        output_exact: restored_ok,
    })
}

fn execute_expected_refusal(fixture: &Fixture, kind: RefusalKind) -> Result<LaneExecution> {
    validate_malformed_owner_state(fixture, kind)?;

    let detach_workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let detach_before = detach_workbook.to_plain_bytes()?;
    let detach_result = {
        let mut detach_edit = detach_workbook.edit()?;
        let mut detach_sheet = detach_edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
        detach_sheet
            .detach_svg(PictureSelector::new(0, fixture.picture))
            .map(|_| ())
    };
    let detach_after = detach_workbook.to_plain_bytes()?;
    let detach_refusal = match checked_refusal(kind, detach_result, &detach_before, &detach_after)?
    {
        LaneExecution::ExpectedRefusal(refusal) => refusal,
        LaneExecution::Success(_) => unreachable!("checked_refusal cannot accept an operation"),
    };

    let attach_workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let attach_before = attach_workbook.to_plain_bytes()?;
    let attach_result = {
        let mut attach_edit = attach_workbook.edit()?;
        let mut attach_sheet = attach_edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
        attach_sheet
            .attach_svg(
                PictureSelector::new(0, fixture.picture),
                SvgInput::borrowed(fixture.payload.as_ref()),
            )
            .map(|_| ())
    };
    let attach_after = attach_workbook.to_plain_bytes()?;
    let attach_refusal = match checked_refusal(kind, attach_result, &attach_before, &attach_after)?
    {
        LaneExecution::ExpectedRefusal(refusal) => refusal,
        LaneExecution::Success(_) => unreachable!("checked_refusal cannot accept an operation"),
    };
    Ok(LaneExecution::ExpectedRefusal(ExpectedRefusal {
        kind,
        message: format!(
            "detach: {}; attach: {}",
            detach_refusal.message, attach_refusal.message
        ),
    }))
}

fn composite_caller_limit_ceilings(
    bytes: &[u8],
    payload_len: usize,
) -> Result<Vec<CallerLimitCeiling>> {
    let package = OpcPackage::from_bytes(bytes)?;
    let part_count = package.part_count();
    let part_bytes = package
        .iter_parts()
        .map(|part| part.blob().len())
        .try_fold(0usize, usize::checked_add)
        .ok_or("part byte count overflow")?;
    let relationship_count = package.rels().len()
        + package
            .iter_parts()
            .map(|part| part.rels().len())
            .try_fold(0usize, usize::checked_add)
            .ok_or("relationship count overflow")?;
    let physical = PhysPkgReader::new(bytes)?;
    let member_names = physical.member_names()?;
    let relationship_members = member_names
        .iter()
        .filter(|name| {
            name == &"_rels/.rels" || (name.contains("/_rels/") && name.ends_with(".rels"))
        })
        .cloned()
        .collect::<Vec<_>>();
    let relationship_xml_bytes = relationship_members
        .iter()
        .map(|name| physical.read_member(name).map(|bytes| bytes.len()))
        .collect::<std::result::Result<Vec<_>, _>>()?
        .into_iter()
        .try_fold(0usize, usize::checked_add)
        .ok_or("relationship XML byte count overflow")?;
    let relationship_xml_events = relationship_members
        .iter()
        .map(|name| {
            let bytes = physical.read_member(name)?;
            let mut reader = Reader::from_reader(bytes.as_slice());
            let mut buffer = Vec::new();
            let mut count = 0usize;
            loop {
                count = count
                    .checked_add(1)
                    .ok_or("relationship XML event overflow")?;
                let end = matches!(reader.read_event_into(&mut buffer)?, Event::Eof);
                buffer.clear();
                if end {
                    break;
                }
            }
            Ok::<usize, BoxError>(count)
        })
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .try_fold(0usize, usize::checked_add)
        .ok_or("relationship XML event count overflow")?;
    let content_types = String::from_utf8(physical.read_member("[Content_Types].xml")?)?;
    let content_type_mappings =
        content_types.matches("<Default ").count() + content_types.matches("<Override ").count();
    Ok(vec![
        CallerLimitCeiling {
            name: "parts",
            value: u64::try_from(part_count)?,
        },
        CallerLimitCeiling {
            name: "total_part_bytes",
            value: u64::try_from(
                part_bytes
                    .checked_add(payload_len)
                    .ok_or("total part byte ceiling overflow")?,
            )?,
        },
        CallerLimitCeiling {
            name: "total_relationships",
            value: u64::try_from(relationship_count)?,
        },
        CallerLimitCeiling {
            name: "relationship_xml_bytes",
            value: u64::try_from(relationship_xml_bytes)?,
        },
        CallerLimitCeiling {
            name: "relationship_xml_events",
            value: u64::try_from(relationship_xml_events)?,
        },
        CallerLimitCeiling {
            name: "content_type_mappings",
            value: u64::try_from(content_type_mappings)?,
        },
    ])
}

fn read_limits_for_ceiling(ceiling: CallerLimitCeiling) -> Result<ReadLimits> {
    let value_as_usize = || -> Result<usize> {
        usize::try_from(ceiling.value)
            .map_err(|_| format!("caller limit {} does not fit usize", ceiling.name).into())
    };
    let limits = match ceiling.name {
        "parts" => ReadLimits::builder().max_parts(value_as_usize()?)?,
        "total_part_bytes" => ReadLimits::builder().max_total_part_bytes(ceiling.value)?,
        "total_relationships" => ReadLimits::builder().max_total_relationships(value_as_usize()?)?,
        "relationship_xml_bytes" => {
            ReadLimits::builder().max_total_relationship_xml_bytes(value_as_usize()?)?
        },
        "relationship_xml_events" => {
            ReadLimits::builder().max_total_relationship_xml_events(value_as_usize()?)?
        },
        "content_type_mappings" => {
            ReadLimits::builder().max_content_type_mappings(value_as_usize()?)?
        },
        other => return Err(format!("unknown composite caller limit: {other}").into()),
    };
    Ok(limits.build()?)
}

fn execute_mixed_caps(fixture: &Fixture) -> Result<LaneExecution> {
    let ceilings = fixture
        .caller_limit_ceilings
        .as_ref()
        .ok_or("mixed cap lane did not record requested ceilings")?;
    let limits = ceilings
        .iter()
        .copied()
        .map(read_limits_for_ceiling)
        .collect::<Result<Vec<_>>>()?;
    let mut validated_refusal = None;
    for limit in limits {
        let workbook = Workbook::from_bytes_with_limits(fixture.package.to_vec(), limit)?;
        let before = workbook.to_plain_bytes()?;
        let operation = (|| {
            let mut edit = workbook.edit()?;
            let mut sheet = edit
                .sheet("Sheet1")?
                .ok_or_else(|| XlsxError::Invalid(String::from("Sheet1 is missing")))?;
            sheet.detach_svg(PictureSelector::new(0, fixture.picture))?;
            sheet.attach_svg(
                PictureSelector::new(0, fixture.picture),
                SvgInput::borrowed(fixture.payload.as_ref()),
            )?;
            edit.commit().map(|_| ())
        })();
        let after = workbook.to_plain_bytes()?;
        let refusal = match checked_refusal(RefusalKind::MixedLimit, operation, &before, &after)? {
            LaneExecution::ExpectedRefusal(refusal) => refusal,
            LaneExecution::Success(_) => {
                unreachable!("checked_refusal cannot accept an operation")
            },
        };
        if validated_refusal.is_none() {
            validated_refusal = Some(refusal);
        }
    }
    Ok(LaneExecution::ExpectedRefusal(
        validated_refusal.ok_or("mixed limit lane did not exercise any limit")?,
    ))
}

fn execute_multi_picture_attach(fixture: &Fixture) -> Result<Execution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut edit = workbook.edit()?;
    for picture in 0..fixture.picture_count {
        let mut sheet = edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
        sheet.attach_svg(
            PictureSelector::new(0, picture),
            SvgInput::borrowed(fixture.payload.as_ref()),
        )?;
    }
    let output = edit.commit()?.into_workbook().to_plain_bytes()?;
    let reopened = Workbook::from_bytes(output.clone())?;
    let reopened_bytes = reopened.to_plain_bytes()?;
    let drawing = drawing_part(&output)?;
    let source = SourceDrawing::scan(&drawing)?;
    let pictures_ok = source.pictures().len() == fixture.picture_count
        && source.pictures().iter().all(|picture| {
            picture.is_direct_embedded_svg() && picture.raster_relationship_id() == "rIdRaster"
        });
    let relationship_ids = source
        .pictures()
        .iter()
        .filter_map(|picture| picture.svg_owner().owner()?.embedded_relationship_id())
        .collect::<Vec<_>>();
    let relationships_ok = relationship_ids.len() == fixture.picture_count
        && relationship_ids.iter().enumerate().all(|(index, id)| {
            relationship_ids[..index]
                .iter()
                .all(|previous| previous != id)
        });
    let svg_parts = svg_parts(&output);
    let svg_parts_ok = svg_parts.len() == fixture.picture_count
        && svg_parts
            .iter()
            .all(|part| part.as_slice() == fixture.payload.as_ref());
    let semantic_ok = output != fixture.package.as_ref()
        && reopened_bytes == output
        && pictures_ok
        && relationships_ok
        && svg_parts_ok
        && validate_graph_closure(&output, fixture.picture_count).is_ok()
        && (0..fixture.picture_count)
            .all(|picture| validate_picture_closure(&output, picture, true).is_ok())
        && validate_opaque_fragments(&output, fixture).is_ok();
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        output_exact: semantic_ok,
    })
}

fn validate_multi_picture_attach_output(
    fixture: &Fixture,
    output: &[u8],
    reopened_bytes: &[u8],
) -> Result<bool> {
    let drawing = drawing_part(output)?;
    let source = SourceDrawing::scan(&drawing)?;
    let pictures_ok = source.pictures().len() == fixture.picture_count
        && source.pictures().iter().all(|picture| {
            picture.is_direct_embedded_svg() && picture.raster_relationship_id() == "rIdRaster"
        });
    let relationship_ids = source
        .pictures()
        .iter()
        .filter_map(|picture| picture.svg_owner().owner()?.embedded_relationship_id())
        .collect::<Vec<_>>();
    let relationships_ok = relationship_ids.len() == fixture.picture_count
        && relationship_ids.iter().enumerate().all(|(index, id)| {
            relationship_ids[..index]
                .iter()
                .all(|previous| previous != id)
        });
    let svg_parts = svg_parts(output);
    let svg_parts_ok = svg_parts.len() == fixture.picture_count
        && svg_parts
            .iter()
            .all(|part| part.as_slice() == fixture.payload.as_ref());
    Ok(output != fixture.package.as_ref()
        && reopened_bytes == output
        && pictures_ok
        && relationships_ok
        && svg_parts_ok
        && validate_graph_closure(output, fixture.picture_count).is_ok()
        && (0..fixture.picture_count)
            .all(|picture| validate_picture_closure(output, picture, true).is_ok())
        && validate_opaque_fragments(output, fixture).is_ok())
}

#[derive(Clone, Copy, Debug)]
struct PhaseMeasurement {
    elapsed_ns: u64,
    allocation: AllocDelta,
}

#[derive(Clone, Copy, Debug)]
struct PhaseSample {
    open: PhaseMeasurement,
    stages: PhaseMeasurement,
    commit: PhaseMeasurement,
    firstsave: PhaseMeasurement,
    reopen_secondsave: PhaseMeasurement,
    validation: PhaseMeasurement,
    semantic_ok: bool,
    output_exact: bool,
}

fn measure_phase<T>(
    name: &str,
    operation: impl FnOnce() -> Result<T>,
) -> Result<(T, PhaseMeasurement)> {
    support::reset_counters();
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let result = operation();
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
    let after = AllocSnapshot::now();
    let measurement = PhaseMeasurement {
        elapsed_ns,
        allocation: before.delta(after),
    };
    if !measurement.allocation.balanced()
        || measurement.allocation.invalid
        || measurement.allocation.failed != 0
    {
        return Err(format!(
            "{name} allocator accounting failed: balanced={} invalid={} failed={}",
            measurement.allocation.balanced(),
            measurement.allocation.invalid,
            measurement.allocation.failed,
        )
        .into());
    }
    Ok((result?, measurement))
}

fn signed_live_delta(allocation: AllocDelta) -> i128 {
    i128::from(allocation.live_after) - i128::from(allocation.live_before)
}

fn phase_measurement_json(measurement: PhaseMeasurement) -> String {
    let allocation = measurement.allocation;
    format!(
        concat!(
            "{{\"elapsed_ns\":{},\"requested_alloc_bytes\":{},",
            "\"direct_allocated_bytes\":{},\"realloc_new_bytes\":{},",
            "\"realloc_old_bytes\":{},\"deallocated_bytes\":{},",
            "\"live_before_bytes\":{},\"live_after_bytes\":{},",
            "\"live_delta_bytes\":{},\"retained_live_bytes_after\":{},",
            "\"peak_live_delta_bytes\":{},\"alloc_balance_ok\":{},",
            "\"alloc_invalid\":{},\"alloc_failed\":{}}}"
        ),
        measurement.elapsed_ns,
        allocation.requested(),
        allocation.direct,
        allocation.realloc_new,
        allocation.realloc_old,
        allocation.deallocated,
        allocation.live_before,
        allocation.live_after,
        signed_live_delta(allocation),
        allocation.live_after,
        allocation.peak_delta,
        allocation.balanced(),
        allocation.invalid,
        allocation.failed,
    )
}

fn phase_sample_json(sample: PhaseSample) -> String {
    format!(
        concat!(
            "{{\"semantic_ok\":{},\"output_exact\":{},\"phases\":{{",
            "\"open\":{},\"stages\":{},\"commit\":{},",
            "\"firstsave\":{},\"reopen_secondsave\":{},\"validation\":{}",
            "}}}}"
        ),
        sample.semantic_ok,
        sample.output_exact,
        phase_measurement_json(sample.open),
        phase_measurement_json(sample.stages),
        phase_measurement_json(sample.commit),
        phase_measurement_json(sample.firstsave),
        phase_measurement_json(sample.reopen_secondsave),
        phase_measurement_json(sample.validation),
    )
}

fn phase_decomposition_sample(fixture: &Fixture) -> Result<PhaseSample> {
    let (source_workbook, open) = measure_phase("open", || {
        Ok(Workbook::from_bytes(fixture.package.to_vec())?)
    })?;
    let (edit, stages) = measure_phase("stages", || {
        let mut edit = source_workbook.edit()?;
        for picture in 0..fixture.picture_count {
            let mut sheet = edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
            sheet.attach_svg(
                PictureSelector::new(0, picture),
                SvgInput::borrowed(fixture.payload.as_ref()),
            )?;
        }
        Ok(edit)
    })?;
    let (committed, commit) = measure_phase("commit", || Ok(edit.commit()?.into_workbook()))?;
    let (first_saved, firstsave) = measure_phase("firstsave", || Ok(committed.to_plain_bytes()?))?;
    // The acceptance path drops the committed workbook as the chained
    // first-save temporary ends. Keep this boundary outside both phase clocks;
    // the resulting live set matches that lifecycle before reopen.
    drop(committed);
    let ((reopened, second_saved), reopen_secondsave) = measure_phase("reopen_secondsave", || {
        let reopened = Workbook::from_bytes(first_saved.clone())?;
        let second_saved = reopened.to_plain_bytes()?;
        Ok((reopened, second_saved))
    })?;
    let (semantic_ok, validation) = measure_phase("validation", || {
        validate_multi_picture_attach_output(fixture, &first_saved, &second_saved)
    })?;
    let sample = PhaseSample {
        open,
        stages,
        commit,
        firstsave,
        reopen_secondsave,
        validation,
        semantic_ok,
        output_exact: semantic_ok,
    };
    drop(reopened);
    drop(source_workbook);
    Ok(sample)
}

fn phase_receipt_json(
    lane: &str,
    fixture: &Fixture,
    warmup: usize,
    samples: &[PhaseSample],
) -> String {
    let mut output = String::new();
    write!(
        output,
        concat!(
            "{{\n",
            "  \"schema\":\"xlsx-svg-lifecycle-phase-profile-v1\",\n",
            "  \"lane\":\"{}\",\n",
            "  \"picture_count\":{},\n",
            "  \"input_bytes\":{},\n",
            "  \"input_hash_fnv1a64\":{},\n",
            "  \"input_sha256\":\"{}\",\n",
            "  \"warmup\":{},\n",
            "  \"sample_count\":{},\n",
            "  \"expected_success\":true,\n",
            "  \"phase_order\":[\"open\",\"stages\",\"commit\",\"firstsave\",\"reopen_secondsave\",\"validation\"],\n",
            "  \"allocation_note\":\"requested_alloc_bytes is phase-local; live_after includes objects retained for later phases; phase values must not be summed or subtracted across unlike retained live sets\",\n",
            "  \"semantic_checks\":[\"drawing_picture_inventory\",\"anchor_and_raster_relationships\",\"unique_svg_relationships\",\"svg_media_bytes\",\"reopen_byte_identity\",\"graph_closure\",\"picture_closure\",\"opaque_fragments\"],\n",
            "  \"samples\":[\n"
        ),
        support::json_escape(lane),
        fixture.picture_count,
        fixture.input_bytes,
        fixture.input_hash,
        support::json_escape(&fixture.input_sha256),
        warmup,
        samples.len(),
    )
    .expect("String cannot fail");
    for (index, sample) in samples.iter().enumerate() {
        let comma = if index + 1 == samples.len() { "" } else { "," };
        writeln!(output, "    {}{comma}", phase_sample_json(*sample)).expect("String cannot fail");
    }
    output.push_str("  ]\n}\n");
    output
}

pub fn run_phase_decomposition(
    picture_count: usize,
    warmup: usize,
    samples: usize,
) -> Result<String> {
    if !matches!(picture_count, 16 | 64 | 256) {
        return Err("phase decomposition picture count must be 16, 64, or 256".into());
    }
    if warmup < 2 || samples < 20 {
        return Err("phase decomposition requires at least 2 warmups and 20 samples".into());
    }
    let lane = format!("multi_picture_same_drawing_{picture_count}");
    let fixture = fixture_for_lane(&lane)?;
    for _ in 0..warmup {
        let live_before = AllocSnapshot::now().live;
        let sample = phase_decomposition_sample(&fixture)?;
        if !sample.semantic_ok || !sample.output_exact {
            return Err("phase decomposition warm-up semantic validation failed".into());
        }
        assert_sample_dropped(live_before)?;
    }
    let mut receipts = Vec::with_capacity(samples);
    let live_before_samples = AllocSnapshot::now().live;
    for _ in 0..samples {
        let sample = phase_decomposition_sample(&fixture)?;
        if !sample.semantic_ok || !sample.output_exact {
            return Err("phase decomposition semantic validation failed".into());
        }
        assert_sample_dropped(live_before_samples)?;
        receipts.push(sample);
    }
    Ok(phase_receipt_json(&lane, &fixture, warmup, &receipts))
}

fn assert_sample_dropped(live_before: u64) -> Result<()> {
    let after = AllocSnapshot::now();
    if after.live != live_before || after.invalid || after.failed != 0 {
        return Err(format!(
            "phase sample did not return to its live boundary: before={live_before} after={} invalid={} failed={}",
            after.live, after.invalid, after.failed
        )
        .into());
    }
    Ok(())
}

fn execute_multi_picture_detach(fixture: &Fixture) -> Result<Execution> {
    let before_drawing = drawing_part(&fixture.package)?;
    let before_source = SourceDrawing::scan(&before_drawing)?;
    let before_ok = before_source.pictures().len() == fixture.picture_count
        && before_source
            .pictures()
            .iter()
            .all(|picture| picture.is_direct_embedded_svg());
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut edit = workbook.edit()?;
    for picture in 0..fixture.picture_count {
        let mut sheet = edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
        sheet.detach_svg(PictureSelector::new(0, picture))?;
    }
    let output = edit.commit()?.into_workbook().to_plain_bytes()?;
    let reopened = Workbook::from_bytes(output.clone())?;
    let reopened_bytes = reopened.to_plain_bytes()?;
    let after_drawing = drawing_part(&output)?;
    let after_source = SourceDrawing::scan(&after_drawing)?;
    let pictures_ok = after_source.pictures().len() == fixture.picture_count
        && after_source.pictures().iter().all(|picture| {
            !picture.is_direct_embedded_svg() && picture.raster_relationship_id() == "rIdRaster"
        });
    let semantic_ok = before_ok
        && output != fixture.package.as_ref()
        && reopened_bytes == output
        && pictures_ok
        && !has_svg_part(&output)
        && validate_graph_closure(&output, 0).is_ok()
        && (0..fixture.picture_count)
            .all(|picture| validate_picture_closure(&output, picture, false).is_ok());
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        output_exact: semantic_ok,
    })
}

fn drawing_part(bytes: &[u8]) -> Result<Vec<u8>> {
    drawing_part_at(bytes, DRAWING)
}

fn drawing_part_at(bytes: &[u8], name: &str) -> Result<Vec<u8>> {
    Ok(OpcPackage::from_bytes(bytes)?
        .get_part(&PackURI::new(name)?)?
        .blob()
        .to_vec())
}

fn content_types_member(bytes: &[u8]) -> Result<Vec<u8>> {
    Ok(PhysPkgReader::new(bytes)?.read_member("[Content_Types].xml")?)
}

fn source_anchors_match(original: &SourceDrawing, candidate: &SourceDrawing) -> bool {
    original.pictures().len() == candidate.pictures().len()
        && original
            .pictures()
            .iter()
            .zip(candidate.pictures())
            .all(|(before, after)| before.anchor() == after.anchor())
}

fn fragment_count(haystack: &[u8], needle: &[u8]) -> usize {
    if needle.is_empty() || haystack.len() < needle.len() {
        return 0;
    }
    haystack
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

fn opaque_extension_fragment() -> &'static [u8] {
    br#"<a:ext uri="urn:litchi:future-extension"><!--future-comment--><future:payload future:keep="yes"/></a:ext>"#
}

fn validate_opaque_fragments(bytes: &[u8], fixture: &Fixture) -> Result<()> {
    if !fixture.opaque_extension && !fixture.opaque_descendants {
        return Ok(());
    }
    let drawing = drawing_part(bytes)?;
    let expected_occurrences = fixture
        .picture_count
        .checked_mul(fixture.drawing_count)
        .ok_or("opaque-fragment occurrence count overflow")?;
    if fixture.opaque_extension {
        let count = fragment_count(&drawing, opaque_extension_fragment());
        if count != fixture.picture_count {
            return Err(format!(
                "opaque extension count changed: expected {}, found {count}",
                fixture.picture_count
            )
            .into());
        }
    }
    if fixture.opaque_descendants {
        let descendants = namespace_descendants();
        let count = fragment_count(&drawing, descendants.as_bytes());
        if count != fixture.picture_count {
            return Err(format!(
                "opaque descendant fragment count changed: expected {}, found {count}",
                fixture.picture_count
            )
            .into());
        }
    }
    if fixture.drawing_count > 1 {
        let second = drawing_part_at(bytes, DRAWING2)?;
        if fixture.opaque_extension {
            let count = fragment_count(&second, opaque_extension_fragment());
            if count != expected_occurrences.saturating_sub(fixture.picture_count) {
                return Err(format!(
                    "multisheet opaque extension count changed: expected {}, found {count}",
                    expected_occurrences.saturating_sub(fixture.picture_count)
                )
                .into());
            }
        }
        if fixture.opaque_descendants {
            let count = fragment_count(&second, namespace_descendants().as_bytes());
            if count != expected_occurrences.saturating_sub(fixture.picture_count) {
                return Err(format!(
                    "multisheet opaque descendant count changed: expected {}, found {count}",
                    expected_occurrences.saturating_sub(fixture.picture_count)
                )
                .into());
            }
        }
    }
    Ok(())
}

fn validate_strict_host(bytes: &[u8], expected_svg_parts: usize) -> Result<()> {
    let package = OpcPackage::from_bytes(bytes)?;
    let sheet = package.get_part(&PackURI::new(SHEET)?)?;
    let drawing = package.get_part(&PackURI::new(DRAWING)?)?;
    let sheet_xml = sheet.blob();
    let drawing_xml = drawing.blob();
    for marker in [
        b"http://purl.oclc.org/ooxml/spreadsheetml/main".as_slice(),
        STRICT_REL.as_bytes(),
    ] {
        if !drawing_xml
            .windows(marker.len())
            .any(|window| window == marker)
            && !sheet_xml
                .windows(marker.len())
                .any(|window| window == marker)
        {
            return Err(format!(
                "strict host marker missing: {}",
                String::from_utf8_lossy(marker)
            )
            .into());
        }
    }
    for marker in [STRICT_XDR.as_bytes(), STRICT_A.as_bytes()] {
        if !drawing_xml
            .windows(marker.len())
            .any(|window| window == marker)
        {
            return Err(format!(
                "strict drawing marker missing: {}",
                String::from_utf8_lossy(marker)
            )
            .into());
        }
    }
    let drawing_relationship = sheet
        .rels()
        .get("rIdDrawing")
        .ok_or("strict worksheet drawing relationship is missing")?;
    if drawing_relationship.reltype() != rt::STRICT_DRAWING {
        return Err("strict worksheet drawing relationship was rewritten".into());
    }
    if drawing.rels().len() != expected_svg_parts.saturating_add(1)
        || drawing
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() != rt::STRICT_IMAGE)
    {
        return Err("strict drawing image relationships were rewritten".into());
    }
    Ok(())
}

fn validate_incoming_svg_edge(bytes: &[u8]) -> Result<()> {
    let package = OpcPackage::from_bytes(bytes)?;
    let owner = package.get_part(&PackURI::new(OTHER)?)?;
    let retained = owner.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|target| target.as_str() == "/xl/media/image2.svg")
    });
    if retained {
        Ok(())
    } else {
        Err("incoming SVG edge was lost".into())
    }
}

fn validate_selected_drawing_svg_edge_removed(bytes: &[u8]) -> Result<()> {
    let package = OpcPackage::from_bytes(bytes)?;
    let drawing = package.get_part(&PackURI::new(DRAWING)?)?;
    let retained = drawing.rels().iter().any(|relationship| {
        relationship.r_id() == "rIdSvg"
            || relationship
                .target_ref()
                .split(['?', '#'])
                .next()
                .is_some_and(|target| target.to_ascii_lowercase().ends_with(".svg"))
    });
    if retained {
        Err("selected drawing SVG relationship was not removed".into())
    } else {
        Ok(())
    }
}

/// Validate the complete package closure used by the lifecycle operation.
///
/// The public XLSX transaction owns this proof; the adapter repeats the
/// observable graph checks after publication so a receipt cannot report a
/// successful semantic edit merely because the drawing XML happened to parse.
fn validate_graph_closure(bytes: &[u8], expected_svg_parts: usize) -> Result<()> {
    let package = OpcPackage::from_bytes(bytes)?;
    let mut targets = Vec::new();
    for relationship in package.rels().iter() {
        if !relationship.is_external() {
            targets.push(relationship.target_partname()?);
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external() {
                targets.push(relationship.target_partname()?);
            }
        }
    }
    for target in &targets {
        package
            .get_part(target)
            .map_err(|error| format!("dangling relationship target {target}: {error}"))?;
    }

    let svg_names = package
        .iter_parts()
        .filter(|part| part.content_type() == "image/svg+xml")
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    if svg_names.len() != expected_svg_parts {
        return Err(format!(
            "expected {expected_svg_parts} SVG parts, found {}",
            svg_names.len()
        )
        .into());
    }
    let mut incoming = HashMap::<String, usize>::new();
    let mut check_svg_relationship = |relationship: &litchi_opc::Relationship| -> Result<()> {
        if relationship.is_external() {
            return Ok(());
        }
        let target = relationship.target_partname()?;
        if svg_names.contains(target.as_str()) {
            if relationship.reltype() != rt::IMAGE && relationship.reltype() != rt::STRICT_IMAGE {
                return Err(format!(
                    "SVG target {} has non-image relationship {}",
                    target,
                    relationship.reltype()
                )
                .into());
            }
            *incoming.entry(target.to_string()).or_default() += 1;
        }
        Ok(())
    };
    for relationship in package.rels().iter() {
        check_svg_relationship(relationship)?;
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            check_svg_relationship(relationship)?;
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == "image/svg+xml")
    {
        if !part.rels().is_empty() {
            return Err(format!("SVG part {} has outbound relationships", part.partname()).into());
        }
        if incoming.get(part.partname().as_str()).copied().unwrap_or(0) == 0 {
            return Err(format!("SVG part {} is orphaned", part.partname()).into());
        }
    }

    // A source-backed package may use either a default .svg mapping or one
    // override per leaf.  Whichever form is present, no leaf may acquire a
    // duplicate override during a transaction.
    let physical = PhysPkgReader::new(bytes)?;
    let content_types = String::from_utf8(physical.read_member("[Content_Types].xml")?)?;
    for name in &svg_names {
        let needle = format!("PartName=\"{name}\"");
        if content_types.matches(&needle).count() > 1 {
            return Err(format!("duplicate content-type override for {name}").into());
        }
    }
    Ok(())
}

fn validate_picture_closure(bytes: &[u8], picture: usize, expected_attached: bool) -> Result<()> {
    validate_picture_closure_at(bytes, DRAWING, picture, expected_attached)
}

fn validate_picture_closure_at(
    bytes: &[u8],
    drawing_name: &str,
    picture: usize,
    expected_attached: bool,
) -> Result<()> {
    let drawing = drawing_part_at(bytes, drawing_name)?;
    let source = SourceDrawing::scan(&drawing)?;
    let selected = source.picture(picture)?;
    if selected.anchor_range().start >= selected.anchor_range().end {
        return Err("selected anchor range is not ordered".into());
    }
    let package = OpcPackage::from_bytes(bytes)?;
    let drawing_part = package.get_part(&PackURI::new(drawing_name)?)?;
    let raster = drawing_part
        .rels()
        .get(selected.raster_relationship_id())
        .ok_or("selected raster relationship is missing")?;
    if raster.is_external()
        || (raster.reltype() != rt::IMAGE && raster.reltype() != rt::STRICT_IMAGE)
    {
        return Err("selected raster relationship is not an internal image".into());
    }
    let raster_target = raster.target_partname()?;
    let raster_part = package.get_part(&raster_target)?;
    if raster_part.content_type() != "image/png"
        || !raster_target.as_str().starts_with("/xl/media/")
        || raster_part.blob().is_empty()
    {
        return Err("selected raster fallback is not an internal PNG leaf".into());
    }
    if selected.is_direct_embedded_svg() != expected_attached {
        return Err("selected SVG owner state differs from expected state".into());
    }
    if expected_attached {
        let owner = selected
            .svg_owner()
            .owner()
            .ok_or("selected embedded SVG owner is missing")?;
        let relationship_id = owner
            .embedded_relationship_id()
            .ok_or("selected SVG owner has no embedded relationship")?;
        let relationship = drawing_part
            .rels()
            .get(relationship_id)
            .ok_or("selected SVG relationship is missing")?;
        if relationship.is_external()
            || (relationship.reltype() != rt::IMAGE && relationship.reltype() != rt::STRICT_IMAGE)
        {
            return Err("selected SVG relationship is not an internal image".into());
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type() != "image/svg+xml"
            || !target.as_str().starts_with("/xl/media/")
            || !part.rels().is_empty()
        {
            return Err("selected SVG target is not a closed internal SVG leaf".into());
        }
    }
    Ok(())
}

fn svg_payload_matches(bytes: &[u8], expected: &[u8]) -> bool {
    let Ok(package) = OpcPackage::from_bytes(bytes) else {
        return false;
    };
    package
        .iter_parts()
        .find(|part| part.content_type() == "image/svg+xml")
        .is_some_and(|part| part.blob() == expected)
}

fn has_svg_part(bytes: &[u8]) -> bool {
    !svg_parts(bytes).is_empty()
}

fn svg_parts(bytes: &[u8]) -> Vec<Vec<u8>> {
    let Ok(package) = OpcPackage::from_bytes(bytes) else {
        return Vec::new();
    };
    package
        .iter_parts()
        .filter(|part| part.content_type() == "image/svg+xml")
        .map(|part| part.blob().to_vec())
        .collect()
}

fn anchor_kind(anchor: &DrawingAnchor) -> &'static str {
    match anchor {
        DrawingAnchor::TwoCell { .. } => "two_cell",
        DrawingAnchor::OneCell { .. } => "one_cell",
        DrawingAnchor::Absolute { .. } => "absolute",
    }
}

fn lane_anchor_kind(lane: &str) -> Option<&'static str> {
    if lane.contains("two_cell") {
        Some("two_cell")
    } else if lane.contains("one_cell") {
        Some("one_cell")
    } else if lane.contains("absolute") {
        Some("absolute")
    } else {
        None
    }
}

fn sample_json(
    _lane: &str,
    expected_success: bool,
    elapsed_ns: u64,
    allocation: AllocDelta,
    outcome: Outcome,
) -> String {
    let (actual_success, semantic_ok, output_exact, error) = match outcome {
        Outcome::Success {
            semantic_ok,
            output_exact,
        } => {
            if expected_success {
                (true, semantic_ok, output_exact, None)
            } else {
                (
                    true,
                    false,
                    false,
                    Some(ErrorInfo {
                        class: String::from("unexpected-success"),
                        message: String::from("lane accepted an input expected to refuse"),
                    }),
                )
            }
        },
        Outcome::ExpectedRefusal(refusal) => {
            if expected_success {
                (
                    false,
                    false,
                    false,
                    Some(ErrorInfo {
                        class: String::from("unexpected-refusal"),
                        message: refusal.message,
                    }),
                )
            } else {
                (
                    false,
                    true,
                    true,
                    Some(ErrorInfo {
                        class: refusal.kind.as_str().to_owned(),
                        message: refusal.message,
                    }),
                )
            }
        },
        Outcome::Failure(error) => (
            false,
            false,
            false,
            Some(ErrorInfo {
                class: String::from("harness-error"),
                message: error.to_string(),
            }),
        ),
    };
    let error_json = error.map_or_else(
        || String::from("null"),
        |error| {
            format!(
                r#"{{"class":"{}","message":"{}"}}"#,
                support::json_escape(&error.class),
                support::json_escape(&error.message)
            )
        },
    );
    format!(
        r#"{{"elapsed_ns":{elapsed_ns},"requested_alloc_bytes":{},"direct_allocated_bytes":{},"realloc_new_bytes":{},"realloc_old_bytes":{},"deallocated_bytes":{},"live_before":{},"live_after":{},"peak_live_delta":{},"alloc_balance_ok":{},"alloc_invalid":{},"alloc_failed":{},"expected_success":{},"actual_success":{},"semantic_ok":{},"output_exact":{},"error":{error_json},"stage_ns":null,"source_bytes_read":null,"range_read_count":null,"range_request_bytes":null,"bytes_decompressed":null,"bytes_recompressed":null,"bytes_copied":null,"source_copy_bytes":null,"staged_bytes":null,"output_bytes":null,"write_call_count":null,"write_call_bytes":null,"hardware_counters":null}}"#,
        allocation.requested(),
        allocation.direct,
        allocation.realloc_new,
        allocation.realloc_old,
        allocation.deallocated,
        allocation.live_before,
        allocation.live_after,
        allocation.peak_delta,
        allocation.balanced(),
        allocation.invalid,
        allocation.failed,
        expected_success,
        actual_success,
        semantic_ok,
        output_exact,
    )
}

fn receipt_json(lane: &str, fixture: &Fixture, warmup: usize, samples: &[String]) -> String {
    let mut output = String::new();
    let namespace_generated_bindings = fixture
        .namespace_generated_bindings
        .map_or_else(|| String::from("null"), |value| value.to_string());
    let namespace_active_bindings = fixture
        .namespace_active_bindings
        .map_or_else(|| String::from("null"), |value| value.to_string());
    let namespace_active_limit = fixture
        .namespace_active_limit
        .map_or_else(|| String::from("null"), |value| value.to_string());
    let caller_limits = match fixture.caller_limit_profile {
        None => String::from("null"),
        Some("svg_input_bytes") => {
            format!(r#"{{"kind":"svg_input_bytes","max":{SVG_INPUT_LIMIT}}}"#)
        },
        Some("composite_read_limits") => {
            let ceilings = fixture
                .caller_limit_ceilings
                .as_ref()
                .expect("mixed cap fixture records caller-limit ceilings");
            let ceilings_json = ceilings
                .iter()
                .map(|ceiling| {
                    format!(r#"{{"name":"{}","value":{}}}"#, ceiling.name, ceiling.value)
                })
                .collect::<Vec<_>>()
                .join(",");
            format!(r#"{{"kind":"composite_read_limits","ceilings":[{ceilings_json}]}}"#)
        },
        Some(other) => format!(
            r#"{{"kind":"unknown","name":"{}"}}"#,
            support::json_escape(other)
        ),
    };
    write!(
        output,
        "{{\n  \"schema\":\"xlsx-svg-lifecycle-profile-v1\",\n  \"lane\":\"{}\",\n  \"input_bytes\":{},\n  \"input_hash_fnv1a64\":{},\n  \"input_sha256\":\"{}\",\n  \"namespace_generated_bindings\":{},\n  \"namespace_active_bindings\":{},\n  \"namespace_active_limit\":{},\n  \"caller_limits\":{},\n  \"warmup\":{},\n  \"sample_count\":{},\n  \"expected_success\":{},\n  \"samples\":[\n",
        support::json_escape(lane),
        fixture.input_bytes,
        fixture.input_hash,
        support::json_escape(&fixture.input_sha256),
        namespace_generated_bindings,
        namespace_active_bindings,
        namespace_active_limit,
        caller_limits,
        warmup,
        samples.len(),
        fixture.expected_success,
    )
    .expect("String cannot fail");
    for (index, sample) in samples.iter().enumerate() {
        let comma = if index + 1 == samples.len() { "" } else { "," };
        writeln!(output, "    {sample}{comma}").expect("String cannot fail");
    }
    output.push_str("  ]\n}\n");
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    fn zero_alloc() -> AllocDelta {
        AllocDelta {
            calls: 0,
            realloc_calls: 0,
            dealloc_calls: 0,
            direct: 0,
            realloc_old: 0,
            realloc_new: 0,
            deallocated: 0,
            live_before: 0,
            live_after: 0,
            peak_delta: 0,
            failed: 0,
            invalid: false,
        }
    }

    fn caller_limit_error() -> XlsxError {
        XlsxError::Invalid(String::from("SVG attachment input exceeds 32 bytes"))
    }

    #[test]
    fn exact_opaque_fragment_checks_reject_near_matches() {
        let exact = opaque_extension_fragment();
        let mut haystack = exact.to_vec();
        haystack.extend_from_slice(b"\n");
        let mut near = exact.to_vec();
        let index = near
            .iter()
            .position(|byte| *byte == b'y')
            .expect("opaque marker has a mutable attribute value");
        near[index] = b'n';
        haystack.extend_from_slice(&near);
        assert_eq!(fragment_count(&haystack, exact), 1);

        let descendants = namespace_descendants();
        assert_eq!(
            fragment_count(descendants.as_bytes(), descendants.as_bytes()),
            1
        );
        let mut changed = descendants.into_bytes();
        let changed_index = changed
            .windows(b"index=\"17\"".len())
            .position(|window| window == b"index=\"17\"")
            .expect("descendant marker exists");
        changed[changed_index + b"index=\"".len()] = b'8';
        assert_eq!(
            fragment_count(&changed, namespace_descendants().as_bytes()),
            0
        );
    }

    #[test]
    fn coverage_registry_counts_new_acceptance_and_exploratory_lanes() {
        assert_eq!(LANES.len(), 69);
        assert_eq!(EXPLORATORY_LANES.len(), 7);
        assert_eq!(FIXED_ROOT_NAMESPACE_BINDINGS, 7);
        for lane in [
            "clone_captured_owner_small",
            "clone_captured_owner_large",
            "strict_attach_end_to_end_two_cell_small",
            "strict_detach_end_to_end_shared_final_two_cell",
            "incoming_edge_shared_final_two_cell",
        ] {
            assert!(is_known_lane(lane), "missing coverage lane: {lane}");
        }
    }

    #[test]
    fn removed_selected_edge_cannot_survive_as_an_external_relationship() {
        let bytes = package_bytes(FixtureSpec::default(), b"<svg/>").unwrap();
        validate_selected_drawing_svg_edge_removed(&bytes).unwrap();
        let mut package = OpcPackage::from_bytes(&bytes).unwrap();
        package
            .get_part_mut(&PackURI::new(DRAWING).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "https://example.invalid/renamed-resource.svg".to_owned(),
                "rIdSvgExternal".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let output = PackageWriter::to_bytes(&package).unwrap();
        assert!(validate_selected_drawing_svg_edge_removed(&output).is_err());
    }

    #[test]
    fn checked_refusal_rejects_acceptance_mutation_and_setup_errors() {
        let before = b"source";
        assert!(checked_refusal(RefusalKind::CallerLimit, Ok(()), before, before,).is_err());

        let validated = checked_refusal(
            RefusalKind::CallerLimit,
            Err(caller_limit_error()),
            before,
            before,
        )
        .expect("typed caller-limit refusal should validate");
        assert!(matches!(
            validated,
            LaneExecution::ExpectedRefusal(ExpectedRefusal {
                kind: RefusalKind::CallerLimit,
                ..
            })
        ));

        let changed = checked_refusal(
            RefusalKind::CallerLimit,
            Err(caller_limit_error()),
            before,
            b"changed",
        )
        .expect_err("source mutation must not be accepted as a refusal");
        assert!(changed.to_string().contains("changed source bytes"));

        let setup = checked_refusal(
            RefusalKind::CallerLimit,
            Err(XlsxError::Invalid(String::from("worksheet is missing"))),
            before,
            before,
        )
        .expect_err("unrelated setup errors must not be accepted as a refusal");
        assert!(setup.to_string().contains("unexpected API error"));
    }

    #[test]
    fn sample_json_keeps_untyped_failures_and_unexpected_acceptance_failed() {
        let accepted = sample_json(
            "limit_small",
            false,
            0,
            zero_alloc(),
            Outcome::Success {
                semantic_ok: true,
                output_exact: true,
            },
        );
        assert!(accepted.contains("\"actual_success\":true"));
        assert!(accepted.contains("\"semantic_ok\":false"));
        assert!(accepted.contains("\"class\":\"unexpected-success\""));

        let setup = sample_json(
            "limit_small",
            false,
            0,
            zero_alloc(),
            Outcome::Failure(Box::new(std::io::Error::other("setup failure"))),
        );
        assert!(setup.contains("\"actual_success\":false"));
        assert!(setup.contains("\"semantic_ok\":false"));
        assert!(setup.contains("\"output_exact\":false"));
        assert!(setup.contains("\"class\":\"harness-error\""));

        let refusal = sample_json(
            "limit_small",
            false,
            0,
            zero_alloc(),
            Outcome::ExpectedRefusal(ExpectedRefusal {
                kind: RefusalKind::CallerLimit,
                message: String::from("SVG attachment input exceeds 32 bytes"),
            }),
        );
        assert!(refusal.contains("\"actual_success\":false"));
        assert!(refusal.contains("\"semantic_ok\":true"));
        assert!(refusal.contains("\"output_exact\":true"));
        assert!(refusal.contains("\"class\":\"caller_limit\""));
    }

    #[test]
    fn exploratory_attach_validation_matches_acceptance_predicate() {
        for picture_count in [16, 64, 256] {
            let lane = format!("multi_picture_same_drawing_{picture_count}");
            let fixture = fixture_for_lane(&lane).expect("deterministic decomposition fixture");
            let execution = execute_multi_picture_attach(&fixture).expect("acceptance attach");
            let expected = execution.semantic_ok;
            let output = execution.output.expect("acceptance output");
            let reopened = Workbook::from_bytes(output.clone())
                .expect("reopen acceptance output")
                .to_plain_bytes()
                .expect("second serialization");
            let exploratory = validate_multi_picture_attach_output(&fixture, &output, &reopened)
                .expect("exploratory validation");
            assert_eq!(exploratory, expected, "picture count {picture_count}");
        }
    }
}
