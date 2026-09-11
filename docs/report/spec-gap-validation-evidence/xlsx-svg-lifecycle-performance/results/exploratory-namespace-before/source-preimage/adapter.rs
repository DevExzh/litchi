//! XLSX profile fixture generation, public-API calls, and receipt encoding.
//!
//! The fixture builders are adapted from
//! `crates/litchi-xlsx/tests/drawing_svg_lifecycle.rs`. They remain synthetic,
//! bounded inputs and are kept here so the profile does not depend on test
//! implementation details. The timed paths call only the public workbook,
//! drawing scanner, selector, and worksheet transaction APIs.

use std::error::Error as StdError;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::hint::black_box;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::drawing::{
    DrawingAnchor, PictureSelector, SourceDrawing, SvgInput, SvgOwnerState,
};
use litchi_xlsx::{Package, Workbook};

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
    "detach_end_to_end_shared_first_two_cell",
    "detach_end_to_end_shared_first_one_cell",
    "detach_end_to_end_shared_first_absolute",
    "detach_end_to_end_shared_final_two_cell",
    "detach_end_to_end_shared_final_one_cell",
    "detach_end_to_end_shared_final_absolute",
    "detach_end_to_end_distinct_two_cell_small",
    "detach_end_to_end_distinct_two_cell_large",
    "detach_end_to_end_distinct_one_cell_small",
    "detach_end_to_end_distinct_one_cell_large",
    "detach_end_to_end_distinct_absolute_small",
    "detach_end_to_end_distinct_absolute_large",
    "noop_detach_two_cell",
    "noop_detach_one_cell",
    "noop_detach_absolute",
    "limit_small",
    "limit_large",
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
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const RASTER: &str = "/xl/media/image1.png";
const OTHER: &str = "/xl/opaque-owner.xml";
const SMALL_PAYLOAD: usize = 512;
const LARGE_PAYLOAD: usize = 65_536;
const LARGE_RASTER: usize = 65_536;
const SVG_INPUT_LIMIT: usize = 32 * 1024 * 1024;

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
    incoming_svg_edge: bool,
    namespace_heavy: bool,
    namespace_root_heavy: bool,
    namespace_limit: bool,
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
            incoming_svg_edge: false,
            namespace_heavy: false,
            namespace_root_heavy: false,
            namespace_limit: false,
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
    pub input_bytes: u64,
    pub input_hash: u64,
    pub expected_success: bool,
    pub refusal_class: Option<&'static str>,
    pub picture: usize,
}

#[derive(Clone, Debug)]
struct Execution {
    output: Option<Vec<u8>>,
    semantic_ok: bool,
    output_exact: bool,
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
    Failure(BoxError),
}

pub fn fixture_for_lane(lane: &str) -> Result<Fixture> {
    if lane == "capture_native_fixture" {
        return native_fixture();
    }
    let mut spec = FixtureSpec::default();
    let picture = picture_index(lane);
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
    if lane.starts_with("attach_end_to_end") || lane.starts_with("noop_detach") {
        spec.svg = SvgVariant::None;
        spec.targets = TargetTopology::None;
    }
    if lane.starts_with("detach_end_to_end_shared") {
        spec.svg = SvgVariant::Embedded;
        spec.targets = TargetTopology::Shared;
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
    let mut package = package_bytes(spec, &payload)?;
    if lane.starts_with("detach_end_to_end_shared_final") {
        for index in 0..spec.picture_count {
            if index != picture {
                package = detach_once(&package, index)?;
            }
        }
    }
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
    );
    Ok(Fixture {
        input_bytes: u64::try_from(identity.len())?,
        input_hash: support::fnv1a64(&identity),
        package: Arc::from(package.into_boxed_slice()),
        payload: Arc::from(payload.into_boxed_slice()),
        picture_count: spec.picture_count,
        expected_success,
        refusal_class: refusal_class(lane),
        picture,
    })
}

fn refusal_class(lane: &str) -> Option<&'static str> {
    match lane {
        "namespace_limit_refusal" => Some("namespace_limit"),
        "limit_small" | "limit_large" => Some("caller_limit"),
        "malformed_duplicate_owner" => Some("duplicate_owner"),
        "malformed_mce_owner" => Some("mce_ancestry"),
        "malformed_linked_owner" => Some("linked_owner"),
        _ => None,
    }
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
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../../../")
        .canonicalize()?;
    let path =
        root.join("3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx");
    let bytes = fs::read(path)?;
    let input_bytes = u64::try_from(bytes.len())?;
    let input_hash = support::fnv1a64(&bytes);
    Ok(Fixture {
        input_bytes,
        input_hash,
        package: Arc::from(bytes.into_boxed_slice()),
        payload: Arc::from(Vec::<u8>::new().into_boxed_slice()),
        picture_count: 2,
        expected_success: true,
        refusal_class: None,
        picture: 0,
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

fn namespace_pressure() -> String {
    let mut output = String::new();
    output.push_str("<n:scope>");
    for level in (0..64).rev() {
        write!(output, "<n:l{level}").expect("String cannot fail");
        for binding in 0..256 {
            write!(
                output,
                " xmlns:p{level}_{binding:03}=\"urn:litchi:limit:{level}:{binding}\""
            )
            .expect("String cannot fail");
        }
        output.push('>');
    }
    output.push_str("<n:leaf/>");
    for level in 0..64 {
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
        body.push_str(&namespace_pressure());
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

fn detach_once(bytes: &[u8], picture: usize) -> Result<Vec<u8>> {
    let workbook = Workbook::from_bytes(bytes.to_vec())?;
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or("Sheet1 is missing")?
        .detach_svg(PictureSelector::new(0, picture))?;
    Ok(edit.commit()?.into_workbook().to_plain_bytes()?)
}

pub fn run(lane: &str, warmup: usize, samples: usize) -> Result<String> {
    let fixture = fixture_for_lane(lane)?;
    for _ in 0..warmup {
        let _ = execute_lane(lane, &fixture);
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
            Ok(execution) => {
                let outcome = Outcome::Success {
                    semantic_ok: execution.semantic_ok,
                    output_exact: execution.output_exact,
                };
                drop(execution.output);
                outcome
            },
            Err(error) => Outcome::Failure(error),
        };
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let allocation = before.delta(after);
        receipts.push(sample_json(
            lane,
            fixture.expected_success,
            fixture.refusal_class,
            elapsed_ns,
            allocation,
            outcome,
        ));
    }
    Ok(receipt_json(lane, &fixture, warmup, &receipts))
}

fn execute_lane(lane: &str, fixture: &Fixture) -> Result<Execution> {
    if lane == "namespace_limit_refusal" {
        let drawing = drawing_part(&fixture.package)?;
        let _ = SourceDrawing::scan(&drawing)?;
        return Err("namespace limit fixture was accepted".into());
    }
    if lane.starts_with("limit_") {
        return execute_attach(fixture).map(|_| Execution {
            output: None,
            semantic_ok: false,
            output_exact: false,
        });
    }
    if lane.starts_with("capture_") {
        return execute_capture(lane, fixture);
    }
    if lane == "namespace_heavy" {
        return execute_capture(lane, fixture);
    }
    if lane.starts_with("clone_") {
        return execute_clone(lane, fixture);
    }
    if lane.starts_with("inventory_") {
        if lane == "inventory_shared_root_namespace_32" {
            return execute_inventory_root_namespace(fixture);
        }
        return execute_inventory(lane, fixture);
    }
    if lane.starts_with("attach_end_to_end") {
        return execute_attach(fixture);
    }
    if lane.starts_with("detach_end_to_end") {
        return execute_detach(lane, fixture);
    }
    if lane.starts_with("noop_detach") {
        return execute_detach(lane, fixture);
    }
    if lane == "malformed_duplicate_owner"
        || lane == "malformed_mce_owner"
        || lane == "malformed_linked_owner"
    {
        return execute_expected_refusal(fixture);
    }
    if lane == "malformed_unknown_uri" {
        return execute_detach(lane, fixture);
    }
    if lane.starts_with("multi_picture_same_drawing_") {
        if lane.starts_with("multi_picture_same_drawing_detach_") {
            return execute_multi_picture_detach(fixture);
        }
        return execute_multi_picture_attach(fixture);
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
    Ok(Execution {
        output: None,
        semantic_ok: actual_attached == expected_attached && anchor_ok && native_ok,
        output_exact: true,
    })
}

fn execute_clone(_lane: &str, fixture: &Fixture) -> Result<Execution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let drawing = drawing_part(&fixture.package)?;
    let source = SourceDrawing::scan(&drawing)?;
    let _owner = source.picture(fixture.picture)?;
    let first = workbook.clone();
    let second = first.clone();
    black_box(second);
    Ok(Execution {
        output: None,
        semantic_ok: true,
        output_exact: true,
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
    Ok(Execution {
        output: None,
        semantic_ok: all_present && topology_ok,
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
            SvgOwnerState::Embedded(owner) => owner.value().source(),
            _ => None,
        })
        .map(|source| source.len())
        .sum::<usize>();
    let semantic_ok = source.pictures().len() == fixture.picture_count
        && source.pictures().iter().all(|picture| {
            picture.is_direct_embedded_svg() && picture.raster_relationship_id() == "rIdRaster"
        })
        && relationship_ids.len() == fixture.picture_count
        && relationship_ids.iter().enumerate().all(|(index, id)| {
            relationship_ids[..index]
                .iter()
                .all(|previous| previous != id)
        })
        && retained_svg_source_bytes > fixture.picture_count;
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
    let semantic_ok = output != fixture.package.as_ref()
        && reopened_bytes == output
        && after_picture.anchor() == &before_anchor
        && after_picture.is_direct_embedded_svg()
        && svg_payload_ok;
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        // The candidate is checked against the complete expected closure and
        // reopened bytes. "Exact" here means the adapter's deterministic
        // output assertions passed; attach necessarily changes source bytes.
        output_exact: semantic_ok,
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
    let expected_svg = if lane.starts_with("detach_end_to_end_shared_first") {
        true
    } else if lane.starts_with("detach_end_to_end_shared_final")
        || lane.starts_with("detach_end_to_end_distinct")
    {
        !lane.starts_with("detach_end_to_end_shared_final")
    } else {
        false
    };
    let svg_present = has_svg_part(&output);
    let semantic_ok = reopened_bytes == output
        && selected_absent
        && svg_present == expected_svg
        && (lane
            .starts_with("noop_detach")
            .then_some(output.as_slice() == fixture.package.as_ref())
            .unwrap_or(true));
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        // Changed detach outputs are validated by the selected-owner,
        // reachability, and reopen checks above; source-byte equality is an
        // additional no-op assertion when this lane has no owner.
        output_exact: semantic_ok,
    })
}

fn execute_expected_refusal(fixture: &Fixture) -> Result<Execution> {
    let workbook = Workbook::from_bytes(fixture.package.to_vec())?;
    let mut edit = workbook.edit()?;
    let mut sheet = edit.sheet("Sheet1")?.ok_or("Sheet1 is missing")?;
    let result = sheet.detach_svg(PictureSelector::new(0, fixture.picture));
    match result {
        Ok(_) => Ok(Execution {
            output: None,
            semantic_ok: false,
            output_exact: false,
        }),
        Err(error) => Err(error.into()),
    }
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
        && svg_parts_ok;
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        output_exact: semantic_ok,
    })
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
        && !has_svg_part(&output);
    Ok(Execution {
        output: Some(output),
        semantic_ok,
        output_exact: semantic_ok,
    })
}

fn drawing_part(bytes: &[u8]) -> Result<Vec<u8>> {
    Ok(OpcPackage::from_bytes(bytes)?
        .get_part(&PackURI::new(DRAWING)?)?
        .blob()
        .to_vec())
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
    refusal_class: Option<&'static str>,
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
        Outcome::Failure(error) => {
            let message = error.to_string();
            let class = if !expected_success {
                if refusal_class == Some("namespace_limit") && !message.contains("active namespace")
                {
                    String::from("unexpected-refusal")
                } else {
                    refusal_class.unwrap_or("expected-refusal").to_owned()
                }
            } else {
                message.clone()
            };
            (
                false,
                !expected_success,
                !expected_success,
                Some(ErrorInfo { class, message }),
            )
        },
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
    write!(
        output,
        "{{\n  \"schema\":\"xlsx-svg-lifecycle-profile-v1\",\n  \"lane\":\"{}\",\n  \"input_bytes\":{},\n  \"input_hash_fnv1a64\":{},\n  \"warmup\":{},\n  \"sample_count\":{},\n  \"expected_success\":{},\n  \"samples\":[\n",
        support::json_escape(lane),
        fixture.input_bytes,
        fixture.input_hash,
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
