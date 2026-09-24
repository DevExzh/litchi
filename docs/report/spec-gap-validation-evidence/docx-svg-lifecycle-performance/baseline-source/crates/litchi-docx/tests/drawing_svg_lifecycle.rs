use std::io::{self, Cursor};
use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::path::Path;
use std::sync::Arc;
use std::sync::RwLock;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
    SourceVersion,
};
use litchi_docx::drawing::DrawingPlacement;
use litchi_docx::source_backed::{
    self, PictureSelector, StorySelector, SvgInput, SvgPictureOwnerState,
};
use litchi_docx::{Package, ReadLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_WP: &str = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_PIC: &str = "http://purl.oclc.org/ooxml/drawingml/picture";
const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const ASVG: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const PICTURE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const INLINE_FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/svg.docx"
);
const FLOATING_FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf164835_nonDummyLineHeight.docx"
);

const PNG: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 16, 80, 48, 248, 15, 0, 2,
    4, 1, 96, 141, 188, 187, 113, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

#[derive(Clone)]
struct MutableArchiveSource {
    bytes: Arc<RwLock<Vec<u8>>>,
    revision: Arc<AtomicU64>,
}

impl MutableArchiveSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(RwLock::new(bytes)),
            revision: Arc::new(AtomicU64::new(0)),
        }
    }

    fn replace_and_bump(&self, bytes: Vec<u8>) {
        *self.bytes.write().unwrap() = bytes;
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for MutableArchiveSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.read().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        let bytes = self.bytes.read().unwrap();
        if offset >= bytes.len() {
            return Ok(0);
        }
        let end = offset
            .checked_add(output.len())
            .unwrap_or(bytes.len())
            .min(bytes.len());
        output[..end - offset].copy_from_slice(&bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5_5644,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

/// A positional source whose revision can be advanced by the test and whose
/// reads can be attributed to selected ZIP payload spans.  The lifecycle
/// contract observes only the public `ReadAt` version and read operations;
/// the fixture does not depend on package internals or cache state.
#[derive(Clone)]
struct ObservedSource {
    bytes: Arc<Vec<u8>>,
    revision: Arc<AtomicU64>,
    selected_payload: Option<Range<usize>>,
    unrelated_payload: Option<Range<usize>>,
    selected_reads: Arc<AtomicUsize>,
    unrelated_reads: Arc<AtomicUsize>,
}

impl ObservedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            revision: Arc::new(AtomicU64::new(0)),
            selected_payload: None,
            unrelated_payload: None,
            selected_reads: Arc::new(AtomicUsize::new(0)),
            unrelated_reads: Arc::new(AtomicUsize::new(0)),
        }
    }

    fn with_media_ranges(mut self, selected: Range<usize>, unrelated: Range<usize>) -> Self {
        self.selected_payload = Some(selected);
        self.unrelated_payload = Some(unrelated);
        self
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn selected_reads(&self) -> usize {
        self.selected_reads.load(Ordering::SeqCst)
    }

    fn unrelated_reads(&self) -> usize {
        self.unrelated_reads.load(Ordering::SeqCst)
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
        if let Some(range) = &self.selected_payload {
            if offset < range.end && range.start < end {
                self.selected_reads.fetch_add(1, Ordering::SeqCst);
            }
        }
        if let Some(range) = &self.unrelated_payload {
            if offset < range.end && range.start < end {
                self.unrelated_reads.fetch_add(1, Ordering::SeqCst);
            }
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5_5643,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn payload_range(zip: &[u8], name: &str) -> Range<usize> {
    let name = name.as_bytes();
    for (offset, _) in zip
        .windows(4)
        .enumerate()
        .filter(|(_, signature)| *signature == b"PK\x01\x02")
    {
        if offset + 46 > zip.len() {
            continue;
        }
        let compressed =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_len =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        if offset + 46 + name_len > zip.len() || &zip[offset + 46..offset + 46 + name_len] != name {
            continue;
        }
        let local = u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
        let local_name =
            u16::from_le_bytes(zip[local + 26..local + 28].try_into().unwrap()) as usize;
        let local_extra =
            u16::from_le_bytes(zip[local + 28..local + 30].try_into().unwrap()) as usize;
        let start = local + 30 + local_name + local_extra;
        return start..start + compressed;
    }
    panic!("ZIP member was not found: {name:?}");
}

fn read_fixture(path: &str) -> Vec<u8> {
    assert!(
        Path::new(path).is_file(),
        "native fixture is missing: {path}"
    );
    std::fs::read(path).expect("read native DOCX fixture")
}

fn open(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_reader(Cursor::new(bytes)).expect("open source-backed DOCX")
}

fn part_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    let package = Package::from_reader(Cursor::new(bytes)).expect("open published DOCX");
    package
        .opc_package()
        .get_part(&PackURI::new(name).expect("part URI"))
        .expect("part exists")
        .blob()
        .to_vec()
}

fn relationship_bytes(bytes: &[u8], owner: &str) -> Vec<u8> {
    let package = Package::from_reader(Cursor::new(bytes)).expect("open DOCX relationships");
    package
        .opc_package()
        .source_relationships(&PackURI::new(owner).expect("relationship owner URI"))
        .expect("relationship source bytes")
        .bytes()
        .to_vec()
}

fn content_types_bytes(bytes: &[u8]) -> Vec<u8> {
    let package = Package::from_reader(Cursor::new(bytes)).expect("open DOCX content types");
    package
        .opc_package()
        .source_content_types()
        .expect("content-types source bytes")
        .bytes()
        .to_vec()
}

fn has_part(bytes: &[u8], name: &str) -> bool {
    Package::from_reader(Cursor::new(bytes))
        .expect("open published DOCX")
        .opc_package()
        .get_part(&PackURI::new(name).expect("part URI"))
        .is_ok()
}

fn main_part_relationship(bytes: &[u8], id: &str) -> Option<(String, String)> {
    let package = Package::from_reader(Cursor::new(bytes)).expect("open published DOCX");
    let part = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    part.rels().get(id).map(|relationship| {
        (
            relationship.reltype().to_owned(),
            relationship
                .target_partname()
                .expect("internal relationship")
                .as_str()
                .to_owned(),
        )
    })
}

fn root_has_target(bytes: &[u8], target: &str) -> bool {
    let package = Package::from_reader(Cursor::new(bytes)).expect("open published DOCX");
    package
        .opc_package()
        .rels()
        .iter()
        .filter_map(|relationship| relationship.target_partname().ok())
        .any(|part| part.as_str() == target)
}

fn root_relationship_targets(bytes: &[u8]) -> Vec<String> {
    let package = Package::from_reader(Cursor::new(bytes)).expect("open published DOCX");
    package
        .opc_package()
        .rels()
        .iter()
        .filter_map(|relationship| relationship.target_partname().ok())
        .map(|part| part.as_str().to_owned())
        .collect()
}

fn picture(placement: &str, id: &str, raster_id: &str, extension: &str) -> String {
    let geometry = if placement == "anchor" {
        r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2524125" cy="2524125"/><wp:wrapSquare wrapText="bothSides"/>"#
    } else {
        r#"<wp:extent cx="2895600" cy="2762250"/>"#
    };
    let ext_list = if extension.is_empty() {
        String::new()
    } else {
        format!(r#"<a:extLst>{extension}</a:extLst>"#)
    };
    format!(
        r#"<w:drawing><wp:{placement}>{geometry}<wp:docPr id="{id}" name="picture {id}"/><a:graphic><a:graphicData uri="{PICTURE_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="picture {id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{raster_id}">{ext_list}</a:blip></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:{placement}></w:drawing>"#
    )
}

fn linked_raster_picture(placement: &str, id: &str, raster_id: &str) -> String {
    picture(placement, id, raster_id, "").replace(
        &format!("r:embed=\"{raster_id}\""),
        &format!("r:link=\"{raster_id}\""),
    )
}

fn strict_picture(placement: &str, id: &str, raster_id: &str) -> String {
    let geometry = if placement == "anchor" {
        r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2524125" cy="2524125"/><wp:wrapSquare wrapText="bothSides"/>"#
    } else {
        r#"<wp:extent cx="2895600" cy="2762250"/>"#
    };
    format!(
        r#"<w:drawing><wp:{placement}>{geometry}<wp:docPr id="{id}" name="strict picture {id}"/><a:graphic><a:graphicData uri="{STRICT_PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="strict picture {id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{raster_id}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:{placement}></w:drawing>"#
    )
}

fn document(pictures: &str, marker: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}" xmlns:asvg="{ASVG}"><w:body><!--{marker}--><w:p><w:r>{pictures}</w:r></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn strict_document(pictures: &str, marker: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{STRICT_W}" xmlns:wp="{STRICT_WP}" xmlns:a="{STRICT_A}" xmlns:pic="{STRICT_PIC}" xmlns:r="{STRICT_R}" xmlns:asvg="{ASVG}"><w:body><!--{marker}--><w:p><w:r>{pictures}</w:r></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn package_with_document(
    document_xml: Vec<u8>,
    svg: Option<&[u8]>,
    root_svg_edge: bool,
) -> Vec<u8> {
    package_with_document_root_target(
        document_xml,
        svg,
        root_svg_edge.then_some("word/media/image2.svg"),
    )
}

fn package_with_document_root_target(
    document_xml: Vec<u8>,
    svg: Option<&[u8]>,
    root_svg_target: Option<&str>,
) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document_xml,
    );
    main.rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "media/image1.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    if svg.is_some() {
        main.rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "media/image2.svg".to_owned(),
                "rIdSvg".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/media/image1.png").unwrap(),
            ct::PNG.to_owned(),
            PNG.to_vec(),
        )))
        .unwrap();
    if let Some(svg) = svg {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/media/image2.svg").unwrap(),
                "image/svg+xml".to_owned(),
                svg.to_vec(),
            )))
            .unwrap();
    }
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    if let Some(root_svg_target) = root_svg_target {
        package.relate_to(root_svg_target, rt::IMAGE);
    }
    PackageWriter::to_bytes(&package).expect("write synthetic DOCX")
}

fn package_with_two_rasters(document_xml: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document_xml,
    );
    for (id, target) in [
        ("rIdRaster0", "media/image1.png"),
        ("rIdRaster1", "media/image2.png"),
    ] {
        main.rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                target.to_owned(),
                id.to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/media/image1.png").unwrap(),
            ct::PNG.to_owned(),
            PNG.to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/media/image2.png").unwrap(),
            ct::PNG.to_owned(),
            b"unrelated-raster-payload".to_vec(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).expect("write two-raster synthetic DOCX")
}

fn two_raster_fixture(marker: &str) -> Vec<u8> {
    let pictures = format!(
        "{}{}",
        picture("inline", "1", "rIdRaster0", ""),
        picture("anchor", "2", "rIdRaster1", ""),
    );
    package_with_two_rasters(document(&pictures, marker))
}

fn quote_variant(bytes: &[u8]) -> Vec<u8> {
    bytes
        .iter()
        .map(|byte| if *byte == b'"' { b'\'' } else { *byte })
        .collect()
}

fn lexical_stale_fixture() -> (Vec<u8>, Vec<u8>) {
    let document_xml = document(&picture("inline", "7", "rIdRaster", ""), "lexical-stale");
    let content_types = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="{wml}"/></Types>"#,
        wml = ct::WML_DOCUMENT_MAIN,
    )
    .into_bytes();
    let root_relationships = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDocument" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    )
    .into_bytes();
    let document_relationships = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRaster" Type="{IMAGE_REL}" Target="media/image1.png"/></Relationships>"#
    )
    .into_bytes();

    let build = |content_types: &[u8], document_relationships: &[u8]| {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("[Content_Types].xml", content_types)
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships.as_slice())
            .unwrap();
        writer
            .write_stored("word/_rels/document.xml.rels", document_relationships)
            .unwrap();
        writer
            .write_stored("word/document.xml", &document_xml)
            .unwrap();
        writer.write_stored("word/media/image1.png", PNG).unwrap();
        writer.finish_to_bytes().unwrap()
    };

    let initial = build(&content_types, &document_relationships);
    let altered = build(
        &quote_variant(&content_types),
        &quote_variant(&document_relationships),
    );
    assert_eq!(initial.len(), altered.len());
    (initial, altered)
}

fn total_part_bytes(bytes: &[u8]) -> u64 {
    Package::from_reader(Cursor::new(bytes))
        .expect("open package for declared part total")
        .opc_package()
        .iter_parts()
        .map(|part| part.blob().len() as u64)
        .sum()
}

fn managed_package(
    source: ObservedSource,
    memory_limit: u64,
    work_limit: u64,
) -> (Budget, source_backed::Package) {
    let budget = Budget::root(
        "docx-svg-lifecycle-test",
        Limits::new(
            memory_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            work_limit,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory_limit.max(1)).unwrap(),
        0,
    )
    .expect("valid managed execution limits");
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(source),
        ReadLimits::default(),
        ExecutionContext::new(budget.clone(), cancellation, execution_limits),
    )
    .expect("open managed source-backed package");
    (budget, package)
}

fn strict_package_with_document(document_xml: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document_xml,
    );
    main.rels_mut()
        .try_add_relationship(
            rt::STRICT_IMAGE.to_owned(),
            "media/image1.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/media/image1.png").unwrap(),
            ct::PNG.to_owned(),
            PNG.to_vec(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::STRICT_OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).expect("write synthetic Strict DOCX")
}

fn raster_only_fixture(placement: &str, marker: &str) -> Vec<u8> {
    package_with_document(
        document(&picture(placement, "7", "rIdRaster", ""), marker),
        None,
        false,
    )
}

fn svg_owner(id: &str) -> String {
    format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="{id}"/></a:ext>"#)
}

#[test]
fn native_inline_and_floating_views_expose_graph_and_placement() {
    for (path, placement, raster_extent, raster_part) in [
        (
            INLINE_FIXTURE,
            DrawingPlacement::Inline,
            "2895600",
            "/word/media/image1.png",
        ),
        (
            FLOATING_FIXTURE,
            DrawingPlacement::Floating,
            "2524125",
            "/word/media/image1.png",
        ),
    ] {
        let source = read_fixture(path);
        let package = open(&source);
        let views = package
            .svg_pictures(StorySelector::Main)
            .expect("native SVG picture inventory");
        assert_eq!(views.len(), 1);
        let view = &views[0];
        assert_eq!(view.selector(), PictureSelector::new(0, 0));
        assert_eq!(view.placement(), placement);
        assert_eq!(view.raster_relationship_id(), "rId4");
        assert_eq!(view.raster_part_uri().as_str(), raster_part);
        assert_eq!(view.raster_content_type(), ct::PNG);
        assert_eq!(view.owner_state(), SvgPictureOwnerState::Embedded);
        let svg = view.svg().expect("native SVG view");
        assert_eq!(svg.relationship_id(), "rId5");
        assert_eq!(svg.part_uri().as_str(), "/word/media/image2.svg");
        assert_eq!(svg.relationship_type(), rt::IMAGE);
        assert_eq!(svg.bytes(), part_bytes(&source, "/word/media/image2.svg"));
        let document = part_bytes(&source, "/word/document.xml");
        let extent = format!("cx=\"{raster_extent}\"");
        assert!(
            document
                .windows(extent.len())
                .any(|window| window == extent.as_bytes())
        );
    }
}

#[test]
fn detach_native_svg_publishes_and_reopens_without_touching_raster_or_anchor() {
    for (path, anchor_marker) in [
        (
            INLINE_FIXTURE,
            b"<wp:extent cx=\"2895600\" cy=\"2762250\"/".as_slice(),
        ),
        (
            FLOATING_FIXTURE,
            b"<wp:wrapSquare wrapText=\"bothSides\"/>".as_slice(),
        ),
    ] {
        let source = read_fixture(path);
        let raster_before = part_bytes(&source, "/word/media/image1.png");
        let package = open(&source);
        let mut edit = package
            .edit_svg_attachment(PictureSelector::new(0, 0))
            .expect("select native SVG picture");
        let before = edit.source();
        assert_eq!(before.svg().unwrap().relationship_id(), "rId5");
        assert!(edit.detach_svg().expect("detach native SVG"));
        let commit = edit.commit().expect("commit native detach");
        assert!(commit.is_changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.diagnostics().selected_pictures(), 1);
        assert!(commit.diagnostics().changed());
        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("exact inverse snapshot");
        assert_eq!(inverse.story_xml(), before.story_xml());
        assert_eq!(inverse.svg().unwrap().relationship_id(), "rId5");

        let mut published = Vec::new();
        let detached = package
            .publish_svg_attachment_commit_to_stream(&mut published, &commit)
            .expect("publish native detach");
        assert!(detached.svg().is_none());
        let reopened = open(&published);
        let view = &reopened.svg_pictures(StorySelector::Main).unwrap()[0];
        assert_eq!(view.owner_state(), SvgPictureOwnerState::None);
        assert!(view.svg().is_none());
        assert_eq!(view.raster_relationship_id(), "rId4");
        assert_eq!(
            part_bytes(&published, "/word/media/image1.png"),
            raster_before
        );
        assert!(!has_part(&published, "/word/media/image2.svg"));
        assert!(main_part_relationship(&published, "rId5").is_none());
        let document = part_bytes(&published, "/word/document.xml");
        assert!(
            document
                .windows(anchor_marker.len())
                .any(|window| window == anchor_marker)
        );
        assert!(
            !document
                .windows(b"rId5".len())
                .any(|window| window == b"rId5")
        );
        assert!(
            !document
                .windows(b"litchi-svg-detached".len())
                .any(|window| window == b"litchi-svg-detached")
        );
    }
}

#[test]
fn borrowed_attach_adds_graph_closure_and_reopens_with_raster_preserved() {
    let payload = br#"<svg xmlns="http://www.w3.org/2000/svg"><path/></svg>"#;
    for placement in ["inline", "anchor"] {
        let source = raster_only_fixture(placement, "attach");
        let package = open(&source);
        let view = &package.svg_pictures(StorySelector::Main).unwrap()[0];
        assert_eq!(view.owner_state(), SvgPictureOwnerState::None);
        assert!(view.svg().is_none());

        let mut edit = package
            .edit_svg_attachment(PictureSelector::new(0, 0))
            .unwrap();
        let before = edit.source();
        assert!(edit.attach_svg(SvgInput::borrowed(payload)).unwrap());
        let commit = edit.commit().expect("commit borrowed attach");
        assert!(commit.is_changed());
        assert_eq!(commit.diagnostics().operation_count(), 1);
        assert_eq!(commit.diagnostics().selected_picture_count(), 1);
        assert!(commit.diagnostics().changed());
        assert!(
            commit
                .snapshot()
                .story_xml()
                .windows(b"rIdSvg".len())
                .any(|window| window == b"rIdSvg")
        );
        assert_eq!(commit.snapshot().svg().unwrap().bytes(), payload);
        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse attach snapshot");
        assert_eq!(inverse.story_xml(), before.story_xml());
        assert!(inverse.svg().is_none());

        let mut published = Vec::new();
        let target = package
            .publish_svg_attachment_commit_to_stream(&mut published, &commit)
            .expect("publish borrowed attach");
        assert_eq!(target.svg().unwrap().bytes(), payload);
        let reopened = open(&published);
        let view = &reopened.svg_pictures(StorySelector::Main).unwrap()[0];
        assert_eq!(view.owner_state(), SvgPictureOwnerState::Embedded);
        let svg = view.svg().unwrap();
        assert_eq!(svg.relationship_id(), "rIdSvg");
        assert_eq!(svg.part_uri().as_str(), "/word/media/image.svg");
        assert_eq!(svg.bytes(), payload);
        assert_eq!(view.raster_relationship_id(), "rIdRaster");
        assert_eq!(part_bytes(&published, "/word/media/image1.png"), PNG);
        assert_eq!(part_bytes(&published, "/word/media/image.svg"), payload);
        assert_eq!(
            main_part_relationship(&published, "rIdSvg").unwrap(),
            (rt::IMAGE.to_owned(), "/word/media/image.svg".to_owned())
        );
        let document = part_bytes(&published, "/word/document.xml");
        let preserved = if placement == "anchor" {
            b"<wp:wrapSquare wrapText=\"bothSides\"/>".as_slice()
        } else {
            b"<wp:extent cx=\"2895600\" cy=\"2762250\"/>".as_slice()
        };
        assert!(
            document
                .windows(preserved.len())
                .any(|window| window == preserved)
        );
    }
}

#[test]
fn lifecycle_commit_diagnostics_retain_noop_scope() {
    let single_package = open(&raster_only_fixture("inline", "diagnostics-single-noop"));
    let single_edit = single_package.edit_svg_attachment(0).unwrap();
    let single_commit = single_edit.commit().expect("commit single no-op");
    assert!(!single_commit.is_changed());
    assert_eq!(single_commit.diagnostics().operations(), 0);
    assert_eq!(single_commit.diagnostics().selected_pictures(), 1);
    assert!(!single_commit.diagnostics().changed());

    let pictures = format!(
        "{}{}",
        picture("inline", "1", "rIdRaster", ""),
        picture("anchor", "2", "rIdRaster", ""),
    );
    let batch_source =
        package_with_document(document(&pictures, "diagnostics-batch-noop"), None, false);
    let batch_package = open(&batch_source);
    let batch_edit = batch_package
        .edit_svg_attachments([0usize, 1usize])
        .unwrap();
    let batch_commit = batch_edit.commit().expect("commit batch no-op");
    assert!(!batch_commit.is_changed());
    assert_eq!(batch_commit.diagnostics().operation_count(), 0);
    assert_eq!(batch_commit.diagnostics().selected_picture_count(), 2);
    assert!(!batch_commit.diagnostics().changed());
}

#[test]
fn strict_core_keeps_strict_raster_edge_and_transitional_svg_child() {
    let source = strict_package_with_document(strict_document(
        &strict_picture("inline", "7", "rIdRaster"),
        "strict",
    ));
    let package = open(&source);
    let view = &package.svg_pictures(StorySelector::Main).unwrap()[0];
    assert_eq!(view.owner_state(), SvgPictureOwnerState::None);
    assert_eq!(view.raster_relationship_id(), "rIdRaster");

    let payload = br#"<svg xmlns="http://www.w3.org/2000/svg"><path/></svg>"#;
    let mut edit = package.edit_svg_attachment(0).unwrap();
    edit.attach_svg(SvgInput::borrowed(payload)).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();

    let document = part_bytes(&output, "/word/document.xml");
    let transitional_child_namespace = format!(r#"xmlns:r="{R}""#);
    assert!(
        document
            .windows(transitional_child_namespace.len())
            .any(|window| window == transitional_child_namespace.as_bytes())
    );
    let strict_core_namespace = format!(r#"xmlns:r="{STRICT_R}""#);
    assert!(
        document
            .windows(strict_core_namespace.len())
            .any(|window| window == strict_core_namespace.as_bytes())
    );
    assert!(
        document
            .windows(b"r:embed=\"rIdSvg\"".len())
            .any(|window| window == b"r:embed=\"rIdSvg\"")
    );
    assert_eq!(
        main_part_relationship(&output, "rIdRaster").unwrap().0,
        rt::STRICT_IMAGE
    );
    assert_eq!(
        main_part_relationship(&output, "rIdSvg").unwrap().0,
        rt::STRICT_IMAGE
    );
    let view = &open(&output).svg_pictures(StorySelector::Main).unwrap()[0];
    assert_eq!(view.owner_state(), SvgPictureOwnerState::Embedded);
    assert_eq!(view.svg().unwrap().bytes(), payload);
}

#[test]
fn batch_attach_and_shared_detach_are_one_atomic_edit() {
    let pictures = format!(
        "{}{}",
        picture("inline", "1", "rIdRaster", ""),
        picture("anchor", "2", "rIdRaster", ""),
    );
    let source = package_with_document(document(&pictures, "batch-attach"), None, false);
    let source_raster = part_bytes(&source, "/word/media/image1.png");
    let source_relationships = relationship_bytes(&source, "/word/document.xml");
    let source_root_relationships = relationship_bytes(&source, "/");
    let source_content_types = content_types_bytes(&source);
    let package = open(&source);
    let mut edit = package.edit_svg_attachments([0usize, 1usize]).unwrap();
    let before = edit.source().clone();
    let first_payload = br#"<svg xmlns="http://www.w3.org/2000/svg"><path id="one"/></svg>"#;
    let second_payload = br#"<svg xmlns="http://www.w3.org/2000/svg"><path id="two"/></svg>"#;
    assert!(
        edit.attach_svg(0, SvgInput::borrowed(first_payload))
            .unwrap()
    );
    let projected_after_first = edit.projected();
    assert_eq!(
        projected_after_first
            .picture(0)
            .expect("first selected projection")
            .svg()
            .expect("first projected SVG")
            .bytes(),
        first_payload
    );
    assert!(
        projected_after_first
            .picture(1)
            .expect("second selected projection")
            .svg()
            .is_none(),
        "a first operation must not leak into another selected picture"
    );
    let projected_before_rejected_operation = edit.projected().story_xml().unwrap().to_vec();
    assert!(
        edit.attach_svg(0, SvgInput::borrowed(second_payload))
            .is_err()
    );
    assert_eq!(
        edit.projected().story_xml().unwrap(),
        projected_before_rejected_operation.as_slice()
    );
    assert!(
        edit.attach_svg(1, SvgInput::borrowed(second_payload))
            .unwrap()
    );
    let projected_after_second = edit.projected();
    assert_eq!(
        projected_after_second
            .picture(0)
            .unwrap()
            .svg()
            .unwrap()
            .bytes(),
        first_payload
    );
    assert_eq!(
        projected_after_second
            .picture(1)
            .unwrap()
            .svg()
            .unwrap()
            .bytes(),
        second_payload
    );
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    assert_eq!(commit.diagnostics().operations(), 2);
    assert_eq!(commit.diagnostics().selected_pictures(), 2);
    assert!(commit.diagnostics().changed());
    let inverse = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("exact inverse batch snapshot");
    assert_eq!(inverse.story_xml(), before.story_xml());
    assert!(inverse.picture(0).unwrap().svg().is_none());
    assert!(inverse.picture(1).unwrap().svg().is_none());

    let foreign_pictures = format!(
        "{}{}",
        picture("inline", "1", "rIdRaster", ""),
        picture("anchor", "2", "rIdRaster", ""),
    );
    let foreign_source =
        package_with_document(document(&foreign_pictures, "batch-foreign"), None, false);
    let foreign = open(&foreign_source)
        .edit_svg_attachments([0usize, 1usize])
        .unwrap()
        .source()
        .clone();
    assert!(commit.patch().apply(&foreign).is_err());

    let mut output = Vec::new();
    package
        .publish_svg_attachment_batch_commit_to_stream(&mut output, &commit)
        .unwrap();
    let views = open(&output).svg_pictures(StorySelector::Main).unwrap();
    assert_eq!(views.len(), 2);
    assert_eq!(views[0].svg().unwrap().bytes(), first_payload);
    assert_eq!(views[1].svg().unwrap().bytes(), second_payload);
    assert_eq!(
        main_part_relationship(&output, "rIdSvg").unwrap().1,
        "/word/media/image.svg"
    );
    assert_eq!(
        main_part_relationship(&output, "rIdSvg1").unwrap().1,
        "/word/media/image1.svg"
    );

    let reopened_attach = open(&output);
    let mut inverse_edit = reopened_attach
        .edit_svg_attachments([0usize, 1usize])
        .unwrap();
    assert!(inverse_edit.detach_svg(0).unwrap());
    assert!(inverse_edit.detach_svg(1).unwrap());
    let inverse_commit = inverse_edit.commit().unwrap();
    let mut inverse_output = Vec::new();
    reopened_attach
        .publish_svg_attachment_batch_commit_to_stream(&mut inverse_output, &inverse_commit)
        .unwrap();
    assert_eq!(
        relationship_bytes(&inverse_output, "/word/document.xml"),
        source_relationships
    );
    assert_eq!(
        relationship_bytes(&inverse_output, "/"),
        source_root_relationships
    );
    assert_eq!(content_types_bytes(&inverse_output), source_content_types);
    // The in-memory inverse above proves exact story bytes. After reopening,
    // a fresh detach is a new edit and the design permits lexical story
    // normalization; the physical relationship/content-type/raster closure
    // remains exact here.
    assert_eq!(
        part_bytes(&inverse_output, "/word/media/image1.png"),
        source_raster
    );
    assert!(
        open(&inverse_output)
            .svg_pictures(StorySelector::Main)
            .unwrap()
            .iter()
            .all(|view| view.svg().is_none())
    );

    let shared_pictures = format!(
        "{}{}",
        picture("inline", "1", "rIdRaster", &svg_owner("rIdSvg")),
        picture("anchor", "2", "rIdRaster", &svg_owner("rIdSvg")),
    );
    let shared_source = package_with_document(
        document(&shared_pictures, "batch-shared"),
        Some(b"<svg/>"),
        true,
    );
    let shared_package = open(&shared_source);
    let mut shared_edit = shared_package
        .edit_svg_attachments([0usize, 1usize])
        .unwrap();
    let shared_before = shared_edit.source().clone();
    assert!(shared_edit.detach_svg(0).unwrap());
    assert!(shared_edit.detach_svg(1).unwrap());
    let shared_commit = shared_edit.commit().unwrap();
    let shared_inverse = shared_commit
        .patch()
        .inverse()
        .apply(shared_commit.snapshot())
        .expect("exact inverse shared batch snapshot");
    assert_eq!(shared_inverse.story_xml(), shared_before.story_xml());
    let mut shared_output = Vec::new();
    shared_package
        .publish_svg_attachment_batch_commit_to_stream(&mut shared_output, &shared_commit)
        .unwrap();
    let shared_views = open(&shared_output)
        .svg_pictures(StorySelector::Main)
        .unwrap();
    assert!(shared_views.iter().all(|view| view.svg().is_none()));
    assert!(main_part_relationship(&shared_output, "rIdSvg").is_none());
    assert!(has_part(&shared_output, "/word/media/image2.svg"));
    assert!(root_has_target(&shared_output, "/word/media/image2.svg"));
    let shared_document = part_bytes(&shared_output, "/word/document.xml");
    assert!(
        !shared_document
            .windows(b"litchi-svg-detached".len())
            .any(|window| window == b"litchi-svg-detached")
    );
}

#[test]
fn attach_preserves_unknown_extension_and_rejects_stale_snapshot() {
    let unknown = r#"<a:ext uri="urn:future" future:flag="keep" xmlns:future="urn:future"><!--opaque sibling--><future:payload keep="yes"/></a:ext>"#;
    let source = package_with_document(
        document(&picture("inline", "7", "rIdRaster", unknown), "opaque"),
        None,
        false,
    );
    let payload = b"<svg>opaque-svg-payload</svg>";
    let package = open(&source);
    assert_eq!(
        package.svg_pictures(StorySelector::Main).unwrap()[0].owner_state(),
        SvgPictureOwnerState::Opaque
    );
    let mut edit = package.edit_svg_attachment(0).unwrap();
    edit.attach_svg(payload.as_slice()).unwrap();
    let commit = edit.commit().unwrap();
    let foreign = open(&raster_only_fixture("inline", "foreign"))
        .edit_svg_attachment(0)
        .unwrap()
        .source();
    assert!(commit.patch().apply(&foreign).is_err());

    let mut published = Vec::new();
    package
        .publish_svg_attachment_commit_to_stream(&mut published, &commit)
        .unwrap();
    let document = part_bytes(&published, "/word/document.xml");
    for expected in [
        b"urn:future".as_slice(),
        b"future:flag=\"keep\"".as_slice(),
        b"<!--opaque sibling-->".as_slice(),
        b"<future:payload keep=\"yes\"/>".as_slice(),
        b"rIdSvg".as_slice(),
    ] {
        assert!(
            document
                .windows(expected.len())
                .any(|window| window == expected)
        );
    }
    let view = &open(&published).svg_pictures(StorySelector::Main).unwrap()[0];
    assert_eq!(view.owner_state(), SvgPictureOwnerState::Embedded);
    assert_eq!(view.svg().unwrap().bytes(), payload);
}

#[test]
fn linked_raster_picture_is_refused_by_svg_lifecycle_edits() {
    let source = package_with_document(
        document(
            &linked_raster_picture("inline", "7", "rIdRaster"),
            "linked-raster",
        ),
        None,
        false,
    );
    let package = open(&source);
    let views = package
        .svg_pictures(StorySelector::Main)
        .expect("linked raster remains inspectable");
    assert_eq!(views.len(), 1);
    assert_eq!(views[0].owner_state(), SvgPictureOwnerState::Refused);
    assert!(package.edit_svg_attachment(0).is_err());
}

#[test]
fn shared_svg_and_package_root_incoming_edge_survive_until_final_detach() {
    let pictures = format!(
        "{}{}",
        picture("inline", "1", "rIdRaster", &svg_owner("rIdSvg")),
        picture("anchor", "2", "rIdRaster", &svg_owner("rIdSvg")),
    );
    let source = package_with_document(document(&pictures, "shared"), Some(b"<svg/>"), true);
    let package = open(&source);
    let views = package.svg_pictures(StorySelector::Main).unwrap();
    assert_eq!(views.len(), 2);
    assert!(
        views
            .iter()
            .all(|view| view.owner_state() == SvgPictureOwnerState::Embedded)
    );

    let mut first_edit = package.edit_svg_attachment(0).unwrap();
    assert!(first_edit.detach_svg().unwrap());
    let first_commit = first_edit.commit().unwrap();
    let mut first_output = Vec::new();
    package
        .publish_svg_attachment_commit_to_stream(&mut first_output, &first_commit)
        .unwrap();
    let first_reopened = open(&first_output);
    let first_views = first_reopened.svg_pictures(StorySelector::Main).unwrap();
    assert_eq!(first_views[0].owner_state(), SvgPictureOwnerState::None);
    assert_eq!(first_views[1].owner_state(), SvgPictureOwnerState::Embedded);
    assert!(has_part(&first_output, "/word/media/image2.svg"));
    assert!(main_part_relationship(&first_output, "rIdSvg").is_some());
    assert!(root_has_target(&first_output, "/word/media/image2.svg"));

    let mut second_edit = first_reopened.edit_svg_attachment(1).unwrap();
    assert!(second_edit.detach_svg().unwrap());
    let second_commit = second_edit.commit().unwrap();
    let mut second_output = Vec::new();
    first_reopened
        .publish_svg_attachment_commit_to_stream(&mut second_output, &second_commit)
        .unwrap();
    let second_views = open(&second_output)
        .svg_pictures(StorySelector::Main)
        .unwrap();
    assert!(second_views.iter().all(|view| view.svg().is_none()));
    assert!(main_part_relationship(&second_output, "rIdSvg").is_none());
    assert!(has_part(&second_output, "/word/media/image2.svg"));
    assert!(root_has_target(&second_output, "/word/media/image2.svg"));
    let document = part_bytes(&second_output, "/word/document.xml");
    assert!(
        !document
            .windows(b"litchi-svg-detached".len())
            .any(|window| window == b"litchi-svg-detached")
    );
}

#[test]
fn case_variant_package_incoming_edge_keeps_detached_svg_leaf() {
    let source = package_with_document_root_target(
        document(
            &picture("inline", "1", "rIdRaster", &svg_owner("rIdSvg")),
            "case-variant",
        ),
        Some(b"<svg/>"),
        Some("word/media/IMAGE2.SVG"),
    );
    let package = open(&source);
    let mut edit = package.edit_svg_attachment(0).unwrap();
    assert!(edit.detach_svg().unwrap());
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();

    assert!(has_part(&output, "/word/media/image2.svg"));
    assert!(main_part_relationship(&output, "rIdSvg").is_none());
    assert!(
        root_relationship_targets(&output)
            .iter()
            .any(|target| target == "/word/media/IMAGE2.SVG")
    );
}

#[test]
fn lifecycle_read_and_attach_honor_input_and_story_output_limits() {
    let native = read_fixture(INLINE_FIXTURE);
    let input_limits = ReadLimits::builder()
        .max_input_bytes((native.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source_backed::Package::from_reader_with_limits(Cursor::new(native), input_limits).is_err()
    );

    let source = raster_only_fixture("inline", "limits");
    let document_len = part_bytes(&source, "/word/document.xml").len();
    let output_limits = ReadLimits::builder()
        .max_part_bytes(document_len as u64)
        .unwrap()
        .build()
        .unwrap();
    let package =
        source_backed::Package::from_reader_with_limits(Cursor::new(source), output_limits)
            .unwrap();
    let mut edit = package.edit_svg_attachment(0).unwrap();
    assert!(edit.attach_svg(b"payload".as_slice()).is_err());
}

#[test]
fn lifecycle_commit_refuses_source_revision_changed_after_staging() {
    let source_bytes = raster_only_fixture("inline", "source-version");
    let source = ObservedSource::new(source_bytes);
    let source_handle = source.clone();
    let package = source_backed::Package::from_read_at(Arc::new(source))
        .expect("open mutable-version source");
    let mut edit = package
        .edit_svg_attachment(0)
        .expect("select source-version picture");
    edit.attach_svg(b"<svg/>".as_slice())
        .expect("stage source-version attachment");

    source_handle.bump_revision();
    assert!(
        edit.commit().is_err(),
        "a source revision change after staging must invalidate the commit"
    );
}

#[test]
fn lifecycle_commit_refuses_lexical_relationship_and_content_type_source_change() {
    let (initial, altered) = lexical_stale_fixture();
    assert_eq!(
        part_bytes(&initial, "/word/document.xml"),
        part_bytes(&altered, "/word/document.xml")
    );
    assert_ne!(
        relationship_bytes(&initial, "/word/document.xml"),
        relationship_bytes(&altered, "/word/document.xml")
    );
    assert_ne!(content_types_bytes(&initial), content_types_bytes(&altered));
    open(&altered)
        .svg_pictures(StorySelector::Main)
        .expect("lexical relationship/content-type variant remains readable");

    let source = MutableArchiveSource::new(initial);
    let source_handle = source.clone();
    let package =
        source_backed::Package::from_read_at(Arc::new(source)).expect("open lexical stale source");
    let mut edit = package.edit_svg_attachment(0).unwrap();
    edit.attach_svg(b"<svg/>".as_slice()).unwrap();

    source_handle.replace_and_bump(altered);
    assert!(
        edit.commit().is_err(),
        "a relationship/content-type lexical source change must invalidate the staged commit"
    );
}

#[test]
fn lifecycle_low_total_part_limit_refuses_before_raster_retention() {
    let source_bytes = raster_only_fixture("inline", "total-part-limit");
    let selected_range = payload_range(&source_bytes, "word/media/image1.png");
    let source = ObservedSource::new(source_bytes.clone()).with_media_ranges(selected_range, 0..0);
    let source_handle = source.clone();
    let total = total_part_bytes(&source_bytes);
    let limits = ReadLimits::builder()
        .max_total_part_bytes(total.saturating_sub(1))
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source_backed::Package::from_read_at_with_limits(Arc::new(source), limits).is_err(),
        "a low total-part profile must refuse before retaining media"
    );
    assert_eq!(
        source_handle.selected_reads(),
        0,
        "the selected raster payload must not be read after aggregate refusal"
    );
}

#[test]
fn lifecycle_low_managed_memory_refuses_before_raster_retention() {
    let source_bytes = raster_only_fixture("inline", "managed-memory-limit");
    let document_len = part_bytes(&source_bytes, "/word/document.xml").len() as u64;
    let selected_range = payload_range(&source_bytes, "word/media/image1.png");
    let source = ObservedSource::new(source_bytes).with_media_ranges(selected_range, 0..0);
    let source_handle = source.clone();
    let (budget, package) = managed_package(source, document_len + PNG.len() as u64 - 1, u64::MAX);

    assert!(
        package.edit_svg_attachment(0).is_err(),
        "managed memory exhaustion must reject before retaining the raster"
    );
    assert_eq!(source_handle.selected_reads(), 0);
    assert!(budget.used(Resource::Memory) <= document_len);
}

#[test]
fn lifecycle_low_managed_work_refuses_before_raster_read() {
    let source_bytes = raster_only_fixture("inline", "managed-work-limit");
    let document_len = part_bytes(&source_bytes, "/word/document.xml").len() as u64;
    let selected_range = payload_range(&source_bytes, "word/media/image1.png");
    let source = ObservedSource::new(source_bytes).with_media_ranges(selected_range, 0..0);
    let source_handle = source.clone();
    let (budget, package) = managed_package(source, u64::MAX, document_len + PNG.len() as u64 - 1);

    assert!(
        package.edit_svg_attachment(0).is_err(),
        "the cumulative managed work limit must reject before raster read"
    );
    assert_eq!(source_handle.selected_reads(), 0);
    assert!(budget.used(Resource::Work) <= document_len);
}

#[test]
fn managed_attach_publish_reopen_physical_inverse_preserves_exact_source() {
    let source_bytes = raster_only_fixture("inline", "managed-physical-inverse");
    let source = ObservedSource::new(source_bytes.clone());
    let (budget, package) = managed_package(source, u64::MAX, u64::MAX);
    let payload = br#"<svg xmlns="http://www.w3.org/2000/svg"><path id="managed-inverse"/></svg>"#;

    let mut edit = package
        .edit_svg_attachment(PictureSelector::new(0, 0))
        .expect("select managed source picture");
    let before_attach = budget.used(Resource::Memory);
    assert!(edit.attach_svg(SvgInput::borrowed(payload)).unwrap());
    let after_attach = budget.used(Resource::Memory);
    assert!(
        after_attach >= before_attach + payload.len() as u64,
        "the staged SVG payload must remain charged to the managed budget"
    );

    let commit = edit.commit().expect("commit managed SVG attach");
    let before_semantic_inverse = budget.used(Resource::Memory);
    let inverse_snapshot = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("semantic inverse remains valid with a retained reservation");
    assert!(inverse_snapshot.svg().is_none());
    drop(inverse_snapshot);
    assert_eq!(
        budget.used(Resource::Memory),
        before_semantic_inverse,
        "semantic inverse inspection must not release the staged reservation"
    );

    let mut published = Vec::new();
    let publication = package
        .publish_svg_attachment_commit_with_publication_to_stream(&mut published, &commit)
        .expect("publish managed SVG attach with physical inverse authorization");
    assert!(
        budget.used(Resource::Memory) >= payload.len() as u64,
        "the publication must retain the staged payload reservation after the source package is consumed"
    );
    drop(commit);

    let (_reopen_budget, reopened) =
        managed_package(ObservedSource::new(published.clone()), u64::MAX, u64::MAX);
    let mut restored = Vec::new();
    let restored_snapshot = reopened
        .publish_svg_attachment_inverse_to_stream(&mut restored, &publication)
        .expect("reopened published package must authorize physical inverse");
    assert!(restored_snapshot.svg().is_none());
    assert_eq!(
        restored, source_bytes,
        "physical inverse must restore exact ZIP bytes"
    );
    drop(restored_snapshot);

    let retained_after_reopen = budget.used(Resource::Memory);
    assert!(
        retained_after_reopen >= payload.len() as u64,
        "publication-owned snapshots must keep the SVG reservation live"
    );
    drop(publication);
    assert!(
        budget.used(Resource::Memory) < retained_after_reopen,
        "dropping publication-owned semantic snapshots must release their reservation"
    );
}

#[test]
fn lifecycle_single_picture_selection_does_not_read_unrelated_media() {
    let source_bytes = two_raster_fixture("selective-media");
    let selected_range = payload_range(&source_bytes, "word/media/image1.png");
    let unrelated_range = payload_range(&source_bytes, "word/media/image2.png");
    let source =
        ObservedSource::new(source_bytes).with_media_ranges(selected_range, unrelated_range);
    let source_handle = source.clone();
    let package =
        source_backed::Package::from_read_at(Arc::new(source)).expect("open two-raster source");

    // The source-backed selector captures one inventory entry lazily; its
    // selected raster is read for authoring while the other media payload
    // remains cold.
    let edit = package
        .edit_svg_attachment(0)
        .expect("select only the first picture");
    assert_eq!(
        edit.source().raster_part_uri().as_str(),
        "/word/media/image1.png"
    );
    assert!(
        source_handle.selected_reads() > 0,
        "selecting a picture must read its own raster fallback"
    );
    assert_eq!(
        source_handle.unrelated_reads(),
        0,
        "selecting one picture must not read the other picture's media payload"
    );
}

#[test]
fn source_picture_inventory_keeps_unrequested_rasters_cold() {
    let bytes = two_raster_fixture("lazy-resource-inventory");
    let first = payload_range(&bytes, "word/media/image1.png");
    let second = payload_range(&bytes, "word/media/image2.png");
    let source = ObservedSource::new(bytes).with_media_ranges(first, second);
    let observed = source.clone();
    let package = source_backed::Package::from_read_at(Arc::new(source)).unwrap();
    let pictures = package.svg_picture_sources(StorySelector::Main).unwrap();
    assert_eq!(pictures.len(), 2);
    assert_eq!(observed.selected_reads(), 0);
    assert_eq!(observed.unrelated_reads(), 0);
    assert_eq!(pictures[0].raster().data().unwrap().as_bytes(), PNG);
    assert!(observed.selected_reads() > 0);
    assert_eq!(observed.unrelated_reads(), 0);
    observed.bump_revision();
    assert!(pictures[1].raster().data().is_err());
    assert_eq!(observed.unrelated_reads(), 0);
    assert!(package.svg_picture_sources(StorySelector::Main).is_err());
}

#[test]
fn source_picture_inventory_reads_svg_only_on_explicit_data_request() {
    let svg = br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0L1 1"/></svg>"#;
    let bytes = package_with_document(
        document(
            &picture("inline", "1", "rIdRaster", &svg_owner("rIdSvg")),
            "lazy-svg",
        ),
        Some(svg),
        false,
    );
    let raster_range = payload_range(&bytes, "word/media/image1.png");
    let svg_range = payload_range(&bytes, "word/media/image2.svg");
    let source = ObservedSource::new(bytes).with_media_ranges(raster_range, svg_range);
    let observed = source.clone();
    let package = source_backed::Package::from_read_at(Arc::new(source)).unwrap();
    let pictures = package.svg_picture_sources(StorySelector::Main).unwrap();
    assert_eq!(pictures.len(), 1);
    assert!(matches!(
        pictures[0].owner_state(),
        SvgPictureOwnerState::Embedded
    ));
    assert_eq!(observed.selected_reads(), 0);
    assert_eq!(observed.unrelated_reads(), 0);
    assert_eq!(pictures[0].svg().unwrap().data().unwrap().as_bytes(), svg);
    assert!(observed.unrelated_reads() > 0);
    assert_eq!(observed.selected_reads(), 0);
}
