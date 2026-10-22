#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the in-memory package fixtures and assertions intentionally panic"
)]

//! Source-backed presentation master/layout catalog projections.

mod in_place;
use in_place::InPlace as _;

use std::io::{self, Cursor};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{Position, ReadAt, SourceVersion};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_pptx::master_layout::{PlaceholderKind, PlaceholderSpec, SlideLayoutKind};
use litchi_pptx::presentation::{raw as presentation_raw, related};
use litchi_pptx::{Error, Package, SourceBackedPresentation};

const MASTER_ONE: &str = "/ppt/slideMasters/slideMaster1.xml";

#[derive(Clone)]
struct CatalogFixture {
    bytes: Vec<u8>,
    master_one: PackURI,
    master_two: PackURI,
    layout_three: PackURI,
}

fn catalog_fixture() -> CatalogFixture {
    let mut package = Package::new().unwrap();
    let master_one = PackURI::new(MASTER_ONE).unwrap();
    let authored_master_two = package.add_slide_master_in_place().unwrap();
    let master_two = authored_master_two.part_name.clone();

    // Keep the fixture small and deterministic: the package template starts
    // with eleven layouts on master one; remove those and author two layouts
    // on each master.
    for index in 1..=11 {
        package
            .remove_slide_layout_in_place(
                &PackURI::new(format!("/ppt/slideLayouts/slideLayout{index}.xml")).unwrap(),
            )
            .unwrap();
    }
    let layout_one = package
        .add_slide_layout_in_place(&master_one, SlideLayoutKind::Title, "One Alpha", &[])
        .unwrap()
        .part_name;
    let layout_two = package
        .add_slide_layout_in_place(&master_one, SlideLayoutKind::Blank, "One Beta", &[])
        .unwrap()
        .part_name;
    let layout_three = package
        .add_slide_layout_in_place(
            &master_two,
            SlideLayoutKind::TitleOnly,
            "Two Alpha",
            &[PlaceholderSpec::new(PlaceholderKind::Body)],
        )
        .unwrap()
        .part_name;
    let layout_four = package
        .add_slide_layout_in_place(&master_two, SlideLayoutKind::SectionHeader, "Two Beta", &[])
        .unwrap()
        .part_name;

    let mut opc = OpcPackage::from_bytes(
        &package.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    set_c_sld_name(&mut opc, &master_one, "Master One");
    set_c_sld_name(&mut opc, &master_two, "Master Two");
    set_master_preserve(&mut opc, &master_one, true);
    set_master_preserve(&mut opc, &master_two, false);
    set_c_sld_name(&mut opc, &layout_one, "One Alpha");
    set_c_sld_name(&mut opc, &layout_two, "One Beta");
    set_c_sld_name(&mut opc, &layout_three, "Two Alpha");
    set_c_sld_name(&mut opc, &layout_four, "Two Beta");

    // Deliberately make XML order differ from lexical relationship-ID order.
    // The source facade must retain the ordered ID-list semantics.
    swap_id_entries(
        &mut opc,
        &PackURI::new("/ppt/presentation.xml").unwrap(),
        "sldMasterId",
    );
    swap_id_entries(&mut opc, &master_one, "sldLayoutId");
    swap_id_entries(&mut opc, &master_two, "sldLayoutId");

    CatalogFixture {
        bytes: PackageWriter::to_bytes(&opc).unwrap(),
        master_one,
        master_two,
        layout_three,
    }
}

fn edit_part(opc: &mut OpcPackage, uri: &PackURI, edit: impl FnOnce(String) -> String) {
    let part = opc.get_part_mut(uri).unwrap();
    let xml = String::from_utf8(part.blob().to_vec()).unwrap();
    part.set_blob(edit(xml).into_bytes());
}

fn set_c_sld_name(opc: &mut OpcPackage, uri: &PackURI, name: &str) {
    edit_part(opc, uri, |xml| {
        let start = xml.find("<p:cSld").unwrap();
        let end = start + xml[start..].find('>').unwrap();
        let tag = &xml[start..end];
        if let Some(name_start) = tag.find("name=\"") {
            let value_start = start + name_start + "name=\"".len();
            let value_end = value_start + xml[value_start..].find('"').unwrap();
            format!("{}{}{}", &xml[..value_start], name, &xml[value_end..])
        } else {
            format!("{} name=\"{name}\"{}", &xml[..end], &xml[end..])
        }
    });
}

fn set_master_preserve(opc: &mut OpcPackage, uri: &PackURI, preserve: bool) {
    edit_part(opc, uri, |xml| {
        let start = xml.find("<p:sldMaster").unwrap();
        let end = start + xml[start..].find('>').unwrap();
        let tag = &xml[start..end];
        if let Some(attr_start) = tag.find("preserve=\"") {
            let value_start = start + attr_start + "preserve=\"".len();
            let value_end = value_start + xml[value_start..].find('"').unwrap();
            format!(
                "{}{}{}",
                &xml[..value_start],
                if preserve { "1" } else { "0" },
                &xml[value_end..]
            )
        } else if preserve {
            format!("{} preserve=\"1\"{}", &xml[..end], &xml[end..])
        } else {
            xml
        }
    });
}

fn swap_id_entries(opc: &mut OpcPackage, uri: &PackURI, local: &str) {
    edit_part(opc, uri, |xml| swap_id_entries_xml(&xml, local));
}

fn swap_id_entries_xml(xml: &str, local: &str) -> String {
    let needle = format!("<p:{local} ");
    let Some(first_start) = xml.find(&needle) else {
        return xml.to_owned();
    };
    let first_end = first_start + xml[first_start..].find("/>").unwrap() + 2;
    let Some(relative_second) = xml[first_end..].find(&needle) else {
        return xml.to_owned();
    };
    let second_start = first_end + relative_second;
    let second_end = second_start + xml[second_start..].find("/>").unwrap() + 2;
    let first = &xml[first_start..first_end];
    let second = &xml[second_start..second_end];
    let mut output = String::with_capacity(xml.len());
    output.push_str(&xml[..first_start]);
    output.push_str(second);
    output.push_str(&xml[first_end..second_start]);
    output.push_str(first);
    output.push_str(&xml[second_end..]);
    output
}

fn open(bytes: &[u8]) -> SourceBackedPresentation {
    SourceBackedPresentation::from_reader(Cursor::new(bytes)).unwrap()
}

fn rewrite(bytes: &[u8], edit: impl FnOnce(&mut OpcPackage)) -> Vec<u8> {
    let mut opc = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    edit(&mut opc);
    PackageWriter::to_bytes(&opc).unwrap()
}

#[test]
fn source_catalog_uses_xml_order_and_exposes_raw_identity_explicitly() {
    let fixture = catalog_fixture();
    let source = open(&fixture.bytes);
    let masters = source.masters().unwrap();

    assert_eq!(masters.len(), 2);
    assert_eq!(masters[0].position(), Position::new(0));
    assert_eq!(masters[1].position(), Position::new(1));
    assert_eq!(masters[0].name(), Some("Master Two"));
    assert_eq!(masters[1].name(), Some("Master One"));
    assert!(!masters[0].is_preserved());
    assert!(masters[1].is_preserved());
    assert_eq!(related::raw::part_uri(&masters[0]), &fixture.master_two);
    assert!(!related::raw::relationship_id(&masters[0]).is_empty());
    assert!(!related::raw::data(&masters[0]).is_empty());
    assert_eq!(
        presentation_raw::part_uri(&source).as_str(),
        "/ppt/presentation.xml"
    );
    assert!(!presentation_raw::data(&source).is_empty());

    let layouts = source.layouts(Position::new(0)).unwrap().unwrap();
    assert_eq!(layouts.len(), 2);
    assert_eq!(layouts[0].name(), Some("Two Beta"));
    assert_eq!(layouts[1].name(), Some("Two Alpha"));
    assert!(!layouts[1].scene().is_empty());
    assert_eq!(
        related::raw::part_uri(
            &source
                .layout(Position::new(0), Position::new(1))
                .unwrap()
                .unwrap()
        ),
        &fixture.layout_three
    );
}

#[test]
fn source_catalog_master_and_layout_selectors_are_bounded() {
    let fixture = catalog_fixture();
    let source = open(&fixture.bytes);
    assert!(source.master(Position::new(2)).unwrap().is_none());
    assert!(source.layouts(Position::new(2)).unwrap().is_none());
    assert!(
        source
            .layout(Position::new(0), Position::new(2))
            .unwrap()
            .is_none()
    );
    assert!(source.master(Position::new(0)).unwrap().is_some());
}

#[test]
fn detached_master_scene_survives_drop_and_reuses_scene_cache() {
    let fixture = catalog_fixture();
    let source = open(&fixture.bytes);
    let first = source.master(Position::new(0)).unwrap().unwrap();
    let first_scene_len = first.scene().len();
    let before = source.scene_cache_diagnostics();
    let second = source.master(Position::new(0)).unwrap().unwrap();
    let after = source.scene_cache_diagnostics();
    assert!(after.hits > before.hits);
    assert_eq!(second.scene().len(), first_scene_len);
    drop(source);
    assert_eq!(first.scene().len(), first_scene_len);
}

#[test]
fn missing_and_empty_layout_lists_remain_distinct() {
    let fixture = catalog_fixture();
    let empty = rewrite(&fixture.bytes, |opc| {
        edit_part(opc, &fixture.master_two, |xml| {
            replace_layout_list(&xml, Some("<p:sldLayoutIdLst/>"))
        });
    });
    let source = open(&empty);
    assert!(source.layouts(Position::new(0)).unwrap().is_some());
    assert!(source.layouts(Position::new(1)).unwrap().is_some());
    assert!(
        source
            .layouts(Position::new(0))
            .unwrap()
            .unwrap()
            .is_empty()
    );

    let missing = rewrite(&fixture.bytes, |opc| {
        edit_part(opc, &fixture.master_two, |xml| {
            replace_layout_list(&xml, None)
        });
    });
    let source = open(&missing);
    assert!(source.layouts(Position::new(0)).unwrap().is_some());
    assert!(
        source
            .layouts(Position::new(0))
            .unwrap()
            .unwrap()
            .is_empty()
    );
    assert!(source.layouts(Position::new(2)).unwrap().is_none());
}

fn replace_layout_list(xml: &str, replacement: Option<&str>) -> String {
    let Some(start) = xml.find("<p:sldLayoutIdLst") else {
        return xml.to_owned();
    };
    let end = if xml[start..].starts_with("<p:sldLayoutIdLst/>") {
        start + "<p:sldLayoutIdLst/>".len()
    } else {
        start + xml[start..].find("</p:sldLayoutIdLst>").unwrap() + "</p:sldLayoutIdLst>".len()
    };
    format!(
        "{}{}{}",
        &xml[..start],
        replacement.unwrap_or(""),
        &xml[end..]
    )
}

#[test]
fn invalid_unselected_master_is_deferred_by_positional_read() {
    let fixture = catalog_fixture();
    let bytes = rewrite(&fixture.bytes, |opc| {
        opc.get_part_mut(&fixture.master_one)
            .unwrap()
            .set_blob(
                br#"<p:wrongRoot xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
                    .to_vec(),
            );
    });
    let source = open(&bytes);
    assert_eq!(
        source.master(Position::new(0)).unwrap().unwrap().name(),
        Some("Master Two")
    );
    assert!(source.master(Position::new(1)).is_err());
}

#[test]
fn bad_master_and_layout_relationships_are_refused_strictly() {
    let fixture = catalog_fixture();
    let bad_master = rewrite(&fixture.bytes, |opc| {
        let presentation = PackURI::new("/ppt/presentation.xml").unwrap();
        let part = opc.get_part_mut(&presentation).unwrap();
        let relationship = part
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rt::SLIDE_MASTER)
            .unwrap();
        let id = relationship.r_id().to_owned();
        part.rels_mut().remove(&id).unwrap();
        part.rels_mut()
            .try_add_relationship(
                rt::SLIDE_MASTER.into(),
                "https://example.invalid/master.xml".into(),
                id,
                TargetMode::External,
            )
            .unwrap();
    });
    assert!(open(&bad_master).master(Position::new(0)).is_err());

    let bad_layout = rewrite(&fixture.bytes, |opc| {
        let part = opc.get_part_mut(&fixture.master_two).unwrap();
        let relationship = part
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
            .unwrap();
        let id = relationship.r_id().to_owned();
        part.rels_mut().remove(&id).unwrap();
        part.rels_mut()
            .try_add_relationship(
                rt::SLIDE_LAYOUT.into(),
                "https://example.invalid/layout.xml".into(),
                id,
                TargetMode::External,
            )
            .unwrap();
    });
    assert!(
        open(&bad_layout)
            .layout(Position::new(0), Position::new(0))
            .is_err()
    );
}

#[test]
fn dimensions_reject_wrong_namespace_and_duplicate_sld_sz() {
    let fixture = catalog_fixture();
    let wrong_namespace = rewrite(&fixture.bytes, |opc| {
        let presentation = PackURI::new("/ppt/presentation.xml").unwrap();
        edit_part(opc, &presentation, |xml| {
            xml.replacen(
                "<p:sldSz",
                "<x:sldSz xmlns:x=\"urn:litchi:wrong-dimensions\"",
                1,
            )
        });
    });
    assert!(open(&wrong_namespace).slide_size().is_err());

    let duplicate = rewrite(&fixture.bytes, |opc| {
        let presentation = PackURI::new("/ppt/presentation.xml").unwrap();
        edit_part(opc, &presentation, |xml| {
            let start = xml.find("<p:sldSz").unwrap();
            let end = start + xml[start..].find("/>").unwrap() + 2;
            format!("{}<p:sldSz cx=\"1\" cy=\"1\"/>{}", &xml[..end], &xml[end..])
        });
    });
    assert!(open(&duplicate).slide_size().is_err());
}

#[test]
fn post_publication_master_fence_discards_every_built_scene() {
    let probe = Arc::new(ArmedRevision::new(catalog_fixture().bytes));
    let presentation = SourceBackedPresentation::from_read_at(probe.clone()).unwrap();
    let before = probe.checks.load(Ordering::SeqCst);
    presentation.masters().unwrap();
    let checks = probe.checks.load(Ordering::SeqCst) - before;
    assert!(checks > 2);

    let source = Arc::new(ArmedRevision::new(catalog_fixture().bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    source.change_on.store(
        source.checks.load(Ordering::SeqCst) + checks,
        Ordering::SeqCst,
    );
    assert!(matches!(
        presentation.masters(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.fence_discards, diagnostics.builds);
}

#[test]
fn post_publication_layout_fence_discards_every_built_scene() {
    let fixture = catalog_fixture();
    let probe = Arc::new(ArmedRevision::new(fixture.bytes.clone()));
    let presentation = SourceBackedPresentation::from_read_at(probe.clone()).unwrap();
    let before = probe.checks.load(Ordering::SeqCst);
    presentation.layouts(Position::new(0)).unwrap();
    let checks = probe.checks.load(Ordering::SeqCst) - before;
    assert!(checks > 2);

    let source = Arc::new(ArmedRevision::new(fixture.bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    source.change_on.store(
        source.checks.load(Ordering::SeqCst) + checks,
        Ordering::SeqCst,
    );
    assert!(matches!(
        presentation.layouts(Position::new(0)),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.fence_discards, diagnostics.builds);
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

struct ArmedRevision {
    inner: VersionedSource,
    checks: AtomicUsize,
    change_on: AtomicUsize,
}

impl ArmedRevision {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: VersionedSource::new(bytes),
            checks: AtomicUsize::new(0),
            change_on: AtomicUsize::new(usize::MAX),
        }
    }
}

impl ReadAt for VersionedSource {
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
        Ok(SourceVersion::new(61, self.revision.load(Ordering::SeqCst)))
    }
}

impl ReadAt for ArmedRevision {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let check = self.checks.fetch_add(1, Ordering::SeqCst) + 1;
        if check == self.change_on.load(Ordering::SeqCst) {
            self.inner.changed();
        }
        self.inner.version()
    }
}
