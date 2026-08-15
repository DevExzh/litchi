#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use std::collections::BTreeMap;
use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
    SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, SourceCacheLimits, TargetMode,
};
use litchi_xlsx::Error;
use litchi_xlsx::conditional_formatting::{
    DifferentialRef, Formatting, Kind, Range, Rule, Snapshot, SourceBackedEditor,
    replace_conditional_formattings,
};
use soapberry_zip::{PreservationIndex, ZipArchive};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SECOND: &str = "/xl/worksheets/sheet2.xml";
const STYLES: &str = "/xl/styles.xml";
const MEDIA: &str = "/xl/media/opaque.bin";

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
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            331,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

struct FailingSink(usize);

#[derive(Debug, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

fn raw_members(bytes: &[u8]) -> BTreeMap<String, RawMember> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    index
        .entries()
        .iter()
        .map(|preserved| {
            let record = records.next_entry().unwrap().unwrap();
            let name = record
                .file_path()
                .try_normalize()
                .unwrap()
                .as_ref()
                .to_owned();
            let local = bytes
                [preserved.local_span().start as usize..preserved.local_span().end as usize]
                .to_vec();
            let central_range = preserved.central_record();
            let mut central_without_offset =
                bytes[central_range.start as usize..central_range.end as usize].to_vec();
            central_without_offset[42..46].fill(0);
            (
                name,
                RawMember {
                    local,
                    central_without_offset,
                },
            )
        })
        .collect()
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.0 == 0 {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.0);
        self.0 -= written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn fixture(sheet_xml: &str, strict: bool, signed: bool) -> Vec<u8> {
    let namespace = if strict { STRICT_SML } else { SML };
    let rel_namespace = if strict { STRICT_REL } else { REL };
    let worksheet_relationship = if strict {
        rt::STRICT_WORKSHEET
    } else {
        rt::WORKSHEET
    };
    let styles_relationship = if strict {
        rt::STRICT_STYLES
    } else {
        rt::STYLES
    };
    let owner_relationship = if strict {
        rt::STRICT_OFFICE_DOCUMENT
    } else {
        rt::OFFICE_DOCUMENT
    };
    let workbook = format!(
        r#"<workbook xmlns="{namespace}" xmlns:r="{rel_namespace}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Untouched" sheetId="2" r:id="rIdUntouched"/></sheets></workbook>"#
    );
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook.into_bytes()),
        (SHEET, ct::SML_WORKSHEET, sheet_xml.as_bytes().to_vec()),
        (
            SECOND,
            ct::SML_WORKSHEET,
            format!(
                r#"<worksheet xmlns="{namespace}"><sheetData/><!--{}--></worksheet>"#,
                "untouched".repeat(4096)
            )
            .into_bytes(),
        ),
        (
            STYLES,
            ct::SML_STYLES,
            format!(
                r#"<styleSheet xmlns="{namespace}"><dxfs count="2"><dxf><fill/></dxf><dxf><font/></dxf></dxfs></styleSheet>"#
            )
            .into_bytes(),
        ),
        (
            MEDIA,
            "application/octet-stream",
            (0..32 * 1024).map(|value| (value % 251) as u8).collect(),
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                content_type.to_owned(),
                bytes,
            )))
            .unwrap();
    }
    let workbook_part = package.get_part_mut(&PackURI::new(MAIN).unwrap()).unwrap();
    workbook_part
        .rels_mut()
        .try_add_relationship(
            worksheet_relationship.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rIdSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    workbook_part
        .rels_mut()
        .try_add_relationship(
            worksheet_relationship.to_owned(),
            "worksheets/sheet2.xml".to_owned(),
            "rIdUntouched".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    workbook_part
        .rels_mut()
        .try_add_relationship(
            styles_relationship.to_owned(),
            "styles.xml".to_owned(),
            "rIdStyles".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/opaque.bin".to_owned(),
            "rIdOpaque".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("xl/workbook.xml", owner_relationship);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn worksheet_xml(namespace: &str, owners: &str) -> String {
    format!(
        r#"<worksheet xmlns="{namespace}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>{owners}<hyperlinks/><extLst><ext uri="opaque"><opaque xmlns="urn:vendor">keep</opaque></ext></extLst></worksheet>"#
    )
}

fn expression(range: &str, priority: i32, formula: &str, dxf: Option<u32>) -> Formatting {
    let mut rule = Rule::new(Kind::Expression, priority).unwrap();
    rule.push_formula(formula).unwrap();
    rule.differential_format = dxf.map(DifferentialRef::StylesIndex);
    Formatting::new(vec![Range::new(range).unwrap()], vec![rule]).unwrap()
}

fn initial_owners(namespace: &str) -> String {
    r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="expression" priority="2"><formula>B1&gt;0</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="C1"><cfRule type="expression" priority="3" dxfId="1"><formula>C1&gt;0</formula></cfRule></conditionalFormatting>"#
    .to_owned()
    .replace("<conditionalFormatting", &format!("<conditionalFormatting xmlns=\"{namespace}\""))
}

fn part_len(bytes: &[u8], member: &str) -> u64 {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(member).unwrap())
        .unwrap()
        .blob()
        .len() as u64
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-conditional-formatting-managed-test",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn relationships(relationships: &litchi_opc::Relationships) -> Vec<String> {
    let mut values = relationships
        .iter()
        .map(|value| {
            format!(
                "{}|{}|{}|{:?}",
                value.r_id(),
                value.reltype(),
                value.target_ref(),
                value.target_mode()
            )
        })
        .collect::<Vec<_>>();
    values.sort_unstable();
    values
}

#[test]
fn whole_collection_add_replace_clear_noop_and_order_are_exact() {
    for namespace in [SML, STRICT_SML] {
        let strict = namespace == STRICT_SML;
        let source = fixture(&worksheet_xml(namespace, ""), strict, false);
        let editor =
            SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source))).unwrap();
        let mut edit = editor.edit("Sheet1").unwrap();
        let first = expression("A1:A4", 1, "A1>2", Some(0));
        let middle = expression("B1:B4", 2, "B1<8", None);
        let last = expression("C1:C4", 3, "C1=4", Some(1));
        assert!(
            edit.set_collections(vec![first.clone(), middle.clone(), last.clone()])
                .unwrap()
        );
        assert!(
            !edit
                .set_collections(vec![first.clone(), middle.clone(), last.clone()])
                .unwrap()
        );
        let commit = edit.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            commit.snapshot().collections(),
            &[first.clone(), middle.clone(), last.clone()]
        );

        let mut package =
            OpcPackage::from_bytes(&fixture(&worksheet_xml(namespace, ""), strict, false)).unwrap();
        commit.patch().apply(&mut package).unwrap();
        assert_eq!(
            Snapshot::load(&package, 0usize).unwrap().collections(),
            &[first.clone(), middle.clone(), last.clone()]
        );
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(
            Snapshot::load(&package, "Sheet1")
                .unwrap()
                .collections()
                .is_empty()
        );

        let existing = fixture(
            &worksheet_xml(namespace, &initial_owners(namespace)),
            strict,
            false,
        );
        let editor =
            SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(existing))).unwrap();
        let mut edit = editor.edit("Sheet1").unwrap();
        assert!(
            edit.set_collections(vec![last.clone(), first.clone()])
                .unwrap()
        );
        let commit = edit.commit().unwrap();
        assert_eq!(
            commit.snapshot().collections(),
            &[last.clone(), first.clone()]
        );

        for expected in [
            vec![middle.clone(), last.clone()],
            vec![first.clone(), last.clone()],
            vec![first.clone(), middle.clone()],
            Vec::new(),
        ] {
            let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
                &worksheet_xml(namespace, &initial_owners(namespace)),
                strict,
                false,
            ))))
            .unwrap();
            let mut edit = editor.edit("Sheet1").unwrap();
            assert!(edit.set_collections(expected.clone()).unwrap());
            assert_eq!(edit.commit().unwrap().snapshot().collections(), expected);
        }

        let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &worksheet_xml(namespace, &initial_owners(namespace)),
            strict,
            false,
        ))))
        .unwrap();
        let mut edit = editor.edit("Sheet1").unwrap();
        assert!(edit.clear());
        assert!(edit.commit().unwrap().snapshot().collections().is_empty());

        let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &worksheet_xml(namespace, ""),
            strict,
            false,
        ))))
        .unwrap();
        assert!(!editor.edit("Sheet1").unwrap().commit().unwrap().changed());
    }
}

#[test]
fn publication_reopens_and_preserves_every_unselected_part_and_relationship() {
    let source = fixture(&worksheet_xml(SML, ""), false, false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    let expected = vec![expression("A1:C9", 1, "$A1>5", Some(1))];
    edit.set_collections(expected.clone()).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();

    let source_package = OpcPackage::from_bytes(&source).unwrap();
    let output_package = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(
        Snapshot::load(&output_package, "Sheet1")
            .unwrap()
            .collections(),
        expected
    );
    for source_part in source_package.iter_parts() {
        let output_part = output_package.get_part(source_part.partname()).unwrap();
        assert_eq!(output_part.content_type(), source_part.content_type());
        assert_eq!(
            relationships(output_part.rels()),
            relationships(source_part.rels())
        );
        if source_part.partname() != &PackURI::new(SHEET).unwrap() {
            assert_eq!(output_part.blob(), source_part.blob());
        }
    }
    let source_raw = raw_members(&source);
    let output_raw = raw_members(&output);
    assert_eq!(
        source_raw.keys().collect::<Vec<_>>(),
        output_raw.keys().collect::<Vec<_>>()
    );
    for (name, member) in source_raw {
        if name != "xl/worksheets/sheet1.xml" {
            assert_eq!(output_raw.get(&name), Some(&member), "raw member {name}");
        }
    }
}

#[test]
fn managed_snapshot_changed_publication_retains_styles_and_releases_budget() {
    let source_bytes = fixture(&worksheet_xml(SML, ""), false, false);
    let exact = part_len(&source_bytes, MAIN)
        + part_len(&source_bytes, SHEET)
        + part_len(&source_bytes, STYLES);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor =
        SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(VersionedSource::new(source_bytes.clone())),
            litchi_xlsx::ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap(),
            context,
        )
        .unwrap();
    let snapshot = editor.snapshot("Sheet1").unwrap();
    assert!(snapshot.collections().is_empty());
    assert_eq!(budget.used(Resource::Memory), exact);
    drop(snapshot);

    let expected = vec![expression("A1:C9", 1, "$A1>5", Some(1))];
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_collections(expected.clone()).unwrap();
    let commit = edit.commit().unwrap();
    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        Snapshot::load(&replay, "Sheet1").unwrap().collections(),
        expected.as_slice()
    );
    assert!(matches!(
        commit.patch().inverse().apply(&mut replay),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(
        Snapshot::load(&published, "Sheet1").unwrap().collections(),
        expected.as_slice()
    );
    assert_eq!(
        published
            .get_part(&PackURI::new(MEDIA).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&source_bytes)
            .unwrap()
            .get_part(&PackURI::new(MEDIA).unwrap())
            .unwrap()
            .blob()
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn closure_conflicts_signatures_protection_and_partial_sinks_fail_closed() {
    let unsigned = fixture(&worksheet_xml(SML, ""), false, false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(unsigned.clone()))).unwrap();
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
    let mut exact = Vec::new();
    editor
        .publish_commit_to_stream(&mut exact, &commit)
        .unwrap();
    assert_eq!(exact, unsigned);

    let signed = fixture(&worksheet_xml(SML, ""), false, true);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
    let mut exact = Vec::new();
    editor
        .publish_commit_to_stream(&mut exact, &commit)
        .unwrap();
    assert_eq!(exact, signed);

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
        &worksheet_xml(SML, ""),
        false,
        true,
    ))))
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_collections(vec![expression("A1", 1, "1", None)])
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );

    let source = Arc::new(VersionedSource::new(fixture(
        &worksheet_xml(SML, ""),
        false,
        false,
    )));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_collections(vec![expression("A1", 1, "1", None)])
        .unwrap();
    let commit = edit.commit().unwrap();
    source.changed();
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
        &worksheet_xml(SML, ""),
        false,
        false,
    ))))
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_collections(vec![expression("A1", 1, "1", None)])
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        editor
            .publish_commit_to_stream(FailingSink(32), &commit)
            .is_err()
    );

    let mut foreign =
        OpcPackage::from_bytes(&fixture(&worksheet_xml(SML, ""), false, false)).unwrap();
    foreign
        .get_part_mut(&PackURI::new(STYLES).unwrap())
        .unwrap()
        .set_blob(
            format!(r#"<styleSheet xmlns="{SML}"><dxfs count="1"><dxf/></dxfs></styleSheet>"#)
                .into_bytes(),
        );
    assert!(matches!(
        commit.patch().apply(&mut foreign),
        Err(Error::PatchConflict { .. })
    ));

    let protected = fixture(
        &worksheet_xml(SML, r#"<sheetProtection sheet="1" formatCells="1"/>"#),
        false,
        false,
    );
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(protected))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_collections(vec![expression("A1", 1, "1", None)])
        .unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn x14_mce_opaque_dxf_priority_malformed_and_limits_are_refused() {
    let x14 = format!(
        r#"<worksheet xmlns="{SML}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData/><extLst><ext uri="{{78C0D931-6437-407D-A8EE-F0AAD7539E65}}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"/><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#,
    );
    assert!(
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &x14, false, false
        ))))
        .unwrap()
        .snapshot("Sheet1")
        .is_err()
    );

    let mce = format!(
        r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><sheetData/><mc:AlternateContent><mc:Choice Requires="x"><x:conditionalFormatting/></mc:Choice><mc:Fallback/></mc:AlternateContent></worksheet>"#
    );
    assert!(
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &mce, false, false
        ))))
        .unwrap()
        .snapshot("Sheet1")
        .is_err()
    );

    for owner in [
        r#"<conditionalFormatting sqref="A1" vendor="opaque"><cfRule type="expression" priority="1"/></conditionalFormatting>"#,
        r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><vendor/></cfRule></conditionalFormatting>"#,
        r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="2"/></conditionalFormatting>"#,
        r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"/></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="expression" priority="1"/></conditionalFormatting>"#,
    ] {
        let bytes = fixture(&worksheet_xml(SML, owner), false, false);
        assert!(
            SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes)))
                .unwrap()
                .snapshot("Sheet1")
                .is_err()
        );
    }

    let empty_formula_second_owner = format!(
        r#"<worksheet xmlns="{SML}"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="expression" priority="2"><formula/></cfRule></conditionalFormatting></worksheet>"#
    );
    assert!(
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &empty_formula_second_owner,
            false,
            false,
        ))))
        .unwrap()
        .snapshot("Sheet1")
        .is_err()
    );

    let malformed = format!(r#"<?bad value?><worksheet xmlns="{SML}"><sheetData/></worksheet>"#);
    assert!(
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &malformed, false, false
        ))))
        .unwrap()
        .snapshot("Sheet1")
        .is_err()
    );

    let oversized = vec![b' '; 32 * 1024 * 1024 + 1];
    assert!(replace_conditional_formattings(&oversized, &[], 0).is_err());

    for worksheet in [
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="future"><sheetData/></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:q="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><sheetData q:future="1"/></worksheet>"#
        ),
        format!(r#"<worksheet xmlns="{SML}"><hyperlinks/><sheetData/></worksheet>"#),
        format!(r#"<worksheet xmlns="{SML}"><sheetData/><unknown/></worksheet>"#),
        format!(r#"<worksheet xmlns="{SML}"><sheetData/><sheetData/></worksheet>"#),
        format!(r#"<worksheet xmlns="{SML}"><sheetPr/></worksheet>"#),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:q="urn:vendor"><sheetData/><q:opaque/></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:q="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheetData/><q:Choice/></worksheet>"#
        ),
    ] {
        let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
            &worksheet, false, false,
        ))))
        .unwrap();
        assert!(editor.snapshot("Sheet1").is_err());
    }
}
