//! Batch 0/1 read-only integration coverage for worksheet form controls.
//!
//! The fixture corpus is deliberately exercised through the public worksheet
//! facades.  ZIP rewrites below only create bounded, structurally valid test
//! inputs for a single targeted refusal or fallback rule; no test writes an
//! XLSX result or claims native application acceptance.
#![allow(
    clippy::unwrap_used,
    reason = "focused fixture assertions panic on failure"
)]
#![allow(
    clippy::expect_used,
    reason = "focused fixture assertions panic on failure"
)]

use std::io::{self, Cursor};
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, EvidenceDigest, ExecutionContext, ExecutionLimits,
    Limits as BudgetLimits, OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_opc::{OpcError, ReadLimits};
use litchi_xlsx::form_control::{
    ControlSelector, FormControlDiagnosticCode, KnownOrUnknown, OwnerLimits, OwnerProfile,
    Properties,
};
use litchi_xlsx::{Error, SourceBackedWorkbook, Workbook};
use serde_json::Value;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const FORM_CONTROL_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SML_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const CONTROL_PROPERTIES_CONTENT_TYPE: &str = "application/vnd.ms-excel.controlproperties+xml";
const ACTIVEX_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const CONTROL_PROPERTIES_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const ACTIVEX_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
const STRICT_ACTIVEX_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
const UNKNOWN_CONTROL_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/unknownControl";
const ACTIVEX_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/activeX";
const ACTIVEX_DESCRIPTOR_MEMBER: &str = "xl/activeX/activeX1.xml";
const NATIVE_CORPUS: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json"
));

/// Deterministically flips the source revision immediately before one version
/// observation, so the owner final fence can be exercised without a timing
/// race in the test.
struct VersionFlipSource {
    bytes: Vec<u8>,
    version_calls: AtomicUsize,
    revision: AtomicU64,
    flip_before_call: Option<usize>,
}

impl VersionFlipSource {
    fn new(bytes: Vec<u8>, flip_before_call: Option<usize>) -> Self {
        Self {
            bytes,
            version_calls: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
            flip_before_call,
        }
    }
}

impl ReadAt for VersionFlipSource {
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
        let call = self.version_calls.fetch_add(1, Ordering::SeqCst) + 1;
        if self.flip_before_call == Some(call) {
            self.revision.store(1, Ordering::SeqCst);
        }
        Ok(SourceVersion::new(92, self.revision.load(Ordering::SeqCst)))
    }
}

#[derive(Clone, Copy)]
struct Fixture {
    file: &'static str,
    names: &'static [&'static str],
    shape_ids: &'static [u64],
    object_types: &'static [&'static str],
    vml_ids: &'static [&'static str],
    vml_spids: &'static [Option<&'static str>],
}

const FIXTURES: &[Fixture] = &[
    Fixture {
        file: "button-form-control.xlsx",
        names: &["Button 1"],
        shape_ids: &[1025],
        object_types: &["Button"],
        vml_ids: &["_x0000_s1025"],
        vml_spids: &[None],
    },
    Fixture {
        file: "checkbox-form-control.xlsx",
        names: &["Check Box 1"],
        shape_ids: &[1025],
        object_types: &["CheckBox"],
        vml_ids: &["_x0000_s1025"],
        vml_spids: &[None],
    },
    Fixture {
        file: "singlecontrol.xlsx",
        names: &["Check Box 2"],
        shape_ids: &[1026],
        object_types: &["CheckBox"],
        vml_ids: &["_x0000_s1026"],
        vml_spids: &[None],
    },
    Fixture {
        file: "tdf120301_xmlSpaceParsing.xlsx",
        names: &["Check Box 1", "Option Button 2"],
        shape_ids: &[1025, 1026],
        object_types: &["CheckBox", "Radio"],
        vml_ids: &["_x0000_s1025", "_x0000_s1026"],
        vml_spids: &[None, None],
    },
    Fixture {
        file: "tdf134769.xlsx",
        names: &["Check Box 1"],
        shape_ids: &[1025],
        object_types: &["CheckBox"],
        vml_ids: &["_x0000_s1025"],
        vml_spids: &[None],
    },
    Fixture {
        file: "tdf161365.xlsx",
        names: &["Check Box 4", "Check Box 4"],
        shape_ids: &[1026, 1027],
        object_types: &["CheckBox", "CheckBox"],
        vml_ids: &["Check_x0020_Box_x0020_4", "_x0000_s1027"],
        vml_spids: &[Some("_x0000_s1026"), None],
    },
    Fixture {
        file: "tdf60673.xlsx",
        names: &["Button 1", "Button 2"],
        shape_ids: &[1025, 1026],
        object_types: &["Button", "Button"],
        vml_ids: &["_x0000_s1025", "_x0000_s1026"],
        vml_spids: &[None, None],
    },
];

fn fixture_path(file: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/form_control_properties")
        .join(file)
}

fn fixture_bytes(file: &str) -> Vec<u8> {
    std::fs::read(fixture_path(file)).unwrap()
}

fn json_string<'a>(value: &'a Value, key: &str) -> &'a str {
    value[key].as_str().unwrap()
}

fn known_token<T: std::fmt::Display>(value: Option<&KnownOrUnknown<T>>) -> Option<String> {
    value.map(|value| match value {
        KnownOrUnknown::Known(value) => value.to_string(),
        KnownOrUnknown::Unknown(value) => value.clone(),
    })
}

fn rewrite_member(source: &[u8], member: &str, edit: impl FnOnce(String) -> String) -> Vec<u8> {
    let archive = ArchiveReader::new(source).unwrap();
    let updated = edit(String::from_utf8(archive.read(member).unwrap()).unwrap()).into_bytes();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let bytes = if name == member {
            updated.clone()
        } else {
            archive.read(name).unwrap()
        };
        writer.write_deflated(name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn add_member(source: &[u8], member: &str, bytes: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(source).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        writer
            .write_deflated(name, &archive.read(name).unwrap())
            .unwrap();
    }
    writer.write_deflated(member, bytes).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn add_activex_descriptor(source: &[u8]) -> Vec<u8> {
    let source = rewrite_member(source, "[Content_Types].xml", |xml| {
        xml.replace(
            "</Types>",
            &format!(
                "<Override PartName=\"/{ACTIVEX_DESCRIPTOR_MEMBER}\" ContentType=\"{ACTIVEX_CONTENT_TYPE}\"/></Types>"
            ),
        )
    });
    let descriptor = format!(
        "<ax:ocx xmlns:ax=\"{ACTIVEX_NAMESPACE}\" ax:classid=\"form-control-test\" ax:persistence=\"persistPropertyBag\"><ax:ocxPr ax:name=\"Opaque\" ax:value=\"retained\"/></ax:ocx>"
    );
    add_member(&source, ACTIVEX_DESCRIPTOR_MEMBER, descriptor.as_bytes())
}

fn with_extra_activex_control(
    source: &[u8],
    relationship_type: &str,
    shape_id: u64,
    name: &str,
) -> Vec<u8> {
    let source = rewrite_member(source, "xl/worksheets/sheet1.xml", |xml| {
        let start = xml.find("<control shapeId=").unwrap();
        let end = xml[start..].find("</control>").unwrap() + start + "</control>".len();
        let control = format!("<control shapeId=\"{shape_id}\" r:id=\"rId5\" name=\"{name}\"/>");
        format!("{}{}{}", &xml[..end], control, &xml[end..])
    });
    let source = rewrite_member(&source, "xl/worksheets/_rels/sheet1.xml.rels", |xml| {
        xml.replace(
                "</Relationships>",
                &format!(
                    "<Relationship Id=\"rId5\" Type=\"{relationship_type}\" Target=\"../activeX/activeX1.xml\"/></Relationships>"
                ),
            )
    });
    add_activex_descriptor(&source)
}

fn with_activex_control_replacing_form_control(source: &[u8], relationship_type: &str) -> Vec<u8> {
    let source = rewrite_member(source, "xl/worksheets/_rels/sheet1.xml.rels", |xml| {
        xml.replace(
                &format!(
                    "Id=\"rId3\" Type=\"{CONTROL_PROPERTIES_RELATIONSHIP}\" Target=\"../ctrlProps/ctrlProp1.xml\""
                ),
                &format!(
                    "Id=\"rId3\" Type=\"{relationship_type}\" Target=\"../activeX/activeX1.xml\""
                ),
            )
    });
    add_activex_descriptor(&source)
}

fn with_ctrlprop_edge_targeting_activex_descriptor(source: &[u8]) -> Vec<u8> {
    let source = rewrite_member(source, "xl/worksheets/_rels/sheet1.xml.rels", |xml| {
        xml.replace("../ctrlProps/ctrlProp1.xml", "../activeX/activeX1.xml")
    });
    add_activex_descriptor(&source)
}

fn rewrite_members(source: &[u8], edits: &[(&str, &dyn Fn(String) -> String)]) -> Vec<u8> {
    let archive = ArchiveReader::new(source).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let mut value = String::from_utf8(archive.read(name).unwrap())
            .unwrap_or_else(|_| panic!("member {name} is not text; omit it from rewrite_members"));
        for (target, edit) in edits {
            if *target == name {
                value = edit(value);
            }
        }
        writer.write_deflated(name, value.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn entity_escape_uri(xml: String, uri: &str, escaped: &str) -> String {
    assert!(
        xml.contains(uri),
        "namespace or relationship URI {uri:?} was absent from the fixture member"
    );
    xml.replace(uri, escaped)
}

fn expected_fixture(file: &str) -> Fixture {
    FIXTURES
        .iter()
        .copied()
        .find(|fixture| fixture.file == file)
        .unwrap()
}

fn assert_native_attributes(properties: &Properties, attributes: &serde_json::Map<String, Value>) {
    for (name, expected) in attributes {
        let expected = expected.as_str().unwrap();
        match name.as_str() {
            "objectType" => assert_eq!(
                known_token(properties.object_type()),
                Some(expected.to_owned())
            ),
            "checked" => assert_eq!(known_token(properties.checked()), Some(expected.to_owned())),
            "fmlaLink" => assert_eq!(
                properties.fmla_link().map(|value| value.as_str()),
                Some(expected)
            ),
            "lockText" => assert_eq!(properties.lock_text(), Some(expected == "1")),
            "noThreeD" => assert_eq!(properties.no_three_d(), Some(expected == "1")),
            "firstButton" => assert_eq!(properties.first_button(), Some(expected == "1")),
            other => panic!("uncovered native form-control attribute {other}"),
        }
    }
}

fn assert_shape(
    view: &litchi_xlsx::form_control::FormControlView,
    fixture: Fixture,
    position: usize,
) {
    let shape = view.shape();
    assert_eq!(shape.shape_id(), fixture.shape_ids[position]);
    assert_eq!(shape.drawing_name(), Some(fixture.names[position]));
    assert_eq!(
        shape.compat_spid(),
        format!("_x0000_s{}", fixture.shape_ids[position])
    );
    assert_eq!(shape.vml_id(), fixture.vml_ids[position]);
    assert_eq!(shape.vml_spid(), fixture.vml_spids[position]);
    assert_eq!(shape.object_type(), fixture.object_types[position]);
}

fn source_collection(bytes: &[u8]) -> litchi_xlsx::form_control::FormControlCollection {
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    workbook.sheet(0).unwrap().unwrap().form_controls().unwrap()
}

fn managed_owner_context(
    objects: u64,
    depth: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-form-control-owner-managed-test",
        BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, objects, depth, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn managed_owner_usage(budget: &Budget) -> [u64; 5] {
    [
        budget.used(Resource::Memory),
        budget.used(Resource::InputBytes),
        budget.used(Resource::OutputBytes),
        budget.used(Resource::Objects),
        budget.used(Resource::Depth),
    ]
}

fn assert_execution_depth_limit<T>(result: litchi_xlsx::Result<T>, label: &str) {
    match result {
        Err(Error::ResourceLimit(limit)) => assert_eq!(
            limit.resource,
            Resource::Depth,
            "{label} returned the wrong resource: {limit:?}"
        ),
        Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
        Ok(_) => panic!("{label} was admitted"),
    }
}

fn assert_owner_refuses(source: Vec<u8>, label: &str) {
    assert!(
        Workbook::from_bytes(source.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err(),
        "{label} was admitted by the eager owner"
    );
    assert!(
        SourceBackedWorkbook::from_reader(Cursor::new(source))
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err(),
        "{label} was admitted by the source owner"
    );
}

fn assert_resource_limit<T>(result: litchi_xlsx::Result<T>, label: &str) {
    match result {
        Err(Error::ResourceLimit(_)) => {},
        Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
        Ok(_) => panic!("{label} was admitted"),
    }
}

fn projection_cap_admits_eager(source: &[u8], cap: usize, label: &str) -> bool {
    let result = Workbook::from_bytes(source.to_vec())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls_with_limits(OwnerLimits::new().with_max_projection_bytes(cap));
    match result {
        Ok(_) => true,
        Err(Error::ResourceLimit(limit)) => {
            assert_eq!(limit.resource, Resource::Memory, "{label}: {limit:?}");
            assert_eq!(
                limit.limit,
                u64::try_from(cap).unwrap(),
                "{label}: {limit:?}"
            );
            assert!(limit.scope.contains("projection"), "{label}: {limit:?}");
            false
        },
        Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
    }
}

fn projection_cap_admits_source(source: &[u8], cap: usize, label: &str) -> bool {
    let result = SourceBackedWorkbook::from_reader(Cursor::new(source.to_vec()))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls_with_limits(OwnerLimits::new().with_max_projection_bytes(cap));
    match result {
        Ok(_) => true,
        Err(Error::ResourceLimit(limit)) => {
            assert_eq!(limit.resource, Resource::Memory, "{label}: {limit:?}");
            assert_eq!(
                limit.limit,
                u64::try_from(cap).unwrap(),
                "{label}: {limit:?}"
            );
            assert!(limit.scope.contains("projection"), "{label}: {limit:?}");
            false
        },
        Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
    }
}

fn minimum_projection_cap<F>(mut admits: F, label: &str) -> usize
where
    F: FnMut(usize) -> bool,
{
    let mut lower = 0;
    let mut upper = OwnerLimits::new().max_projection_bytes();
    assert!(admits(upper), "{label} was not admitted at the default cap");
    while lower < upper {
        let middle = lower + (upper - lower) / 2;
        if admits(middle) {
            upper = middle;
        } else {
            lower = middle + 1;
        }
    }
    assert!(admits(lower), "{label} minimum cap was not admitted");
    if lower > 0 {
        assert!(
            !admits(lower - 1),
            "{label} was admitted one byte below its minimum cap"
        );
    }
    lower
}

fn assert_form_control_error<T>(result: litchi_xlsx::Result<T>, label: &str) {
    match result {
        Err(Error::FormControl(_)) => {},
        Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
        Ok(_) => panic!("{label} was admitted"),
    }
}

fn assert_typed_limit<T>(result: litchi_xlsx::Result<T>, label: &str) {
    match result {
        Err(Error::ResourceLimit(_))
        | Err(Error::FormControl(litchi_xlsx::form_control::FormControlError::Limit { .. })) => {},
        Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
        Ok(_) => panic!("{label} was admitted"),
    }
}

#[test]
fn native_corpus_inventory_shape_closure_and_source_bytes() {
    let corpus: Value = serde_json::from_str(NATIVE_CORPUS).unwrap();
    let fixtures = corpus["fixtures"].as_array().unwrap();
    assert_eq!(fixtures.len(), FIXTURES.len());
    let mut part_count = 0;

    for manifest in fixtures {
        let file = Path::new(json_string(manifest, "retained_path"))
            .file_name()
            .unwrap()
            .to_str()
            .unwrap();
        let fixture = expected_fixture(file);
        let source = fixture_bytes(file);
        assert_eq!(
            source.len(),
            manifest["bytes"].as_u64().unwrap() as usize,
            "retained fixture byte count drifted for {file}"
        );
        assert_eq!(
            EvidenceDigest::of(&source).to_string(),
            json_string(manifest, "sha256"),
            "retained fixture SHA-256 drifted for {file}"
        );
        let archive = ArchiveReader::new(&source).unwrap();
        let eager = Workbook::from_bytes(source.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .unwrap();
        let deferred = source_collection(&source);
        assert_eq!(eager.profile(), OwnerProfile::canonical());
        assert_eq!(deferred.profile(), OwnerProfile::canonical());
        assert!(eager.diagnostics().is_empty());
        assert!(deferred.diagnostics().is_empty());
        assert!(eager.read_set().is_none());
        assert!(deferred.read_set().is_some());
        assert_eq!(eager.len(), fixture.names.len());
        assert_eq!(deferred.len(), eager.len());

        let parts = manifest["form_control_parts"].as_array().unwrap();
        assert_eq!(parts.len(), eager.len());
        for (position, part) in parts.iter().enumerate() {
            part_count += 1;
            let member = json_string(part, "member");
            let raw_properties = archive.read(member).unwrap();
            assert_eq!(
                raw_properties.len(),
                part["bytes"].as_u64().unwrap() as usize,
                "retained member byte count drifted for {file}:{member}"
            );
            assert_eq!(
                EvidenceDigest::of(&raw_properties).to_string(),
                json_string(part, "sha256"),
                "retained member SHA-256 drifted for {file}:{member}"
            );
            let attributes = part["attributes"].as_object().unwrap();
            let eager_view = eager
                .get(ControlSelector::position(position))
                .unwrap()
                .unwrap();
            let deferred_view = deferred
                .get(ControlSelector::position(position))
                .unwrap()
                .unwrap();

            assert_eq!(eager_view.position(), position);
            assert_eq!(eager_view.name(), Some(fixture.names[position]));
            assert_eq!(eager_view, deferred_view);
            assert_eq!(eager_view.anchor_profile(), Some("LoSmlAnchorV1"));
            assert_eq!(
                eager_view.properties().source_bytes(),
                Some(raw_properties.as_slice())
            );
            assert_native_attributes(eager_view.properties(), attributes);
            assert_eq!(
                known_token(eager_view.properties().object_type()),
                Some(fixture.object_types[position].to_owned())
            );
            assert_shape(eager_view, fixture, position);

            let selected = eager.get(ControlSelector::name(fixture.names[position]));
            if fixture
                .names
                .iter()
                .filter(|name| **name == fixture.names[position])
                .count()
                == 1
            {
                assert_eq!(selected.unwrap().unwrap(), eager_view);
            } else {
                assert!(selected.is_err());
            }
        }
    }
    assert_eq!(
        part_count, 10,
        "the owner must reopen all ten retained native parts"
    );
}

#[test]
fn duplicate_name_selector_is_ambiguous_but_positions_remain_checked() {
    let source = fixture_bytes("tdf161365.xlsx");
    let workbook = Workbook::from_bytes(source.clone()).unwrap();
    let sheet = workbook.sheet(0).unwrap().unwrap();
    let controls = sheet.form_controls().unwrap();
    assert_eq!(controls.len(), 2);
    assert!(controls.get(ControlSelector::name("Check Box 4")).is_err());
    assert!(
        sheet
            .form_control(ControlSelector::name("Check Box 4"))
            .is_err()
    );
    assert_eq!(
        sheet
            .form_control(ControlSelector::position(0))
            .unwrap()
            .unwrap()
            .name(),
        Some("Check Box 4")
    );
    assert_eq!(
        sheet.form_control(ControlSelector::position(2)).unwrap(),
        None
    );

    let source_sheet = SourceBackedWorkbook::from_reader(Cursor::new(source))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert!(
        source_sheet
            .form_controls()
            .unwrap()
            .get(ControlSelector::name("Check Box 4"))
            .is_err()
    );
    assert!(
        source_sheet
            .form_control(ControlSelector::name("Check Box 4"))
            .is_err()
    );
}

#[test]
fn aliases_and_xml_entities_preserve_semantic_owner_reads() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let aliased = rewrite_members(
        &source,
        &[
            ("xl/worksheets/sheet1.xml", &|xml: String| {
                xml.replace("Requires=\"x14\"", "Requires=\"x14Alias\"")
                        .replace(
                            "<mc:Choice Requires=\"x14Alias\">",
                            &format!(
                                "<mc:Choice xmlns:x14Alias=\"{FORM_CONTROL_NAMESPACE}\" Requires=\"x14Alias\">"
                            ),
                        )
                        .replace("name=\"Check Box 1\"", "name=\"Check &amp; Box 1\"")
            }),
            ("xl/drawings/drawing1.xml", &|xml: String| {
                xml.replace(
                        "xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\"",
                        "xmlns:z14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"z14\"",
                    )
                    .replace("mc:Ignorable=\"a14\"", "mc:Ignorable=\"z14\"")
                    .replace("name=\"Check Box 1\"", "name=\"Check &amp; Box 1\"")
                    .replace("a14:", "z14:")
            }),
        ],
    );
    let workbook = Workbook::from_bytes(aliased.clone()).unwrap();
    let view = workbook
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(view.name(), Some("Check & Box 1"));
    assert_eq!(view.shape().drawing_name(), Some("Check & Box 1"));

    let source_view = SourceBackedWorkbook::from_reader(Cursor::new(aliased))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(source_view.name(), Some("Check & Box 1"));

    let properties_alias = rewrite_member(&source, "xl/ctrlProps/ctrlProp1.xml", |xml| {
        xml.replace(
            &format!("<formControlPr xmlns=\"{FORM_CONTROL_NAMESPACE}\""),
            &format!("<fc:formControlPr xmlns:fc=\"{FORM_CONTROL_NAMESPACE}\""),
        )
        .replace("</formControlPr>", "</fc:formControlPr>")
    });
    let properties_view = Workbook::from_bytes(properties_alias)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(properties_view.properties().no_three_d(), Some(true));
}

#[test]
fn entity_escaped_namespace_uris_preserve_semantics_and_source_bytes() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let escaped = rewrite_members(
        &source,
        &[
            ("xl/worksheets/sheet1.xml", &|xml: String| {
                let xml = entity_escape_uri(
                    xml,
                    SML_NAMESPACE,
                    "http://schemas.openxmlformats.org&#x2F;spreadsheetml&#x2F;2006&#x2F;main",
                );
                let xml = entity_escape_uri(
                    xml,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "http://schemas.openxmlformats.org&#x2F;officeDocument&#x2F;2006&#x2F;relationships",
                );
                entity_escape_uri(
                    xml,
                    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    "http://schemas.openxmlformats.org&#x2F;drawingml&#x2F;2006&#x2F;spreadsheetDrawing",
                )
            }),
            ("xl/worksheets/_rels/sheet1.xml.rels", &|xml: String| {
                entity_escape_uri(
                    xml,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "http://schemas.openxmlformats.org&#x2F;officeDocument&#x2F;2006&#x2F;relationships",
                )
            }),
            ("xl/drawings/drawing1.xml", &|xml: String| {
                let xml = entity_escape_uri(
                    xml,
                    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    "http://schemas.openxmlformats.org&#x2F;drawingml&#x2F;2006&#x2F;spreadsheetDrawing",
                );
                entity_escape_uri(
                    xml,
                    "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "http://schemas.openxmlformats.org&#x2F;drawingml&#x2F;2006&#x2F;main",
                )
            }),
            ("xl/ctrlProps/ctrlProp1.xml", &|xml: String| {
                entity_escape_uri(
                    xml,
                    FORM_CONTROL_NAMESPACE,
                    "http://schemas.microsoft.com&#x2F;office&#x2F;spreadsheetml&#x2F;2009&#x2F;9&#x2F;main",
                )
            }),
            ("xl/drawings/vmlDrawing1.vml", &|xml: String| {
                let xml = entity_escape_uri(
                    xml,
                    "urn:schemas-microsoft-com:vml",
                    "urn&#x3A;schemas-microsoft-com&#x3A;vml",
                );
                let xml = entity_escape_uri(
                    xml,
                    "urn:schemas-microsoft-com:office:office",
                    "urn&#x3A;schemas-microsoft-com&#x3A;office&#x3A;office",
                );
                entity_escape_uri(
                    xml,
                    "urn:schemas-microsoft-com:office:excel",
                    "urn&#x3A;schemas-microsoft-com&#x3A;office&#x3A;excel",
                )
            }),
        ],
    );

    let source_archive = ArchiveReader::new(&source).unwrap();
    let escaped_archive = ArchiveReader::new(&escaped).unwrap();
    for member in [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/drawings/drawing1.xml",
        "xl/ctrlProps/ctrlProp1.xml",
        "xl/drawings/vmlDrawing1.vml",
    ] {
        assert_ne!(
            source_archive.read(member).unwrap(),
            escaped_archive.read(member).unwrap(),
            "namespace entity mutation did not change {member}"
        );
    }

    let baseline = Workbook::from_bytes(source)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    let eager_controls = Workbook::from_bytes(escaped.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    let eager = eager_controls
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    let source_workbook = SourceBackedWorkbook::from_reader(Cursor::new(escaped.clone())).unwrap();
    let source_controls = source_workbook
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    let deferred = source_controls
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap();

    assert_eq!(eager_controls.profile(), OwnerProfile::canonical());
    assert_eq!(source_controls.profile(), OwnerProfile::canonical());
    assert_eq!(eager_controls.diagnostics(), source_controls.diagnostics());
    assert_eq!(eager, deferred);
    assert_eq!(eager, &baseline);
    assert_eq!(eager.name(), Some("Check Box 1"));
    assert_eq!(eager.anchor_profile(), Some("LoSmlAnchorV1"));
    assert_eq!(
        eager.properties().object_type(),
        baseline.properties().object_type()
    );
    assert_eq!(
        eager.properties().lock_text(),
        baseline.properties().lock_text()
    );
    assert_eq!(
        eager.properties().no_three_d(),
        baseline.properties().no_three_d()
    );

    let raw_worksheet = escaped_archive.read("xl/worksheets/sheet1.xml").unwrap();
    let raw_drawing = escaped_archive.read("xl/drawings/drawing1.xml").unwrap();
    let raw_vml = escaped_archive.read("xl/drawings/vmlDrawing1.vml").unwrap();
    let raw_properties = escaped_archive.read("xl/ctrlProps/ctrlProp1.xml").unwrap();
    assert_eq!(
        eager.properties().source_bytes(),
        Some(raw_properties.as_slice())
    );
    assert_eq!(
        deferred.properties().source_bytes(),
        Some(raw_properties.as_slice())
    );

    let read_set = source_controls.read_set().unwrap();
    assert_eq!(read_set.worksheet().bytes(), raw_worksheet.as_slice());
    assert_eq!(read_set.drawing().unwrap().bytes(), raw_drawing.as_slice());
    assert_eq!(read_set.vml().unwrap().bytes(), raw_vml.as_slice());
    assert_eq!(read_set.properties()[0].bytes(), raw_properties.as_slice());
}

#[test]
fn strict_and_activex_relationship_dialects_are_refused() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let strict = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        xml.replacen(SML_NAMESPACE, STRICT_SML_NAMESPACE, 1)
    });
    let strict_result = Workbook::from_bytes(strict)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls();
    assert!(strict_result.is_err());

    for relationship in [ACTIVEX_RELATIONSHIP, STRICT_ACTIVEX_RELATIONSHIP] {
        let mutated = rewrite_member(&source, "xl/worksheets/_rels/sheet1.xml.rels", |xml| {
            xml.replacen(CONTROL_PROPERTIES_RELATIONSHIP, relationship, 1)
        });
        let result = Workbook::from_bytes(mutated)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls();
        assert!(
            result.unwrap().is_empty(),
            "ActiveX relationship {relationship} was admitted by the form-control owner"
        );
    }

    let wrong_content_type = rewrite_member(&source, "[Content_Types].xml", |xml| {
        xml.replace(CONTROL_PROPERTIES_CONTENT_TYPE, ACTIVEX_CONTENT_TYPE)
    });
    assert!(
        Workbook::from_bytes(wrong_content_type)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );
}

#[test]
fn unrelated_activex_owner_does_not_block_form_control_dispatch() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    for relationship_type in [ACTIVEX_RELATIONSHIP, STRICT_ACTIVEX_RELATIONSHIP] {
        let source = with_extra_activex_control(&source, relationship_type, 2048, "ActiveX 1");
        let workbook = Workbook::from_bytes(source).unwrap();
        let sheet = workbook.sheet(0).unwrap().unwrap();
        let controls = sheet.form_controls().unwrap();
        assert_eq!(
            controls.len(),
            1,
            "{relationship_type} changed form dispatch"
        );
        assert_eq!(
            controls
                .get(ControlSelector::position(0))
                .unwrap()
                .unwrap()
                .shape()
                .shape_id(),
            1025
        );
    }
}

#[test]
fn active_x_selection_and_shared_persistence_edges_refuse_as_distinct_owners() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    for relationship_type in [ACTIVEX_RELATIONSHIP, STRICT_ACTIVEX_RELATIONSHIP] {
        let active_x_only = with_activex_control_replacing_form_control(&source, relationship_type);
        let eager = Workbook::from_bytes(active_x_only.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap();
        assert!(
            eager.form_controls().unwrap().is_empty(),
            "ActiveX-only control was admitted by the form-control owner"
        );
        assert_eq!(
            eager.form_control(ControlSelector::position(0)).unwrap(),
            None
        );
        let deferred = SourceBackedWorkbook::from_reader(Cursor::new(active_x_only.clone()))
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap();
        assert!(
            deferred.form_controls().unwrap().is_empty(),
            "ActiveX-only control was admitted by the source form-control owner"
        );

        let active_x = eager.active_x().unwrap();
        assert_eq!(active_x.controls.len(), 1);
        assert_eq!(active_x.controls[0].control.shape_id, 1025);
        assert_eq!(
            active_x.controls[0].descriptor.class_id,
            "form-control-test"
        );
        assert!(active_x.controls[0].binaries.is_empty());

        let shared = with_extra_activex_control(&source, relationship_type, 1025, "Check Box 1");
        assert_owner_refuses(
            shared,
            "one effective shape with both form-control and ActiveX owners",
        );
    }
}

#[test]
fn unknown_control_edges_and_activex_payload_targets_are_never_ctrlprop_fallbacks() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let unknown =
        with_extra_activex_control(&source, UNKNOWN_CONTROL_RELATIONSHIP, 2048, "Unknown 1");
    assert_owner_refuses(unknown, "an unknown worksheet persistence relationship");

    let fallback = with_ctrlprop_edge_targeting_activex_descriptor(&source);
    assert_owner_refuses(
        fallback,
        "an ActiveX descriptor reached through a ctrlProp edge",
    );
}

#[test]
fn duplicate_canonical_form_shape_identity_is_refused() {
    let source = fixture_bytes("tdf161365.xlsx");
    let duplicate_shape = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        xml.replacen("shapeId=\"1027\"", "shapeId=\"1026\"", 1)
    });
    assert_owner_refuses(
        duplicate_shape,
        "two canonical ctrlProp owners sharing one worksheet shapeId",
    );
}

#[test]
fn identity_geometry_and_object_profile_mismatches_are_refused() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let missing_drawing_identity = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        xml.replace("cNvPr id=\"1025\"", "cNvPr id=\"1026\"")
    });
    assert!(
        Workbook::from_bytes(missing_drawing_identity)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );

    let bad_object_type = rewrite_member(&source, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replace("ObjectType=\"Checkbox\"", "ObjectType=\"Button\"")
    });
    let bad_object_type_view = Workbook::from_bytes(bad_object_type)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(bad_object_type_view.shape().vml_object_type(), "Button");
    assert!(
        bad_object_type_view
            .diagnostics()
            .iter()
            .any(|diagnostic| diagnostic.code() == FormControlDiagnosticCode::UnprovenMirror)
    );

    let missing_anchor_endpoint = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        let start = xml.find("<xdr:to>").unwrap();
        let end = xml[start..].find("</xdr:to>").unwrap() + start + "</xdr:to>".len();
        format!("{}{}", &xml[..start], &xml[end..])
    });
    assert!(
        Workbook::from_bytes(missing_anchor_endpoint)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );

    let malformed_spid = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        xml.replace("spid=\"_x0000_s1025\"", "spid=\"_x0000_sbad\"")
    });
    assert!(
        Workbook::from_bytes(malformed_spid)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );

    let disagreement = rewrite_member(
        &fixture_bytes("tdf161365.xlsx"),
        "xl/drawings/vmlDrawing1.vml",
        |xml| {
            xml.replace(
                "id=\"_x0000_s1027\" type=\"#_x0000_t201\"",
                "id=\"_x0000_s1027\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t201\"",
            )
        },
    );
    assert!(
        Workbook::from_bytes(disagreement)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );
}

#[test]
fn duplicate_vml_identity_attributes_are_refused() {
    let checkbox = fixture_bytes("checkbox-form-control.xlsx");
    let tdf = fixture_bytes("tdf161365.xlsx");
    let candidates = [
        (
            "duplicate VML shape id",
            rewrite_member(&checkbox, "xl/drawings/vmlDrawing1.vml", |xml| {
                xml.replacen(
                    "<v:shape id=\"_x0000_s1025\"",
                    "<v:shape id=\"_x0000_s1025\" id=\"_x0000_s1025\"",
                    1,
                )
            }),
        ),
        (
            "duplicate VML ClientData ObjectType",
            rewrite_member(&checkbox, "xl/drawings/vmlDrawing1.vml", |xml| {
                xml.replacen(
                    "<x:ClientData ObjectType=\"Checkbox\">",
                    "<x:ClientData ObjectType=\"Checkbox\" ObjectType=\"Checkbox\">",
                    1,
                )
            }),
        ),
        (
            "duplicate VML namespace declaration",
            rewrite_member(&checkbox, "xl/drawings/vmlDrawing1.vml", |xml| {
                xml.replacen(
                    "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"",
                    "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:v=\"urn:schemas-microsoft-com:vml\"",
                    1,
                )
            }),
        ),
        (
            "duplicate VML o:spid",
            rewrite_member(&tdf, "xl/drawings/vmlDrawing1.vml", |xml| {
                xml.replacen(
                    "o:spid=\"_x0000_s1026\"",
                    "o:spid=\"_x0000_s1026\" o:spid=\"_x0000_s1026\"",
                    1,
                )
            }),
        ),
    ];
    for (label, candidate) in candidates {
        assert_owner_refuses(candidate, label);
    }
}

fn has_mirror_detail(
    collection: &litchi_xlsx::form_control::FormControlCollection,
    position: usize,
    fragment: &str,
) -> bool {
    collection
        .get(ControlSelector::position(position))
        .unwrap()
        .unwrap()
        .diagnostics()
        .iter()
        .any(|diagnostic| {
            diagnostic.code() == FormControlDiagnosticCode::UnprovenMirror
                && diagnostic.detail().contains(fragment)
        })
}

fn assert_mirror_warning_in_eager_and_source(
    source: &[u8],
    position: usize,
    field: &str,
    label: &str,
) {
    let eager = Workbook::from_bytes(source.to_vec())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert!(
        has_mirror_detail(&eager, position, field),
        "{label} did not produce an eager UnprovenMirror diagnostic for {field}"
    );

    let deferred = source_collection(source);
    assert!(
        has_mirror_detail(&deferred, position, field),
        "{label} did not produce a source-backed UnprovenMirror diagnostic for {field}"
    );
}

#[test]
fn native_mirror_pairs_are_clean_and_changed_vml_values_are_diagnosed() {
    let checked = fixture_bytes("tdf161365.xlsx");
    let checked_baseline = Workbook::from_bytes(checked.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(1))
        .unwrap()
        .unwrap();
    assert_eq!(
        known_token(checked_baseline.properties().checked()),
        Some("Checked".to_owned())
    );
    assert!(!checked_baseline.diagnostics().iter().any(|diagnostic| {
        diagnostic.code() == FormControlDiagnosticCode::UnprovenMirror
            && diagnostic.detail().contains("checked")
    }));
    let checked_changed = rewrite_member(&checked, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replace("<x:Checked>1</x:Checked>", "<x:Checked>0</x:Checked>")
    });
    let checked_changed_view = Workbook::from_bytes(checked_changed.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(1))
        .unwrap()
        .unwrap();
    assert_eq!(checked_changed_view.shape().shape_id(), 1027);
    assert_eq!(checked_changed_view.shape().vml_id(), "_x0000_s1027");
    assert_eq!(
        known_token(checked_changed_view.properties().checked()),
        Some("Checked".to_owned())
    );
    assert_mirror_warning_in_eager_and_source(&checked_changed, 1, "checked", "Checked 1 -> 0");

    let formula = fixture_bytes("tdf134769.xlsx");
    let formula_baseline = Workbook::from_bytes(formula.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(
        formula_baseline
            .properties()
            .fmla_link()
            .map(|value| value.as_str()),
        Some("#REF!")
    );
    assert!(!formula_baseline.diagnostics().iter().any(|diagnostic| {
        diagnostic.code() == FormControlDiagnosticCode::UnprovenMirror
            && diagnostic.detail().contains("fmlaLink")
    }));
    let formula_changed = rewrite_member(&formula, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replace(
            "<x:FmlaLink>#REF!</x:FmlaLink>",
            "<x:FmlaLink>#A1</x:FmlaLink>",
        )
    });
    let formula_changed_view = Workbook::from_bytes(formula_changed.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(formula_changed_view.shape().shape_id(), 1025);
    assert_eq!(formula_changed_view.shape().vml_id(), "_x0000_s1025");
    assert_eq!(
        formula_changed_view
            .properties()
            .fmla_link()
            .map(|value| value.as_str()),
        Some("#REF!")
    );
    assert_mirror_warning_in_eager_and_source(
        &formula_changed,
        0,
        "fmlaLink",
        "FmlaLink #REF! -> #A1",
    );

    let aligned = rewrite_members(
        &fixture_bytes("button-form-control.xlsx"),
        &[
            ("xl/ctrlProps/ctrlProp1.xml", &|xml: String| {
                xml.replace(
                    "objectType=\"Button\" lockText=\"1\"/>",
                    "objectType=\"Button\" lockText=\"1\" textHAlign=\"center\" textVAlign=\"center\"/>",
                )
            }),
            ("xl/drawings/vmlDrawing1.vml", &|xml: String| xml),
        ],
    );
    let aligned_baseline = Workbook::from_bytes(aligned.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert_eq!(
        aligned_baseline
            .get(ControlSelector::position(0))
            .unwrap()
            .unwrap()
            .shape()
            .shape_id(),
        1025
    );
    assert!(!has_mirror_detail(&aligned_baseline, 0, "textHAlign"));
    assert!(!has_mirror_detail(&aligned_baseline, 0, "textVAlign"));
    let aligned_changed = rewrite_member(&aligned, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replace(
            "<x:TextHAlign>Center</x:TextHAlign>",
            "<x:TextHAlign>Right</x:TextHAlign>",
        )
        .replace(
            "<x:TextVAlign>Center</x:TextVAlign>",
            "<x:TextVAlign>Top</x:TextVAlign>",
        )
    });
    let aligned_changed_eager = Workbook::from_bytes(aligned_changed.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert_eq!(
        aligned_changed_eager
            .get(ControlSelector::position(0))
            .unwrap()
            .unwrap()
            .shape()
            .shape_id(),
        1025
    );
    assert!(has_mirror_detail(&aligned_changed_eager, 0, "textHAlign"));
    assert!(has_mirror_detail(&aligned_changed_eager, 0, "textVAlign"));
    let aligned_changed_source = source_collection(&aligned_changed);
    assert!(has_mirror_detail(&aligned_changed_source, 0, "textHAlign"));
    assert!(has_mirror_detail(&aligned_changed_source, 0, "textVAlign"));

    let list_baseline = rewrite_members(
        &fixture_bytes("button-form-control.xlsx"),
        &[
            ("xl/ctrlProps/ctrlProp1.xml", &|xml: String| {
                xml.replace(
                    "objectType=\"Button\" lockText=\"1\"/>",
                    "objectType=\"Drop\"><itemLst><item val=\"Alpha\"/><item val=\"Beta\"/></itemLst></formControlPr>",
                )
            }),
            ("xl/drawings/vmlDrawing1.vml", &|xml: String| {
                xml.replace("ObjectType=\"Button\"", "ObjectType=\"Drop\"")
                    .replace(
                        "</x:ClientData>",
                        "<x:ListItem>Alpha</x:ListItem><x:ListItem>Beta</x:ListItem></x:ClientData>",
                    )
            }),
        ],
    );
    let list_baseline_eager = Workbook::from_bytes(list_baseline.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert!(has_mirror_detail(&list_baseline_eager, 0, "objectType"));
    assert!(!has_mirror_detail(&list_baseline_eager, 0, "ListItem"));
    let list_baseline_source = source_collection(&list_baseline);
    assert!(has_mirror_detail(&list_baseline_source, 0, "objectType"));
    assert!(!has_mirror_detail(&list_baseline_source, 0, "ListItem"));
    let list_changed = rewrite_member(&list_baseline, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replace(
            "<x:ListItem>Beta</x:ListItem>",
            "<x:ListItem>Gamma</x:ListItem>",
        )
    });
    let list_changed_eager = Workbook::from_bytes(list_changed.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert!(has_mirror_detail(&list_changed_eager, 0, "ListItem"));
    let list_changed_source = source_collection(&list_changed);
    assert!(has_mirror_detail(&list_changed_source, 0, "ListItem"));
}

#[test]
fn native_mirror_defaults_and_explicit_boolean_forms_keep_presence_distinct() {
    let cases = [
        ("tdf161365.xlsx", 0, "checked", false),
        ("tdf161365.xlsx", 1, "checked", false),
        ("tdf161365.xlsx", 1, "noThreeD", false),
        ("tdf161365.xlsx", 0, "lockText", false),
        ("tdf161365.xlsx", 1, "lockText", false),
        ("tdf120301_xmlSpaceParsing.xlsx", 1, "firstButton", false),
        ("tdf120301_xmlSpaceParsing.xlsx", 1, "noThreeD", false),
        ("tdf134769.xlsx", 0, "fmlaLink", false),
        ("tdf134769.xlsx", 0, "lockText", false),
        ("button-form-control.xlsx", 0, "lockText", false),
    ];
    for (file, position, field, expected) in cases {
        let source = fixture_bytes(file);
        let eager = Workbook::from_bytes(source.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .unwrap();
        assert_eq!(
            has_mirror_detail(&eager, position, field),
            expected,
            "unexpected eager {field} mirror status for {file} position {position}"
        );
        let deferred = source_collection(&source);
        assert_eq!(
            has_mirror_detail(&deferred, position, field),
            expected,
            "unexpected source {field} mirror status for {file} position {position}"
        );
    }

    let effective_mismatch = rewrite_member(
        &fixture_bytes("tdf134769.xlsx"),
        "xl/ctrlProps/ctrlProp1.xml",
        |xml| xml.replace("lockText=\"1\"", "lockText=\"0\""),
    );
    assert_mirror_warning_in_eager_and_source(
        &effective_mismatch,
        0,
        "lockText",
        "x14 lockText=false against omitted VML LockText",
    );
}

#[test]
fn invalid_boolean_and_vml_token_lexicals_are_refused_or_diagnosed() {
    let invalid_x14_boolean = rewrite_member(
        &fixture_bytes("tdf134769.xlsx"),
        "xl/ctrlProps/ctrlProp1.xml",
        |xml| xml.replace("lockText=\"1\"", "lockText=\"maybe\""),
    );
    let eager = Workbook::from_bytes(invalid_x14_boolean.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_form_control_error(eager.form_controls(), "invalid x14 lockText boolean");
    let deferred = SourceBackedWorkbook::from_reader(Cursor::new(invalid_x14_boolean))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_form_control_error(
        deferred.form_controls(),
        "source invalid x14 lockText boolean",
    );

    let invalid_vml_boolean = rewrite_member(
        &fixture_bytes("singlecontrol.xlsx"),
        "xl/drawings/vmlDrawing1.vml",
        |xml| xml.replace("<x:NoThreeD/>", "<x:NoThreeD>maybe</x:NoThreeD>"),
    );
    let eager = Workbook::from_bytes(invalid_vml_boolean.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert!(has_mirror_detail(&eager, 0, "noThreeD"));
    let deferred = source_collection(&invalid_vml_boolean);
    assert!(has_mirror_detail(&deferred, 0, "noThreeD"));

    let aligned = rewrite_members(
        &fixture_bytes("button-form-control.xlsx"),
        &[
            ("xl/ctrlProps/ctrlProp1.xml", &|xml: String| {
                xml.replace(
                    "objectType=\"Button\" lockText=\"1\"/>",
                    "objectType=\"Button\" lockText=\"1\" textVAlign=\"center\"/>",
                )
            }),
            ("xl/drawings/vmlDrawing1.vml", &|xml: String| xml),
        ],
    );
    let invalid_vml_token = rewrite_member(&aligned, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replace(
            "<x:TextVAlign>Center</x:TextVAlign>",
            "<x:TextVAlign>sideways</x:TextVAlign>",
        )
    });
    let eager = Workbook::from_bytes(invalid_vml_token.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert!(has_mirror_detail(&eager, 0, "textVAlign"));
    let deferred = source_collection(&invalid_vml_token);
    assert!(has_mirror_detail(&deferred, 0, "textVAlign"));
}

#[test]
fn unanchored_drawing_shape_with_unrelated_valid_anchor_is_refused() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let unanchored = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        let anchor_start = xml.find("<xdr:twoCellAnchor").unwrap();
        let anchor_end = xml[anchor_start..].find("</xdr:twoCellAnchor>").unwrap()
            + anchor_start
            + "</xdr:twoCellAnchor>".len();
        let anchor = &xml[anchor_start..anchor_end];
        let shape_start = anchor.find("<xdr:sp ").unwrap();
        let shape_end =
            anchor[shape_start..].find("</xdr:sp>").unwrap() + shape_start + "</xdr:sp>".len();
        let control_shape = &anchor[shape_start..shape_end];
        let unrelated_shape = control_shape
            .replace("id=\"1025\"", "id=\"2048\"")
            .replace("name=\"Check Box 1\"", "name=\"Unrelated shape\"")
            .replace("spid=\"_x0000_s1025\"", "spid=\"_x0000_s2048\"");
        let unrelated_anchor = anchor.replacen(control_shape, &unrelated_shape, 1);
        let replacement = format!("{unrelated_anchor}{control_shape}");
        format!(
            "{}{}{}",
            &xml[..anchor_start],
            replacement,
            &xml[anchor_end..]
        )
    });

    let eager = Workbook::from_bytes(unanchored.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls();
    assert!(
        eager.is_err(),
        "unanchored DrawingML control shape was admitted"
    );

    let source_result = SourceBackedWorkbook::from_reader(Cursor::new(unanchored))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls();
    assert!(
        source_result.is_err(),
        "source owner admitted an unanchored DrawingML control shape"
    );
}

#[test]
fn worksheet_and_drawing_control_ancestry_requires_direct_admitted_parents() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    assert_eq!(source_collection(&source).len(), 1);

    let worksheet_foreign_wrapper = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        let start = xml.find("<control shapeId=").unwrap();
        let end = xml[start..].find("</control>").unwrap() + start + "</control>".len();
        let control = &xml[start..end];
        format!(
            "{}<f:wrapper xmlns:f=\"urn:test:foreign\">{control}</f:wrapper>{}",
            &xml[..start],
            &xml[end..]
        )
    });
    assert_owner_refuses(
        worksheet_foreign_wrapper,
        "worksheet control under a foreign wrapper",
    );

    let worksheet_nested_controls = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        let start = xml.find("<control shapeId=").unwrap();
        let end = xml[start..].find("</control>").unwrap() + start + "</control>".len();
        let control = &xml[start..end];
        format!(
            "{}<controls>{control}</controls>{}",
            &xml[..start],
            &xml[end..]
        )
    });
    assert_owner_refuses(
        worksheet_nested_controls,
        "worksheet control under nested controls",
    );

    let drawing_foreign_wrapper = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        let start = xml.find("<xdr:sp ").unwrap();
        let end = xml[start..].find("</xdr:sp>").unwrap() + start + "</xdr:sp>".len();
        let shape = &xml[start..end];
        format!(
            "{}<f:wrapper xmlns:f=\"urn:test:foreign\">{shape}</f:wrapper>{}",
            &xml[..start],
            &xml[end..]
        )
    });
    assert_owner_refuses(
        drawing_foreign_wrapper,
        "DrawingML xdr:sp under a foreign wrapper",
    );

    let drawing_nested_group = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        let start = xml.find("<xdr:sp ").unwrap();
        let end = xml[start..].find("</xdr:sp>").unwrap() + start + "</xdr:sp>".len();
        let shape = &xml[start..end];
        let group = format!(
            "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"9000\" name=\"Group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>{shape}</xdr:grpSp>"
        );
        format!("{}{}{}", &xml[..start], group, &xml[end..])
    });
    assert_owner_refuses(
        drawing_nested_group,
        "DrawingML xdr:sp nested inside a group",
    );
}

#[test]
fn optional_control_pr_anchor_is_read_with_a_diagnostic() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let absent = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        let start = xml.find("<controlPr").unwrap();
        let end = xml[start..].find("</controlPr>").unwrap() + start + "</controlPr>".len();
        format!("{}{}", &xml[..start], &xml[end..])
    });
    let malformed = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        let start = xml.find("<to>").unwrap();
        let end = xml[start..].find("</to>").unwrap() + start + "</to>".len();
        format!("{}{}", &xml[..start], &xml[end..])
    });

    for candidate in [absent, malformed] {
        let eager = Workbook::from_bytes(candidate.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .unwrap();
        let eager_view = eager.get(ControlSelector::position(0)).unwrap().unwrap();
        assert_eq!(eager_view.anchor_profile(), None);
        assert!(
            eager_view
                .diagnostics()
                .iter()
                .any(|diagnostic| diagnostic.code() == FormControlDiagnosticCode::ControlPrAnchor)
        );
        assert_eq!(eager_view.properties().no_three_d(), Some(true));

        let deferred = source_collection(&candidate);
        let deferred_view = deferred.get(ControlSelector::position(0)).unwrap().unwrap();
        assert_eq!(deferred_view, eager_view);
        assert_eq!(deferred_view.anchor_profile(), None);
        assert!(
            deferred_view
                .diagnostics()
                .iter()
                .any(|diagnostic| diagnostic.code() == FormControlDiagnosticCode::ControlPrAnchor)
        );
    }
}

#[test]
fn losml_anchor_coordinate_grammar_is_diagnosed_without_discarding_owner() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let candidates = [
        (
            "missing coordinate child",
            rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
                xml.replacen("<xdr:colOff>209550</xdr:colOff>", "", 1)
            }),
        ),
        (
            "invalid coordinate lexical form",
            rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
                xml.replacen("<xdr:row>3</xdr:row>", "<xdr:row>not-a-row</xdr:row>", 1)
            }),
        ),
        (
            "coordinate overflow",
            rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
                xml.replacen(
                    "<xdr:rowOff>28575</xdr:rowOff>",
                    "<xdr:rowOff>18446744073709551616</xdr:rowOff>",
                    1,
                )
            }),
        ),
    ];

    for (label, candidate) in candidates {
        let eager = Workbook::from_bytes(candidate.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .unwrap();
        let eager_view = eager.get(ControlSelector::position(0)).unwrap().unwrap();
        assert_eq!(
            eager_view.anchor_profile(),
            None,
            "{label} kept the profile"
        );
        assert!(
            eager_view
                .diagnostics()
                .iter()
                .any(|diagnostic| diagnostic.code() == FormControlDiagnosticCode::ControlPrAnchor),
            "{label} did not produce the ControlPrAnchor diagnostic"
        );

        let deferred = source_collection(&candidate);
        let deferred_view = deferred.get(ControlSelector::position(0)).unwrap().unwrap();
        assert_eq!(deferred_view, eager_view, "{label} eager/deferred mismatch");
        assert_eq!(
            deferred_view.anchor_profile(),
            None,
            "{label} kept source profile"
        );
    }
}

#[test]
fn malformed_vml_root_and_incomplete_eof_are_refused() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let malformed_root = rewrite_member(&source, "xl/drawings/vmlDrawing1.vml", |xml| {
        xml.replacen("<xml ", "<notVml ", 1)
            .replacen("</xml>", "</notVml>", 1)
    });
    let incomplete_eof = rewrite_member(&source, "xl/drawings/vmlDrawing1.vml", |mut xml| {
        assert!(xml.ends_with("</xml>"));
        xml.truncate(xml.len() - "</xml>".len());
        xml
    });

    for (label, candidate) in [
        ("malformed VML root", malformed_root),
        ("incomplete VML EOF", incomplete_eof),
    ] {
        let eager = Workbook::from_bytes(candidate.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls();
        assert!(eager.is_err(), "{label} was admitted by the eager owner");

        let deferred = SourceBackedWorkbook::from_reader(Cursor::new(candidate))
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls();
        assert!(
            deferred.is_err(),
            "{label} was admitted by the source owner"
        );
    }
}

#[test]
fn control_name_over_32_characters_is_refused_by_the_native_profile() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let long_name = "N".repeat(33);
    let candidate = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("name=\"Check Box 1\"", &format!("name=\"{long_name}\""))
    });
    let candidate = rewrite_member(&candidate, "xl/drawings/drawing1.xml", |xml| {
        xml.replace("name=\"Check Box 1\"", &format!("name=\"{long_name}\""))
    });

    assert!(
        Workbook::from_bytes(candidate.clone())
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err(),
        "eager owner admitted a control name longer than the 32-character profile"
    );
    assert!(
        SourceBackedWorkbook::from_reader(Cursor::new(candidate))
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err(),
        "source owner admitted a control name longer than the 32-character profile"
    );
}

fn unsupported_choice_with_fallback(xml: String, capability: &str) -> String {
    let marker = format!("Requires=\"{capability}\"");
    let choice_start = xml.find(&marker).unwrap();
    let open_end = xml[choice_start..].find('>').unwrap() + choice_start + 1;
    // The worksheet's outer AlternateContent has no native fallback, while
    // the drawing fixture has an empty one.  In both cases copy the selected
    // branch into a valid synthetic fallback.
    let mut replaced = xml;
    let choice_end = replaced.rfind("</mc:Choice>").unwrap();
    let branch = replaced[open_end..choice_end].to_owned();
    let has_fallback = replaced[choice_end..].contains("<mc:Fallback/>");
    replaced.replace_range(
        choice_start..choice_start + marker.len(),
        "xmlns:future=\"urn:test:future\" Requires=\"future\"",
    );
    let fallback = if capability == "a14" {
        format!(
            "<mc:Fallback xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">{branch}</mc:Fallback>"
        )
    } else {
        format!("<mc:Fallback>{branch}</mc:Fallback>")
    };
    if has_fallback {
        let fallback_start = replaced.rfind("<mc:Fallback/>").unwrap();
        replaced.replace_range(
            fallback_start..fallback_start + "<mc:Fallback/>".len(),
            &fallback,
        );
    } else {
        let alternate_end = replaced.rfind("</mc:AlternateContent>").unwrap();
        replaced.insert_str(alternate_end, &fallback);
    }
    replaced
}

#[test]
fn unsupported_mce_choice_uses_fallback_and_missing_fallback_refuses() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let worksheet_fallback = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        unsupported_choice_with_fallback(xml, "x14")
    });
    let worksheet = Workbook::from_bytes(worksheet_fallback)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_eq!(worksheet.form_controls().unwrap().len(), 1);

    let drawing_fallback = rewrite_member(&source, "xl/drawings/drawing1.xml", |xml| {
        unsupported_choice_with_fallback(xml, "a14")
    });
    let drawing = Workbook::from_bytes(drawing_fallback)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_eq!(drawing.form_controls().unwrap().len(), 1);

    let missing_fallback = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        let marker = "Requires=\"x14\"";
        xml.replace(
            marker,
            "xmlns:future=\"urn:test:future\" Requires=\"future\"",
        )
    });
    assert!(
        Workbook::from_bytes(missing_fallback)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );
}

#[test]
fn properties_relationship_member_presence_is_read_and_malformed_member_refused() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let relationship_member = "xl/ctrlProps/_rels/ctrlProp1.xml.rels";
    let drawing_relationship_member = "xl/drawings/_rels/drawing1.xml.rels";
    let vml_relationship_member = "xl/drawings/_rels/vmlDrawing1.vml.rels";
    assert!(
        !ArchiveReader::new(&source)
            .unwrap()
            .file_names()
            .any(|name| name == relationship_member)
    );
    assert!(
        !ArchiveReader::new(&source)
            .unwrap()
            .file_names()
            .any(|name| name == drawing_relationship_member)
    );
    assert!(
        !ArchiveReader::new(&source)
            .unwrap()
            .file_names()
            .any(|name| name == vml_relationship_member)
    );
    let empty_relationships = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
    let valid = add_member(&source, relationship_member, empty_relationships);
    let valid = add_member(&valid, drawing_relationship_member, empty_relationships);
    let valid = add_member(&valid, vml_relationship_member, empty_relationships);
    let source_controls = source_collection(&valid);
    let source_view = source_controls
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert!(source_view.properties_relationships_present());
    assert!(source_view.shape().drawing_relationships_present());
    assert!(source_view.shape().vml_relationships_present());
    let read_set = source_controls.read_set().unwrap();
    let property_read = &read_set.properties()[0];
    assert_eq!(
        property_read.relationship_bytes(),
        Some(empty_relationships.as_slice())
    );
    assert!(property_read.relationship_is_source_backed());
    assert!(property_read.relationships().present());
    assert_eq!(
        property_read.relationships().bytes(),
        empty_relationships.len()
    );
    assert_eq!(property_read.relationships().edges(), 0);
    assert_eq!(
        source_view.properties_relationships(),
        property_read.relationships()
    );
    let drawing_read = read_set.drawing().unwrap();
    assert_eq!(
        drawing_read.relationship_bytes(),
        Some(empty_relationships.as_slice())
    );
    assert!(drawing_read.relationship_is_source_backed());
    let vml_read = read_set.vml().unwrap();
    assert_eq!(
        vml_read.relationship_bytes(),
        Some(empty_relationships.as_slice())
    );
    assert!(vml_read.relationship_is_source_backed());
    let limited_source = SourceBackedWorkbook::from_reader(Cursor::new(valid.clone()))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert!(
        limited_source
            .form_controls_with_limits(
                OwnerLimits::new().with_max_control_properties_relationship_bytes(1)
            )
            .is_err()
    );
    assert!(
        limited_source
            .form_controls_with_limits(OwnerLimits::new().with_max_sidecar_relationship_bytes(1))
            .is_err()
    );
    assert_eq!(
        Workbook::from_bytes(valid)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .unwrap()
            .len(),
        1
    );

    let malformed = add_member(&source, relationship_member, b"<Relationships>");
    if let Ok(workbook) = Workbook::from_bytes(malformed) {
        assert!(workbook.sheet(0).unwrap().unwrap().form_controls().is_err());
    }
    let malformed = add_member(&source, drawing_relationship_member, b"<Relationships>");
    if let Ok(workbook) = SourceBackedWorkbook::from_reader(Cursor::new(malformed)) {
        assert!(workbook.sheet(0).unwrap().unwrap().form_controls().is_err());
    }
}

#[test]
fn sidecar_relationships_retain_exact_target_data_and_fingerprint() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let property_member = "xl/ctrlProps/_rels/ctrlProp1.xml.rels";
    let drawing_member = "xl/drawings/_rels/drawing1.xml.rels";
    let vml_member = "xl/drawings/_rels/vmlDrawing1.vml.rels";
    let property_relationships = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rCtrlTarget" Type="urn:litchi:form-control-target" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>"#;
    let drawing_relationships = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDrawingTarget" Type="urn:litchi:drawing-target" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>"#;
    let vml_relationships = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rVmlTarget" Type="urn:litchi:vml-target" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>"#;
    let source = add_member(&source, property_member, property_relationships);
    let source = add_member(&source, drawing_member, drawing_relationships);
    let source = add_member(&source, vml_member, vml_relationships);
    let properties_xml = ArchiveReader::new(&source)
        .unwrap()
        .read("xl/ctrlProps/ctrlProp1.xml")
        .unwrap();

    let source_controls = source_collection(&source);
    let source_read_set = source_controls.read_set().unwrap();
    assert_eq!(source_read_set.sidecar_targets().len(), 1);
    let sidecar_target = &source_read_set.sidecar_targets()[0];
    assert_eq!(sidecar_target.bytes(), properties_xml.as_slice());
    assert!(sidecar_target.is_source_backed());
    assert_eq!(sidecar_target.range(), 0..properties_xml.len());
    assert_eq!(
        sidecar_target.relationship_bytes(),
        Some(property_relationships.as_slice())
    );
    assert!(sidecar_target.relationship_is_source_backed());
    assert_eq!(sidecar_target.relationships().edges(), 1);
    assert_eq!(
        sidecar_target.relationships().bytes(),
        property_relationships.len()
    );
    let property_read = &source_read_set.properties()[0];
    assert_eq!(
        property_read.relationship_bytes(),
        Some(property_relationships.as_slice())
    );
    assert!(property_read.relationship_is_source_backed());
    assert_eq!(property_read.relationships().edges(), 1);
    assert_eq!(
        property_read.relationships().bytes(),
        property_relationships.len()
    );
    assert_ne!(property_read.relationships().digest(), 0);
    let drawing_read = source_read_set.drawing().unwrap();
    assert_eq!(
        drawing_read.relationship_bytes(),
        Some(drawing_relationships.as_slice())
    );
    assert_eq!(drawing_read.relationships().edges(), 1);
    assert_ne!(drawing_read.relationships().digest(), 0);
    let vml_read = source_read_set.vml().unwrap();
    assert_eq!(
        vml_read.relationship_bytes(),
        Some(vml_relationships.as_slice())
    );
    assert_eq!(vml_read.relationships().edges(), 1);
    assert_ne!(vml_read.relationships().digest(), 0);

    let source_view = source_controls
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(
        source_view.properties_relationships(),
        property_read.relationships()
    );
    assert_eq!(
        source_view.shape().drawing_relationships(),
        drawing_read.relationships()
    );
    assert_eq!(
        source_view.shape().vml_relationships(),
        vml_read.relationships()
    );

    let eager_view = Workbook::from_bytes(source)
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(eager_view.properties_relationships().edges(), 1);
    assert_eq!(eager_view.shape().drawing_relationships().edges(), 1);
    assert_eq!(eager_view.shape().vml_relationships().edges(), 1);
}

#[test]
fn unreferenced_properties_are_reported_without_becoming_selectable_controls() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let source = rewrite_member(&source, "[Content_Types].xml", |xml| {
        xml.replace(
            "</Types>",
            &format!(
                "<Override PartName=\"/xl/ctrlProps/ctrlProp-unused.xml\" ContentType=\"{CONTROL_PROPERTIES_CONTENT_TYPE}\"/></Types>"
            ),
        )
    });
    let source = add_member(
        &source,
        "xl/ctrlProps/ctrlProp-unused.xml",
        format!("<formControlPr xmlns=\"{FORM_CONTROL_NAMESPACE}\" objectType=\"Button\"/>")
            .as_bytes(),
    );
    let controls = source_collection(&source);
    assert_eq!(controls.len(), 1);
    assert!(
        controls
            .get(ControlSelector::position(1))
            .unwrap()
            .is_none()
    );
    assert!(controls.diagnostics().iter().any(|diagnostic| {
        diagnostic.code() == FormControlDiagnosticCode::UnreferencedControlProperties
    }));
}

#[test]
fn unreferenced_properties_are_deterministic_and_bounded_by_mirror_cap() {
    let base = fixture_bytes("checkbox-form-control.xlsx");
    let build = |order: &[&str]| {
        let source = rewrite_member(&base, "[Content_Types].xml", |xml| {
            order.iter().fold(xml, |xml, name| {
                xml.replace(
                    "</Types>",
                    &format!(
                        "<Override PartName=\"/xl/ctrlProps/{name}\" ContentType=\"{CONTROL_PROPERTIES_CONTENT_TYPE}\"/></Types>"
                    ),
                )
            })
        });
        order.iter().fold(source, |source, name| {
            add_member(
                &source,
                &format!("xl/ctrlProps/{name}"),
                format!(
                    "<formControlPr xmlns=\"{FORM_CONTROL_NAMESPACE}\" objectType=\"Button\"/>"
                )
                .as_bytes(),
            )
        })
    };
    let archive_a = build(&["ctrlProp-z.xml", "ctrlProp-a.xml"]);
    let archive_b = build(&["ctrlProp-a.xml", "ctrlProp-z.xml"]);

    let eager_a = Workbook::from_bytes(archive_a.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    let eager_b = Workbook::from_bytes(archive_b.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    let eager_unreferenced = eager_a
        .diagnostics()
        .iter()
        .filter(|diagnostic| {
            diagnostic.code() == FormControlDiagnosticCode::UnreferencedControlProperties
        })
        .count();
    assert_eq!(eager_unreferenced, 2);
    assert_eq!(eager_a.diagnostics(), eager_b.diagnostics());

    let source_a = source_collection(&archive_a);
    let source_b = source_collection(&archive_b);
    let source_unreferenced = source_a
        .diagnostics()
        .iter()
        .filter(|diagnostic| {
            diagnostic.code() == FormControlDiagnosticCode::UnreferencedControlProperties
        })
        .count();
    assert_eq!(source_unreferenced, 2);
    assert_eq!(source_a.diagnostics(), source_b.diagnostics());

    let capped = OwnerLimits::new().with_max_mirror_nodes(1);
    let eager_sheet = Workbook::from_bytes(archive_a.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_resource_limit(
        eager_sheet.form_controls_with_limits(capped),
        "eager diagnostic mirror-node cap",
    );
    let source_sheet = SourceBackedWorkbook::from_reader(Cursor::new(archive_a))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_resource_limit(
        source_sheet.form_controls_with_limits(capped),
        "source diagnostic mirror-node cap",
    );
}

#[test]
fn unreferenced_diagnostic_detail_is_charged_by_projection_limit() {
    let base = fixture_bytes("checkbox-form-control.xlsx");
    let suffix = format!("{}{}", "u".repeat(46), ".xml");
    let partname = format!("/xl/ctrlProps/{suffix}");
    let member = partname.strip_prefix('/').unwrap().to_owned();
    assert_eq!(partname.len(), 64);
    assert_eq!(
        format!("unreferenced control-properties part was retained as opaque source: {partname}")
            .len(),
        132
    );
    let source = rewrite_member(&base, "[Content_Types].xml", |xml| {
        xml.replace(
            "</Types>",
            &format!(
                "<Override PartName=\"{partname}\" ContentType=\"{CONTROL_PROPERTIES_CONTENT_TYPE}\"/></Types>"
            ),
        )
    });
    let source = add_member(
        &source,
        &member,
        format!("<formControlPr xmlns=\"{FORM_CONTROL_NAMESPACE}\" objectType=\"Button\"/>")
            .as_bytes(),
    );

    // Measure the same package with and without the unreferenced diagnostic
    // instead of pinning the old 3334-byte boundary.  The exact floor may
    // move when the owner ledger changes its allocation strategy, but the
    // additional retained detail and its diagnostic metadata must still be
    // charged in both public read paths.
    let eager_base = minimum_projection_cap(
        |cap| projection_cap_admits_eager(&base, cap, "eager baseline projection"),
        "eager baseline projection",
    );
    let eager_with_diagnostic = minimum_projection_cap(
        |cap| projection_cap_admits_eager(&source, cap, "eager unreferenced projection"),
        "eager unreferenced projection",
    );
    let source_base = minimum_projection_cap(
        |cap| projection_cap_admits_source(&base, cap, "source baseline projection"),
        "source baseline projection",
    );
    let source_with_diagnostic = minimum_projection_cap(
        |cap| projection_cap_admits_source(&source, cap, "source unreferenced projection"),
        "source unreferenced projection",
    );
    assert!(eager_with_diagnostic > eager_base);
    assert!(
        eager_with_diagnostic - eager_base >= 132,
        "unreferenced diagnostic delta must include its 132-byte detail: baseline={eager_base}, with={eager_with_diagnostic}"
    );
    assert!(source_with_diagnostic > source_base);
    assert!(
        source_with_diagnostic - source_base >= 132,
        "unreferenced diagnostic delta must include its 132-byte detail: baseline={source_base}, with={source_with_diagnostic}"
    );
}

#[test]
fn mirror_diagnostic_projection_budget_counts_every_detail() {
    let source = rewrite_member(
        &fixture_bytes("checkbox-form-control.xlsx"),
        "xl/drawings/vmlDrawing1.vml",
        |xml| {
            let fields = [
                ("Colored", "0"),
                ("DropLines", "0"),
                ("DropStyle", "Simple"),
                ("Dx", "0"),
                ("FirstButton", "0"),
                ("FmlaGroup", "0"),
                ("FmlaLink", "0"),
                ("FmlaRange", "0"),
                ("FmlaTxbx", "0"),
                ("Horiz", "False"),
                ("Inc", "0"),
                ("JustLastX", "0"),
                ("LockText", "False"),
                ("Max", "0"),
                ("Min", "0"),
                ("MultiSel", "0"),
                ("NoThreeD", ""),
                ("NoThreeD2", ""),
                ("Page", "0"),
                ("Sel", "0"),
                ("SelType", "Single"),
                ("TextHAlign", "Left"),
                ("TextVAlign", "Top"),
                ("Val", "0"),
                ("WidthMin", "0"),
                ("VTEdit", "0"),
                ("MultiLine", "0"),
                ("VScroll", "False"),
                ("SecretEdit", "0"),
                ("ListItem", "0"),
            ];
            let mut extra = String::new();
            for (name, value) in fields {
                if value.is_empty() {
                    extra.push_str(&format!("<x:{name}/><x:{name}/>"));
                } else {
                    extra.push_str(&format!(
                        "<x:{name}>{value}</x:{name}><x:{name}>{value}</x:{name}>"
                    ));
                }
            }
            xml.replacen("</x:ClientData>", &format!("{extra}</x:ClientData>"), 1)
        },
    );
    let eager_workbook = Workbook::from_bytes(source.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    let source_workbook = SourceBackedWorkbook::from_reader(Cursor::new(source.clone())).unwrap();
    let source_sheet = source_workbook.sheet(0).unwrap().unwrap();

    // The worksheet relationship namespace URI is 67 bytes in this retained
    // fixture, so leave one byte for that canonical identity while keeping
    // the projection ceiling at its normal generous default for the exact
    // diagnostic inventory below.
    let accepted = OwnerLimits::new().with_max_name_bytes(68);
    let eager = eager_workbook.form_controls_with_limits(accepted).unwrap();
    let source_controls = source_sheet.form_controls_with_limits(accepted).unwrap();
    let eager_view = eager.get(ControlSelector::position(0)).unwrap().unwrap();
    let source_view = source_controls
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    let detail_bytes = eager_view
        .diagnostics()
        .iter()
        .map(|diagnostic| diagnostic.detail().len())
        .sum::<usize>();
    assert_eq!(eager_view.diagnostics().len(), 30);
    assert_eq!(detail_bytes, 1567);
    assert_eq!(eager_view, source_view);
    assert_eq!(eager_view.diagnostics(), source_view.diagnostics());
    assert!(source_controls.read_set().is_some());

    // 3643 was the old undercounting boundary: it admitted this 30-diagnostic
    // fixture even though the retained detail bytes alone total 1567.  The
    // owner must charge each detail and refuse the same bound for both
    // worksheet facades.
    let eager_workbook = Workbook::from_bytes(source.clone()).unwrap();
    let eager_sheet = eager_workbook.sheet(0).unwrap().unwrap();
    assert_resource_limit(
        eager_sheet.form_controls_with_limits(
            OwnerLimits::new()
                .with_max_name_bytes(68)
                .with_max_projection_bytes(3643),
        ),
        "eager mirror diagnostic projection at historical boundary",
    );
    let source_workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).unwrap();
    let source_sheet = source_workbook.sheet(0).unwrap().unwrap();
    assert_resource_limit(
        source_sheet.form_controls_with_limits(
            OwnerLimits::new()
                .with_max_name_bytes(68)
                .with_max_projection_bytes(3643),
        ),
        "source mirror diagnostic projection at historical boundary",
    );
}

#[test]
fn source_read_set_retains_exact_owner_relationship_and_mce_provenance() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let archive = ArchiveReader::new(&source).unwrap();
    let worksheet_xml = archive.read("xl/worksheets/sheet1.xml").unwrap();
    let worksheet_rels = archive.read("xl/worksheets/_rels/sheet1.xml.rels").unwrap();
    let drawing_xml = archive.read("xl/drawings/drawing1.xml").unwrap();
    let vml_xml = archive.read("xl/drawings/vmlDrawing1.vml").unwrap();
    let properties_xml = archive.read("xl/ctrlProps/ctrlProp1.xml").unwrap();

    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).unwrap();
    let collection = workbook.sheet(0).unwrap().unwrap().form_controls().unwrap();
    let read_set = collection.read_set().unwrap();
    assert_eq!(collection.source_version(), read_set.source_version());
    assert!(read_set.source_version().is_some());

    let worksheet = read_set.worksheet();
    assert_eq!(worksheet.bytes(), worksheet_xml.as_slice());
    assert!(worksheet.is_source_backed());
    assert_eq!(worksheet.range(), 0..worksheet_xml.len());
    assert_eq!(
        worksheet.relationship_bytes(),
        Some(worksheet_rels.as_slice())
    );
    assert!(worksheet.relationship_is_source_backed());
    assert_eq!(
        worksheet.relationship_range(),
        Some(0..worksheet_rels.len())
    );
    assert!(worksheet.relationships().present());
    assert_eq!(worksheet.relationships().bytes(), worksheet_rels.len());
    assert!(worksheet.relationships().edges() > 0);

    let drawing = read_set.drawing().unwrap();
    assert_eq!(drawing.bytes(), drawing_xml.as_slice());
    assert!(drawing.is_source_backed());
    assert!(!drawing.relationships().present());
    assert_eq!(drawing.relationship_bytes(), None);
    let vml = read_set.vml().unwrap();
    assert_eq!(vml.bytes(), vml_xml.as_slice());
    assert!(vml.is_source_backed());
    assert!(!vml.relationships().present());
    assert_eq!(vml.relationship_bytes(), None);

    assert_eq!(read_set.properties().len(), 1);
    let properties = &read_set.properties()[0];
    assert_eq!(properties.bytes(), properties_xml.as_slice());
    assert!(properties.is_source_backed());
    assert!(!properties.relationships().present());
    assert_eq!(properties.relationship_bytes(), None);

    let worksheet_mce = read_set.worksheet_mce();
    assert_eq!(worksheet_mce.raw_bytes(), worksheet_xml.as_slice());
    assert!(!worksheet_mce.selected_bytes().is_empty());
    assert!(worksheet_mce.selected_choices() >= 2);
    assert!(!worksheet_mce.selected_ranges().is_empty());
    assert_eq!(
        worksheet_xml
            .windows(b"xmlns:x14=\"".len())
            .filter(|window| *window == b"xmlns:x14=\"")
            .count(),
        1,
        "the native worksheet uses one baseline x14 namespace declaration"
    );
    assert!(
        worksheet_mce.selected_ranges().iter().all(|range| {
            worksheet_xml[range.clone()]
                .windows(b"Requires=\"x14\"".len())
                .any(|window| window == b"Requires=\"x14\"")
        }),
        "selected worksheet MCE ranges lost their baseline Requires branch"
    );
    assert!(
        worksheet_mce
            .selected_bytes()
            .windows(b"<control ".len())
            .any(|window| window == b"<control ")
    );
    let drawing_mce = read_set.drawing_mce();
    assert_eq!(drawing_mce.raw_bytes(), drawing_xml.as_slice());
    assert!(!drawing_mce.selected_bytes().is_empty());
    assert_eq!(drawing_mce.selected_choices(), 1);
    assert!(!drawing_mce.selected_ranges().is_empty());

    let view = collection
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert!(!view.properties_relationships().present());
    assert!(!view.shape().drawing_relationships().present());
    assert!(!view.shape().vml_relationships().present());
}

#[test]
fn shared_target_and_external_edges_are_refused_before_projection() {
    let source = fixture_bytes("tdf161365.xlsx");
    let shared = rewrite_member(&source, "xl/worksheets/_rels/sheet1.xml.rels", |xml| {
        xml.replace(
            "Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp\" Target=\"../ctrlProps/ctrlProp2.xml\"",
            "Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp\" Target=\"../ctrlProps/ctrlProp1.xml\"",
        )
    });
    assert!(
        Workbook::from_bytes(shared)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );

    let external = rewrite_member(&source, "xl/worksheets/_rels/sheet1.xml.rels", |xml| {
        xml.replace(
            "Target=\"../ctrlProps/ctrlProp1.xml\"",
            "Target=\"https://example.invalid/ctrlProp.xml\" TargetMode=\"External\"",
        )
    });
    assert!(
        Workbook::from_bytes(external)
            .unwrap()
            .sheet(0)
            .unwrap()
            .unwrap()
            .form_controls()
            .is_err()
    );
}

#[test]
fn caller_part_limit_refuses_before_form_control_source_projection() {
    let source = fixture_bytes("tdf134769.xlsx");
    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(SourceBackedWorkbook::from_reader_with_limits(Cursor::new(source), limits).is_err());
}

#[test]
fn caller_owner_limits_are_applied_through_eager_and_source_facades() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let eager_workbook = Workbook::from_bytes(source.clone()).unwrap();
    let eager_sheet = eager_workbook.sheet(0).unwrap().unwrap();
    let source_workbook = SourceBackedWorkbook::from_reader(Cursor::new(source)).unwrap();
    let source_sheet = source_workbook.sheet(0).unwrap().unwrap();
    let cases = [
        ("controls", OwnerLimits::new().with_max_controls(0)),
        ("MCE branches", OwnerLimits::new().with_max_mce_branches(0)),
        ("MCE depth", OwnerLimits::new().with_max_mce_depth(0)),
        ("MCE bytes", OwnerLimits::new().with_max_mce_bytes(1)),
        (
            "relationship edges",
            OwnerLimits::new().with_max_relationship_edges(1),
        ),
        (
            "DrawingML bytes",
            OwnerLimits::new().with_max_drawing_bytes(1),
        ),
        ("VML bytes", OwnerLimits::new().with_max_vml_bytes(1)),
        ("shape identities", OwnerLimits::new().with_max_shapes(0)),
        ("mirror nodes", OwnerLimits::new().with_max_mirror_nodes(0)),
        (
            "identity name bytes",
            OwnerLimits::new().with_max_name_bytes(1),
        ),
    ];
    for (label, limits) in cases {
        assert_resource_limit(
            eager_sheet.form_controls_with_limits(limits),
            &format!("eager {label} limit"),
        );
        assert_resource_limit(
            source_sheet.form_controls_with_limits(limits),
            &format!("source {label} limit"),
        );
    }
}

#[test]
fn caller_name_limit_is_typed_before_control_projection() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let limits = OwnerLimits::new().with_max_name_bytes(1);
    let eager_sheet = Workbook::from_bytes(source.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_resource_limit(
        eager_sheet.form_controls_with_limits(limits),
        "eager authored-name limit",
    );
    let source_sheet = SourceBackedWorkbook::from_reader(Cursor::new(source))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    assert_resource_limit(
        source_sheet.form_controls_with_limits(limits),
        "source authored-name limit",
    );
}

#[test]
fn worksheet_control_identity_name_limit_reports_object_resource() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    // The canonical worksheet carries an XDR namespace QName whose bounded
    // attribute identity is 69 bytes.  Keep this independent name-limit
    // probe above that valid 69/64 namespace refusal so it reaches the
    // worksheet control identity itself.
    let long_name = "n".repeat(71);
    let source = rewrite_member(&source, "xl/worksheets/sheet1.xml", |xml| {
        xml.replacen("name=\"Check Box 1\"", &format!("name=\"{long_name}\""), 1)
    });
    let limits = OwnerLimits::new().with_max_name_bytes(70);
    let eager_sheet = Workbook::from_bytes(source.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    let source_sheet = SourceBackedWorkbook::from_reader(Cursor::new(source))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    for (label, result) in [
        (
            "eager worksheet control identity name",
            eager_sheet.form_controls_with_limits(limits),
        ),
        (
            "source worksheet control identity name",
            source_sheet.form_controls_with_limits(limits),
        ),
    ] {
        match result {
            Err(Error::ResourceLimit(limit)) => {
                assert_eq!(limit.resource, Resource::Objects, "{label}: {limit:?}");
                assert_eq!(limit.observed, 71, "{label}: {limit:?}");
                assert_eq!(limit.limit, 70, "{label}: {limit:?}");
                assert!(
                    limit.scope.contains("worksheet control identity"),
                    "{label}: {limit:?}"
                );
            },
            Err(error) => panic!("{label} returned the wrong public error: {error:?}"),
            Ok(_) => panic!("{label} admitted a 71-byte identity under a 70-byte limit"),
        }
    }
}

#[test]
fn caller_leaf_depth_event_and_opaque_caps_are_forwarded() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let eager_sheet = Workbook::from_bytes(source.clone())
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    let source_sheet = SourceBackedWorkbook::from_reader(Cursor::new(source))
        .unwrap()
        .sheet(0)
        .unwrap()
        .unwrap();
    let cases = [
        ("leaf depth", OwnerLimits::new().with_max_mce_depth(1)),
        ("leaf events", OwnerLimits::new().with_max_mce_events(1)),
        (
            "leaf opaque bytes",
            OwnerLimits::new().with_max_mce_bytes(1),
        ),
    ];
    for (label, limits) in cases {
        assert_typed_limit(
            eager_sheet.form_controls_with_limits(limits),
            &format!("eager lowered {label} cap"),
        );
        assert_typed_limit(
            source_sheet.form_controls_with_limits(limits),
            &format!("source lowered {label} cap"),
        );
    }
}

#[test]
fn managed_owner_clones_and_read_set_keep_budget_charges_alive() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let (budget, _cancellation_source, context) = managed_owner_context(u64::MAX, u64::MAX);
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let sheet = workbook.sheet(0).unwrap().unwrap();
    let package_baseline = managed_owner_usage(&budget);
    let collection = sheet.form_controls().unwrap();
    let before_clones = managed_owner_usage(&budget);
    let read_set = collection.read_set().unwrap().clone();
    let view = collection
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap()
        .clone();
    assert!(view.has_execution_budget());
    let cloned_collection = collection.clone();
    let cloned_view = view.clone();
    assert_eq!(managed_owner_usage(&budget), before_clones);
    let before_workbook_drop = managed_owner_usage(&budget);
    assert!(
        before_workbook_drop
            .iter()
            .zip(package_baseline)
            .any(|(usage, baseline)| usage > &baseline),
        "managed owner did not retain any charged resource"
    );
    drop(sheet);
    drop(collection);
    drop(cloned_collection);
    drop(view);
    drop(cloned_view);
    assert_eq!(
        managed_owner_usage(&budget),
        before_workbook_drop,
        "the cloned read set did not retain owner charges after collection/view drops"
    );
    drop(read_set);
    let after_read_set_drop = managed_owner_usage(&budget);
    let package_cache = workbook.cache_diagnostics();
    assert_eq!(
        after_read_set_drop[0], package_cache.budget_cache_reserved_bytes,
        "owner memory delta did not release to the package cache baseline"
    );
    assert_eq!(
        after_read_set_drop[1], package_cache.budget_input_bytes_used,
        "input-byte accounting diverged from the package's cumulative baseline"
    );
    assert_eq!(after_read_set_drop[2], package_baseline[2]);
    assert_eq!(
        after_read_set_drop[3], package_cache.budget_objects_used,
        "owner object delta did not release to the package cache/catalog baseline"
    );
    assert_eq!(after_read_set_drop[4], package_baseline[4]);
    let package_input_baseline = package_cache.budget_input_bytes_used;
    drop(workbook);
    let after_workbook_drop = managed_owner_usage(&budget);
    assert_eq!(
        after_workbook_drop,
        [0, package_input_baseline, 0, 0, 0],
        "only cumulative package input-byte usage may outlive the package"
    );
}

#[test]
fn managed_owner_tiny_execution_depth_is_a_typed_resource_refusal() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    let (budget, _cancellation_source, context) = managed_owner_context(u64::MAX, 0);
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let sheet = workbook.sheet(0).unwrap().unwrap();
    assert_execution_depth_limit(sheet.form_controls(), "managed owner depth budget");
    drop(sheet);
    drop(workbook);
    assert_eq!(budget.used(Resource::Depth), 0);
}

#[test]
fn mirror_node_caps_refuse_without_retaining_control_shape_charges() {
    let source = fixture_bytes("checkbox-form-control.xlsx");
    for maximum in [0, 1] {
        let (budget, _cancellation_source, context) = managed_owner_context(u64::MAX, u64::MAX);
        let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(source.clone())),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let sheet = workbook.sheet(0).unwrap().unwrap();
        let before_owner = managed_owner_usage(&budget);
        let result =
            sheet.form_controls_with_limits(OwnerLimits::new().with_max_mirror_nodes(maximum));
        match result {
            Err(Error::ResourceLimit(limit)) => {
                assert_eq!(limit.resource, Resource::Objects);
                assert!(
                    limit.scope.contains("form-control owner"),
                    "mirror-node refusal came from an unexpected scope: {limit:?}"
                );
            },
            Err(error) => {
                panic!("mirror-node cap {maximum} returned the wrong public error: {error:?}")
            },
            Ok(_) => panic!("mirror-node cap {maximum} was admitted"),
        }
        // The public budget reports retained charges, not allocator peaks.  A
        // refused scan must leave no owner control/shape charge behind.
        assert_eq!(
            managed_owner_usage(&budget),
            before_owner,
            "mirror-node cap {maximum} retained owner charges after refusal"
        );
        drop(sheet);
        drop(workbook);
        assert_eq!(
            managed_owner_usage(&budget),
            [0, before_owner[1], 0, 0, 0],
            "only cumulative package input-byte usage may outlive the package"
        );
    }
}

#[test]
fn source_reads_are_read_only_and_repeated_semantic_selection_is_stable() {
    let source = fixture_bytes("tdf134769.xlsx");
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(source.clone())).unwrap();
    let sheet = workbook.sheet(0).unwrap().unwrap();
    let first = sheet
        .form_control(ControlSelector::name("Check Box 1"))
        .unwrap()
        .unwrap();
    let second = sheet
        .form_control(ControlSelector::position(0))
        .unwrap()
        .unwrap();
    assert_eq!(first, second);
    assert_eq!(first.properties().fmla_link().unwrap().as_str(), "#REF!");
    let archive = ArchiveReader::new(&source).unwrap();
    let raw = archive.read("xl/ctrlProps/ctrlProp1.xml").unwrap();
    assert_eq!(first.properties().source_bytes(), Some(raw.as_slice()));
}

#[test]
fn source_owner_final_fence_rejects_revision_flip_after_scan() {
    let source_bytes = fixture_bytes("tdf134769.xlsx");
    let baseline_source = Arc::new(VersionFlipSource::new(source_bytes.clone(), None));
    let baseline = SourceBackedWorkbook::from_read_at(baseline_source.clone()).unwrap();
    baseline.sheet(0).unwrap().unwrap().form_controls().unwrap();
    let final_check_call = baseline_source.version_calls.load(Ordering::SeqCst);
    drop(baseline);

    let source = Arc::new(VersionFlipSource::new(source_bytes, Some(final_check_call)));
    let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
    let sheet = workbook.sheet(0).unwrap().unwrap();
    assert!(matches!(
        sheet.form_controls(),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
}
