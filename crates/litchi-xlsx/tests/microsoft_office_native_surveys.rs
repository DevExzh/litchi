//! Explicit inert Survey parts attached to native Excel tables; no rendering or submission.
use litchi_opc::{PackURI, SourceLimits, phys_pkg::OwnedPhysPkgReader};
use litchi_xlsx::{
    Error, Package,
    survey::{
        self, Binding, ElementProperties, Guid, Id, Position, Question, QuestionType, Survey,
    },
};
use std::{collections::BTreeMap, path::Path};
const TABLE: &str = "/xl/tables/table1.xml";
const TABLE_RELS: &str = "xl/tables/_rels/table1.xml.rels";
const PARENTS: [(&str, &str, &str, &str); 4] = [
    (
        "excel-values-formulas-chart.xlsx",
        "Table1",
        "Item",
        "Quantity",
    ),
    ("excel-visual-analytics.xltx", "Table1", "Date", "Region"),
    (
        "excel-visual-analytics-strict.xlsx",
        "Table1",
        "Date",
        "Region",
    ),
    (
        "excel-connections.xlsx",
        "native_query_010",
        "ID",
        "Category",
    ),
];
const GUID: &str = "{01234567-89AB-CDEF-0123-456789ABCDEF}";
const RAW: &str = r#"<s:survey xmlns:s='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main' xmlns='urn:native-survey:foreign-default' xmlns:k='urn:native-survey:kept' id='71' guid='{01234567-89AB-CDEF-0123-456789ABCDEF}' title='Before _x0041_' k:opaque='root'>
<!-- root comment -->
<s:surveyPr cssClass='card' top='-2' width='320' position='relative'><s:extLst><x:ext xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' uri='urn:property'><k:property keep='yes'/></x:ext></s:extLst></s:surveyPr>
<s:questions k:opaque='questions'>
<!-- before first question -->
<s:question binding='1' text='Name _x005F_x0041_' type='singleLineOfText' required='1' k:opaque='question'><s:questionPr left='7'/><!-- question comment --><s:extLst><x:ext xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' uri='urn:question'><opaque keep='foreign default'/></x:ext></s:extLst></s:question>
<!-- between questions -->
<s:question binding='2' text='Choice' type='choice' defaultValue='_x0000_' rowSource='one;&quot;two;three&quot;;'/>
<!-- after last question -->
</s:questions>
<s:extLst><x:ext xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main' uri='urn:survey'><k:payload target='https://example.invalid/never-submitted'/></x:ext></s:extLst>
</s:survey>"#;
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = OwnedPhysPkgReader::from_bytes(p.to_bytes().unwrap(), SourceLimits::default()).unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
type Graph = BTreeMap<String, (String, Vec<(String, String, String, bool)>)>;
fn graph(p: &Package) -> Graph {
    let p = p.clone().into_plain_opc();
    let edges = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .map(|r| {
                (
                    r.r_id().into(),
                    r.reltype().into(),
                    r.target_ref().into(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    let mut g = BTreeMap::from([("/".into(), (String::new(), edges(p.rels())))]);
    for p in p.iter_parts() {
        g.insert(
            p.partname().to_string(),
            (p.content_type().into(), edges(p.rels())),
        );
    }
    g
}
fn preserve(a: &Package, b: &Package, owned: &[&str]) {
    let a = members(a);
    let b = members(b);
    for n in a.keys().chain(b.keys()) {
        if !owned.contains(&n.as_str()) {
            assert_eq!(a.get(n), b.get(n), "changed {n}");
        }
    }
}
fn same_members(a: &Package, b: &Package) {
    let a = members(a);
    let b = members(b);
    let changed = a
        .keys()
        .chain(b.keys())
        .filter(|name| a.get(*name) != b.get(*name))
        .collect::<std::collections::BTreeSet<_>>();
    assert!(changed.is_empty(), "different package members: {changed:?}");
}
fn publish(p: &Package, c: &survey::Commit, owned: &[&str]) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    c.package().save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    preserve(p, &q, owned);
    same_members(&q, c.package());
    same_members(p.apply_surveys(c.patch()).unwrap().package(), &q);
    let inverse = q.apply_surveys(&c.patch().inverse()).unwrap();
    same_members(inverse.package(), p);
    assert_eq!(graph(inverse.package()), graph(p));
    let catalog = q.surveys().unwrap();
    let view = q.workbook_ref().surveys().unwrap();
    assert_eq!(view.len(), catalog.len());
    for (a, b) in view.iter().zip(catalog.entries()) {
        assert_eq!(a.survey(), b.survey());
        assert_eq!(a.part_name(), b.part_name());
    }
    q
}
fn authored(first: &str, second: &str) -> Survey {
    let mut a = Question::for_column(first).unwrap();
    a.set_text(Some("Name 🍎".into()));
    a.set_question_type(Some(QuestionType::SingleLineOfText));
    a.set_required(true);
    let mut b = Question::for_column(second).unwrap();
    b.set_text(Some("Choose".into()));
    b.set_question_type(Some(QuestionType::Choice));
    b.set_row_source(Some("one;\"two;three\";".into())).unwrap();
    let mut s = Survey::new(Id::new(71), Guid::new(GUID).unwrap(), vec![a, b]).unwrap();
    s.set_title(Some("Native table survey".into()));
    s.set_description(Some("Inert metadata".into()));
    let mut props = ElementProperties::new();
    props.set_top(Some(-2));
    props.set_width(Some(320));
    props.set_position(Some(Position::Relative));
    s.set_properties(Some(props));
    s
}
fn attach(file: &str, table: &str, first: &str, second: &str) -> Package {
    let bytes = native(file);
    let p = Package::from_bytes(&bytes).unwrap();
    assert!(p.surveys().unwrap().is_empty());
    let mut edit = p.edit_surveys().unwrap();
    assert_eq!(
        edit.insert_for_table(table, authored(first, second))
            .unwrap(),
        0
    );
    let entry = edit.entries().unwrap()[0].clone();
    assert_eq!(
        entry
            .survey()
            .questions()
            .values()
            .iter()
            .map(|q| q.binding().get())
            .collect::<Vec<_>>(),
        [1, 2]
    );
    let name = entry
        .part_name()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let c = edit.commit().unwrap();
    let q = publish(&p, &c, &["[Content_Types].xml", TABLE_RELS, &name]);
    let mut expected = graph(&p);
    expected.insert(
        entry.part_name().to_string(),
        (survey::CONTENT_TYPE.into(), vec![]),
    );
    let rels = &mut expected.get_mut(TABLE).unwrap().1;
    rels.push((
        entry.relationship_id().into(),
        survey::RELATIONSHIP_TYPE.into(),
        entry.relationship_target().into(),
        false,
    ));
    rels.sort();
    assert_eq!(graph(&q), expected);
    assert_eq!(
        members(&q)["xl/tables/table1.xml"],
        members(&p)["xl/tables/table1.xml"]
    );
    assert_eq!(native(file), bytes);
    q
}
fn raw_source(file: &str, table: &str, first: &str, second: &str) -> Package {
    let p = attach(file, table, first, second);
    let entry = p.surveys().unwrap().entries()[0].clone();
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.get_part_mut(entry.part_name())?
            .set_blob(RAW.as_bytes().to_vec());
        Ok(())
    })
    .unwrap();
    let c = edit.commit().unwrap();
    let bytes = c.package().to_bytes().unwrap();
    let q = Package::from_bytes(&bytes).unwrap();
    preserve(
        &p,
        &q,
        &[entry.part_name().as_str().trim_start_matches('/')],
    );
    assert_eq!(graph(&p), graph(&q));
    assert_eq!(
        q.surveys().unwrap().entries()[0].survey(),
        &survey::parse(RAW.as_bytes(), &survey::Limits::default()).unwrap()
    );
    q
}

#[test]
fn native_xlsx_surveys_resolve_columns_and_preserve_tables_through_lifecycle() {
    for (file, table, first, second) in PARENTS {
        let original = Package::from_bytes(&native(file)).unwrap();
        let p = attach(file, table, first, second);
        let source = p.surveys().unwrap();
        let path = source.entries()[0]
            .part_name()
            .as_str()
            .trim_start_matches('/')
            .to_owned();
        let mut noop = p.edit_surveys().unwrap();
        assert!(!noop.set(0, source.entries()[0].survey().clone()).unwrap());
        let c = noop.commit().unwrap();
        assert!(c.patch().is_noop());
        assert_eq!(c.package().to_bytes().unwrap(), p.to_bytes().unwrap());
        let mut edit = p.edit_surveys().unwrap();
        edit.edit(0, |s| {
            s.set_title(Some("Updated & 🍎".into()));
            Ok(())
        })
        .unwrap();
        let staged = edit.entries().unwrap()[0].survey().clone();
        assert!(
            matches!(edit.edit(0,|s|{s.questions_mut().values_mut()[0].set_binding(Binding::new(999));Ok(())}),Err(Error::Invalid(m)) if m=="survey 0 has invalid table binding: invalid XLSX structure: survey question binding 999 does not match an owner table column")
        );
        assert!(
            matches!(edit.edit(0,|s|{s.questions_mut().values_mut()[0].set_decimal_places(Some(16));Ok(())}),Err(Error::Invalid(m)) if m=="survey decimalPlaces must be at most 15")
        );
        assert_eq!(edit.entries().unwrap()[0].survey(), &staged);
        let c = edit.commit().unwrap();
        let q = publish(&p, &c, &[&path]);
        assert_eq!(graph(&p), graph(&q));
        assert_eq!(q.surveys().unwrap().entries()[0].survey(), &staged);
        let mut edit = q.edit_surveys().unwrap();
        assert!(edit.remove(0).unwrap().is_some());
        assert!(edit.remove(0).unwrap().is_none());
        let c = edit.commit().unwrap();
        let removed = publish(&q, &c, &["[Content_Types].xml", TABLE_RELS, &path]);
        assert!(removed.surveys().unwrap().is_empty());
        assert_eq!(graph(&removed), graph(&original));
        preserve(&original, &removed, &["[Content_Types].xml", TABLE_RELS]);
    }
}

#[test]
fn native_xlsx_survey_scalar_edits_preserve_namespaces_comments_and_exact_unowned_xml() {
    for (file, table, first, second) in PARENTS {
        let p = raw_source(file, table, first, second);
        let catalog = p.surveys().unwrap();
        let source = catalog.entries()[0].survey();
        let path = catalog.entries()[0]
            .part_name()
            .as_str()
            .trim_start_matches('/')
            .to_owned();
        assert_eq!(source.title(), Some("Before A"));
        assert_eq!(source.questions().values()[0].text(), Some("Name _x0041_"));
        assert_eq!(source.questions().values()[1].default_value(), Some("\0"));
        let mut edit = p.edit_surveys().unwrap();
        edit.edit(0, |s| {
            s.set_title(Some("After & 🍎 _x0041_".into()));
            s.questions_mut().values_mut()[0].set_text(Some("Changed & 日本語".into()));
            Ok(())
        })
        .unwrap();
        let c = edit.commit().unwrap();
        let q = publish(&p, &c, &[&path]);
        let expected = RAW
            .replace(
                "title='Before _x0041_'",
                "title='After &amp; 🍎 _x005F_x0041_'",
            )
            .replace("text='Name _x005F_x0041_'", "text='Changed &amp; 日本語'");
        assert_eq!(members(&q)[&path], expected.as_bytes());
        assert_eq!(graph(&p), graph(&q));
        assert!(matches!(
            q.apply_surveys(c.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }
}

#[test]
fn native_xlsx_survey_question_structure_retains_opaque_source_fragments() {
    for (file, table, first, second) in PARENTS {
        let p = raw_source(file, table, first, second);
        let cat = p.surveys().unwrap();
        let path = cat.entries()[0]
            .part_name()
            .as_str()
            .trim_start_matches('/')
            .to_owned();
        let mut edit = p.edit_surveys().unwrap();
        edit.edit(0, |s| {
            let q = s.questions_mut().remove(0).unwrap();
            let mut added = Question::new(Binding::new(3));
            added.set_text(Some("Number 🍎".into()));
            added.set_question_type(Some(QuestionType::Number));
            added.set_format(Some(survey::QuestionFormat::Fixed));
            added.set_decimal_places(Some(2));
            s.questions_mut().insert(0, added)?;
            s.questions_mut().push(q)?;
            Ok(())
        })
        .unwrap();
        let expected = edit.entries().unwrap()[0].survey().clone();
        let c = edit.commit().unwrap();
        let q = publish(&p, &c, &[&path]);
        assert_eq!(graph(&p), graph(&q));
        let changed = q.surveys().unwrap();
        let survey = changed.entries()[0].survey();
        assert_eq!(survey, &expected);
        assert_eq!(
            survey
                .questions()
                .values()
                .iter()
                .map(|q| q.binding().get())
                .collect::<Vec<_>>(),
            [3, 2, 1]
        );
        assert_eq!(
            survey.questions().values()[2].extension_xml(),
            cat.entries()[0].survey().questions().values()[0].extension_xml()
        );
        let xml = members(&q).remove(&path).unwrap();
        let text = std::str::from_utf8(&xml).unwrap();
        let begin = RAW.find("<s:questions ").unwrap();
        let end = RAW.find("</s:questions>").unwrap() + "</s:questions>".len();
        assert!(text.starts_with(&RAW[..begin]));
        assert!(text.ends_with(&RAW[end..]));
        let question = RAW.find("<s:question binding='1'").unwrap();
        let question_end =
            RAW[question..].find("</s:question>").unwrap() + question + "</s:question>".len();
        assert!(text.contains(&RAW[question..question_end]));
        for marker in [
            "<!-- root comment -->",
            "<!-- before first question -->",
            "<!-- between questions -->",
            "<!-- after last question -->",
        ] {
            assert!(text.contains(marker));
        }
    }
}

#[test]
fn native_xlsx_survey_patch_checks_owner_table_bytes_even_when_survey_is_unchanged() {
    let (file, table, first, second) = PARENTS[3];
    let p = raw_source(file, table, first, second);
    let before = members(&p);
    let mut edit = p.edit_surveys().unwrap();
    edit.edit(0, |s| {
        s.set_description(Some("Staged description".into()));
        Ok(())
    })
    .unwrap();
    let c = edit.commit().unwrap();
    let mut raw = p.edit_opc().unwrap();
    raw.update(|opc| {
        let owner = opc.get_part_mut(&PackURI::new(TABLE).unwrap())?;
        let xml = std::str::from_utf8(owner.blob())
            .unwrap()
            .replace("</table>", "<!-- source changed --></table>");
        owner.set_blob(xml.into_bytes());
        Ok(())
    })
    .unwrap();
    let different = raw.commit().unwrap();
    let q = different.package();
    assert_eq!(
        p.surveys().unwrap().entries()[0].survey(),
        q.surveys().unwrap().entries()[0].survey()
    );
    assert!(
        matches!(q.apply_surveys(c.patch()),Err(Error::PatchConflict{part}) if part=="Survey catalog")
    );
    assert_eq!(members(&p), before);
}

#[test]
fn native_xlsx_survey_relationship_targets_cannot_silently_disappear_from_the_catalog() {
    let (file, table, first, second) = PARENTS[0];
    let p = attach(file, table, first, second);
    let entry = p.surveys().unwrap().entries()[0].clone();
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.get_part_mut(entry.part_name())?
            .set_content_type("application/xml".into())?;
        Ok(())
    })
    .unwrap();
    let wrong = edit.commit().unwrap();
    match wrong.package().surveys() {
        Ok(catalog) => panic!(
            "wrong Survey content type returned {} surveys",
            catalog.len()
        ),
        Err(Error::Invalid(message)) => assert_eq!(
            message,
            "Survey relationship target has invalid content type"
        ),
        Err(error) => panic!("unexpected Survey error: {error}"),
    }
    assert!(
        matches!(wrong.package().workbook_ref().surveys(),Err(Error::Invalid(m)) if m=="Survey relationship target has invalid content type")
    );
    for (target, mode, expected) in [
        (
            "https://example.invalid/not-loaded.xml",
            litchi_opc::TargetMode::External,
            "Survey relationship must be internal",
        ),
        (
            "../surveys/missing.xml",
            litchi_opc::TargetMode::Internal,
            "Survey relationship target is absent",
        ),
    ] {
        let mut opc = p.clone().into_plain_opc();
        let relationships = opc
            .get_part_mut(entry.table_part_name())
            .unwrap()
            .rels_mut();
        relationships.remove(entry.relationship_id()).unwrap();
        relationships
            .try_add_relationship(
                survey::RELATIONSHIP_TYPE.into(),
                target.into(),
                entry.relationship_id().into(),
                mode,
            )
            .unwrap();
        assert!(
            matches!(survey::Snapshot::load(&opc), Err(Error::Invalid(message)) if message == expected)
        );
    }
    assert_eq!(p.surveys().unwrap().len(), 1);
}

#[test]
fn native_xlsx_survey_question_limits_and_duplicate_ids_leave_staging_intact() {
    let (file, table, first, second) = PARENTS[3];
    let original = Package::from_bytes(&native(file)).unwrap();
    let mut bounded = original
        .edit_surveys_with_options(
            survey::EditOptions::default().limits(survey::Limits::default().with_max_questions(1)),
        )
        .unwrap();
    assert!(matches!(
        bounded.insert_for_table(table, authored(first, second)),
        Err(Error::Limit {
            kind: litchi_xlsx::LimitKind::Survey(survey::LimitKind::Questions),
            observed: 2,
            maximum: 1
        })
    ));
    assert!(bounded.entries().unwrap().is_empty());
    assert!(bounded.commit().unwrap().patch().is_noop());
    let p = attach(file, table, first, second);
    let mut edit = p.edit_surveys().unwrap();
    edit.edit(0, |s| {
        s.set_title(Some("Valid staged title".into()));
        Ok(())
    })
    .unwrap();
    assert!(
        matches!(edit.insert_for_table(table,authored(first,second)),Err(Error::Invalid(m)) if m=="survey IDs must be unique within a workbook")
    );
    assert_eq!(edit.entries().unwrap().len(), 1);
    assert_eq!(
        edit.entries().unwrap()[0].survey().title(),
        Some("Valid staged title")
    );
    let path = edit.entries().unwrap()[0]
        .part_name()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let c = edit.commit().unwrap();
    publish(&p, &c, &[&path]);
}

#[test]
fn native_xlsx_survey_creation_restores_explicit_empty_relationships_and_checks_new_owner() {
    let (file, table, first, second) = PARENTS[0];
    let native_package = Package::from_bytes(&native(file)).unwrap();
    let mut entries = members(&native_package);
    assert!(!entries.contains_key(TABLE_RELS));
    let empty=br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><!-- explicit empty table relationships --></Relationships>"#;
    entries.insert(TABLE_RELS.into(), empty.to_vec());
    let mut zip = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    for (name, bytes) in &entries {
        zip.write_stored(name, bytes).unwrap();
    }
    let p = Package::from_bytes(&zip.finish_to_bytes().unwrap()).unwrap();
    assert!(p.surveys().unwrap().is_empty());
    let mut edit = p.edit_surveys().unwrap();
    edit.insert_for_table(table, authored(first, second))
        .unwrap();
    let name = edit.entries().unwrap()[0]
        .part_name()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let c = edit.commit().unwrap();
    let q = publish(&p, &c, &["[Content_Types].xml", TABLE_RELS, &name]);
    let restored = q.apply_surveys(&c.patch().inverse()).unwrap();
    assert_eq!(members(restored.package())[TABLE_RELS], empty);
    let mut raw = p.edit_opc().unwrap();
    raw.update(|opc| {
        let part = opc.get_part_mut(&PackURI::new(TABLE).unwrap())?;
        let xml = std::str::from_utf8(part.blob())
            .unwrap()
            .replace("</table>", "<!-- fresh owner changed --></table>");
        part.set_blob(xml.into_bytes());
        Ok(())
    })
    .unwrap();
    let changed = raw.commit().unwrap();
    assert!(
        matches!(changed.package().apply_surveys(c.patch()),Err(Error::PatchConflict{part}) if part=="Survey owner table")
    );
    same_members(&p, restored.package());
}
