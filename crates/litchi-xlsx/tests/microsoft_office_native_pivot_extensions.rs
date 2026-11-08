//! Explicit non-worksheet PivotTable metadata on native parents; no model loading or refresh.
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, PackURI, SourceLimits, TargetMode, part::Part};
use litchi_xlsx::{
    Error, Package,
    pivot::{cached_unique_names as names, server_formats as formats},
};
use quick_xml::{Reader, events::Event};
use std::{collections::BTreeMap, path::Path};
const PARENTS: [&str; 5] = [
    "excel-metadata.xlsx",
    "excel-recorded-macro.xlsm",
    "excel-visual-analytics-strict.xlsx",
    "excel-visual-analytics.xltx",
    "excel-connections.xlsx",
];
const BOOK: &str = "/xl/workbook.xml";
const TABLE: &str = "/xl/pivotTables/nativeMetadata.xml";
const CACHE: &str = "/xl/pivotCache/nativeMetadata.xml";
const CONNECTIONS: &str = "/xl/connections.xml";
const PIVOT: &str = "NativeMetadata700";
const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const EXT: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const SERVER_URI: &str = "{C510F80B-63DE-4267-81D5-13C33094786E}";
const NAMES_URI: &str = "{4F2E5C28-24EA-4EB8-9CBF-B6C8F9C3D259}";
const REFERENCES_URI: &str = "{983426D0-5260-488c-9760-48F4B6AC55F4}";
const VERSION_URI: &str = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";
const MODEL_URI: &str = "{DE250136-89BD-433C-8126-D09CA5730AF9}";
const SOURCE_URI: &str = "{F057638F-6D5F-4E77-A914-E7F072B9BCA8}";
const CONNECTION_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
fn uri(s: &str) -> PackURI {
    PackURI::new(s).unwrap()
}
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
        SourceLimits::default(),
    )
    .unwrap();
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
fn same(a: &Package, b: &Package) {
    let a = members(a);
    let b = members(b);
    let changed = a
        .keys()
        .chain(b.keys())
        .filter(|n| a.get(*n) != b.get(*n))
        .collect::<std::collections::BTreeSet<_>>();
    assert!(changed.is_empty(), "different members: {changed:?}");
}
fn preserve(a: &Package, b: &Package, owned: &[&str]) {
    let a = members(a);
    let b = members(b);
    for n in a.keys().chain(b.keys()) {
        if !owned.contains(&n.as_str()) {
            assert!(a.get(n) == b.get(n), "changed {n}");
        }
    }
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    p.save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    same(p, &q);
    q
}
/// Direct child ranges of the native workbook root; retained bytes are never reserialized.
fn child(xml: &str, name: &str) -> Option<(usize, usize, usize)> {
    let mut r = Reader::from_str(xml);
    let mut depth = 0;
    let mut start = None;
    loop {
        let at = r.buffer_position() as usize;
        match r.read_event().unwrap() {
            Event::Start(e) => {
                if depth == 1 && e.name().as_ref() == name.as_bytes() {
                    start = Some(at);
                }
                depth += 1;
            },
            Event::Empty(e) if depth == 1 && e.name().as_ref() == name.as_bytes() => {
                return Some((at, r.buffer_position() as usize, at));
            },
            Event::End(_) => {
                depth -= 1;
                if depth == 1
                    && let Some(start) = start
                {
                    return Some((start, r.buffer_position() as usize, at));
                }
            },
            Event::Eof => return None,
            _ => {},
        }
    }
}
fn append(xml: &str, name: &str, addition: &str) -> String {
    let (start, end, close) = child(xml, name).unwrap();
    if start == close {
        let open = &xml[start..end - 2];
        format!(
            "{}{}>{}</{}>{}",
            &xml[..start],
            open,
            addition,
            name,
            &xml[end..]
        )
    } else {
        format!("{}{}{}", &xml[..close], addition, &xml[close..])
    }
}
fn append_root(xml: &str, addition: &str) -> String {
    let at = xml.rfind("</").unwrap();
    format!("{}{}{}", &xml[..at], addition, &xml[at..])
}
fn table_xml(main: &str) -> String {
    format!(
        r#"<pivotTableDefinition xmlns='{main}' xmlns:m='{EXT}' xmlns:mc='{MCE}' mc:Ignorable='m' name='{PIVOT}' cacheId='700' dataCaption='Stored values'>
<!-- table prefix -->
<location ref='A1:B2' firstHeaderRow='0' firstDataRow='0' firstDataCol='0'/><pivotFields count='1'><pivotField/></pivotFields><rowItems count='1'><i/></rowItems><colItems count='1'><i/></colItems>
<extLst><ext uri='{SERVER_URI}'><m:pivotTableServerFormats count='2'>
<!-- format zero --><m:serverFormat culture='en-US' format='#,##0.00'/><!-- format one --><m:serverFormat culture='fr-FR' format='_x005F_x0041_'/>
</m:pivotTableServerFormats></ext><ext uri='{{44433962-1CF7-4059-B4EE-95C3D5FFCF73}}'><m:pivotTableData rowCount='1' columnCount='1' cacheId='700'><m:pivotRow count='1' r='0'><m:c i='0' t='n'><m:v>42.5</m:v><m:x in='1'/></m:c></m:pivotRow></m:pivotTableData></ext><ext uri='urn:native:untouched'><k:kept xmlns:k='urn:native:pivot-opaque' value='A &amp; B'/></ext></extLst>
<!-- table tail --></pivotTableDefinition>"#
    )
}
fn cache_xml(main: &str) -> String {
    format!(
        r#"<pivotCacheDefinition xmlns='{main}' xmlns:m='{EXT}' xmlns:x14='{X14}' xmlns:mc='{MCE}' mc:Ignorable='m x14'>
<cacheSource type='external' connectionId='700'><extLst><ext uri='{SOURCE_URI}'><x14:sourceConnection name='{PIVOT}'/></ext></extLst></cacheSource>
<cacheFields count='1'><cacheField name='Region'><sharedItems count='2'><s v='North'/><s v='South'/></sharedItems><extLst><ext uri='{NAMES_URI}'><m:cachedUniqueNames>
<!-- name zero --><m:cachedUniqueName index='0' name='[Region].[North]'/><!-- name one --><m:cachedUniqueName index='1' name='_x005F_x0041_'/>
</m:cachedUniqueNames></ext><ext uri='urn:native:field-opaque'><k:kept xmlns:k='urn:native:pivot-opaque' value='field'/></ext></extLst></cacheField></cacheFields>
<extLst><ext uri='{VERSION_URI}'><m:pivotCacheIdVersion cacheIdSupportedVersion='15' cacheIdCreatedVersion='15'/></ext><ext uri='{{725AE2AE-9491-48BE-B2B4-4EB974FC3084}}'><x14:pivotCacheDefinition pivotCacheId='700'/></ext></extLst>
</pivotCacheDefinition>"#
    )
}
fn seed(file: &str) -> Package {
    let bytes = native(file);
    let p = Package::from_bytes(&bytes).unwrap();
    let old = members(&p);
    let mut expected = graph(&p);
    let strict = file.contains("strict");
    let main = if strict { STRICT } else { MAIN };
    let rel = if strict {
        "http://purl.oclc.org/ooxml/officeDocument/relationships"
    } else {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    };
    let connection = format!(
        r#"<connection xmlns='{main}' xmlns:m='{EXT}' id='700' name='{PIVOT}' type='5' refreshedVersion='7'><extLst><ext uri='{MODEL_URI}'><m:connection model='true' id=''/></ext></extLst></connection>"#
    );
    let new_connections = if let Some(b) = old.get("xl/connections.xml") {
        append_root(std::str::from_utf8(b).unwrap(), &connection)
    } else {
        format!("<connections xmlns='{main}'>{connection}</connections>")
    };
    let cache_ref = format!(
        r#"<pivotCache xmlns='{main}' xmlns:r='{rel}' cacheId='700' r:id='rIdNativeCache700'/>"#
    );
    let mut book = String::from_utf8(old["xl/workbook.xml"].clone()).unwrap();
    if child(&book, "pivotCaches").is_some() {
        book = append(&book, "pivotCaches", &cache_ref);
    } else {
        let at = [
            "smartTagPr",
            "smartTagTypes",
            "webPublishing",
            "fileRecoveryPr",
            "webPublishObjects",
            "extLst",
        ]
        .iter()
        .filter_map(|n| child(&book, n).map(|r| r.0))
        .min()
        .unwrap();
        book.insert_str(at, &format!("<pivotCaches>{cache_ref}</pivotCaches>"));
    }
    let reference = format!(
        r#"<ext xmlns='{main}' xmlns:m='{EXT}' xmlns:r='{rel}' xmlns:mc='{MCE}' mc:Ignorable='m' uri='{REFERENCES_URI}'><m:pivotTableReferences><m:pivotTableReference r:id='rIdNativePivot700'/></m:pivotTableReferences></ext>"#
    );
    book = append(&book, "extLst", &reference);
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.try_add_part(Box::new(BlobPart::new(
            uri(CACHE),
            ct::SML_PIVOT_CACHE_DEFINITION.into(),
            cache_xml(main).into_bytes(),
        )))?;
        let mut table = BlobPart::new(
            uri(TABLE),
            ct::SML_PIVOT_TABLE.into(),
            table_xml(main).into_bytes(),
        );
        table.rels_mut().try_add_relationship(
            format!("{rel}/pivotCacheDefinition"),
            "../pivotCache/nativeMetadata.xml".into(),
            "rIdCache".into(),
            TargetMode::Internal,
        )?;
        opc.try_add_part(Box::new(table))?;
        if old.contains_key("xl/connections.xml") {
            opc.get_part_mut(&uri(CONNECTIONS))?
                .set_blob(new_connections.clone().into_bytes());
        } else {
            opc.try_add_part(Box::new(BlobPart::new(
                uri(CONNECTIONS),
                CONNECTION_TYPE.into(),
                new_connections.clone().into_bytes(),
            )))?;
        }
        let owner = opc.get_part_mut(&uri(BOOK))?;
        owner.set_blob(book.clone().into_bytes());
        owner.rels_mut().try_add_relationship(
            format!("{rel}/pivotCacheDefinition"),
            "pivotCache/nativeMetadata.xml".into(),
            "rIdNativeCache700".into(),
            TargetMode::Internal,
        )?;
        owner.rels_mut().try_add_relationship(
            format!("{rel}/pivotTable"),
            "pivotTables/nativeMetadata.xml".into(),
            "rIdNativePivot700".into(),
            TargetMode::Internal,
        )?;
        if !old.contains_key("xl/connections.xml") {
            owner.rels_mut().try_add_relationship(
                format!("{rel}/connections"),
                "connections.xml".into(),
                "rIdNativeConnection700".into(),
                TargetMode::Internal,
            )?;
        }
        Ok(())
    })
    .unwrap();
    let c = edit.commit().unwrap();
    let q = save(c.package());
    preserve(
        &p,
        &q,
        &[
            "[Content_Types].xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/connections.xml",
            &CACHE[1..],
            &TABLE[1..],
            "xl/pivotTables/_rels/nativeMetadata.xml.rels",
        ],
    );
    same(p.apply_opc(c.patch()).unwrap().package(), &q);
    same(q.apply_opc(&c.patch().inverse()).unwrap().package(), &p);
    expected.insert(
        CACHE.into(),
        (ct::SML_PIVOT_CACHE_DEFINITION.into(), vec![]),
    );
    expected.insert(
        TABLE.into(),
        (
            ct::SML_PIVOT_TABLE.into(),
            vec![(
                "rIdCache".into(),
                format!("{rel}/pivotCacheDefinition"),
                "../pivotCache/nativeMetadata.xml".into(),
                false,
            )],
        ),
    );
    if !old.contains_key("xl/connections.xml") {
        expected.insert(CONNECTIONS.into(), (CONNECTION_TYPE.into(), vec![]));
        expected.get_mut(BOOK).unwrap().1.push((
            "rIdNativeConnection700".into(),
            format!("{rel}/connections"),
            "connections.xml".into(),
            false,
        ));
    }
    expected.get_mut(BOOK).unwrap().1.extend([
        (
            "rIdNativeCache700".into(),
            format!("{rel}/pivotCacheDefinition"),
            "pivotCache/nativeMetadata.xml".into(),
            false,
        ),
        (
            "rIdNativePivot700".into(),
            format!("{rel}/pivotTable"),
            "pivotTables/nativeMetadata.xml".into(),
            false,
        ),
    ]);
    expected.get_mut(BOOK).unwrap().1.sort();
    assert_eq!(graph(&q), expected);
    assert_eq!(native(file), bytes);
    q
}
fn publish_formats(p: &Package, c: &formats::Commit) -> Package {
    let q = save(c.package());
    preserve(p, &q, &[&TABLE[1..]]);
    assert_eq!(graph(p), graph(&q));
    same(p.apply_pivot_table(c.patch()).unwrap().package(), &q);
    same(
        q.apply_pivot_table(&c.patch().inverse()).unwrap().package(),
        p,
    );
    q
}
fn publish_names(p: &Package, c: &names::Commit) -> Package {
    let q = save(c.package());
    preserve(p, &q, &[&CACHE[1..]]);
    assert_eq!(graph(p), graph(&q));
    same(p.apply_pivot_cache(c.patch()).unwrap().package(), &q);
    same(
        q.apply_pivot_cache(&c.patch().inverse()).unwrap().package(),
        p,
    );
    q
}
fn names(p: &Package) -> names::CachedUniqueNames {
    p.workbook_ref()
        .pivot_cache(700u32)
        .field(names::PivotCacheFieldSelector::Name("Region"))
        .cached_unique_names()
        .unwrap()
}

#[test]
fn native_xlsx_pivot_extensions_read_exact_metadata_and_retain_noop_packages() {
    for file in PARENTS {
        let p = seed(file);
        let view = p.workbook_ref().pivot_table_server_formats(PIVOT).unwrap();
        assert_eq!(view.cache_id(), 700);
        assert_eq!(
            view.server_formats().formats(),
            [
                formats::ServerFormat::new(Some("en-US".into()), Some("#,##0.00".into())),
                formats::ServerFormat::new(Some("fr-FR".into()), Some("_x0041_".into()))
            ]
        );
        let cache = names(&p);
        assert_eq!(cache.cache_id(), names::PivotCacheId(700));
        assert_eq!(cache.field_name(), "Region");
        assert_eq!(cache.field_ordinal(), 0);
        assert_eq!(
            cache.index_bound_status(),
            names::IndexBoundStatus::Unresolved
        );
        assert_eq!(
            cache.names(),
            [
                names::CachedUniqueName {
                    index: 0,
                    name: "[Region].[North]".into()
                },
                names::CachedUniqueName {
                    index: 1,
                    name: "_x0041_".into()
                }
            ]
        );
        let c = p.edit_pivot_table(PIVOT).unwrap().commit().unwrap();
        assert!(c.patch().is_noop());
        same(&p, c.package());
        let c = p
            .edit_pivot_cache(700u32, names::PivotCacheFieldSelector::Name("Region"))
            .unwrap()
            .commit()
            .unwrap();
        assert!(c.patch().is_noop());
        same(&p, c.package());
    }
}
#[test]
fn native_xlsx_server_format_scalar_changes_preserve_exact_unowned_xml() {
    for file in PARENTS {
        let p = seed(file);
        let before = members(&p);
        let mut edit = p.edit_pivot_table(PIVOT).unwrap();
        assert!(edit.set_culture(0, Some("ja-JP".into())).unwrap());
        assert!(edit.set_format(1, Some("_x0041_ & 日本語".into())).unwrap());
        let c = edit.commit().unwrap();
        let q = publish_formats(&p, &c);
        let expected = String::from_utf8(before[&TABLE[1..]].clone())
            .unwrap()
            .replace("culture='en-US'", "culture='ja-JP'")
            .replace(
                "format='_x005F_x0041_'",
                "format='_x005F_x0041_ &amp; 日本語'",
            );
        assert_eq!(members(&q)[&TABLE[1..]], expected.as_bytes());
        assert!(matches!(
            q.apply_pivot_table(c.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }
}
#[test]
fn native_xlsx_cached_unique_name_edits_preserve_indexes_and_refuse_unproven_structure() {
    for file in PARENTS {
        let p = seed(file);
        let before = members(&p);
        let mut edit = p
            .edit_pivot_cache(700u32, names::PivotCacheFieldSelector::Name("Region"))
            .unwrap();
        assert!(
            edit.set_cached_unique_name(1, "[Region].[東京🍎] & _x0041_")
                .unwrap()
        );
        assert!(
            matches!(edit.insert_cached_unique_name(0,names::CachedUniqueName{index:2,name:"new".into()}),Err(Error::Invalid(m)) if m=="cachedUniqueNames structural insertion requires the unresolved CT_Items bound")
        );
        assert!(
            matches!(edit.remove_cached_unique_name(0),Err(Error::Invalid(m)) if m=="cachedUniqueNames structural removal requires the unresolved CT_Items bound")
        );
        assert!(
            matches!(edit.set_cached_unique_name_index(0,1),Err(Error::Invalid(m)) if m=="cachedUniqueNames index edits require the unresolved CT_Items bound")
        );
        let c = edit.commit().unwrap();
        let q = publish_names(&p, &c);
        assert_eq!(
            names(&q).names()[1],
            names::CachedUniqueName {
                index: 1,
                name: "[Region].[東京🍎] & _x0041_".into()
            }
        );
        let expected = String::from_utf8(before[&CACHE[1..]].clone())
            .unwrap()
            .replace(
                "name='_x005F_x0041_'",
                "name='[Region].[東京🍎] &amp; _x005F_x0041_'",
            );
        assert_eq!(members(&q)[&CACHE[1..]], expected.as_bytes());
        assert!(matches!(
            q.apply_pivot_cache(c.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }
}

fn replace_xml(p: &Package, path: &str, from: &str, to: &str) -> Package {
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        let part = opc.get_part_mut(&uri(path))?;
        let text = std::str::from_utf8(part.blob()).unwrap();
        assert!(text.contains(from));
        part.set_blob(text.replace(from, to).into_bytes());
        Ok(())
    })
    .unwrap();
    let c = edit.commit().unwrap();
    let q = save(c.package());
    preserve(p, &q, &[&path[1..]]);
    assert_eq!(graph(p), graph(&q));
    same(q.apply_opc(&c.patch().inverse()).unwrap().package(), p);
    q
}
fn malformed_opc(p: &Package, path: &str, from: &str, to: &str) -> litchi_opc::OpcPackage {
    let mut opc = p.clone().into_plain_opc();
    let part = opc.get_part_mut(&uri(path)).unwrap();
    let text = std::str::from_utf8(part.blob()).unwrap();
    assert!(text.contains(from));
    part.set_blob(text.replace(from, to).into_bytes());
    opc
}
fn table_unowned(p: &Package) -> String {
    let mut xml = String::from_utf8(members(p)[&TABLE[1..]].clone()).unwrap();
    let start = xml.find("<m:pivotTableServerFormats").unwrap();
    let end =
        xml.find("</m:pivotTableServerFormats>").unwrap() + "</m:pivotTableServerFormats>".len();
    xml.replace_range(start..end, "");
    let start = xml.find("<m:x in='").unwrap() + "<m:x in='".len();
    let end = start + xml[start..].find('\'').unwrap();
    xml.replace_range(start..end, "INDEX");
    xml
}
fn referenced_format(p: &Package) -> formats::ServerFormat {
    let parts = members(p);
    let xml = std::str::from_utf8(&parts[&TABLE[1..]]).unwrap();
    let start = xml.find("<m:x in='").unwrap() + "<m:x in='".len();
    let end = start + xml[start..].find('\'').unwrap();
    let index = xml[start..end].parse::<usize>().unwrap();
    p.workbook_ref()
        .pivot_table_server_formats(PIVOT)
        .unwrap()
        .server_formats()
        .formats()[index]
        .clone()
}
#[test]
fn native_xlsx_server_format_structure_preserves_cell_index_associations() {
    for file in PARENTS {
        let p = seed(file);
        let rest = table_unowned(&p);
        let original = p
            .workbook_ref()
            .pivot_table_server_formats(PIVOT)
            .unwrap()
            .server_formats()
            .formats()
            .to_vec();
        let mut edit = p.edit_pivot_table(PIVOT).unwrap();
        edit.insert_server_format(
            0,
            formats::ServerFormat::new(Some("new".into()), Some("0".into())),
        )
        .unwrap();
        let staged = edit.server_formats().unwrap().to_vec();
        assert!(
            matches!(edit.reorder_server_formats(&[0,0,2]),Err(Error::Invalid(m)) if m=="server-format reorder contains a duplicate index")
        );
        assert_eq!(edit.server_formats().unwrap(), staged);
        let c = edit.commit().unwrap();
        let q = publish_formats(&p, &c);
        assert_eq!(table_unowned(&q), rest);
        assert_eq!(referenced_format(&q), original[1]);
        assert!(
            std::str::from_utf8(&members(&q)[&TABLE[1..]])
                .unwrap()
                .contains("<m:x in='2'/>")
        );
        let mut edit = q.edit_pivot_table(PIVOT).unwrap();
        edit.reorder_server_formats(&[2, 0, 1]).unwrap();
        let c = edit.commit().unwrap();
        let r = publish_formats(&q, &c);
        assert_eq!(table_unowned(&r), rest);
        assert_eq!(referenced_format(&r), original[1]);
        assert!(
            std::str::from_utf8(&members(&r)[&TABLE[1..]])
                .unwrap()
                .contains("<m:x in='0'/>")
        );
        let mut edit = r.edit_pivot_table(PIVOT).unwrap();
        let removed = edit.remove_server_format(1).unwrap();
        assert_eq!(removed.culture.as_deref(), Some("new"));
        let c = edit.commit().unwrap();
        let q = publish_formats(&r, &c);
        assert_eq!(table_unowned(&q), rest);
        assert_eq!(referenced_format(&q), original[1]);
        let before = members(&q);
        let mut invalid = q.edit_pivot_table(PIVOT).unwrap();
        invalid.remove_server_format(0).unwrap();
        assert!(
            matches!(invalid.commit(),Err(Error::Invalid(m)) if m=="removing a server-format leaf would orphan pivotValueCellExtra@in")
        );
        assert!(members(&q) == before);
        let mut edit = q.edit_pivot_table(PIVOT).unwrap();
        edit.move_server_format(0, 1).unwrap();
        let c = edit.commit().unwrap();
        let restored = publish_formats(&q, &c);
        assert_eq!(
            restored
                .workbook_ref()
                .pivot_table_server_formats(PIVOT)
                .unwrap()
                .server_formats()
                .formats(),
            original
        );
        assert_eq!(table_unowned(&restored), rest);
    }
}

fn wrap(xml: &str, element: &str) -> (String, String) {
    let start = xml.find(&format!("<m:{element}")).unwrap();
    let end = xml.find(&format!("</m:{element}>")).unwrap() + element.len() + 5;
    let from = xml[start..end].to_owned();
    let to = format!(
        "<mc:AlternateContent><mc:Choice Requires='m'>{from}</mc:Choice><mc:Fallback><m:futureFallback/></mc:Fallback></mc:AlternateContent>"
    );
    (from, to)
}
#[test]
fn native_xlsx_pivot_diagnostic_boundaries_remain_readable_and_refuse_changed_publication() {
    let p = seed(PARENTS[2]);
    let diagnostic = replace_xml(&p, TABLE, "<m:x in='1'/>", "<m:x in='2'/>");
    let view = diagnostic
        .workbook_ref()
        .pivot_table_server_formats(PIVOT)
        .unwrap();
    assert!(view.server_formats().has_diagnostic_index_boundary());
    let mut edit = diagnostic.edit_pivot_table(PIVOT).unwrap();
    assert!(
        matches!(edit.set_culture(0,Some("changed".into())),Err(Error::Invalid(m)) if m=="pivotTableServerFormats has an ambiguous pivotValueCellExtra@in == count boundary")
    );
    let data = members(&p);
    let text = std::str::from_utf8(&data[&TABLE[1..]]).unwrap();
    let (from, to) = wrap(text, "pivotTableServerFormats");
    let mce = replace_xml(&p, TABLE, &from, &to);
    let view = mce
        .workbook_ref()
        .pivot_table_server_formats(PIVOT)
        .unwrap();
    assert!(view.server_formats().has_ambiguous_mce_owner());
    let mut edit = mce.edit_pivot_table(PIVOT).unwrap();
    assert!(
        matches!(edit.set_format(0,Some("changed".into())),Err(Error::Invalid(m)) if m=="pivotTableServerFormats selected through an MCE branch is read-only")
    );
    let text = std::str::from_utf8(&data[&CACHE[1..]]).unwrap();
    let (from, to) = wrap(text, "cachedUniqueNames");
    let mce = replace_xml(&p, CACHE, &from, &to);
    let source = mce
        .workbook_ref()
        .pivot_cache(700u32)
        .field(names::PivotCacheFieldSelector::Name("Region"))
        .cached_unique_names_source()
        .unwrap();
    assert!(source.has_ambiguous_mce_owner());
    assert_eq!(source.cached_unique_names(), &names(&p));
    let noop = mce
        .edit_pivot_cache(700u32, names::PivotCacheFieldSelector::Name("Region"))
        .unwrap()
        .commit()
        .unwrap();
    assert!(noop.patch().is_noop());
    same(noop.package(), &mce);
    let before = members(&mce);
    let mut edit = mce
        .edit_pivot_cache(700u32, names::PivotCacheFieldSelector::Name("Region"))
        .unwrap();
    edit.set_cached_unique_name(0, "changed").unwrap();
    assert!(
        matches!(edit.commit(),Err(Error::Invalid(m)) if m=="cachedUniqueNames selected through an MCE branch is read-only")
    );
    assert!(members(&mce) == before);
}

#[test]
fn native_xlsx_pivot_owner_eligibility_and_mce_declarations_are_validated_in_context() {
    let p = seed(PARENTS[2]);
    let baseline = members(&p);
    let misplaced = replace_xml(&p, TABLE, "ref='A1:B2'", "ref='A3:B4'");
    assert!(
        matches!(misplaced.workbook_ref().pivot_table_server_formats(PIVOT),Err(Error::Invalid(m)) if m=="Non-Worksheet PivotTable location@ref must begin with A1")
    );
    let definition = format!(
        "<ext uri='{{962EF5D1-5CA2-4C93-8EF4-DBF5C05439D2}}' xmlns:x14='{X14}'><x14:pivotTableDefinition enableEdit='true'/></ext>"
    );
    let editable = replace_xml(&p, TABLE, "</extLst>", &format!("{definition}</extLst>"));
    assert!(
        matches!(editable.workbook_ref().pivot_table_server_formats(PIVOT),Err(Error::Invalid(m)) if m=="Non-Worksheet PivotTable enableEdit must be false")
    );
    let permitted = definition.replace("enableEdit='true'", "enableEdit='false'");
    let valid = replace_xml(&p, TABLE, "</extLst>", &format!("{permitted}</extLst>"));
    assert_eq!(
        valid
            .workbook_ref()
            .pivot_table_server_formats(PIVOT)
            .unwrap(),
        p.workbook_ref().pivot_table_server_formats(PIVOT).unwrap()
    );
    for child in ["pivotEdits", "pivotChanges", "conditionalFormats"] {
        let definition = format!(
            "<ext uri='{{962EF5D1-5CA2-4C93-8EF4-DBF5C05439D2}}' xmlns:x14='{X14}'><x14:pivotTableDefinition><x14:{child}/></x14:pivotTableDefinition></ext>"
        );
        let invalid = replace_xml(&p, TABLE, "</extLst>", &format!("{definition}</extLst>"));
        assert!(
            matches!(invalid.workbook_ref().pivot_table_server_formats(PIVOT),Err(Error::Invalid(m)) if m=="PivotTable is not eligible for pivotTableReferences")
        );
    }
    let collision = replace_xml(&p, TABLE, &format!("name='{PIVOT}'"), "name='PivotTable1'");
    assert!(
        matches!(collision.workbook_ref().pivot_table_server_formats("PivotTable1"),Err(Error::Invalid(m)) if m=="PivotTable name is not unique in the workbook")
    );
    for (value, message) in [
        ("missing", "mc:Ignorable contains an unbound prefix"),
        ("mc", "mc:Ignorable cannot name the MCE namespace"),
    ] {
        let bad = malformed_opc(
            &p,
            BOOK,
            "mc:Ignorable='m'",
            &format!("mc:Ignorable='{value}'"),
        );
        assert!(
            matches!(formats::Snapshot::load(&bad, PIVOT),Err(Error::Invalid(m)) if m==message)
        );
    }
    let wrong_namespace = malformed_opc(
        &p,
        BOOK,
        &format!("xmlns:mc='{MCE}'"),
        "xmlns:mc='urn:foreign-annotation'",
    );
    assert!(
        matches!(formats::Snapshot::load(&wrong_namespace, PIVOT),Err(Error::Invalid(m)) if m=="pivotTableReferences extension has an unknown attribute")
    );
    let nonmodel = replace_xml(&p, CONNECTIONS, "model='true'", "model='false'");
    assert!(
        matches!(nonmodel.workbook_ref().pivot_cache(700u32).field(names::PivotCacheFieldSelector::Name("Region")).cached_unique_names(),Err(Error::Invalid(m)) if m=="cachedUniqueNames connection is not a DE250 model connection")
    );
    assert!(members(&p) == baseline);
}

#[test]
fn native_xlsx_pivot_patches_check_connection_source_bytes_without_losing_other_parts() {
    let p = seed(PARENTS[4]);
    let mut edit = p.edit_pivot_table(PIVOT).unwrap();
    edit.set_culture(0, Some("staged".into())).unwrap();
    let formats = edit.commit().unwrap();
    let mut edit = p
        .edit_pivot_cache(700u32, names::PivotCacheFieldSelector::Name("Region"))
        .unwrap();
    edit.set_cached_unique_name(0, "staged").unwrap();
    let cache = edit.commit().unwrap();
    let changed = replace_xml(
        &p,
        CONNECTIONS,
        "</connection>",
        "<!-- source closure changed --></connection>",
    );
    assert_eq!(names(&p), names(&changed));
    assert_eq!(
        p.workbook_ref().pivot_table_server_formats(PIVOT).unwrap(),
        changed
            .workbook_ref()
            .pivot_table_server_formats(PIVOT)
            .unwrap()
    );
    assert!(matches!(
        changed.apply_pivot_table(formats.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert!(matches!(
        changed.apply_pivot_cache(cache.patch()),
        Err(Error::PatchConflict { .. })
    ));
}
