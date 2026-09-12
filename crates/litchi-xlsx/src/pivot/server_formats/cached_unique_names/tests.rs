use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const PIVOT_CACHE_ID_VERSION_URI: &str = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let cache_uri = PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap();
    let connections_uri = PackURI::new("/xl/connections.xml").unwrap();
    let workbook_xml = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><pivotCaches><pivotCache cacheId="7" r:id="rId1"/></pivotCaches></workbook>"#
    );
    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        workbook_xml.into_bytes(),
    );
    workbook.relate_to(
        "pivotCache/pivotCacheDefinition1.xml",
        rt::PIVOT_CACHE_DEFINITION,
    );
    workbook.relate_to("connections.xml", CONNECTIONS_RELATIONSHIP);
    let cache_xml = format!(
        r#"<pivotCacheDefinition xmlns="{SML}" xmlns:x14="{X14}" xmlns:x15="{X15}"><cacheSource type="external" connectionId="8"><extLst><ext uri="{CACHE_SOURCE_URI}"><x14:sourceConnection name="External"/></ext></extLst></cacheSource><cacheFields count="1"><cacheField name="Field"><sharedItems count="2"/><extLst><ext uri="{CACHED_UNIQUE_NAMES_URI}"><x15:cachedUniqueNames><x15:cachedUniqueName index="0" name="[A]&amp;B"/><x15:cachedUniqueName index="1" name="B"/></x15:cachedUniqueNames></ext></extLst></cacheField></cacheFields><extLst><ext uri="{PIVOT_CACHE_ID_VERSION_URI}"><x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></ext></extLst></pivotCacheDefinition>"#
    );
    let cache = BlobPart::new(
        cache_uri,
        ct::SML_PIVOT_CACHE_DEFINITION.into(),
        cache_xml.into_bytes(),
    );
    let connections_xml = format!(
        r#"<connections xmlns="{SML}"><connection id="8" name="External" type="5"><extLst><ext uri="{CONNECTION_MODEL_URI}"><x15:connection xmlns:x15="{X15}" model="true" id=""/></ext></extLst></connection></connections>"#
    );
    let connections = BlobPart::new(
        connections_uri,
        CONNECTIONS_CONTENT_TYPE.into(),
        connections_xml.into_bytes(),
    );
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(cache));
    package.add_part(Box::new(connections));
    package
}

#[test]
fn reads_model_cached_unique_names_without_an_item_bound() {
    let package = package();
    let snapshot = Snapshot::load(
        &package,
        CacheSelector::Id(PivotCacheId(7)),
        FieldSelector::Ordinal(0),
    )
    .unwrap();
    assert_eq!(snapshot.entries()[0].index, 0);
    assert_eq!(snapshot.entries()[0].name, "[A]&B");
    assert_eq!(snapshot.diagnostic_status(), DiagnosticStatus::Unresolved);
}

#[test]
fn scalar_edit_preserves_unchanged_lexical_neighbor_and_inverse() {
    let mut package = package();
    let uri = PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap();
    let before = package.get_part(&uri).unwrap().blob().to_vec();
    let mut transaction = Transaction::new(
        &mut package,
        CacheSelector::Id(PivotCacheId(7)),
        FieldSelector::Ordinal(0),
    )
    .unwrap();
    transaction
        .set_cached_unique_name(0, "Changed <name>")
        .unwrap();
    let commit = transaction.commit().unwrap();
    let changed = package.get_part(&uri).unwrap().blob().to_vec();
    assert!(
        std::str::from_utf8(&changed)
            .unwrap()
            .contains("Changed &lt;name>")
    );
    assert!(
        std::str::from_utf8(&changed)
            .unwrap()
            .contains(r#"name="B""#)
    );
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(package.get_part(&uri).unwrap().blob(), before);
}

fn element(index: usize, parent_index: Option<usize>, ns: &[u8], local: &[u8]) -> XmlElement {
    XmlElement {
        index,
        parent_index,
        ns: Arc::from(ns),
        local: local.to_vec(),
        start: 0..0,
        end: 0,
        attrs: Vec::new(),
        mce_branch: None,
        ignorable_scope: 0,
        mce_context: false,
        has_element_child: false,
        has_cdata: false,
        has_text: false,
        has_non_whitespace_cdata: false,
        has_non_whitespace_text: false,
    }
}

#[test]
fn mce_payload_requires_choice_or_fallback_ancestry() {
    let elements = vec![
        element(0, None, EXT_NS, b"ext"),
        element(1, Some(0), MCE_NS, b"AlternateContent"),
        element(2, Some(1), EXT_NS, CACHED_UNIQUE_NAMES_PAYLOAD),
    ];
    let scan = XmlScan {
        elements,
        limits: XmlScanLimits::from_read_limits(ReadLimits::default()),
        ignorable_scopes: Vec::new(),
    };
    assert!(
        owned_cache_payload(&scan, &scan.elements[2], &scan.elements[0])
            .unwrap()
            .is_none()
    );

    let elements = vec![
        element(0, None, EXT_NS, b"ext"),
        element(1, Some(0), MCE_NS, b"AlternateContent"),
        element(2, Some(1), MCE_NS, b"Choice"),
        element(3, Some(2), MCE_NS, b"AlternateContent"),
        element(4, Some(3), EXT_NS, CACHED_UNIQUE_NAMES_PAYLOAD),
    ];
    let scan = XmlScan {
        elements,
        limits: XmlScanLimits::from_read_limits(ReadLimits::default()),
        ignorable_scopes: Vec::new(),
    };
    assert!(
        owned_cache_payload(&scan, &scan.elements[4], &scan.elements[0])
            .unwrap()
            .is_none()
    );
}
