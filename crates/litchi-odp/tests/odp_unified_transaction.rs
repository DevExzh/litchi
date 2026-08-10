#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::chart::ChartClass;
use litchi_odf_common::chart::authoring::{
    CachedCell, CachedRow, CachedTable, CachedValue, Definition, SeriesSpec, Text,
};
use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::rdf::{Object, Subject, Triple};
use litchi_odp::{Builder, MasterPage, edit};
use soapberry_zip::office::StreamingArchiveWriter;
use std::path::Path;

const CHART: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:future="urn:example:future"><office:body><office:chart><chart:chart chart:class="chart:bar"><future:retained/></chart:chart></office:chart></office:body></office:document-content>"#;
const DEPENDENT_CHART: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:chart><chart:chart chart:class="chart:bar" xlink:href="Pictures/chart.png"><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#;

fn typed_chart() -> Definition {
    let mut definition = Definition::new(ChartClass::line());
    definition.title = Some(Text::new("Quarterly revenue"));
    definition.plot_area.series.push(SeriesSpec {
        values_cell_range_address: Some("local-table.B2:.B3".to_string()),
        label_cell_address: Some("local-table.B1".to_string()),
        ..SeriesSpec::default()
    });
    let mut table = CachedTable::new("local-table", 2);
    table.header_rows.push(CachedRow::new(vec![
        CachedCell::new(CachedValue::String("Quarter".to_string())),
        CachedCell::new(CachedValue::String("Revenue".to_string())),
    ]));
    table.rows.push(CachedRow::new(vec![
        CachedCell::new(CachedValue::String("Q1".to_string())),
        CachedCell::new(CachedValue::Float(10.0)),
    ]));
    table.rows.push(CachedRow::new(vec![
        CachedCell::new(CachedValue::String("Q2".to_string())),
        CachedCell::new(CachedValue::Float(14.5)),
    ]));
    definition.cached_table = Some(table);
    definition
}

fn literal(value: &str) -> Triple {
    Triple {
        subject: Subject::Iri("https://example.test/deck".to_string()),
        predicate: "https://example.test/title".to_string(),
        object: Object::Literal {
            value: value.to_string(),
            datatype: None,
            language: Some("en".to_string()),
        },
    }
}

#[test]
fn slide_and_rdf_edits_publish_as_one_reversible_package_commit() {
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles/><office:body><office:presentation><draw:page draw:name="Slide 1" draw:master-page-name="Default"><draw:rect draw:name="retained"/></draw:page></office:presentation></office:body></office:document-content>"#;
    const STYLES: &[u8] = br#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles/><office:automatic-styles><style:page-layout style:name="pm1"/></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer.add_file("styles.xml", STYLES).unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        source.security_policy().unwrap(),
        edit::SecurityPolicy::Editable
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::ClearSignatures)
            .unwrap(),
        edit::CryptoCapability::AvailableNoOp
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::DecryptForEditing)
            .unwrap(),
        edit::CryptoCapability::AvailableNoOp
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::AddSignature)
            .unwrap()
            .refusal(),
        Some(edit::CryptoRefusal::SignatureAuthoringUnavailable)
    );
    let source_bytes = source.bytes().to_vec();

    let mut transaction = source.transaction().unwrap();
    transaction
        .add("Second", "Slide and graph are atomic")
        .unwrap();
    let layout = litchi_odp::layout::Layout::new("unified-layout").unwrap();
    transaction.add_layout(&layout).unwrap();
    let mut master = MasterPage::new("unified-master", "pm1").unwrap();
    master.page_layout_name = Some("unified-layout".to_string());
    transaction.add_master_page(&master).unwrap();
    transaction
        .assign_slide_master_page(0, Some("unified-master"))
        .unwrap();
    transaction
        .assign_slide_page_layout(0, Some("unified-layout"))
        .unwrap();
    let mut annotation = litchi_odp::annotation::Annotation::new("Atomic review");
    annotation.set_name(Some("review-1"));
    transaction
        .add_annotation(&litchi_odp::annotation::Anchor::page(0), &annotation)
        .unwrap();
    transaction
        .add_chart(
            0usize,
            "Unified Chart",
            litchi_odp::charts::Storage::InlineXml,
            litchi_odp::charts::Part::from_xml(CHART).unwrap(),
        )
        .unwrap();
    let graph_path = transaction
        .add_rdf_graph(None, &[literal("First")])
        .unwrap();
    assert_eq!(graph_path, "Metadata/metadata_1.rdf");
    assert_eq!(
        transaction
            .add_rdf_triple(&graph_path, &literal("Second"))
            .unwrap(),
        1
    );
    transaction.move_rdf_triple(&graph_path, 1, 0).unwrap();
    transaction
        .replace_rdf_triple(&graph_path, 0, &literal("Published"))
        .unwrap();

    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(source.bytes(), source_bytes);
    assert_eq!(commit.snapshot().slides().len(), 2);

    let presentation = commit.snapshot().to_presentation().unwrap();
    let graphs = presentation.rdf_graphs().unwrap();
    assert_eq!(graphs.len(), 1);
    assert_eq!(graphs[0].path, graph_path);
    assert_eq!(graphs[0].triples.len(), 2);
    assert_eq!(graphs[0].triples[0], literal("Published"));
    assert_eq!(graphs[0].triples[1], literal("First"));
    assert!(
        presentation
            .layouts()
            .unwrap()
            .get("unified-layout")
            .is_some()
    );
    assert!(
        presentation
            .master_pages()
            .unwrap()
            .iter()
            .any(|candidate| candidate.name() == "unified-master")
    );
    assert_eq!(presentation.annotations().unwrap().len(), 1);
    assert_eq!(
        presentation
            .chart("Unified Chart")
            .unwrap()
            .unwrap()
            .storage(),
        litchi_odp::charts::Storage::InlineXml
    );

    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let rdf = String::from_utf8(package.get_file(&graph_path).unwrap()).unwrap();
    assert!(!rdf.contains('\n'));
    assert!(!rdf.contains("> <"));

    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    let restored = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(restored.bytes(), source.bytes());

    assert_eq!(
        commit.patch().domains(),
        &[
            edit::Domain::Slides,
            edit::Domain::Rdf,
            edit::Domain::Charts,
            edit::Domain::Design,
            edit::Domain::Annotations,
        ]
    );
    let durable = commit.patch().to_durable_bytes().unwrap();
    let durable_patch = edit::Patch::from_durable_bytes(&durable).unwrap();
    assert_eq!(durable_patch.domains(), commit.patch().domains());
    assert_eq!(
        durable_patch.apply(&source).unwrap().bytes(),
        commit.snapshot().bytes()
    );
    let mut malformed = durable;
    malformed.push(0);
    assert!(edit::Patch::from_durable_bytes(&malformed).is_err());

    let history_budget = source
        .bytes()
        .len()
        .checked_add(commit.snapshot().bytes().len())
        .unwrap();
    let mut history = edit::History::new(source.clone(), 2, history_budget).unwrap();
    history.record(&commit).unwrap();
    assert_eq!(history.current().bytes(), commit.snapshot().bytes());
    assert_eq!(history.undo().unwrap().bytes(), source.bytes());
    assert_eq!(history.redo().unwrap().bytes(), commit.snapshot().bytes());
    let durable_history = history.to_durable_bytes().unwrap();
    let mut reopened_history = edit::History::from_durable_bytes(&durable_history).unwrap();
    assert_eq!(
        reopened_history.current().bytes(),
        commit.snapshot().bytes()
    );
    assert_eq!(reopened_history.undo().unwrap().bytes(), source.bytes());
    assert_eq!(
        reopened_history.redo().unwrap().bytes(),
        commit.snapshot().bytes()
    );
    let mut malformed_history = durable_history;
    malformed_history.push(0);
    assert!(edit::History::from_durable_bytes(&malformed_history).is_err());
}

#[test]
fn merge_planning_is_independent_only_for_disjoint_rdf_work() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut slide_edit = source.transaction().unwrap();
    slide_edit.add("Slide", "content").unwrap();
    let slide_commit = slide_edit.commit().unwrap();

    let mut rdf_edit = source.transaction().unwrap();
    rdf_edit
        .add_rdf_graph(Some("Metadata/merge.rdf"), &[literal("metadata")])
        .unwrap();
    let rdf_commit = rdf_edit.commit().unwrap();

    let plan = slide_commit.patch().join(rdf_commit.patch()).unwrap();
    assert!(plan.is_independent());
    assert!(plan.conflicts().is_empty());
    assert!(
        edit::Patch::three_way(&source, slide_commit.patch(), rdf_commit.patch())
            .unwrap()
            .is_independent()
    );
    let durable_slide =
        edit::Patch::from_durable_bytes(&slide_commit.patch().to_durable_bytes().unwrap()).unwrap();
    let durable_rdf =
        edit::Patch::from_durable_bytes(&rdf_commit.patch().to_durable_bytes().unwrap()).unwrap();
    let joined = durable_slide.join_snapshot(&durable_rdf).unwrap();
    assert_eq!(joined.slides().len(), 1);
    assert_eq!(
        joined
            .to_presentation()
            .unwrap()
            .rdf_graphs()
            .unwrap()
            .len(),
        1
    );
    let three_way = edit::Patch::three_way_snapshot(&source, &durable_slide, &durable_rdf).unwrap();
    assert_eq!(three_way.bytes(), joined.bytes());

    let mut other_slide = source.transaction().unwrap();
    other_slide.add("Other", "content").unwrap();
    let other_commit = other_slide.commit().unwrap();
    let conflict = slide_commit
        .patch()
        .plan_join(other_commit.patch())
        .unwrap();
    assert_eq!(conflict.conflicts(), &[edit::Domain::Slides]);
    assert!(
        slide_commit
            .patch()
            .join_snapshot(other_commit.patch())
            .is_err()
    );
}

#[test]
fn typed_chart_data_transfers_and_a_real_impress_deck_fully_reopens() {
    let chart_base = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut source_transaction = chart_base.transaction().unwrap();
    source_transaction
        .add("Chart source", "typed cached data")
        .unwrap();
    source_transaction
        .add_chart_definition(
            0usize,
            "Typed source chart",
            litchi_odp::charts::Storage::PackageSubdocument,
            &typed_chart(),
        )
        .unwrap();
    let chart_source = source_transaction.commit().unwrap().snapshot().clone();

    let real_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/odp/tdf169979.odp");
    let real = edit::Snapshot::open(real_path).unwrap();
    let reopened_real = edit::Snapshot::from_bytes(real.bytes().to_vec()).unwrap();
    assert_eq!(reopened_real.slides(), real.slides());
    assert!(!real.slides().is_empty());

    let mut destination_builder = Builder::new();
    destination_builder.add_slide("Chart destination").unwrap();
    let destination_source =
        edit::Snapshot::from_bytes(destination_builder.build().unwrap()).unwrap();
    let mut destination = destination_source.transaction().unwrap();
    destination
        .transfer_chart_from(
            &chart_source,
            "Typed source chart",
            0usize,
            "Transferred typed chart",
            litchi_odp::charts::Storage::PackageSubdocument,
        )
        .unwrap();
    let commit = destination.commit().unwrap();
    assert_eq!(commit.snapshot().slides().len(), 1);

    let reopened = edit::Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let chart = reopened
        .to_presentation()
        .unwrap()
        .chart("Transferred typed chart")
        .unwrap()
        .unwrap();
    assert!(chart.content_xml().contains("Quarterly revenue"));
    assert!(chart.content_xml().contains("local-table"));
    assert!(chart.content_xml().contains("14.5"));
    assert!(!chart.content_xml().contains("> <"));
    assert!(!chart.content_xml().contains('\n'));

    let dependent_base = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut dependent = dependent_base.transaction().unwrap();
    dependent.add("Dependent", "chart").unwrap();
    dependent
        .embed_media("Pictures/chart.png", b"chart-pixels", "image/png")
        .unwrap();
    dependent
        .add_chart(
            0usize,
            "Dependent chart",
            litchi_odp::charts::Storage::InlineXml,
            litchi_odp::charts::Part::from_xml(DEPENDENT_CHART).unwrap(),
        )
        .unwrap();
    let dependent_source = dependent.commit().unwrap().snapshot().clone();
    let mut dependent_transfer = reopened.transaction().unwrap();
    dependent_transfer
        .embed_media("Pictures/chart.png", b"destination-pixels", "image/png")
        .unwrap();
    dependent_transfer
        .transfer_chart_from(
            &dependent_source,
            "Dependent chart",
            0usize,
            "Transferred dependent chart",
            litchi_odp::charts::Storage::InlineXml,
        )
        .unwrap();
    let dependent_commit = dependent_transfer.commit().unwrap();
    let dependent_package =
        OwnedPackage::from_bytes(dependent_commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        dependent_package.get_file("Pictures/chart.png").unwrap(),
        b"destination-pixels"
    );
    assert_eq!(
        dependent_package
            .get_file("Pictures/chart_litchi_1.png")
            .unwrap(),
        b"chart-pixels"
    );
    assert!(
        dependent_commit
            .snapshot()
            .to_presentation()
            .unwrap()
            .chart("Transferred dependent chart")
            .unwrap()
            .unwrap()
            .content_xml()
            .contains("Pictures/chart_litchi_1.png")
    );
}

#[test]
fn failed_checked_rdf_selector_leaves_an_exact_noop() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let graph_path = transaction
        .add_rdf_graph(Some("Metadata/deck.rdf"), &[literal("Kept")])
        .unwrap();
    let before = transaction.rdf_graphs().unwrap().to_vec();

    assert!(
        transaction
            .remove_rdf_triple(&graph_path, usize::MAX)
            .is_err()
    );
    assert_eq!(transaction.rdf_graphs().unwrap(), before);

    transaction.remove_rdf_graph(&graph_path).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
}

#[test]
fn signed_packages_are_refused_before_a_transaction_can_stage_changes() {
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file(
            "META-INF/documentsignatures.xml",
            br#"<dsig:document-signatures xmlns:dsig="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
        )
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        source.security_policy().unwrap(),
        edit::SecurityPolicy::SignedReadOnly
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::VerifySignatures)
            .unwrap(),
        edit::CryptoCapability::Refused(edit::CryptoRefusal::SignatureVerificationUnavailable)
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::ClearSignatures)
            .unwrap(),
        edit::CryptoCapability::Refused(edit::CryptoRefusal::SignatureRemovalUnavailable)
    );

    let Err(error) = source.transaction() else {
        panic!("signed package unexpectedly admitted an editing transaction");
    };
    assert!(error.to_string().contains("signed packages"));
}

#[test]
fn encrypted_package_entries_are_refused_before_staging() {
    const MIME: &str = "application/vnd.oasis.opendocument.presentation";
    const CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#;
    const MANIFEST: &[u8] = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.presentation"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="secret.bin" m:media-type="application/octet-stream" m:size="1"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIME.as_bytes()).unwrap();
    archive.write_deflated("content.xml", CONTENT).unwrap();
    archive.write_deflated("secret.bin", b"x").unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    let source = edit::Snapshot::from_bytes(archive.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        source.security_policy().unwrap(),
        edit::SecurityPolicy::EncryptedReadOnly
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::DecryptForEditing)
            .unwrap(),
        edit::CryptoCapability::Refused(edit::CryptoRefusal::UsePasswordOpening)
    );
    assert_eq!(
        source
            .crypto_capability(edit::CryptoOperation::Encrypt)
            .unwrap(),
        edit::CryptoCapability::Refused(edit::CryptoRefusal::EncryptionAuthoringUnavailable)
    );

    let Err(error) = source.transaction() else {
        panic!("encrypted package unexpectedly admitted an editing transaction");
    };
    assert!(error.to_string().contains("encrypted package entries"));
}
