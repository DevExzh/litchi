use litchi_odf_common::package::raw_identical_members;
use litchi_odt::{
    Document,
    transaction::{DurablePatch, Position, SealedPatch, Snapshot},
};

mod support;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn content(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:scripts/><o:automatic-styles/><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
    )
}

fn snapshot(body: &str, marker: &[u8]) -> Snapshot {
    let xml = content(body);
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", xml.as_bytes()),
            ("meta.xml", b"<meta>opaque donor or target bytes</meta>"),
            ("Pictures/opaque.bin", marker),
        ],
    ))
    .unwrap()
    .snapshot()
    .unwrap()
}

fn try_snapshot(body: &str, _marker: &[u8]) -> Option<Snapshot> {
    let xml = content(body);
    Document::from_bytes(support::package(
        MIMETYPE,
        &[("content.xml", xml.as_bytes())],
    ))
    .ok()?
    .snapshot()
    .ok()
}

fn try_custom_snapshot(xml: &str, marker: &[u8]) -> Option<Snapshot> {
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", xml.as_bytes()),
            ("meta.xml", b"opaque custom bytes"),
            ("Pictures/opaque.bin", marker),
        ],
    ))
    .ok()?
    .snapshot()
    .ok()
}

fn texts(snapshot: &Snapshot) -> Vec<String> {
    snapshot
        .document()
        .unwrap()
        .paragraphs()
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect()
}

#[test]
fn transfers_first_middle_last_and_only_lexical_fragments() {
    let donor = snapshot(
        "\n<t:p>first&amp;lexical</t:p>\n  <t:p><![CDATA[middle&lt;raw&gt;]]></t:p>\n<t:p>last</t:p>\n",
        b"donor",
    );
    let destination = snapshot("<t:p>before</t:p><t:p>after</t:p>", b"target");
    let donor_before = donor.as_bytes().to_vec();

    for (source, target, expected) in [
        (0, 0, vec!["first&lexical", "before", "after"]),
        (1, 1, vec!["before", "middle&lt;raw&gt;", "after"]),
        (2, 2, vec!["before", "after", "last"]),
    ] {
        let plan = destination
            .plan_plain_paragraph_transfer_from(
                &donor,
                Position::new(source),
                Position::new(target),
            )
            .unwrap();
        assert_eq!(plan.paragraph_count(), 1);
        assert_eq!(plan.source_positions(), &[source]);
        let mut edit = destination.edit();
        plan.apply(&mut edit).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(texts(commit.snapshot()), expected);
        assert_eq!(donor.as_bytes(), donor_before);
        let content = commit
            .snapshot()
            .document()
            .unwrap()
            .get_file("content.xml")
            .unwrap();
        let lexical = match source {
            0 => b"<t:p>first&amp;lexical</t:p>".as_slice(),
            1 => b"<t:p><![CDATA[middle&lt;raw&gt;]]></t:p>".as_slice(),
            _ => b"<t:p>last</t:p>".as_slice(),
        };
        assert!(
            content
                .windows(lexical.len())
                .any(|window| window == lexical)
        );
        let identical =
            raw_identical_members(destination.as_bytes(), commit.snapshot().as_bytes()).unwrap();
        assert!(!identical.contains("content.xml"));
        for path in [
            "mimetype",
            "meta.xml",
            "Pictures/opaque.bin",
            "META-INF/manifest.xml",
        ] {
            assert!(identical.contains(path), "{path}");
        }
    }

    let only_donor = snapshot("<t:p>only</t:p>", b"only-donor");
    let only_target = snapshot("<t:p>target</t:p>", b"only-target");
    let plan = only_target
        .plan_plain_paragraph_transfer_from(&only_donor, Position::new(0), Position::new(1))
        .unwrap();
    let mut edit = only_target.edit();
    plan.apply(&mut edit).unwrap();
    assert_eq!(texts(edit.commit().unwrap().snapshot()), ["target", "only"]);
}

#[test]
fn batch_transfer_is_source_bound_durable_and_reversible() {
    let donor = snapshot("<t:p>A</t:p>\n<t:p>B</t:p>\n<t:p>C</t:p>", b"donor-batch");
    let destination = snapshot("<t:p>X</t:p><t:p>Y</t:p>", b"target-batch");
    let plan = destination
        .plan_plain_paragraphs_transfer_from(
            &donor,
            &[Position::new(0), Position::new(2)],
            Position::new(1),
        )
        .unwrap();
    assert_eq!(plan.paragraph_count(), 2);
    assert_eq!(plan.donor_fingerprint().len(), 64);
    assert_eq!(plan.destination_fingerprint().len(), 64);
    assert_eq!(plan.fragment_digest().len(), 64);

    let mut edit = destination.edit();
    edit.apply_plain_paragraph_transfer(&plan).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(texts(commit.snapshot()), ["X", "A", "C", "Y"]);
    assert_eq!(commit.results().len(), 1);

    let patch = commit.patch();
    assert_eq!(
        patch.apply(&destination).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        patch.inverse().apply(commit.snapshot()).unwrap().as_bytes(),
        destination.as_bytes()
    );
    assert!(patch.apply(commit.snapshot()).is_err());

    let durable = patch.durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    assert!(!String::from_utf8_lossy(&wire).contains("\"donor\""));
    assert_eq!(wire, durable.to_deterministic_json().unwrap());
    let reopened = DurablePatch::from_deterministic_json(&wire).unwrap();
    assert_eq!(
        reopened.apply(&destination).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        reopened
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );
    assert!(reopened.apply(commit.snapshot()).is_err());
}

#[test]
fn durable_transfer_rejects_repeated_blob_refs_before_copying_over_limit() {
    const FRAGMENT_BYTES: usize = 1024 * 1024;
    const EXACT_FRAGMENT_COUNT: usize = 8;
    let open = "<t:p>";
    let close = "</t:p>";
    let payload = "x".repeat(FRAGMENT_BYTES - open.len() - close.len());
    let mut body = String::with_capacity(FRAGMENT_BYTES * EXACT_FRAGMENT_COUNT);
    for _ in 0..EXACT_FRAGMENT_COUNT {
        body.push_str(open);
        body.push_str(&payload);
        body.push_str(close);
    }
    let donor = snapshot(&body, b"durable-large-donor");
    let destination = snapshot("<t:p>target</t:p>", b"durable-large-target");
    let source_positions: Vec<_> = (0..EXACT_FRAGMENT_COUNT).map(Position::new).collect();
    let plan = destination
        .plan_plain_paragraphs_transfer_from(&donor, &source_positions, Position::new(0))
        .unwrap();
    let mut edit = destination.edit();
    plan.apply(&mut edit).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let exact_wire = durable.to_deterministic_json().unwrap();
    assert!(DurablePatch::from_deterministic_json(&exact_wire).is_ok());

    let mut empty_reversible: serde_json::Value = serde_json::from_slice(&exact_wire).unwrap();
    let forward = empty_reversible["operations"][0]["forward"]
        .as_object_mut()
        .unwrap();
    forward.insert(
        "target".to_string(),
        serde_json::Value::String("/body/paragraphs/256/transfer".to_string()),
    );
    let value = forward["value"].as_object_mut().unwrap();
    value.insert("target_position".to_string(), serde_json::Value::from(256));
    value.insert(
        "source_positions".to_string(),
        serde_json::Value::Array(Vec::new()),
    );
    value.insert(
        "fragments".to_string(),
        serde_json::Value::Array(Vec::new()),
    );
    let empty_reversible_wire = serde_json::to_vec(&empty_reversible).unwrap();
    assert!(DurablePatch::from_deterministic_json(&empty_reversible_wire).is_err());

    let sealed = durable.clone().seal();
    let sealed_wire = sealed.to_deterministic_json().unwrap();
    let mut empty_sealed: serde_json::Value = serde_json::from_slice(&sealed_wire).unwrap();
    let forward = empty_sealed["operations"][0].as_object_mut().unwrap();
    forward.insert(
        "target".to_string(),
        serde_json::Value::String("/body/paragraphs/256/transfer".to_string()),
    );
    let value = forward["value"].as_object_mut().unwrap();
    value.insert("target_position".to_string(), serde_json::Value::from(256));
    value.insert(
        "source_positions".to_string(),
        serde_json::Value::Array(Vec::new()),
    );
    value.insert(
        "fragments".to_string(),
        serde_json::Value::Array(Vec::new()),
    );
    let empty_sealed_wire = serde_json::to_vec(&empty_sealed).unwrap();
    assert!(SealedPatch::from_deterministic_json(&empty_sealed_wire).is_err());

    let mut repeated: serde_json::Value = serde_json::from_slice(&exact_wire).unwrap();
    let forward = repeated["operations"][0]["forward"]["value"]
        .as_object_mut()
        .unwrap();
    let first_fragment = forward["fragments"][0].clone();
    forward.insert(
        "fragments".to_string(),
        serde_json::Value::Array(vec![first_fragment; 256]),
    );
    forward.insert(
        "source_positions".to_string(),
        serde_json::Value::Array((0..256).map(serde_json::Value::from).collect()),
    );
    let repeated_wire = serde_json::to_vec(&repeated).unwrap();
    assert!(DurablePatch::from_deterministic_json(&repeated_wire).is_err());

    let mut cross_operation: serde_json::Value = serde_json::from_slice(&exact_wire).unwrap();
    let pair = cross_operation["operations"][0].clone();
    cross_operation["operations"] = serde_json::Value::Array(vec![pair; 1_024]);
    let cross_operation_wire = serde_json::to_vec(&cross_operation).unwrap();
    assert!(DurablePatch::from_deterministic_json(&cross_operation_wire).is_err());
}

#[test]
fn empty_transfer_is_exact_noop_and_target_or_donor_is_bound() {
    let donor = snapshot("<t:p>donor</t:p>", b"donor");
    let target = snapshot("<t:p>target</t:p>", b"target");
    let plan = target
        .plan_plain_paragraphs_transfer_from(&donor, &[], Position::new(0))
        .unwrap();
    let mut edit = target.edit();
    plan.apply(&mut edit).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().as_bytes(), target.as_bytes());
    assert!(commit.results().is_empty());

    let foreign = snapshot("<t:p>foreign</t:p>", b"foreign");
    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
        .unwrap();
    let mut foreign_edit = foreign.edit();
    assert!(plan.apply(&mut foreign_edit).is_err());
}

#[test]
fn transfer_joins_as_a_deterministic_isolation_conflict() {
    let donor = snapshot("<t:p>donor</t:p>", b"join-donor");
    let target = snapshot("<t:p>target</t:p>", b"join-target");
    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
        .unwrap();
    let mut transfer = target.edit();
    plan.apply(&mut transfer).unwrap();
    let mut replacement = target.edit();
    replacement
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();

    let mut joined = target.joined_edit();
    joined.join("transfer", transfer).unwrap();
    assert!(joined.join("replace", replacement).is_err());

    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
        .unwrap();
    let mut transfer = target.edit();
    plan.apply(&mut transfer).unwrap();
    let mut joined = target.joined_edit();
    let mut replacement = target.edit();
    replacement
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();
    joined.join("replace", replacement).unwrap();
    assert!(joined.join("transfer", transfer).is_err());

    let mut mixed = target.edit();
    plan.apply(&mut mixed).unwrap();
    mixed
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();
    let mut joined = target.joined_edit();
    assert!(joined.join("mixed-single-edit", mixed).is_err());

    let mut duplicated = target.edit();
    plan.apply(&mut duplicated).unwrap();
    plan.apply(&mut duplicated).unwrap();
    let mut joined = target.joined_edit();
    assert!(joined.join("duplicate-transfers", duplicated).is_err());

    // Even a paragraph-count-preserving content edit can regenerate
    // content.xml and alter the transfer's lexical fragment spelling.
    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(1))
        .unwrap();
    let mut transfer = target.edit();
    plan.apply(&mut transfer).unwrap();
    let mut replacement = target.edit();
    replacement
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();
    let mut joined = target.joined_edit();
    joined.join("replace", replacement).unwrap();
    assert!(joined.join("transfer", transfer).is_err());

    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(1))
        .unwrap();
    let mut transfer = target.edit();
    plan.apply(&mut transfer).unwrap();
    let mut replacement = target.edit();
    replacement
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();
    let mut joined = target.joined_edit();
    joined.join("transfer", transfer).unwrap();
    assert!(joined.join("replace", replacement).is_err());

    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(1))
        .unwrap();
    let mut transfer = target.edit();
    plan.apply(&mut transfer).unwrap();
    let mut shifted_replacement = target.edit();
    shifted_replacement
        .replace_paragraph(Position::new(1), "changed")
        .unwrap();
    let mut joined = target.joined_edit();
    joined.join("transfer", transfer).unwrap();
    assert!(joined.join("replace", shifted_replacement).is_err());

    let plan = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(1))
        .unwrap();
    let mut transfer = target.edit();
    plan.apply(&mut transfer).unwrap();
    let mut shifted_replacement = target.edit();
    shifted_replacement
        .replace_paragraph(Position::new(1), "changed")
        .unwrap();
    let mut joined = target.joined_edit();
    joined.join("replace", shifted_replacement).unwrap();
    assert!(joined.join("transfer", transfer).is_err());
}

#[test]
fn mixed_transfer_refuses_lexical_regeneration_for_cdata_and_prefixes() {
    let donor = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="{OFFICE}" xmlns:x="{TEXT}" o:version="1.3"><o:scripts/><o:automatic-styles/><o:body><o:text><x:p><![CDATA[middle&lt;raw&gt;]]></x:p><x:p>second&amp;raw</x:p></o:text></o:body></o:document-content>"#
        ),
        b"lexical-prefix-donor",
    )
    .expect("custom donor should be a valid ODT fixture");
    let target = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:x="{TEXT}" o:version="1.3"><o:scripts/><o:automatic-styles/><o:body><o:text><t:p>target</t:p></o:text></o:body></o:document-content>"#
        ),
        b"lexical-prefix-target",
    )
    .expect("custom target should be a valid ODT fixture");
    let donor_before = donor.as_bytes().to_vec();

    let one = target
        .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
        .unwrap();
    let mut standalone = target.edit();
    one.apply(&mut standalone).unwrap();
    let standalone = standalone.commit().unwrap();
    assert_eq!(donor.as_bytes(), donor_before.as_slice());
    let standalone_content = standalone
        .snapshot()
        .document()
        .unwrap()
        .get_file("content.xml")
        .unwrap();
    let cdata = b"<x:p><![CDATA[middle&lt;raw&gt;]]></x:p>";
    assert!(
        standalone_content
            .windows(cdata.len())
            .any(|window| window == cdata)
    );

    let batch = target
        .plan_plain_paragraphs_transfer_from(
            &donor,
            &[Position::new(0), Position::new(1)],
            Position::new(0),
        )
        .unwrap();
    let mut batched = target.edit();
    batch.apply(&mut batched).unwrap();
    let batched = batched.commit().unwrap();
    let batched_content = batched
        .snapshot()
        .document()
        .unwrap()
        .get_file("content.xml")
        .unwrap();
    for lexical in [cdata.as_slice(), b"<x:p>second&amp;raw</x:p>".as_slice()] {
        assert!(
            batched_content
                .windows(lexical.len())
                .any(|window| window == lexical)
        );
    }

    let mut mixed = target.edit();
    one.apply(&mut mixed).unwrap();
    mixed
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();
    assert!(mixed.commit().is_err());

    let mut transfer = target.edit();
    one.apply(&mut transfer).unwrap();
    let mut replacement = target.edit();
    replacement
        .replace_paragraph(Position::new(0), "changed")
        .unwrap();
    let mut joined = target.joined_edit();
    joined.join("transfer", transfer).unwrap();
    assert!(joined.join("replace", replacement).is_err());
}

#[test]
fn refuses_cross_document_text_namespace_rebinding() {
    let target = snapshot("<t:p>target</t:p>", b"namespace-target");
    let donor_other_prefix = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:x="{TEXT}"><o:body><o:text><x:p>donor</x:p></o:text></o:body></o:document-content>"#
        ),
        b"namespace-other",
    );
    if let Some(donor) = donor_other_prefix {
        assert!(
            target
                .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                .is_err()
        );
    }

    let donor_default_prefix = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><o:text xmlns="{TEXT}"><p>donor</p></o:text></o:body></o:document-content>"#
        ),
        b"namespace-default",
    );
    if let Some(donor) = donor_default_prefix {
        assert!(
            target
                .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                .is_err()
        );
    }

    let target_default_prefix = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}"><o:body><o:text xmlns="{TEXT}"><p>target</p></o:text></o:body></o:document-content>"#
        ),
        b"namespace-target-default",
    );
    if let Some(target) = target_default_prefix {
        let donor = snapshot("<t:p>donor</t:p>", b"namespace-donor");
        assert!(
            target
                .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                .is_err()
        );
    }

    let unbound = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><o:text><p>donor</p></o:text></o:body></o:document-content>"#
        ),
        b"namespace-unbound",
    );
    if let Some(donor) = unbound {
        assert!(
            target
                .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                .is_err()
        );
    }
}

#[test]
fn refuses_unknown_preamble_descendants_and_nonwhitespace_outside_text() {
    let target = snapshot("<t:p>target</t:p>", b"scanner-target");
    for (xml, marker) in [
        (
            format!(
                r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:x="urn:foreign"><o:automatic-styles><x:unknown/></o:automatic-styles><o:body><o:text><t:p>donor</t:p></o:text></o:body></o:document-content>"#
            ),
            b"scanner-preamble".as_slice(),
        ),
        (
            format!(
                r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:automatic-styles bad="1"/><o:body><o:text><t:p>donor</t:p></o:text></o:body></o:document-content>"#
            ),
            b"scanner-style-attrs".as_slice(),
        ),
        (
            format!(
                r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:font-face-decls bad="1"/><o:body><o:text><t:p>donor</t:p></o:text></o:body></o:document-content>"#
            ),
            b"scanner-font-attrs".as_slice(),
        ),
    ] {
        if let Some(donor) = try_custom_snapshot(&xml, marker) {
            assert!(
                target
                    .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                    .is_err()
            );
        }
    }

    let nonwhitespace = try_custom_snapshot(
        &format!(
            r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body>unsafe<o:text><t:p>donor</t:p></o:text></o:body></o:document-content>"#
        ),
        b"scanner-text",
    );
    if let Some(donor) = nonwhitespace {
        assert!(
            target
                .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                .is_err()
        );
    }

    for (outside_text, marker) in [
        ("<!-- comment -->", b"scanner-comment".as_slice()),
        ("<?processing instruction?>", b"scanner-pi".as_slice()),
        ("<?xml version=\"1.0\"?>", b"scanner-decl".as_slice()),
    ] {
        let xml = format!(
            r#"<?xml version="1.0"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body>{outside_text}<o:text><t:p>donor</t:p></o:text></o:body></o:document-content>"#
        );
        if let Some(donor) = try_custom_snapshot(&xml, marker) {
            assert!(
                target
                    .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                    .is_err()
            );
        }
    }
}

#[test]
fn refuses_dependency_bearing_fragments_and_security_envelopes() {
    for body in [
        r#"<t:p t:style-name="P">styled</t:p>"#,
        r#"<t:p><t:span>nested</t:span></t:p>"#,
        r#"<t:p><text:a xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">link</text:a></t:p>"#,
        r#"<t:section><t:p>section</t:p></t:section>"#,
        r#"<table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
        r#"<unknown xmlns="urn:unknown"/>"#,
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>"#,
        r#"<t:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x">mce</t:p>"#,
    ] {
        let Some(donor) = try_snapshot(body, b"unsafe") else {
            continue;
        };
        let target = snapshot("<t:p>target</t:p>", b"safe-target");
        assert!(
            target
                .plan_plain_paragraph_transfer_from(&donor, Position::new(0), Position::new(0))
                .is_err(),
            "{body}"
        );
    }

    let target = snapshot("<t:p>target</t:p>", b"safe-target");
    let Some(scripted) = try_snapshot(
        r#"<o:scripts><o:script/></o:scripts><t:p>scripted</t:p>"#,
        b"scripted",
    ) else {
        return;
    };
    assert!(
        target
            .plan_plain_paragraph_transfer_from(&scripted, Position::new(0), Position::new(0))
            .is_err()
    );
}
