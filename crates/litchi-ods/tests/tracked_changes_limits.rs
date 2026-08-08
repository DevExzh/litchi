use litchi_ods::tracked_changes::{
    Acceptance, Change, Changes, Dimension, Info, Insertion, Limits, Metadata, PositiveInteger,
    Snapshot,
};

const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:body><office:spreadsheet>"#;
const SUFFIX: &str = "</office:spreadsheet></office:body></office:document-content>";
const DATE: &str = "2026-08-08T00:00:00Z";

fn one_change(id: &str) -> Changes {
    Changes {
        enabled: false,
        changes: vec![Change::Insertion(Insertion {
            metadata: Metadata {
                id: id.to_string(),
                acceptance: Acceptance::Pending,
                rejecting_change_id: None,
                info: Info {
                    creator: Some("A".to_string()),
                    date: Some(DATE.to_string()),
                    comments: Vec::new(),
                },
                dependencies: Vec::new(),
                deletions: Vec::new(),
            },
            dimension: Dimension::Row,
            position: 0.into(),
            count: PositiveInteger::try_from(1usize).unwrap(),
            table: Some(0.into()),
        })],
    }
}

fn xml(id: &str) -> String {
    format!(
        r#"{PREFIX}<table:tracked-changes><table:insertion table:id="{id}" table:type="row" table:position="0"><office:change-info><dc:creator>A</dc:creator><dc:date>{DATE}</dc:date></office:change-info></table:insertion></table:tracked-changes>{SUFFIX}"#
    )
}

#[test]
fn public_defaults_accessors_and_builders_are_exact() {
    let limits = Limits::new();
    assert_eq!(limits, Limits::default());
    assert_eq!(limits.max_changes(), 1_000_000);
    assert_eq!(limits.max_nodes(), 1_000_000);
    assert_eq!(limits.max_value_bytes(), 65_536);
    assert_eq!(limits.max_aggregate_bytes(), 16 * 1_048_576);
    assert_eq!(limits.max_input_bytes(), 32 * 1_048_576);
    assert_eq!(limits.max_output_bytes(), 32 * 1_048_576);
    assert_eq!(limits.max_integer_digits(), 4_096);

    let custom = limits
        .with_max_changes(7)
        .with_max_nodes(11)
        .with_max_value_bytes(13)
        .with_max_aggregate_bytes(17)
        .with_max_input_bytes(19)
        .with_max_output_bytes(23)
        .with_max_integer_digits(29);
    assert_eq!(custom.max_changes(), 7);
    assert_eq!(custom.max_nodes(), 11);
    assert_eq!(custom.max_value_bytes(), 13);
    assert_eq!(custom.max_aggregate_bytes(), 17);
    assert_eq!(custom.max_input_bytes(), 19);
    assert_eq!(custom.max_output_bytes(), 23);
    assert_eq!(custom.max_integer_digits(), 29);
}

#[test]
fn input_and_output_byte_limits_are_inclusive_boundaries() {
    let source = format!("{PREFIX}<table:tracked-changes/>{SUFFIX}");
    assert!(
        Snapshot::parse_with_limits(
            source.clone(),
            Limits::new().with_max_input_bytes(source.len())
        )
        .is_ok()
    );
    assert!(
        Snapshot::parse_with_limits(
            source.clone(),
            Limits::new().with_max_input_bytes(source.len() - 1)
        )
        .is_err()
    );

    let baseline = Snapshot::parse(source.clone()).unwrap();
    let mut authored = baseline.transaction().unwrap();
    authored.set_tracking(Some(true)).unwrap();
    let target_len = authored.commit().unwrap().content_xml().len();

    let exact = Snapshot::parse_with_limits(
        source.clone(),
        Limits::new().with_max_output_bytes(target_len),
    )
    .unwrap();
    let mut transaction = exact.transaction().unwrap();
    transaction.set_tracking(Some(true)).unwrap();
    assert!(transaction.commit().is_ok());

    let too_small =
        Snapshot::parse_with_limits(source, Limits::new().with_max_output_bytes(target_len - 1))
            .unwrap();
    let mut transaction = too_small.transaction().unwrap();
    transaction.set_tracking(Some(true)).unwrap();
    assert!(transaction.commit().is_err());
}

#[test]
fn repeated_inserts_reject_before_node_or_aggregate_draft_growth() {
    let source = format!("{PREFIX}<table:tracked-changes/>{SUFFIX}");
    // The self-closing owner token retains 24 bytes; each insertion retains
    // 1 id + 1 creator + 20 date + three integer bytes, so two fit at 74.
    let limits = Limits::new()
        .with_max_changes(3)
        .with_max_nodes(12)
        .with_max_aggregate_bytes(74);
    let snapshot = Snapshot::parse_with_limits(source.clone(), limits).unwrap();
    let mut transaction = snapshot.transaction().unwrap();

    for (index, id) in ["a", "b"].into_iter().enumerate() {
        let change = one_change(id).changes.pop().unwrap();
        transaction.insert(index, change).unwrap();
    }
    let before_rejection = transaction.changes().cloned();
    assert!(transaction.is_changed());

    let rejected = one_change("c").changes.pop().unwrap();
    assert!(transaction.insert(2, rejected).is_err());
    assert_eq!(transaction.changes(), before_rejection.as_ref());
    assert!(transaction.is_changed());

    transaction.rollback().unwrap();
    assert!(!transaction.is_changed());
    assert!(transaction.changes().unwrap().changes.is_empty());
    let noop = transaction.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.content_xml(), source);
}

#[test]
fn retained_opaque_resources_reject_insert_before_candidate_or_draft_growth() {
    let owner =
        "<table:tracked-changes><future token='opaque'><child/></future></table:tracked-changes>";
    let source = format!("{PREFIX}{owner}{SUFFIX}");
    // Exact bounded scan: owner 1/23 plus opaque subtree 2/31.
    let limits = Limits::new().with_max_nodes(3).with_max_aggregate_bytes(54);
    let snapshot = Snapshot::parse_with_limits(source.clone(), limits).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    let before = transaction.changes().cloned();
    let donor = one_change("i").changes.pop().unwrap();

    assert!(transaction.insert(0, donor).is_err());
    assert_eq!(transaction.changes(), before.as_ref());
    assert!(!transaction.is_changed());
    let noop = transaction.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.content_xml(), source);
}

#[test]
fn max_changes_accepts_the_boundary_and_rejects_one_beyond() {
    let source = xml("one");
    assert!(Snapshot::parse_with_limits(source.clone(), Limits::new().with_max_changes(1)).is_ok());
    assert!(Snapshot::parse_with_limits(source, Limits::new().with_max_changes(0)).is_err());
}

#[test]
fn max_value_bytes_accepts_exact_utf8_bytes_and_rejects_one_less() {
    let id = "12345678901234567890";
    let source = xml(id);
    assert!(
        Snapshot::parse_with_limits(source.clone(), Limits::new().with_max_value_bytes(20)).is_ok()
    );
    assert!(Snapshot::parse_with_limits(source, Limits::new().with_max_value_bytes(19)).is_err());
}

#[test]
fn semantic_node_and_aggregate_limits_are_inclusive_boundaries() {
    let changes = one_change("i");
    assert!(
        changes
            .validate_with_limits(&Limits::new().with_max_nodes(5))
            .is_ok()
    );
    assert!(
        changes
            .validate_with_limits(&Limits::new().with_max_nodes(4))
            .is_err()
    );

    // id (1) + creator (1) + dateTime (20) + position/count/table (3) = 25.
    assert!(
        changes
            .validate_with_limits(&Limits::new().with_max_aggregate_bytes(25))
            .is_ok()
    );
    assert!(
        changes
            .validate_with_limits(&Limits::new().with_max_aggregate_bytes(24))
            .is_err()
    );
}

#[test]
fn zero_node_budget_rejects_even_a_present_empty_owner() {
    let source = format!("{PREFIX}<table:tracked-changes/>{SUFFIX}");
    assert!(Snapshot::parse_with_limits(source.clone(), Limits::new().with_max_nodes(4)).is_ok());
    assert!(Snapshot::parse_with_limits(source, Limits::new().with_max_nodes(0)).is_err());
}

#[test]
fn exact_semantic_node_budget_creates_and_reopens_an_empty_owner() {
    let absent = format!("{PREFIX}{SUFFIX}");
    let limits = Limits::new().with_max_nodes(1);
    let snapshot = Snapshot::parse_with_limits(absent, limits).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    transaction.replace_all(Some(Changes::default())).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let reopened = Snapshot::parse_with_limits(commit.source_arc(), limits).unwrap();
    assert!(reopened.changes().is_some());
    assert!(reopened.changes().unwrap().changes.is_empty());
}
