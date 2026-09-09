#![allow(
    clippy::expect_used,
    clippy::pedantic,
    clippy::unwrap_used,
    reason = "codec review cases use compact synthetic BIFF12 streams and fail-fast assertions"
)]

//! Adversarial Volatile Dependencies codec cases.
//!
//! The wire cases follow the local [MS-XLSB] descriptions of the
//! `BrtBeginVol*`, `BrtVol*`, and `BrtEndVol*` records in section 2.4.299-
//! 2.4.302 and 2.4.859-2.4.864.  They stay below the package facade so that
//! malformed record payloads and parser quota order are observable directly.

use super::codec;
use super::model::{
    CachedValue, CellReference, Dependencies, DependencyKind, ErrorCode, MainTopic, ReadLimits,
    Topic, VolatileType,
};
use crate::package::error::Error;
use crate::raw::{Kind, Writer, kind};

type Record = (Kind, Vec<u8>);

fn encode(records: &[Record]) -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    for (record_kind, payload) in records {
        writer
            .write_record(*record_kind, payload)
            .expect("synthetic record writes");
    }
    bytes
}

fn wide(value: &str) -> Vec<u8> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let mut payload = Vec::with_capacity(4 + units.len() * 2);
    payload.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for unit in units {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    payload
}

fn wide_units(units: &[u16]) -> Vec<u8> {
    let mut payload = Vec::with_capacity(4 + units.len() * 2);
    payload.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for unit in units {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    payload
}

fn reference(row: i32, column: i32, sheet_index: u32) -> Vec<u8> {
    let mut payload = Vec::with_capacity(12);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&sheet_index.to_le_bytes());
    payload
}

fn value_records(value_kind: Kind, value_payload: Vec<u8>) -> Vec<Record> {
    vec![
        (kind::BEGIN_VOL_DEPS, Vec::new()),
        (kind::BEGIN_VOL_TYPE, 0_u32.to_le_bytes().to_vec()),
        (kind::BEGIN_VOL_MAIN, wide("server")),
        (kind::BEGIN_VOL_TOPIC, Vec::new()),
        (value_kind, value_payload),
        (kind::END_VOL_TOPIC, Vec::new()),
        (kind::END_VOL_MAIN, Vec::new()),
        (kind::END_VOL_TYPE, Vec::new()),
        (kind::END_VOL_DEPS, Vec::new()),
    ]
}

fn valid_records() -> Vec<Record> {
    value_records(kind::VOL_BOOL, vec![1])
}

fn dependencies_for(value: CachedValue) -> Dependencies {
    Dependencies::new(vec![VolatileType {
        kind: DependencyKind::Rtd,
        mains: vec![MainTopic {
            first: "server".to_owned(),
            topics: vec![Topic {
                subtopics: Vec::new(),
                value,
                references: vec![CellReference::new(0, 0, 0).expect("valid reference")],
            }],
        }],
    }])
}

fn replace_payload(records: &mut [Record], target: Kind, payload: Vec<u8>) {
    let record = records
        .iter_mut()
        .find(|(record_kind, _)| *record_kind == target)
        .expect("target record in synthetic stream");
    record.1 = payload;
}

fn append_main(records: &mut Vec<Record>, name: &str, subtopic: &str, row: i32) {
    records.push((kind::BEGIN_VOL_MAIN, wide(name)));
    records.push((kind::BEGIN_VOL_TOPIC, Vec::new()));
    records.push((kind::VOL_SUBTOPIC, wide(subtopic)));
    records.push((kind::VOL_BOOL, vec![1]));
    records.push((kind::VOL_REF, reference(row, row, 0)));
    records.push((kind::END_VOL_TOPIC, Vec::new()));
    records.push((kind::END_VOL_MAIN, Vec::new()));
}

fn aggregate_records() -> Vec<Record> {
    let mut records = vec![
        (kind::BEGIN_VOL_DEPS, Vec::new()),
        (kind::BEGIN_VOL_TYPE, 0_u32.to_le_bytes().to_vec()),
    ];
    append_main(&mut records, "a", "x", 0);
    append_main(&mut records, "b", "y", 1);
    records.extend([
        (kind::END_VOL_TYPE, Vec::new()),
        (kind::END_VOL_DEPS, Vec::new()),
    ]);
    records
}

fn aggregate_dependencies() -> Dependencies {
    Dependencies::new(vec![VolatileType {
        kind: DependencyKind::Rtd,
        mains: vec![
            MainTopic {
                first: "a".to_owned(),
                topics: vec![Topic {
                    subtopics: vec!["x".to_owned()],
                    value: CachedValue::Bool(true),
                    references: vec![CellReference::new(0, 0, 0).expect("reference")],
                }],
            },
            MainTopic {
                first: "b".to_owned(),
                topics: vec![Topic {
                    subtopics: vec!["y".to_owned()],
                    value: CachedValue::Bool(true),
                    references: vec![CellReference::new(1, 1, 0).expect("reference")],
                }],
            },
        ],
    }])
}

fn two_type_records() -> Vec<Record> {
    let mut records = vec![(kind::BEGIN_VOL_DEPS, Vec::new())];
    for (flags, name) in [(0_u32, "rtd"), (1_u32, "cube")] {
        records.push((kind::BEGIN_VOL_TYPE, flags.to_le_bytes().to_vec()));
        records.push((kind::BEGIN_VOL_MAIN, wide(name)));
        records.push((kind::END_VOL_MAIN, Vec::new()));
        records.push((kind::END_VOL_TYPE, Vec::new()));
    }
    records.push((kind::END_VOL_DEPS, Vec::new()));
    records
}

fn two_type_dependencies() -> Dependencies {
    Dependencies::new(vec![
        VolatileType {
            kind: DependencyKind::Rtd,
            mains: vec![MainTopic {
                first: "rtd".to_owned(),
                topics: Vec::new(),
            }],
        },
        VolatileType {
            kind: DependencyKind::Cube,
            mains: vec![MainTopic {
                first: "cube".to_owned(),
                topics: Vec::new(),
            }],
        },
    ])
}

fn assert_wire_or_format_error(result: Result<codec::Parsed, Error>, label: &str) {
    match result {
        Err(Error::Wire(_) | Error::InvalidFormat(_)) => {},
        Err(error) => panic!("{label} returned unexpected error: {error:?}"),
        Ok(_) => panic!("{label} unexpectedly parsed"),
    }
}

fn assert_limit(error: Error, resource: &'static str, actual: usize, maximum: usize) {
    if let Error::LimitExceeded {
        resource: got_resource,
        actual: got_actual,
        maximum: got_maximum,
    } = error
    {
        assert_eq!(got_resource, resource);
        assert_eq!(got_actual, actual);
        assert_eq!(got_maximum, maximum);
    } else {
        panic!("expected {resource} limit error, got {error:?}");
    }
}

#[test]
fn malformed_known_payload_lengths_and_trailing_bytes_are_rejected() {
    struct Case {
        name: &'static str,
        value_kind: Kind,
        value_payload: Vec<u8>,
        target: Kind,
        payload: Vec<u8>,
    }

    let cases = vec![
        Case {
            name: "begin dependencies payload",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::BEGIN_VOL_DEPS,
            payload: vec![1],
        },
        Case {
            name: "begin type truncated",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::BEGIN_VOL_TYPE,
            payload: vec![0, 0, 0],
        },
        Case {
            name: "begin type trailing",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::BEGIN_VOL_TYPE,
            payload: vec![0, 0, 0, 0, 9],
        },
        Case {
            name: "begin main truncated string",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::BEGIN_VOL_MAIN,
            payload: 1_u32.to_le_bytes().to_vec(),
        },
        Case {
            name: "begin topic payload",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::BEGIN_VOL_TOPIC,
            payload: vec![1],
        },
        Case {
            name: "subtopic truncated string",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::VOL_SUBTOPIC,
            payload: 1_u32.to_le_bytes().to_vec(),
        },
        Case {
            name: "reference truncated",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::VOL_REF,
            payload: vec![0; 11],
        },
        Case {
            name: "number trailing",
            value_kind: kind::VOL_NUM,
            value_payload: 1.0_f64.to_le_bytes().to_vec(),
            target: kind::VOL_NUM,
            payload: vec![0; 9],
        },
        Case {
            name: "error trailing",
            value_kind: kind::VOL_ERR,
            value_payload: vec![ErrorCode::Null as u8],
            target: kind::VOL_ERR,
            payload: vec![ErrorCode::Null as u8, 0],
        },
        Case {
            name: "string trailing",
            value_kind: kind::VOL_STR,
            value_payload: wide("x"),
            target: kind::VOL_STR,
            payload: [wide("x"), vec![0]].concat(),
        },
        Case {
            name: "boolean trailing",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::VOL_BOOL,
            payload: vec![1, 0],
        },
        Case {
            name: "end topic payload",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::END_VOL_TOPIC,
            payload: vec![1],
        },
        Case {
            name: "end main payload",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::END_VOL_MAIN,
            payload: vec![1],
        },
        Case {
            name: "end type payload",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::END_VOL_TYPE,
            payload: vec![1],
        },
        Case {
            name: "end dependencies payload",
            value_kind: kind::VOL_BOOL,
            value_payload: vec![1],
            target: kind::END_VOL_DEPS,
            payload: vec![1],
        },
    ];

    for case in cases {
        let mut records = value_records(case.value_kind, case.value_payload);
        if case.target == kind::VOL_SUBTOPIC {
            records.insert(4, (kind::VOL_SUBTOPIC, case.payload.clone()));
        } else if case.target == kind::VOL_REF {
            records.insert(5, (kind::VOL_REF, case.payload.clone()));
        } else {
            replace_payload(&mut records, case.target, case.payload);
        }
        assert_wire_or_format_error(
            codec::read(&encode(&records), ReadLimits::DEFAULT, 1),
            case.name,
        );
    }
}

#[test]
fn cached_scalar_domains_round_trip_and_reject_invalid_wire_values() {
    let values = vec![
        CachedValue::Number(0.0),
        CachedValue::Number(-123.5),
        CachedValue::Number(f64::MAX),
        CachedValue::Error(ErrorCode::Null),
        CachedValue::Error(ErrorCode::Div0),
        CachedValue::Error(ErrorCode::Value),
        CachedValue::Error(ErrorCode::Ref),
        CachedValue::Error(ErrorCode::Name),
        CachedValue::Error(ErrorCode::Num),
        CachedValue::Error(ErrorCode::Na),
        CachedValue::Error(ErrorCode::GettingData),
        CachedValue::String("cached".to_owned()),
        CachedValue::Bool(false),
        CachedValue::Bool(true),
    ];
    for value in values {
        let dependencies = dependencies_for(value);
        let bytes = codec::write(&dependencies, ReadLimits::DEFAULT, 1).expect("valid scalar");
        assert_eq!(
            codec::read(&bytes, ReadLimits::DEFAULT, 1)
                .expect("read valid scalar")
                .dependencies,
            dependencies
        );
    }

    for invalid_bool in [2_u8, 0xff] {
        let result = codec::read(
            &encode(&value_records(kind::VOL_BOOL, vec![invalid_bool])),
            ReadLimits::DEFAULT,
            1,
        );
        assert!(
            matches!(result, Err(Error::Wire(_))),
            "invalid bool {invalid_bool}"
        );
    }
    for invalid_error in [1_u8, 2, 0xff] {
        let result = codec::read(
            &encode(&value_records(kind::VOL_ERR, vec![invalid_error])),
            ReadLimits::DEFAULT,
            1,
        );
        assert!(
            matches!(result, Err(Error::InvalidFormat(_))),
            "invalid error {invalid_error}"
        );
    }
    for invalid_number in [
        f64::NAN,
        f64::INFINITY,
        f64::NEG_INFINITY,
        -0.0,
        f64::from_bits(1),
    ] {
        let result = codec::read(
            &encode(&value_records(
                kind::VOL_NUM,
                invalid_number.to_le_bytes().to_vec(),
            )),
            ReadLimits::DEFAULT,
            1,
        );
        assert!(
            matches!(result, Err(Error::InvalidFormat(_))),
            "invalid Xnum 0x{:016x}",
            invalid_number.to_bits()
        );
        assert!(
            codec::write(
                &dependencies_for(CachedValue::Number(invalid_number)),
                ReadLimits::DEFAULT,
                1
            )
            .is_err()
        );
    }
}

#[test]
fn strict_utf16_and_length_sentinels_are_not_accepted() {
    let cases = [
        ("unpaired high surrogate", wide_units(&[0xd800]), false),
        ("unpaired low surrogate", wide_units(&[0xdc00]), false),
        ("odd declared payload", vec![1, 0, 0, 0, 0x41], false),
        (
            "trailing UTF16 payload",
            [wide("x"), vec![0]].concat(),
            false,
        ),
        ("length sentinel", u32::MAX.to_le_bytes().to_vec(), true),
    ];
    for (name, payload, is_limit) in cases {
        let mut records = valid_records();
        replace_payload(&mut records, kind::BEGIN_VOL_MAIN, payload);
        let result = codec::read(&encode(&records), ReadLimits::DEFAULT, 1);
        if is_limit {
            assert!(matches!(result, Err(Error::LimitExceeded { .. })), "{name}");
        } else {
            assert_wire_or_format_error(result, name);
        }
    }
}

#[test]
fn nesting_duplicate_values_and_type_kinds_are_rejected() {
    let mut cases = Vec::new();

    let mut duplicate_value = valid_records();
    let end_topic = duplicate_value
        .iter()
        .position(|(record_kind, _)| *record_kind == kind::END_VOL_TOPIC)
        .expect("end topic");
    duplicate_value.insert(end_topic, (kind::VOL_BOOL, vec![1]));
    cases.push(("duplicate cached value", duplicate_value));

    let mut nested_topic = valid_records();
    let begin_topic = nested_topic
        .iter()
        .position(|(record_kind, _)| *record_kind == kind::BEGIN_VOL_TOPIC)
        .expect("begin topic");
    nested_topic.insert(begin_topic + 1, (kind::BEGIN_VOL_TOPIC, Vec::new()));
    cases.push(("nested topic", nested_topic));

    let mut ref_before_value = valid_records();
    let value = ref_before_value
        .iter()
        .position(|(record_kind, _)| *record_kind == kind::VOL_BOOL)
        .expect("value");
    ref_before_value.insert(value, (kind::VOL_REF, reference(0, 0, 0)));
    cases.push(("reference before cached value", ref_before_value));

    let mut no_value = valid_records();
    no_value.retain(|(record_kind, _)| *record_kind != kind::VOL_BOOL);
    cases.push(("topic without cached value", no_value));

    let mut duplicate_type = two_type_records();
    duplicate_type[5].1 = 0_u32.to_le_bytes().to_vec();
    cases.push(("duplicate dependency type", duplicate_type));

    cases.push((
        "main outside type",
        vec![
            (kind::BEGIN_VOL_DEPS, Vec::new()),
            (kind::BEGIN_VOL_MAIN, wide("server")),
            (kind::END_VOL_DEPS, Vec::new()),
        ],
    ));
    cases.push((
        "topic outside main",
        vec![
            (kind::BEGIN_VOL_DEPS, Vec::new()),
            (kind::BEGIN_VOL_TYPE, 0_u32.to_le_bytes().to_vec()),
            (kind::BEGIN_VOL_TOPIC, Vec::new()),
            (kind::END_VOL_DEPS, Vec::new()),
        ],
    ));

    let mut record_after_end = valid_records();
    record_after_end.push((kind::BEGIN_VOL_TYPE, 0_u32.to_le_bytes().to_vec()));
    cases.push(("known record after end", record_after_end));

    for (name, records) in cases {
        assert_wire_or_format_error(codec::read(&encode(&records), ReadLimits::DEFAULT, 1), name);
    }
}

fn set_max_types(limits: &mut ReadLimits, value: usize) {
    limits.max_types = value;
}

fn set_max_mains(limits: &mut ReadLimits, value: usize) {
    limits.max_mains = value;
}

fn set_max_topics(limits: &mut ReadLimits, value: usize) {
    limits.max_topics = value;
}

fn set_max_subtopics(limits: &mut ReadLimits, value: usize) {
    limits.max_subtopics = value;
}

fn set_max_references(limits: &mut ReadLimits, value: usize) {
    limits.max_references = value;
}

fn set_max_string_units(limits: &mut ReadLimits, value: usize) {
    limits.max_string_units = value;
}

fn set_max_total_string_units(limits: &mut ReadLimits, value: usize) {
    limits.max_total_string_units = value;
}

#[test]
fn aggregate_quotas_are_global_with_exact_boundary_and_plus_one_refusal() {
    struct Case {
        name: &'static str,
        boundary: usize,
        resource: &'static str,
        set_limit: fn(&mut ReadLimits, usize),
        records: fn() -> Vec<Record>,
        dependencies: fn() -> Dependencies,
    }

    let cases = [
        Case {
            name: "types",
            boundary: 2,
            resource: "Volatile Dependencies type collections",
            set_limit: set_max_types,
            records: two_type_records,
            dependencies: two_type_dependencies,
        },
        Case {
            name: "mains",
            boundary: 2,
            resource: "Volatile main collections",
            set_limit: set_max_mains,
            records: aggregate_records,
            dependencies: aggregate_dependencies,
        },
        Case {
            name: "topics",
            boundary: 2,
            resource: "Volatile topic collections",
            set_limit: set_max_topics,
            records: aggregate_records,
            dependencies: aggregate_dependencies,
        },
        Case {
            name: "subtopics",
            boundary: 2,
            resource: "Volatile subtopic records",
            set_limit: set_max_subtopics,
            records: aggregate_records,
            dependencies: aggregate_dependencies,
        },
        Case {
            name: "references",
            boundary: 2,
            resource: "Volatile cell references",
            set_limit: set_max_references,
            records: aggregate_records,
            dependencies: aggregate_dependencies,
        },
        Case {
            name: "one string",
            boundary: 1,
            resource: "Volatile UTF-16 string",
            set_limit: set_max_string_units,
            records: aggregate_records,
            dependencies: aggregate_dependencies,
        },
        Case {
            name: "all strings",
            boundary: 4,
            resource: "Volatile UTF-16 strings",
            set_limit: set_max_total_string_units,
            records: aggregate_records,
            dependencies: aggregate_dependencies,
        },
    ];

    for case in cases {
        let records = (case.records)();
        let bytes = encode(&records);
        let mut accepted = ReadLimits::DEFAULT;
        (case.set_limit)(&mut accepted, case.boundary);
        codec::read(&bytes, accepted, 1)
            .unwrap_or_else(|error| panic!("{} boundary rejected: {error:?}", case.name));
        let mut refused = ReadLimits::DEFAULT;
        (case.set_limit)(&mut refused, case.boundary - 1);
        let error =
            codec::read(&bytes, refused, 1).expect_err("quota plus-one source must be refused");
        assert_limit(error, case.resource, case.boundary, case.boundary - 1);

        let dependencies = (case.dependencies)();
        let mut write_accepted = ReadLimits::DEFAULT;
        (case.set_limit)(&mut write_accepted, case.boundary);
        codec::write(&dependencies, write_accepted, 1)
            .unwrap_or_else(|error| panic!("{} write boundary rejected: {error:?}", case.name));
        let mut write_refused = ReadLimits::DEFAULT;
        (case.set_limit)(&mut write_refused, case.boundary - 1);
        let error = codec::write(&dependencies, write_refused, 1)
            .expect_err("quota plus-one model must be refused");
        let write_resource = match case.name {
            "types" => "volatile type collections",
            "mains" => "volatile main collections",
            "topics" => "volatile topic collections",
            "subtopics" => "volatile subtopic records",
            "references" => "volatile cell references",
            "one string" => "volatile first string",
            "all strings" => "volatile UTF-16 strings",
            _ => unreachable!("covered table case"),
        };
        assert_limit(error, write_resource, case.boundary, case.boundary - 1);
    }

    let records = aggregate_records();
    let bytes = encode(&records);
    let boundary = records.len();
    let mut accepted = ReadLimits::DEFAULT;
    accepted.max_records = boundary;
    codec::read(&bytes, accepted, 1).expect("record quota boundary");
    let mut refused = ReadLimits::DEFAULT;
    refused.max_records = boundary - 1;
    let error = codec::read(&bytes, refused, 1).expect_err("record quota plus-one");
    assert_limit(
        error,
        "Volatile Dependencies records",
        boundary,
        boundary - 1,
    );
}

#[test]
fn quota_checks_precede_later_malformed_or_oversized_strings() {
    let main_count_records = vec![
        (kind::BEGIN_VOL_DEPS, Vec::new()),
        (kind::BEGIN_VOL_TYPE, 0_u32.to_le_bytes().to_vec()),
        (kind::BEGIN_VOL_MAIN, wide("a")),
        (kind::END_VOL_MAIN, Vec::new()),
        (kind::BEGIN_VOL_MAIN, u32::MAX.to_le_bytes().to_vec()),
    ];
    let mut limits = ReadLimits::DEFAULT;
    limits.max_mains = 1;
    let error = codec::read(&encode(&main_count_records), limits, 1)
        .expect_err("main count must fail before sentinel string");
    assert_limit(error, "Volatile main collections", 2, 1);

    let malformed_later_string = vec![
        (kind::BEGIN_VOL_DEPS, Vec::new()),
        (kind::BEGIN_VOL_TYPE, 0_u32.to_le_bytes().to_vec()),
        (kind::BEGIN_VOL_MAIN, wide("a")),
        (kind::END_VOL_MAIN, Vec::new()),
        (kind::BEGIN_VOL_MAIN, 1_u32.to_le_bytes().to_vec()),
    ];
    let mut limits = ReadLimits::DEFAULT;
    limits.max_total_string_units = 1;
    let error = codec::read(&encode(&malformed_later_string), limits, 1)
        .expect_err("aggregate string count must fail before truncated payload");
    assert_limit(error, "Volatile UTF-16 strings", 2, 1);
}

#[test]
fn writer_preflights_record_quota_for_empty_streams() {
    let dependencies = Dependencies::new(Vec::new());
    let mut exact = ReadLimits::DEFAULT;
    exact.max_records = 2;
    codec::write(&dependencies, exact, 1).expect("empty stream has begin and end records");

    let mut too_small = ReadLimits::DEFAULT;
    too_small.max_records = 1;
    let error = codec::write(&dependencies, too_small, 1)
        .expect_err("writer must reject the second record before allocation");
    assert_limit(error, "Volatile Dependencies records", 2, 1);
}
