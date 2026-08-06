use super::super::generated::{SystemGeneratedData, SystemGeneratedFile, SystemGeneratedKind};
use super::super::native::{DictionaryBody, DictionaryType, NativeData, NativeFile};
use super::codec::{validate_columns, validate_hierarchies};
use super::*;

fn object(class: &str, body: &str) -> String {
    format!(
        "<XMObject class=\"{}\">{body}</XMObject>",
        class.replace('<', "&lt;")
    )
}

fn props(values: &[(&str, &str)]) -> String {
    format!(
        "<Properties>{}</Properties>",
        values
            .iter()
            .map(|(name, value)| format!("<{name}>{value}</{name}>"))
            .collect::<String>()
    )
}

fn member(name: &str, value: String) -> String {
    format!("<Member><Name>{name}</Name>{value}</Member>")
}

fn collection(name: &str, values: &[String]) -> String {
    format!(
        "<Collection><Name>{name}</Name>{}</Collection>",
        values.concat()
    )
}

fn empty_table(extra_columns: &[String]) -> String {
    let segment_map = object("XMSegment1Map", &props(&[("Records", "0")]));
    let stats = object(
        "XMTableStats",
        &props(&[("SegmentSize", "0"), ("Usage", "0")]),
    );
    object(
        "XMSimpleTable",
        &format!(
            "{}<Members>{}{}</Members><Collections>{}{}{}{}</Collections>",
            props(&[
                ("Version", "1"),
                ("Settings", "0"),
                ("RIViolationCount", "0")
            ]),
            member("SegmentMap", segment_map),
            member("TableStats", stats),
            collection("Partitions", &[]),
            collection("Columns", extra_columns),
            collection("Relationships", &[]),
            collection("UserHierarchies", &[])
        ),
    )
}

#[test]
fn parses_borrowed_table_metadata_and_writes_exact_bytes() {
    let xml = empty_table(&[]);
    let file = parse_file("Model.1.db/Table.0.dim/Table.1.tbl.xml", xml.as_bytes()).unwrap();
    assert_eq!(file.kind, MetadataFileKind::Table);
    assert_eq!(file.table.class.as_str(), "XMSimpleTable");
    assert!(std::ptr::eq(file.bytes.as_ptr(), xml.as_ptr()));
    assert_eq!(write_file(&file).unwrap(), xml.as_bytes());
}

#[test]
fn rejects_adversarial_schema_and_scalar_inputs() {
    let valid = empty_table(&[]);
    for invalid in [
        valid.replace("<Version>1</Version>", ""),
        valid.replace(
            "<Version>1</Version>",
            "<Version>1</Version><Version>2</Version>",
        ),
        valid.replace("<Usage>0</Usage>", "<Usage>3</Usage>"),
        valid.replace("XMSegment1Map", "UnknownClass"),
        format!("{valid}<XMObject class=\"XMRelationshipIndex123DIDs\"/>"),
    ] {
        assert!(parse_file("Model.1.db/Table.0.dim/Table.1.tbl.xml", invalid.as_bytes()).is_err());
    }
    assert!(parse_file("Model.1.db/Table.0.dim/Table.1.tbl.xml", b"<!DOCTYPE x [<!ENTITY a 'x'>]><XMObject class='XMRelationshipIndex123DIDs'>&a;</XMObject>").is_err());
}

#[test]
fn validates_dictionary_flags_operating_width_and_relationship_layout() {
    let numeric_bytes = [1u8; 8];
    let native = NativeFile {
        storage_path: "0.Table.Value.dictionary",
        bytes: &numeric_bytes,
        data: NativeData::Dictionary(super::super::native::DictionaryFile {
            dictionary_type: DictionaryType::Long,
            hash: None,
            body: DictionaryBody::Numeric(super::super::native::NumericDictionary {
                element_count: 2,
                element_size: 4,
                values: &numeric_bytes,
            }),
            trailing_zero_padding: &[],
        }),
    };
    let model = MetadataModel {
        files: vec![],
        columns: vec![ColumnPolicy {
            name: "Value".into(),
            data_file: "data.idf".into(),
            segment_count: 0,
            row_count: 0,
            compression_type: 0,
            settings: 1,
            dictionary: Some(DictionaryPolicy {
                storage_name: "0.Table.Value.dictionary".into(),
                class: MetadataClass("XMHashDataDictionary<XM_Long>".into()),
                dictionary_flags: None,
                operating_on_32: Some(true),
            }),
        }],
        relationships: vec![],
        hierarchies: vec![],
    };
    let idf = NativeFile {
        storage_path: "data.idf",
        bytes: &[],
        data: NativeData::Idf(super::super::native::IdfFile {
            segments: vec![],
            trailing_zero_padding: &[],
        }),
    };
    validate_columns(&model, &[idf.clone(), native.clone()]).unwrap();
    let mut bad = model.clone();
    bad.columns[0].dictionary.as_mut().unwrap().operating_on_32 = Some(false);
    assert!(validate_columns(&bad, &[idf, native]).is_err());
}

#[test]
fn enforces_generated_file_presence_and_sparse_relationship_policy() {
    let idf_bytes = 0u64.to_le_bytes();
    let idf = super::super::native::parse_idf(&idf_bytes).unwrap();
    let generated = SystemGeneratedFile {
        storage_path: "1.H$T$C.POS_TO_ID.0.idf",
        kind: SystemGeneratedKind::PositionToIdentifier,
        object_key: "1.H$T$C".into(),
        version: 0,
        bytes: &idf_bytes,
        data: SystemGeneratedData::Idf(idf),
    };
    let model = MetadataModel {
        files: vec![],
        columns: vec![],
        relationships: vec![],
        hierarchies: vec![HierarchyPolicy {
            table_store: "H$T$C".into(),
            processed: true,
            position_to_id: true,
            id_to_position: false,
            id_to_position_hash: false,
            level_ids: vec![],
            level_offsets: vec![],
        }],
    };
    validate_hierarchies(&model, &[], std::slice::from_ref(&generated)).unwrap();
    assert!(validate_hierarchies(&model, &[], &[]).is_err());
    let mut forbidden = model.clone();
    forbidden.hierarchies[0].position_to_id = false;
    assert!(validate_hierarchies(&forbidden, &[], &[generated]).is_err());
}
