//! Native Word parents and explicit inert IRM metadata derivatives, not rights enforcement.
use litchi_cfb::{OleFile, OleFileLimits, OleWriter};
use litchi_doc::{Package, raw::Reader, spaces};
use std::{collections::BTreeMap, io::Cursor, path::Path};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let data = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, data)
        })
        .collect()
}
const TRANSFORM: &str = "0x09DRMTransform";
const PAYLOAD: &[u8] = b"INERT_TEST_BYTES_NOT_ENCRYPTED_OR_AUTHORIZED";
const LICENSE: &str = "<XrML><endpoint>https://example.invalid/never-contacted</endpoint></XrML>";
fn header() -> spaces::Header {
    spaces::Header {
        transform_id: spaces::DRM_ID.into(),
        transform_name: spaces::DRM_NAME.into(),
        reader: spaces::Version::V1_0,
        updater: spaces::Version::V1_0,
        writer: spaces::Version::V1_0,
    }
}
fn derived(bytes: &[u8], license_xml: &str) -> Vec<u8> {
    let mut data = streams(bytes);
    let limits = spaces::Limits::default();
    let mut put = |p: &[&str], b: Vec<u8>| {
        assert!(
            data.insert(p.iter().map(|s| (*s).to_owned()).collect(), b)
                .is_none()
        );
    };
    put(&["0x09DRMContent"], PAYLOAD.to_vec());
    put(
        &[spaces::STORAGE, "Version"],
        spaces::write_version_info(&spaces::VersionInfo::default(), &limits).unwrap(),
    );
    put(
        &[spaces::STORAGE, "DataSpaceMap"],
        spaces::write_map(
            &spaces::Map {
                entries: vec![spaces::MapEntry {
                    references: vec![spaces::Reference {
                        kind: spaces::ReferenceKind::Stream,
                        component: "0x09DRMContent".into(),
                    }],
                    data_space_name: "0x09DRMDataSpace".into(),
                }],
            },
            &limits,
        )
        .unwrap(),
    );
    put(
        &[spaces::STORAGE, "DataSpaceInfo", "0x09DRMDataSpace"],
        spaces::write_definition(
            &spaces::Definition {
                transforms: vec![TRANSFORM.into()],
            },
            &limits,
        )
        .unwrap(),
    );
    put(
        &[spaces::STORAGE, "TransformInfo", TRANSFORM, spaces::PRIMARY],
        spaces::write_irm_transform(
            &spaces::IrmTransform {
                header: header(),
                publishing_license: Some(license_xml.into()),
            },
            &limits,
        )
        .unwrap(),
    );
    let license = spaces::License {
        stream_name: "EUL-ETRHA1143ZLUDD412YTI3M5CTZ".into(),
        encoded_license_id: "VwBpAG4AZABvAHcAOgB1AHMAZQByAEA".into(),
        certificate_chain: Some("<certificatechain/>".into()),
    };
    put(
        &[
            spaces::STORAGE,
            "TransformInfo",
            TRANSFORM,
            &license.stream_name,
        ],
        spaces::write_license(&license, &limits).unwrap(),
    );
    let mut writer = OleWriter::new();
    // WordDocument must be the first large stream allocation.
    writer
        .create_stream(
            &["WordDocument"],
            &data.remove(&vec!["WordDocument".into()]).unwrap(),
        )
        .unwrap();
    for (p, b) in data {
        let p = p.iter().map(String::as_str).collect::<Vec<_>>();
        for end in 1..p.len() {
            writer.create_storage(&p[..end]).unwrap();
        }
        writer.create_stream(&p, &b).unwrap();
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    out.into_inner()
}

#[test]
fn native_binary_word_corpus_has_no_irm_graph_or_license_evidence() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut count = 0;
    for row in manifest["files"].as_array().unwrap() {
        let name = row["path"].as_str().unwrap();
        if !name.starts_with("word/")
            || !matches!(
                Path::new(name).extension().and_then(|v| v.to_str()),
                Some("doc" | "dot")
            )
        {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        let mut reader = Reader::from_reader(Cursor::new(&bytes)).unwrap();
        assert!(
            reader.data_spaces().unwrap().is_none(),
            "new IRM evidence: {name}"
        );
        assert!(reader.data_spaces_snapshot().unwrap().is_none());
        assert!(std::fs::read(root.join(name)).unwrap() == bytes);
        count += 1;
    }
    assert_eq!(count, 20);
}

#[test]
fn native_word_irm_metadata_derivatives_preserve_inert_payloads_and_refuse_reinterpretation() {
    for name in ["word-text-structure.doc", "word-recorded-macro.doc"] {
        let bytes = native(name);
        let input = derived(&bytes, LICENSE);
        let original = streams(&bytes);
        let before = streams(&input);
        for (p, b) in &original {
            assert!(before.get(p) == Some(b), "changed native stream {p:?}");
        }
        assert_eq!(
            Package::from_bytes(&input)
                .unwrap()
                .document()
                .text()
                .unwrap(),
            Package::from_bytes(&bytes)
                .unwrap()
                .document()
                .text()
                .unwrap()
        );
        let mut reader = Reader::from_reader(Cursor::new(&input)).unwrap();
        let snapshot = reader.data_spaces_snapshot().unwrap().unwrap();
        let graph = snapshot.graph();
        let irm = graph.irm.as_ref().unwrap();
        assert_eq!(irm.document_kind, spaces::DocumentKind::LegacyBinary);
        assert_eq!(irm.protected_stream, "0x09DRMContent");
        assert!(irm.viewer_content_stream.is_none());
        assert_eq!(graph.transforms.len(), 1);
        assert_eq!(
            graph.transforms[0]
                .irm
                .as_ref()
                .unwrap()
                .publishing_license
                .as_deref(),
            Some(LICENSE)
        );
        assert_eq!(graph.transforms[0].end_user_licenses.len(), 1);
        assert_eq!(
            graph.transforms[0].end_user_licenses[0]
                .certificate_chain
                .as_deref(),
            Some("<certificatechain/>")
        );
        let mut edit = snapshot.edit();
        let mut changed = header();
        changed.transform_name = "Different Runtime".into();
        assert!(
            matches!(edit.set_transform_header(TRANSFORM,changed),Err(spaces::Error::Invalid(m)) if m=="known transform identity cannot be changed while its payload is inert")
        );
        assert!(
            matches!(edit.set_transform_opaque_tail(TRANSFORM,b"replacement".to_vec()),Err(spaces::Error::Invalid(m)) if m=="known transform payloads are inert and cannot be replaced")
        );
        assert_eq!(edit.graph(), graph);
        let commit = edit.commit().unwrap();
        assert!(commit.patch().is_noop());
        let different = derived(
            &bytes,
            &LICENSE.replace("never-contacted", "other-endpoint"),
        );
        let mut other = Reader::from_reader(Cursor::new(&different)).unwrap();
        assert!(matches!(
            commit
                .patch()
                .apply(&other.data_spaces_snapshot().unwrap().unwrap()),
            Err(spaces::Error::PatchConflict)
        ));
        assert_eq!(
            commit.patch().inverse().apply(&snapshot).unwrap().graph(),
            graph
        );

        let mut output = Cursor::new(Vec::new());
        reader
            .write_data_spaces_patch(commit.patch(), &mut output)
            .unwrap();
        let output = output.into_inner();
        assert!(streams(&output) == before);
        let mut reopened = Reader::from_reader(Cursor::new(&output)).unwrap();
        assert_eq!(reopened.data_spaces().unwrap().as_ref(), Some(graph));
        let file = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
        reopened.save(file.path()).unwrap();
        assert!(streams(&std::fs::read(file.path()).unwrap()) == before);
        assert_eq!(before[&vec!["0x09DRMContent".into()]], PAYLOAD);
        assert!(native(name) == bytes);
    }
}
