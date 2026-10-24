//! Focused signed-native qualification consumer for record 0362ul24.
//!
//! This source is compiled as an external binary against a pinned litchi
//! checkout.  It deliberately contains no allocator or raw-pointer instrumentation:
//! native latency and allocation evidence are separate protocols.

use std::fmt::Write as _;
use std::fs;
use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Instant;

use litchi_core::OwnedSource;
use litchi_ops::error::Source;
use litchi_ops::transaction::{EditOptions, Error as TransactionError, LimitKind, SemanticLimits};
use litchi_ops::{Code, File, Props, open, open_bytes};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi_signed_native_probe.v1";
const TARGET_NS: u128 = 500_000_000;
const MAX_CALLS: usize = 2_000_000;
const ROUTES: &[&str] = &[
    "signed_leaf/noop_commit",
    "signed_facade/noop_commit",
    "signed_leaf/noop_write",
    "signed_facade/noop_write",
    "signed_leaf/noop_replay",
    "signed_facade/noop_replay",
    "signed_leaf/refusal",
    "signed_facade/refusal",
    "signed_leaf/unsign_write",
    "signed_facade/unsign_write",
    "signed_leaf/unsigned_negative",
    "signed_facade/unsigned_negative",
    "signed_leaf/source_freshness",
    "signed_facade/source_freshness",
    "signed_leaf/budget",
    "signed_facade/budget",
];

#[derive(Clone)]
struct Fixture {
    name: String,
    format: String,
    origin: String,
    target: String,
    property: String,
    signed: Vec<u8>,
    source_sha256: String,
}

struct TimedTarget {
    fixture_index: usize,
    endpoint: File,
    recipient: File,
    value: String,
    changed: String,
}

struct StreamDigest {
    digest: Sha256,
    bytes: usize,
    writes: usize,
    max_write_bytes: usize,
}

impl StreamDigest {
    fn new() -> Self {
        Self {
            digest: Sha256::new(),
            bytes: 0,
            writes: 0,
            max_write_bytes: 0,
        }
    }

    fn finish(self) -> (String, usize, usize, usize) {
        let mut text = String::with_capacity(64);
        for byte in self.digest.finalize() {
            let _ = write!(text, "{byte:02x}");
        }
        (text, self.bytes, self.writes, self.max_write_bytes)
    }
}

impl Write for StreamDigest {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.digest.update(bytes);
        self.bytes = self.bytes.saturating_add(bytes.len());
        self.writes = self.writes.saturating_add(1);
        self.max_write_bytes = self.max_write_bytes.max(bytes.len());
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn digest(bytes: &[u8]) -> String {
    let mut text = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn fixture_root() -> PathBuf {
    std::env::var_os("LITCHI_SIGNED_FIXTURE_ROOT")
        .map(PathBuf::from)
        .expect("LITCHI_SIGNED_FIXTURE_ROOT")
}

fn endpoint() -> &'static str {
    match std::env::var("LITCHI_SIGNED_ENDPOINT").as_deref() {
        Ok("before") => "before",
        _ => "after",
    }
}

fn manifest_path(root: &Path) -> PathBuf {
    root.join("docs/unified-ops/results/0362ul24/fixtures.json")
}

fn string_field<'a>(object: &'a serde_json::Map<String, Value>, key: &str) -> &'a str {
    object
        .get(key)
        .and_then(Value::as_str)
        .unwrap_or_else(|| panic!("fixture field {key}"))
}

fn signed_opc(plain: &[u8]) -> Vec<u8> {
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};

    let source =
        OpcPackage::from_bytes(plain, litchi_opc::SourceLimits::default()).expect("OPC source");
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .expect("copy root relationship");
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).expect("source part");
        let mut copy = BlobPart::new(
            name.clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .expect("copy part relationship");
        }
        package.try_add_part(Box::new(copy)).expect("copy part");
    }
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("origin URI");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature URI");
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    package
        .try_add_part(Box::new(origin_part))
        .expect("origin part");
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .expect("signature part");
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .expect("origin relationship");
    PackageWriter::to_bytes(&package).expect("OPC signed package")
}

fn signed_compound_file(plain: Vec<u8>) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source = OleFile::open(Cursor::new(plain), litchi_cfb::OleFileLimits::default())
        .expect("CFB source");
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let path = path.iter().map(String::as_str).collect::<Vec<_>>();
        for depth in 1..path.len() {
            writer.create_storage(&path[..depth]).expect("CFB storage");
        }
        let data = source.open_stream(&path).expect("CFB stream");
        writer.create_stream(&path, &data).expect("CFB copy stream");
    }
    writer
        .create_stream(&["_signatures"], b"legacy signature")
        .expect("legacy signature");
    writer
        .create_storage(&["_xmlsignatures"])
        .expect("XML signature storage");
    writer
        .create_stream(&["_xmlsignatures", "1"], b"<Signature/>")
        .expect("XML signature");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB signed package");
    output.into_inner()
}

fn load_fixtures(root: &Path) -> Vec<Fixture> {
    let manifest = fs::read_to_string(manifest_path(root)).expect("signed fixture manifest");
    let value: Value = serde_json::from_str(&manifest).expect("signed fixture JSON");
    let rows = value
        .get("fixtures")
        .and_then(Value::as_array)
        .expect("fixture rows");
    rows.iter()
        .map(|row| {
            let object = row.as_object().expect("fixture object");
            let relative = string_field(object, "relative_path");
            let plain_path = root.join(relative);
            let plain = fs::read(&plain_path)
                .unwrap_or_else(|error| panic!("{}: {error}", plain_path.display()));
            let expected = string_field(object, "input_sha256");
            assert_eq!(digest(&plain), expected, "{} input hash", relative);
            let mode = string_field(object, "fixture_mode").to_owned();
            let signed = match mode.as_str() {
                "producer_signed" => plain.clone(),
                "synthetic_presence" => match string_field(object, "format") {
                    "XLSB" => signed_opc(&plain),
                    "DOC" | "PPT" | "XLS" => signed_compound_file(plain.clone()),
                    other => panic!("unsupported synthetic signed format {other}"),
                },
                other => panic!("unsupported fixture mode {other}"),
            };
            if mode == "producer_signed" {
                assert_eq!(
                    digest(&signed),
                    string_field(object, "signed_sha256"),
                    "{relative} signed hash"
                );
            }
            Fixture {
                name: string_field(object, "name").to_owned(),
                format: string_field(object, "format").to_owned(),
                origin: string_field(object, "origin").to_owned(),
                target: string_field(object, "target").to_owned(),
                property: string_field(object, "property").to_owned(),
                source_sha256: digest(&signed),
                signed,
            }
        })
        .collect()
}

fn open_signed(fixture: &Fixture) -> File {
    let file = open_bytes(fixture.signed.clone()).expect("signed fixture opens");
    let File::Signed(signed) = &file else {
        panic!("{} did not open as File::Signed", fixture.name);
    };
    assert_eq!(
        format!("{:?}", signed.origin()),
        fixture.origin,
        "{}",
        fixture.name
    );
    let format = signed.format();
    assert_eq!(format!("{format:?}").to_ascii_uppercase(), fixture.format);
    assert!(file.is_signed(), "{} signature presence", fixture.name);
    file
}

fn property_value(file: &File, fixture: &Fixture) -> String {
    if fixture.property == "text" {
        return file
            .get(fixture.target.as_str())
            .expect("native text target")
            .view()
            .text()
            .expect("native text read")
            .unwrap_or_default()
            .into_owned();
    }
    file.get(fixture.target.as_str())
        .expect("native format target")
        .view()
        .format()
        .expect("native format read")
        .find(|property| property.key() == fixture.property)
        .map(|property| property.value().to_string())
        .unwrap_or_else(|| panic!("{} has no {}", fixture.name, fixture.property))
}

fn same_width(value: &str) -> String {
    let mut replacement = value
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() {
                'x'
            } else {
                character
            }
        })
        .collect::<String>();
    if replacement == value {
        replacement.push('x');
    }
    replacement
}

fn props<'a>(fixture: &'a Fixture, value: &'a str) -> Props<'a> {
    Props::new().with(fixture.property.as_str(), value)
}

fn endpoint_file(file: &File, facade: bool) -> File {
    if facade {
        return file.clone();
    }
    let File::Signed(signed) = file else {
        // `Signed::unsign(self)` consumes the wrapper and returns its
        // underlying owner.  There is no second signed wrapper to unwrap on
        // the post-authorization path; both public spellings therefore use
        // that returned owner after the explicit consume.
        return file.clone();
    };
    signed.file().clone()
}

fn signature_parts_absent(fixture: &Fixture, bytes: &[u8]) -> bool {
    match fixture.origin.as_str() {
        "Opc" => {
            let package =
                litchi_opc::OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
                    .expect("unsigned OPC output");
            !package.is_signed()
        },
        "Cfb" => {
            let ole =
                litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
                    .expect("unsigned CFB output");
            !ole.exists(&["_signatures"]) && !ole.directory_exists(&["_xmlsignatures"])
        },
        other => panic!("unknown signature origin {other}"),
    }
}

fn signed_row(fixture: &Fixture, file: &File) -> Value {
    json!({
        "name": fixture.name,
        "format": fixture.format,
        "origin": fixture.origin,
        "signature_presence": "detected",
        "cryptographic_verification": "not_performed",
        "source_sha256": fixture.source_sha256,
        "signed_before": file.is_signed(),
    })
}

fn direct_native_noop(fixture: &Fixture) -> Value {
    let status = match fixture.format.as_str() {
        "DOCX" => {
            let package = litchi_docx::Package::from_reader(Cursor::new(&fixture.signed))
                .expect("DOCX native package");
            let commit = package
                .edit()
                .expect("DOCX native edit")
                .commit()
                .expect("DOCX native no-op");
            let output = commit.package().to_bytes().expect("DOCX native bytes");
            assert_eq!(output, fixture.signed, "DOCX native no-op bytes");
            let source = litchi_docx::source_backed::Package::from_read_at(Arc::new(
                OwnedSource::new(fixture.signed.clone()),
            ))
            .expect("DOCX source native package");
            let commit = source
                .edit()
                .expect("DOCX source native edit")
                .commit()
                .expect("DOCX source native no-op");
            let mut output = Vec::new();
            commit
                .package()
                .write_to(&mut output)
                .expect("DOCX source native bytes");
            assert_eq!(output, fixture.signed, "DOCX source native no-op bytes");
            "exact_noop"
        },
        "PPTX" => {
            let mut package = litchi_pptx::Package::from_reader(&fixture.signed[..])
                .expect("PPTX native package");
            let commit = package
                .edit()
                .expect("PPTX native edit")
                .commit()
                .expect("PPTX native no-op");
            package = commit.into_package();
            assert_eq!(
                package.to_bytes().expect("PPTX native bytes"),
                fixture.signed
            );
            let source = litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(
                fixture.signed.clone(),
            )))
            .expect("PPTX source native package");
            let commit = source
                .edit()
                .expect("PPTX source native edit")
                .commit()
                .expect("PPTX source native no-op");
            let mut output = Vec::new();
            commit
                .package()
                .write_to(&mut output)
                .expect("PPTX source native bytes");
            assert_eq!(output, fixture.signed, "PPTX source native no-op bytes");
            "exact_noop"
        },
        "XLSX" => {
            let workbook =
                litchi_xlsx::Package::from_bytes(&fixture.signed).expect("XLSX native package");
            let commit = workbook
                .edit()
                .expect("XLSX native edit")
                .commit()
                .expect("XLSX native no-op");
            assert_eq!(
                commit.package().to_bytes().expect("XLSX native bytes"),
                fixture.signed
            );
            let source = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(
                OwnedSource::new(fixture.signed.clone()),
            ))
            .expect("XLSX source native package");
            let commit = source
                .edit_defined_names()
                .expect("XLSX source native edit")
                .commit()
                .expect("XLSX source native no-op");
            let mut output = Vec::new();
            commit
                .package()
                .write_to(&mut output)
                .expect("XLSX source native bytes");
            assert_eq!(output, fixture.signed, "XLSX source native no-op bytes");
            "exact_noop"
        },
        "XLSB" => {
            let package =
                litchi_xlsb::Package::from_bytes(&fixture.signed).expect("XLSB native package");
            let commit = package
                .edit_cell_values(0)
                .expect("XLSB native selector")
                .resolve()
                .expect("XLSB native resolve")
                .expect("XLSB native worksheet")
                .commit()
                .expect("XLSB native no-op");
            let mut output = Cursor::new(Vec::new());
            commit
                .into_package()
                .write_to(&mut output)
                .expect("XLSB native bytes");
            assert_eq!(output.into_inner(), fixture.signed);
            "exact_noop"
        },
        "XLS" => {
            let package =
                litchi_xls::Package::from_bytes(&fixture.signed).expect("XLS native package");
            let commit = package
                .edit()
                .expect("XLS native edit")
                .commit()
                .expect("XLS native no-op");
            assert_eq!(
                commit.package().to_bytes().expect("XLS native bytes"),
                fixture.signed
            );
            "exact_noop"
        },
        "DOC" => {
            assert!(
                litchi_doc::source_backed::Package::from_bytes(fixture.signed.clone()).is_err(),
                "signed DOC source editor must refuse"
            );
            "unsupported_signed_source"
        },
        "PPT" => {
            let source = Arc::new(OwnedSource::new(fixture.signed.clone()));
            assert!(
                litchi_ppt::text_edit::SourceSnapshot::open(source).is_err(),
                "signed PPT source editor must refuse"
            );
            "unsupported_signed_source"
        },
        other => panic!("direct native format {other}"),
    };
    json!({"format": fixture.format, "status": status})
}

fn native_preflight(fixtures: &[Fixture]) -> Vec<Value> {
    fixtures.iter().map(direct_native_noop).collect()
}

fn no_op_commit(fixture: &Fixture, facade: bool) -> litchi_ops::Result<Value> {
    let file = open_signed(fixture);
    let endpoint = endpoint_file(&file, facade);
    let value = property_value(&endpoint, fixture);
    let commit = endpoint.edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &value)))?;
    let output = commit.file().to_bytes()?;
    let mut row = signed_row(fixture, &file);
    row["status"] = json!("ok");
    row["diagnostics_changed"] = json!(commit.diagnostics().changed());
    row["patch_noop"] = json!(commit.patch().is_noop());
    row["signed_after"] = json!(commit.file().is_signed());
    row["output_sha256"] = json!(digest(&output));
    row["output_equals_source"] = json!(output == fixture.signed);
    row["source_after_sha256"] = json!(digest(&endpoint.to_bytes()?));
    Ok(row)
}

fn bounded_write(file: &File) -> litchi_ops::Result<(String, usize, usize, usize)> {
    let mut stream = StreamDigest::new();
    file.write_to(&mut stream)?;
    Ok(stream.finish())
}

fn no_op_write(fixture: &Fixture, facade: bool) -> litchi_ops::Result<Value> {
    let file = open_signed(fixture);
    let endpoint = endpoint_file(&file, facade);
    let value = property_value(&endpoint, fixture);
    let commit = endpoint.edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &value)))?;
    let (output_sha256, bytes, writes, max_write_bytes) = bounded_write(commit.file())?;
    let mut row = signed_row(fixture, &file);
    row["status"] = json!("ok");
    row["patch_noop"] = json!(commit.patch().is_noop());
    row["diagnostics_changed"] = json!(commit.diagnostics().changed());
    row["signed_after"] = json!(commit.file().is_signed());
    row["output_sha256"] = json!(output_sha256);
    row["output_equals_source"] =
        json!(bytes == fixture.signed.len() && output_sha256 == fixture.source_sha256);
    row["stream_bounded"] = json!(writes > 0 && max_write_bytes <= 8192);
    row["bytes"] = json!(bytes);
    row["writes"] = json!(writes);
    row["max_write_bytes"] = json!(max_write_bytes);
    Ok(row)
}

fn no_op_replay(fixture: &Fixture, facade: bool) -> litchi_ops::Result<Value> {
    let file = open_signed(fixture);
    let endpoint = endpoint_file(&file, facade);
    let value = property_value(&endpoint, fixture);
    let commit = endpoint.edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &value)))?;
    let reopened = open_signed(fixture);
    let recipient = endpoint_file(&reopened, facade);
    let replayed = commit.patch().apply(&recipient)?;
    let output = replayed.to_bytes()?;
    let mut row = signed_row(fixture, &file);
    row["status"] = json!("ok");
    row["patch_noop"] = json!(commit.patch().is_noop());
    row["replay_base_match"] = json!(output == fixture.signed);
    row["signed_after"] = json!(replayed.is_signed());
    row["output_sha256"] = json!(digest(&output));
    row["output_equals_source"] = json!(output == fixture.signed);
    Ok(row)
}

fn changed_refusal(fixture: &Fixture, facade: bool) -> Value {
    let file = open_signed(fixture);
    let endpoint = endpoint_file(&file, facade);
    let original = property_value(&endpoint, fixture);
    let changed = same_width(&original);
    let before = endpoint.to_bytes().expect("source bytes");
    let error = endpoint
        .edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &changed)))
        .expect_err("changed signed edit must refuse");
    let mut row = signed_row(fixture, &file);
    row["status"] = json!("expected_refusal");
    row["refusal_kind"] = json!("ProtectedDocument");
    row["code"] = json!(error.code().as_str());
    row["source_after_sha256"] = json!(digest(&endpoint.to_bytes().expect("source remains")));
    assert_eq!(error.code(), Code::DocumentProtected);
    assert_eq!(before, endpoint.to_bytes().expect("source remains bytes"));
    row
}

fn unsigned_negative(fixture: &Fixture, facade: bool) -> litchi_ops::Result<Value> {
    let file = open_signed(fixture);
    let signed = match &file {
        File::Signed(signed) => signed.clone(),
        _ => unreachable!(),
    };
    let unsigned = signed.unsign();
    let value = property_value(&unsigned, fixture);
    let commit = unsigned
        .edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &value)))
        .expect("authorized unsigned no-op");
    let error = commit
        .patch()
        .apply(&endpoint_file(&file, facade))
        .expect_err("unsigned no-op patch must conflict with signed publication");
    let recipient = endpoint_file(&file, facade);
    let mut row = signed_row(fixture, &file);
    row["status"] = json!("expected_refusal");
    row["refusal_kind"] = json!("PatchConflict");
    row["code"] = json!(error.code().as_str());
    row["re_marked_signed"] = json!(false);
    row["recipient_signed_after"] = json!(recipient.is_signed());
    row["output_published"] = json!(false);
    assert_eq!(error.code(), Code::InvalidInput);
    Ok(row)
}

fn source_freshness(fixture: &Fixture, facade: bool) -> litchi_ops::Result<Value> {
    let suffix = match fixture.format.as_str() {
        "DOCX" => "docx",
        "PPTX" => "pptx",
        "XLSX" => "xlsx",
        "XLSB" => "xlsb",
        "DOC" => "doc",
        "PPT" => "ppt",
        "XLS" => "xls",
        _ => unreachable!(),
    };
    let path = std::env::temp_dir().join(format!(
        "litchi-signed-source-{}-{}.{}",
        std::process::id(),
        fixture
            .name
            .replace('/', "_")
            .replace('\\', "_")
            .replace('.', "_"),
        suffix
    ));
    fs::write(&path, &fixture.signed).expect("source freshness fixture");
    let opened = open(&path)?;
    let endpoint = endpoint_file(&opened, facade);
    let value = property_value(&endpoint, fixture);
    let patch = endpoint
        .edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &value)))?
        .patch()
        .clone();
    let mut changed = fixture.signed.clone();
    changed[0] ^= 1;
    fs::write(&path, &changed).expect("mutate source fixture");
    let error = patch.apply(&endpoint).expect_err("source freshness fence");
    let _ = fs::remove_file(&path);
    let mut row = signed_row(fixture, &opened);
    row["status"] = json!("expected_refusal");
    row["refusal_kind"] = json!("SourceChanged");
    row["code"] = json!(error.code().as_str());
    row["source_mutated_sha256"] = json!(digest(&changed));
    row["output_published"] = json!(false);
    assert_eq!(error.code(), Code::IoError);
    Ok(row)
}

fn limit_kind_name(kind: LimitKind) -> &'static str {
    match kind {
        LimitKind::Operations => "operations",
        LimitKind::SnapshotBytes => "snapshot_bytes",
        LimitKind::ExpandedTargets => "expanded_targets",
        LimitKind::StagedBytes => "staged_bytes",
        _ => "unknown",
    }
}

fn budget(fixture: &Fixture, facade: bool) -> Value {
    let file = open_signed(fixture);
    let endpoint = endpoint_file(&file, facade);
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .expect("one-operation limit")
        .build()
        .expect("semantic limits");
    let options = EditOptions::new().semantic_limits(limits);
    let value = property_value(&endpoint, fixture);
    let error = endpoint
        .edit_with_options(options, |edit| {
            edit.set(fixture.target.as_str(), props(fixture, &value))?;
            edit.set(fixture.target.as_str(), props(fixture, &value))
        })
        .expect_err("one-operation limit");
    let limit_kind = match error.leaf() {
        Source::Transaction(TransactionError::Limit { kind, .. }) => limit_kind_name(*kind),
        other => panic!("budget refusal was not a transaction limit: {other:?}"),
    };
    let mut row = signed_row(fixture, &file);
    row["status"] = json!("expected_refusal");
    row["refusal_kind"] = json!("LimitExceeded");
    row["code"] = json!(error.code().as_str());
    row["limit_kind"] = json!(limit_kind);
    row["source_after_sha256"] = json!(digest(&endpoint.to_bytes().expect("budget source")));
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(limit_kind, "operations");
    row
}

fn route_operation(route: &str, fixture: &Fixture) -> litchi_ops::Result<Value> {
    let facade = route.starts_with("signed_facade/");
    match route {
        "signed_leaf/noop_commit" | "signed_facade/noop_commit" => no_op_commit(fixture, facade),
        "signed_leaf/noop_write" | "signed_facade/noop_write" => no_op_write(fixture, facade),
        "signed_leaf/noop_replay" | "signed_facade/noop_replay" => no_op_replay(fixture, facade),
        "signed_leaf/refusal" | "signed_facade/refusal" => Ok(changed_refusal(fixture, facade)),
        "signed_leaf/unsigned_negative" | "signed_facade/unsigned_negative" => {
            unsigned_negative(fixture, facade)
        },
        "signed_leaf/source_freshness" | "signed_facade/source_freshness" => {
            source_freshness(fixture, facade)
        },
        "signed_leaf/budget" | "signed_facade/budget" => Ok(budget(fixture, facade)),
        "signed_leaf/unsign_write" | "signed_facade/unsign_write" => {
            Ok(unsign_write(fixture, facade))
        },
        other => panic!("unknown signed route {other}"),
    }
}

fn capability_missing(route: &str, fixture: &Fixture, error: &str) -> Value {
    json!({
        "name": fixture.name,
        "format": fixture.format,
        "origin": fixture.origin,
        "signature_presence": "detected",
        "cryptographic_verification": "not_performed",
        "source_sha256": fixture.source_sha256,
        "status": "capability_missing",
        "error": error,
        "refusal_kind": "ProtectedDocument",
        "refusal_code": "document_protected",
        "signed_before": true,
    })
}

fn baseline_gap_route(route: &str) -> bool {
    matches!(
        route,
        "signed_leaf/noop_commit"
            | "signed_facade/noop_commit"
            | "signed_leaf/noop_write"
            | "signed_facade/noop_write"
            | "signed_leaf/noop_replay"
            | "signed_facade/noop_replay"
    )
}

fn unsign_write(fixture: &Fixture, facade: bool) -> Value {
    let file = open_signed(fixture);
    let signed_clone = file.clone();
    let signed = match file {
        File::Signed(signed) => signed,
        _ => unreachable!(),
    };
    let unsigned = signed.unsign();
    let endpoint = endpoint_file(&unsigned, facade);
    let original = property_value(&endpoint, fixture);
    let changed = same_width(&original);
    let commit = endpoint
        .edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &changed)))
        .expect("authorized changed edit");
    let output = commit.file().to_bytes().expect("authorized output");
    let (stream_sha256, bytes, writes, max_write_bytes) =
        bounded_write(commit.file()).expect("authorized bounded publication");
    let reopened = open_bytes(output.clone()).expect("authorized output reopens");
    let signature_parts_absent = signature_parts_absent(fixture, &output);
    let target_after = property_value(&reopened, fixture);
    let mut row = signed_row(fixture, &signed_clone);
    row["status"] = json!("ok");
    row["signed_after"] = json!(reopened.is_signed());
    row["signature_parts_absent"] = json!(signature_parts_absent);
    row["semantic_changed"] = json!(target_after == changed);
    row["target_changed"] = json!(target_after == changed);
    row["output_sha256"] = json!(digest(&output));
    row["stream_sha256"] = json!(stream_sha256);
    row["bytes"] = json!(bytes);
    row["writes"] = json!(writes);
    row["max_write_bytes"] = json!(max_write_bytes);
    row["stream_bounded"] = json!(writes > 0 && max_write_bytes <= 8192);
    row["output_equals_source"] = json!(output == fixture.signed);
    row["signed_clone_still_refuses"] = json!(
        endpoint_file(&signed_clone, facade)
            .edit(|edit| edit.set(fixture.target.as_str(), props(fixture, &changed)))
            .is_err()
    );
    assert!(signature_parts_absent);
    assert_eq!(target_after, changed);
    assert_ne!(output, fixture.signed);
    row
}

fn prepare_timed(route: &str, fixtures: &[Fixture]) -> Vec<TimedTarget> {
    fixtures
        .iter()
        .enumerate()
        .map(|(fixture_index, fixture)| {
            let file = open_signed(fixture);
            let facade = route.starts_with("signed_facade/");
            let endpoint = if route.ends_with("/unsign_write") {
                let signed = match file {
                    File::Signed(signed) => signed,
                    _ => unreachable!(),
                };
                endpoint_file(&signed.unsign(), facade)
            } else {
                endpoint_file(&file, facade)
            };
            let recipient = if route.ends_with("/noop_replay") {
                let reopened = open_signed(fixture);
                endpoint_file(&reopened, facade)
            } else {
                endpoint.clone()
            };
            let value = property_value(&endpoint, fixture);
            TimedTarget {
                fixture_index,
                endpoint,
                recipient,
                changed: same_width(&value),
                value,
            }
        })
        .collect()
}

fn timed_many(
    route: &str,
    fixtures: &[Fixture],
    targets: &[TimedTarget],
    calls: usize,
) -> (u128, usize) {
    let started = Instant::now();
    for _ in 0..calls {
        for target in targets {
            let fixture = &fixtures[target.fixture_index];
            match route {
                "signed_leaf/noop_commit" | "signed_facade/noop_commit" => {
                    let commit = target
                        .endpoint
                        .edit(|edit| {
                            edit.set(fixture.target.as_str(), props(fixture, &target.value))
                        })
                        .expect("timed signed no-op");
                    std::hint::black_box(commit.diagnostics().changed());
                },
                "signed_leaf/noop_write" | "signed_facade/noop_write" => {
                    let commit = target
                        .endpoint
                        .edit(|edit| {
                            edit.set(fixture.target.as_str(), props(fixture, &target.value))
                        })
                        .expect("timed signed publication");
                    let mut stream = StreamDigest::new();
                    commit.file().write_to(&mut stream).expect("timed write");
                    std::hint::black_box(stream.finish());
                },
                "signed_leaf/noop_replay" | "signed_facade/noop_replay" => {
                    let commit = target
                        .endpoint
                        .edit(|edit| {
                            edit.set(fixture.target.as_str(), props(fixture, &target.value))
                        })
                        .expect("timed signed replay patch");
                    let replayed = commit
                        .patch()
                        .apply(&target.recipient)
                        .expect("timed replay");
                    std::hint::black_box(replayed.is_signed());
                },
                "signed_leaf/unsign_write" | "signed_facade/unsign_write" => {
                    let commit = target
                        .endpoint
                        .edit(|edit| {
                            edit.set(fixture.target.as_str(), props(fixture, &target.changed))
                        })
                        .expect("timed authorized edit");
                    let mut stream = StreamDigest::new();
                    commit
                        .file()
                        .write_to(&mut stream)
                        .expect("timed authorized write");
                    std::hint::black_box(stream.finish());
                },
                _ => panic!("route {route} is correctness-only"),
            }
        }
    }
    (
        started.elapsed().as_nanos(),
        calls.saturating_mul(targets.len()),
    )
}

fn mode_value(mode: &str, route: &str, fixtures: &[Fixture], calls: usize) -> Value {
    if mode == "list" {
        return json!({"schema": SCHEMA, "routes": ROUTES});
    }
    if mode == "validate" {
        let mut rows = Vec::with_capacity(fixtures.len());
        let mut capability = "available";
        for fixture in fixtures {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                route_operation(route, fixture)
            }));
            match result {
                Ok(Ok(row)) => rows.push(row),
                Ok(Err(error))
                    if endpoint() == "before"
                        && baseline_gap_route(route)
                        && error.code() == Code::DocumentProtected =>
                {
                    capability = "capability_missing";
                    rows.push(capability_missing(
                        route,
                        fixture,
                        "old endpoint returned the typed protected-document refusal",
                    ));
                },
                Ok(Err(error)) => panic!(
                    "signed route {route} failed preflight with {}",
                    error.code()
                ),
                Err(_) => panic!("signed route {route} panicked during preflight"),
            }
        }
        let native = if capability == "available" {
            native_preflight(fixtures)
        } else {
            Vec::new()
        };
        return json!({
            "schema": SCHEMA,
            "route": route,
            "capability": capability,
            "native_preflight": native,
            "fixtures": rows,
        });
    }
    let requested = calls.max(1);
    let targets = prepare_timed(route, fixtures);
    let (mut total_elapsed, mut operations) = timed_many(route, fixtures, &targets, requested);
    if mode == "calibrate" {
        let mut common_calls = requested;
        while total_elapsed < TARGET_NS && common_calls < MAX_CALLS {
            common_calls = common_calls
                .saturating_mul(2)
                .max(common_calls.saturating_add(1))
                .min(MAX_CALLS);
            let (elapsed, count) = timed_many(route, fixtures, &targets, common_calls);
            total_elapsed = elapsed;
            operations = count;
        }
        return json!({
            "schema": SCHEMA,
            "route": route,
            "mode": "calibrate",
            "common_calls": common_calls,
            "elapsed_ns": total_elapsed,
            "operations": operations,
            "target_ns": TARGET_NS,
        });
    }
    json!({
        "schema": SCHEMA,
        "route": route,
        "mode": "measure",
        "calls": requested,
        "elapsed_ns": total_elapsed,
        "operations": operations,
    })
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mode = args
        .windows(2)
        .find(|pair| pair[0] == "--mode")
        .map(|pair| pair[1].as_str())
        .unwrap_or("list");
    let route = args
        .windows(2)
        .find(|pair| pair[0] == "--route")
        .map(|pair| pair[1].as_str())
        .unwrap_or("*");
    let calls = args
        .windows(2)
        .find(|pair| pair[0] == "--calls")
        .and_then(|pair| pair[1].parse::<usize>().ok())
        .unwrap_or(1);
    if mode == "list" {
        println!("{}", mode_value(mode, route, &[], calls));
        return;
    }
    assert!(ROUTES.contains(&route), "unknown route {route}");
    let root = fixture_root();
    let fixtures = load_fixtures(&root);
    let value = mode_value(mode, route, &fixtures, calls);
    println!("{value}");
}
