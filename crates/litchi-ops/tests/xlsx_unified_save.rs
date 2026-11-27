//! XLSX creation and filesystem publication through the unified `File` facade.
//!
//! The package is generated here so the tests exercise both owner states with
//! one small data workbook and unrelated ZIP members whose bytes must stay
//! intact during no-op and cell-edit publication.
#![cfg(feature = "xlsx")]

use std::io::{self, Cursor, Read as _, Write as _};
use std::path::{Path, PathBuf};

#[cfg(feature = "encryption")]
use litchi_core::Durability;
use litchi_core::FileFormat;
use litchi_ops::{
    Code, File, OpenOptions, Overwrite, SaveOptions, SaveTo, SourceMode, create, props,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const SIGNATURE_ORIGIN_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const OPAQUE_MEMBER: &str = "xl/customXml/item1.xml";
const OPAQUE_BYTES: &[u8] = b"opaque XLSX member remains byte exact\0\xff";
const OPAQUE_NON_PART_MEMBER: &str = "producer-extensionless";
const OPAQUE_NON_PART_BYTES: &[u8] = b"opaque non-part producer payload\0\xff";

fn zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in entries {
        writer
            .start_file(*name, options)
            .expect("start fixture member");
        writer.write_all(bytes).expect("write fixture member");
    }
    writer.finish().expect("finish fixture package");
    output.into_inner()
}

fn package(signed: bool) -> Vec<u8> {
    let signature_content_types = if signed {
        r#"<Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/><Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>"#
    } else {
        ""
    };
    let signature_relationship = if signed {
        format!(
            r#"<Relationship Id="rSig" Type="{SIGNATURE_ORIGIN_REL}" Target="_xmlsignatures/origin.sigs"/>"#
        )
    } else {
        String::new()
    };
    let signature_origin_relationships = if signed {
        Some(format!(
            r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rSig1" Type="{SIGNATURE_REL}" Target="sig1.xml"/></Relationships>"#
        ))
    } else {
        None
    };
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>{signature_content_types}</Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="{PACKAGE_REL}/metadata/core-properties" Target="docProps/core.xml"/>{signature_relationship}</Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    );
    let workbook_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#
    );
    let worksheet = format!(
        r#"<worksheet xmlns="{MAIN}"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Score</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Ada</t></is></c><c r="B2"><v>42</v></c></row></sheetData></worksheet>"#
    );
    let styles = format!(
        r#"<styleSheet xmlns="{MAIN}"><numFmts count="0"/><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
    );
    let core = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Unified XLSX save fixture</dc:title></cp:coreProperties>"#;

    let mut entries: Vec<(&str, &[u8])> = vec![
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_relationships.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        ),
        ("xl/worksheets/sheet1.xml", worksheet.as_bytes()),
        ("xl/styles.xml", styles.as_bytes()),
        ("docProps/core.xml", core.as_bytes()),
        (OPAQUE_MEMBER, OPAQUE_BYTES),
        (OPAQUE_NON_PART_MEMBER, OPAQUE_NON_PART_BYTES),
    ];
    let origin_relationships;
    if let Some(relationships) = signature_origin_relationships {
        origin_relationships = relationships;
        entries.extend([
            (
                "_xmlsignatures/_rels/origin.sigs.rels",
                origin_relationships.as_bytes(),
            ),
            ("_xmlsignatures/origin.sigs", b"<Origin/>"),
            ("_xmlsignatures/sig1.xml", b"<Signature/>"),
        ]);
    }
    zip(&entries)
}

fn fixture() -> Vec<u8> {
    package(false)
}

fn signed_fixture() -> Vec<u8> {
    package(true)
}

fn open_bytes_owner(bytes: Vec<u8>, eager: bool) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager XLSX fixture")
    } else {
        litchi_ops::open_bytes(bytes).expect("source-backed XLSX fixture")
    }
}

fn open_path_owner(path: &Path, eager: bool) -> File {
    if eager {
        litchi_ops::open_with(path, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager path XLSX fixture")
    } else {
        litchi_ops::open(path).expect("source-backed path XLSX fixture")
    }
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open XLSX ZIP");
    let mut member = archive.by_name(name).expect("ZIP member present");
    let mut contents = Vec::new();
    member.read_to_end(&mut contents).expect("read ZIP member");
    contents
}

fn assert_code(error: litchi_ops::Error, expected: Code) {
    assert_eq!(
        error.code(),
        expected,
        "unexpected XLSX save error: {error}"
    );
}

fn assert_data(file: &File) {
    assert_eq!(
        file.get("/data/a2")
            .expect("read fixture cell")
            .view()
            .text()
            .expect("read cell text")
            .as_deref(),
        Some("Ada")
    );
}

fn saved_path(path: &Path) -> PathBuf {
    std::path::absolute(path).expect("absolute temporary path")
}

struct FailingSink;

impl io::Write for FailingSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn create_xlsx_is_eager_deterministic_and_reopens() {
    let first = create(FileFormat::Xlsx).expect("create XLSX");
    let second = create(FileFormat::Xlsx).expect("create second XLSX");
    let File::Xlsx(xlsx) = &first else {
        panic!("creator returned another format: {first:?}");
    };
    let debug = format!("{xlsx:?}");
    assert!(
        debug.contains("source_backed: false"),
        "created XLSX must own an eager snapshot: {debug}"
    );
    let first_bytes = first.to_bytes().expect("publish created XLSX");
    assert_eq!(
        first_bytes,
        second.to_bytes().expect("publish second created XLSX")
    );
    let reopened = litchi_ops::open_bytes(first_bytes).expect("reopen created XLSX");
    assert_eq!(reopened.format(), Some(FileFormat::Xlsx));
}

#[test]
fn source_and_eager_write_to_match_and_preserve_opaque_members() {
    let bytes = fixture();
    let native_bytes = litchi_xlsx::Package::from_bytes(&bytes)
        .expect("open fixture through the native XLSX package")
        .to_bytes()
        .expect("serialize fixture through the native XLSX workbook");
    let mut outputs = Vec::new();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let file = open_bytes_owner(bytes.clone(), eager);
        assert_data(&file);
        let direct = file
            .to_bytes()
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        let mut written = Vec::new();
        let count = file
            .write_to(&mut written)
            .unwrap_or_else(|error| panic!("{owner} write_to: {error}"));
        assert_eq!(count, written.len() as u64, "{owner} byte count");
        assert_eq!(
            written, direct,
            "{owner} write_to differs from direct publish"
        );
        assert_eq!(member(&written, OPAQUE_MEMBER), OPAQUE_BYTES, "{owner}");
        assert_eq!(
            member(&written, OPAQUE_NON_PART_MEMBER),
            OPAQUE_NON_PART_BYTES,
            "{owner} non-part"
        );
        if eager {
            assert_eq!(
                direct, native_bytes,
                "eager publication diverges from the native XLSX writer"
            );
        }
        outputs.push(direct);
    }
    assert_eq!(
        outputs[0], outputs[1],
        "source/eager no-op publication differs"
    );
    assert_eq!(outputs[0], bytes, "exact no-op changed XLSX ZIP bytes");
}

#[test]
fn cell_edit_preserves_opaque_parts_and_non_part_members() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let original = open_bytes_owner(bytes.clone(), eager);
        let updated = original
            .set("/Data/B2", props! { "value" => 43 })
            .unwrap_or_else(|error| panic!("{owner} cell edit: {error}"));
        let output = updated
            .to_bytes()
            .unwrap_or_else(|error| panic!("{owner} edited publication: {error}"));
        assert_eq!(member(&output, OPAQUE_MEMBER), OPAQUE_BYTES, "{owner} part");
        assert_eq!(
            member(&output, OPAQUE_NON_PART_MEMBER),
            OPAQUE_NON_PART_BYTES,
            "{owner} non-part"
        );
        assert_eq!(
            updated
                .get("/Data/B2")
                .expect("edited cell")
                .view()
                .text()
                .expect("edited cell text")
                .as_deref(),
            Some("43"),
            "{owner} cell value"
        );
        assert_eq!(
            original
                .to_bytes()
                .unwrap_or_else(|error| panic!("{owner} original publication: {error}")),
            bytes,
            "{owner} input snapshot changed"
        );
    }

    for (owner, eager) in [("signed source", false), ("signed eager", true)] {
        let signed = open_bytes_owner(signed_fixture(), eager);
        let unsigned = match signed {
            File::Signed(signed) => signed.unsign(),
            other => panic!("{owner} expected signed XLSX, got {other:?}"),
        };
        let edited = unsigned
            .set("/Data/B2", props! { "value" => 43 })
            .unwrap_or_else(|error| panic!("{owner} cell edit: {error}"));
        let output = edited
            .to_bytes()
            .unwrap_or_else(|error| panic!("{owner} edited publication: {error}"));
        assert_eq!(member(&output, OPAQUE_MEMBER), OPAQUE_BYTES, "{owner} part");
        assert_eq!(
            member(&output, OPAQUE_NON_PART_MEMBER),
            OPAQUE_NON_PART_BYTES,
            "{owner} non-part"
        );
        let archive = zip::ZipArchive::new(Cursor::new(&output)).expect("open edited signed ZIP");
        assert!(
            archive.file_names().all(|name| !name
                .expect("valid ZIP member name")
                .starts_with("_xmlsignatures/")),
            "{owner} signature members remain after explicit removal edit"
        );
        assert!(!edited.is_signed(), "{owner} edit remains signed");
    }
}

#[test]
fn write_to_sink_failure_is_typed_and_does_not_mutate_owner() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let file = open_bytes_owner(bytes.clone(), eager);
        let error = file
            .write_to(FailingSink)
            .expect_err("write_to must report a failing sink");
        assert_code(error, Code::IoError);
        assert_eq!(
            file.to_bytes()
                .unwrap_or_else(|error| panic!("{owner} reread after sink failure: {error}")),
            bytes,
            "{owner} owner changed after a failed sink write"
        );
    }
}

#[test]
fn path_origin_save_and_save_as_keep_origin_immutable() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let directory = tempfile::tempdir().expect("save directory");
        let origin = directory.path().join(format!("{owner}-origin.xlsx"));
        std::fs::write(&origin, &bytes).expect("write XLSX origin");
        let file = open_path_owner(&origin, eager);

        let saved = file
            .save()
            .unwrap_or_else(|error| panic!("{owner} save origin: {error}"));
        assert_eq!(saved.path, saved_path(&origin));
        assert_eq!(std::fs::read(&origin).expect("read origin"), bytes);

        let copy = directory.path().join(format!("{owner}-copy.xlsx"));
        let copied = file
            .save_as(&copy)
            .unwrap_or_else(|error| panic!("{owner} save_as: {error}"));
        assert_eq!(copied.path, saved_path(&copy));
        assert_eq!(std::fs::read(&copy).expect("read copy"), bytes);
        assert_eq!(
            std::fs::read(&origin).expect("read unchanged origin"),
            bytes
        );
        assert_eq!(
            member(&std::fs::read(&copy).unwrap(), OPAQUE_MEMBER),
            OPAQUE_BYTES
        );
    }
}

#[test]
fn byte_handles_require_explicit_destination_and_family_matching() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open_bytes_owner(bytes.clone(), eager);
        assert_code(
            file.save().expect_err("byte handle has no origin"),
            Code::FileRequired,
        );
        assert_code(
            file.save_with(SaveTo::Origin, &SaveOptions::new())
                .expect_err("origin save requires a path"),
            Code::FileRequired,
        );

        let directory = tempfile::tempdir().expect("save directory");
        let wrong = directory.path().join("wrong.docx");
        assert_code(
            file.save_as(&wrong).expect_err("cross-family extension"),
            Code::InvalidArgument,
        );
        assert!(!wrong.exists());

        let destination = directory.path().join("same.xlsx");
        file.save_as(&destination)
            .expect("explicit XLSX destination");
        assert_eq!(std::fs::read(destination).expect("read destination"), bytes);
    }
}

#[test]
fn overwrite_refuse_preserves_existing_and_replace_is_atomic_policy_default() {
    let file = open_bytes_owner(fixture(), false);
    let directory = tempfile::tempdir().expect("overwrite directory");
    let existing = directory.path().join("existing.xlsx");
    std::fs::write(&existing, b"keep this destination").expect("write existing destination");
    assert_code(
        file.save_with(
            SaveTo::Path(&existing),
            &SaveOptions::new().overwrite(Overwrite::Refuse),
        )
        .expect_err("overwrite refusal"),
        Code::FileExists,
    );
    assert_eq!(
        std::fs::read(&existing).expect("read preserved destination"),
        b"keep this destination"
    );

    let fresh = directory.path().join("fresh.xlsx");
    file.save_with(
        SaveTo::Path(&fresh),
        &SaveOptions::new().overwrite(Overwrite::Refuse),
    )
    .expect("refuse permits a new destination");
    assert_eq!(
        std::fs::read(&fresh).expect("read fresh destination"),
        fixture()
    );

    file.save_as(&existing).expect("default replacement policy");
    assert_eq!(
        std::fs::read(&existing).expect("read replaced destination"),
        fixture()
    );
}

#[test]
fn changed_source_refuses_before_replacing_destination_but_eager_is_detached() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let directory = tempfile::tempdir().expect("source freshness directory");
        let origin = directory.path().join(format!("{owner}-origin.xlsx"));
        let destination = directory.path().join(format!("{owner}-destination.xlsx"));
        std::fs::write(&origin, &bytes).expect("write source");
        std::fs::write(&destination, b"destination before stale save").expect("write destination");
        let file = open_path_owner(&origin, eager);
        std::fs::write(&origin, b"source changed after open").expect("change source");

        let result = file.save_as(&destination);
        if eager {
            result.unwrap_or_else(|error| panic!("{owner} detached save: {error}"));
            assert_eq!(std::fs::read(&destination).unwrap(), bytes);
        } else {
            assert_code(result.expect_err("stale source save"), Code::IoError);
            assert_eq!(
                std::fs::read(&destination).unwrap(),
                b"destination before stale save"
            );
        }
    }
}

#[test]
fn signed_xlsx_save_preserves_signature_and_explicit_unsign_removes_it() {
    let original = signed_fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let directory = tempfile::tempdir().expect("signed directory");
        let signed = open_bytes_owner(original.clone(), eager);
        assert!(signed.is_signed(), "{owner} signed route");

        let preserved = directory.path().join(format!("{owner}-preserved.xlsx"));
        signed
            .save_as(&preserved)
            .unwrap_or_else(|error| panic!("{owner} signed save: {error}"));
        assert_eq!(std::fs::read(&preserved).unwrap(), original);
        assert!(litchi_ops::open(&preserved).unwrap().is_signed());

        let unsigned = match signed {
            File::Signed(signed) => signed.unsign(),
            other => panic!("expected signed XLSX, got {other:?}"),
        };
        assert!(!unsigned.is_signed());
        let output = directory.path().join(format!("{owner}-unsigned.xlsx"));
        unsigned
            .save_as(&output)
            .unwrap_or_else(|error| panic!("{owner} unsigned save: {error}"));
        let output_bytes = std::fs::read(&output).expect("read unsigned XLSX");
        let archive = zip::ZipArchive::new(Cursor::new(&output_bytes)).expect("open unsigned ZIP");
        assert!(
            archive.file_names().all(|name| !name
                .expect("valid ZIP member name")
                .starts_with("_xmlsignatures/")),
            "signature members remain after explicit unsign"
        );
        assert!(!litchi_ops::open_bytes(output_bytes).unwrap().is_signed());
    }
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_xlsx_save_requires_reencryption_and_round_trips() {
    let password = "xlsx-save-password";
    let encrypted = litchi_xlsx::Package::from_bytes(&fixture())
        .expect("open clear XLSX through the native package")
        .to_encrypted(password, litchi_xlsx::encryption::Mode::Agile)
        .expect("encrypt XLSX fixture through the native workbook API");
    let directory = tempfile::tempdir().expect("encrypted save directory");
    let origin = directory.path().join("encrypted-origin.xlsx");
    std::fs::write(&origin, &encrypted).expect("write encrypted origin");

    let eager_locked =
        litchi_ops::open_with(&origin, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("open eager encrypted origin");
    let File::Locked(eager_locked) = eager_locked else {
        panic!("expected locked eager XLSX");
    };
    let eager_unlocked = eager_locked
        .unlock(litchi_ops::Password::new(password.to_owned()))
        .expect("unlock eager origin");
    assert_code(
        eager_unlocked
            .save()
            .expect_err("origin retained but plaintext refused"),
        Code::DocumentProtected,
    );

    let locked = litchi_ops::open(&origin).expect("open encrypted XLSX origin");
    let locked = match locked {
        File::Locked(locked) => locked,
        other => panic!("expected locked XLSX, got {other:?}"),
    };
    let unlocked = locked
        .unlock(litchi_ops::Password::new(password.to_owned()))
        .expect("unlock encrypted XLSX origin");
    assert_code(
        unlocked
            .to_bytes()
            .expect_err("ordinary plaintext publication is refused"),
        Code::DocumentProtected,
    );
    assert_code(
        unlocked
            .save()
            .expect_err("ordinary origin save requires explicit re-encryption"),
        Code::DocumentProtected,
    );

    let default_destination = directory.path().join("default.xlsx");
    std::fs::write(&default_destination, b"default destination remains")
        .expect("write default destination");
    assert_code(
        unlocked
            .save_as(&default_destination)
            .expect_err("ordinary encrypted save is refused"),
        Code::DocumentProtected,
    );
    assert_eq!(
        std::fs::read(&default_destination).expect("read preserved destination"),
        b"default destination remains"
    );
    assert_eq!(
        std::fs::read(&origin).expect("read encrypted origin"),
        encrypted
    );

    let reencrypted = directory.path().join("reencrypted.xlsx");
    let saved = unlocked
        .save_with(
            SaveTo::Path(&reencrypted),
            &SaveOptions::new().reencrypt(litchi_ops::Password::new(password.to_owned())),
        )
        .expect("reencrypt XLSX");
    assert_eq!(saved.path, saved_path(&reencrypted));
    let reencrypted_bytes = std::fs::read(&reencrypted).expect("read reencrypted XLSX");
    assert_ne!(reencrypted_bytes, fixture());
    let reopened = litchi_ops::open(&reencrypted).expect("reopen reencrypted XLSX");
    let File::Locked(locked_again) = reopened else {
        panic!("reencrypted XLSX must remain locked");
    };
    let round_trip = locked_again
        .unlock(litchi_ops::Password::new(password.to_owned()))
        .expect("unlock reencrypted XLSX");
    assert_data(&round_trip);

    let weak_destination = directory.path().join("weak.xlsx");
    std::fs::write(&weak_destination, b"weak destination remains").expect("write weak destination");
    assert_code(
        unlocked
            .save_with(
                SaveTo::Path(&weak_destination),
                &SaveOptions::new()
                    .durability(Durability::FileOnly)
                    .reencrypt(litchi_ops::Password::new(password.to_owned())),
            )
            .expect_err("reencryption requires full durability"),
        Code::InvalidArgument,
    );
    assert_eq!(
        std::fs::read(&weak_destination).expect("read preserved weak destination"),
        b"weak destination remains"
    );

    let plain_destination = directory.path().join("plain-reencrypt.xlsx");
    assert_code(
        open_bytes_owner(fixture(), false)
            .save_with(
                SaveTo::Path(&plain_destination),
                &SaveOptions::new().reencrypt(litchi_ops::Password::new(password.to_owned())),
            )
            .expect_err("plain publication cannot request re-encryption"),
        Code::InvalidArgument,
    );
    assert!(!plain_destination.exists());

    let saved_origin = unlocked
        .save_with(
            SaveTo::Origin,
            &SaveOptions::new().reencrypt(litchi_ops::Password::new(password.to_owned())),
        )
        .expect("reencrypt retained origin");
    assert_eq!(saved_origin.path, saved_path(&origin));
    let reopened_origin = litchi_ops::open(&origin).expect("reopen encrypted origin");
    let File::Locked(locked_origin) = reopened_origin else {
        panic!("reencrypted origin must remain locked");
    };
    assert_data(
        &locked_origin
            .unlock(litchi_ops::Password::new(password.to_owned()))
            .expect("unlock reencrypted origin"),
    );
}

#[cfg(any(unix, windows))]
#[test]
fn prepared_eager_open_retains_the_exact_source_archive() {
    let bytes = fixture();
    let directory = tempfile::tempdir().expect("prepared fixture directory");
    let path = directory.path().join("prepared.xlsx");
    std::fs::write(&path, &bytes).expect("write prepared fixture");
    let prepared =
        litchi_ops::detect::detect_and_prepare(&path, &litchi_ops::detect::SourceLimits::default())
            .expect("prepare workbook");
    let file = litchi_ops::open_prepared_with(
        prepared,
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("adopt prepared eager workbook");
    assert_data(&file);
    assert_eq!(file.to_bytes().expect("publish exact eager source"), bytes);
    let saved = file.save().expect("prepared origin survives conversion");
    assert_eq!(saved.path, saved_path(&path));
    assert_eq!(std::fs::read(path).unwrap(), bytes);
}

#[cfg(feature = "encryption")]
#[test]
fn signed_encrypted_xlsx_capabilities_keep_output_protection() {
    use litchi_ops::capabilities::{Support, Verb};
    let encrypted = litchi_xlsx::Package::from_bytes(&signed_fixture())
        .unwrap()
        .to_encrypted("protected", litchi_xlsx::encryption::Mode::Agile)
        .unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(encrypted).unwrap() else {
        panic!("expected locked workbook");
    };
    let signed = locked
        .unlock(litchi_ops::Password::new("protected".to_owned()))
        .unwrap();
    assert!(signed.is_signed());
    assert!(matches!(
        signed.capabilities().verb(Verb::Save),
        Support::Unsupported {
            code: Code::DocumentProtected,
            ..
        }
    ));
    assert_eq!(
        signed.to_bytes().unwrap_err().code(),
        Code::DocumentProtected
    );
}
