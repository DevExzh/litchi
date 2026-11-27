//! DOCX filesystem publication through the unified `File` facade.
//!
//! The fixture is intentionally generated here so each assertion controls the
//! package bytes.  It contains a direct body paragraph, a second paragraph for
//! structural edits, and an unrelated package member whose bytes must survive
//! every publication route.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};
use std::path::{Path, PathBuf};

use litchi_core::Durability;
use litchi_ops::{
    Code, File, NewNode, OpenOptions, Overwrite, Placement, SaveOptions, SaveTo, SourceMode, props,
};

const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const OPAQUE_MEMBER: &str = "word/customXml/item1.xml";

fn package(document: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/></Types>"#
    );
    let root_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let opaque = r#"<opaque xmlns="urn:litchi:opaque" marker="save-fixture"><payload>unrelated XML remains byte exact</payload></opaque>"#;

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root_relationships),
        ("word/document.xml", document),
        (OPAQUE_MEMBER, opaque),
    ] {
        zip.start_file(name, options).expect("start fixture member");
        zip.write_all(contents.as_bytes())
            .expect("write fixture member");
    }
    zip.finish().expect("finish fixture package").into_inner()
}

fn fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
}

fn open_bytes_owner(bytes: Vec<u8>, eager: bool) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager DOCX fixture")
    } else {
        litchi_ops::open_bytes(bytes).expect("source-backed DOCX fixture")
    }
}

fn open_path_owner(path: &Path, eager: bool) -> File {
    if eager {
        litchi_ops::open_with(path, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager path DOCX fixture")
    } else {
        litchi_ops::open(path).expect("source-backed path DOCX fixture")
    }
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name(name).expect("ZIP member present");
    let mut contents = Vec::new();
    member.read_to_end(&mut contents).expect("read ZIP member");
    contents
}

fn entries(directory: &Path) -> Vec<String> {
    let mut entries: Vec<_> = std::fs::read_dir(directory)
        .expect("read temporary directory")
        .map(|entry| {
            entry
                .expect("directory entry")
                .file_name()
                .to_string_lossy()
                .into_owned()
        })
        .collect();
    entries.sort();
    entries
}

fn assert_code(error: litchi_ops::Error, expected: Code) {
    assert_eq!(error.code(), expected, "unexpected save error: {error}");
}

fn assert_bytes(path: &Path, expected: &[u8]) {
    assert_eq!(
        std::fs::read(path).expect("read saved package"),
        expected,
        "saved package bytes"
    );
}

fn node_text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .expect("text value")
        .into_owned()
}

fn saved_path(path: &Path) -> PathBuf {
    std::path::absolute(path).expect("absolute temporary path")
}

#[test]
fn save_debug_is_content_free_for_destinations_and_results() {
    let destination = PathBuf::from("/secret/customer-report.docx");
    assert_eq!(format!("{:?}", SaveTo::Origin), "SaveTo::Origin");
    assert_eq!(
        format!("{:?}", SaveTo::Path(&destination)),
        "SaveTo::Path(..)"
    );

    let directory = tempfile::tempdir().expect("debug save directory");
    let actual = directory.path().join("published.docx");
    let saved = open_bytes_owner(fixture(), false)
        .save_as(&actual)
        .expect("publish for Saved debug");
    let debug = format!("{saved:?}");
    assert!(debug.starts_with("Saved { bytes: "));
    assert!(debug.ends_with(", .. }"));
    assert!(!debug.contains(actual.to_string_lossy().as_ref()));
}

#[test]
fn byte_handles_have_no_origin_and_save_as_is_explicit() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let file = open_bytes_owner(bytes.clone(), eager);
        assert_code(
            file.save().expect_err("bytes handle has no origin"),
            Code::FileRequired,
        );
        assert_code(
            file.save_with(SaveTo::Origin, &SaveOptions::new())
                .expect_err("origin save requires a path"),
            Code::FileRequired,
        );

        let directory = tempfile::tempdir().expect("save directory");
        let destination = directory.path().join(format!("{owner}.docx"));
        let saved = file
            .save_as(&destination)
            .unwrap_or_else(|error| panic!("{owner} save_as: {error}"));
        assert_eq!(saved.path, saved_path(&destination));
        assert_eq!(saved.bytes, bytes.len() as u64);
        assert_bytes(&destination, &bytes);
        assert_eq!(
            member(&bytes, OPAQUE_MEMBER),
            member(
                &std::fs::read(&destination).expect("saved bytes"),
                OPAQUE_MEMBER
            )
        );
        assert_eq!(entries(directory.path()), vec![format!("{owner}.docx")]);
    }
}

#[test]
fn path_handles_save_and_save_as_keeps_the_origin_immutable() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let directory = tempfile::tempdir().expect("save directory");
        let origin = directory.path().join(format!("{owner}-origin.docx"));
        std::fs::write(&origin, &bytes).expect("write origin");
        let file = open_path_owner(&origin, eager);

        let saved = file
            .save()
            .unwrap_or_else(|error| panic!("{owner} save origin: {error}"));
        assert_eq!(saved.path, saved_path(&origin));
        assert_eq!(saved.bytes, bytes.len() as u64);
        assert_bytes(&origin, &bytes);

        let copied = directory.path().join(format!("{owner}-copy.docx"));
        let copied_result = file
            .save_as(&copied)
            .unwrap_or_else(|error| panic!("{owner} save_as: {error}"));
        assert_eq!(copied_result.path, saved_path(&copied));
        assert_bytes(&copied, &bytes);
        assert_bytes(&origin, &bytes);

        // Saving the original handle again still names the original path.
        let repeated = file
            .save()
            .unwrap_or_else(|error| panic!("{owner} repeated origin save: {error}"));
        assert_eq!(repeated.path, saved_path(&origin));
        assert_bytes(&origin, &bytes);
        assert_bytes(&copied, &bytes);
    }
}

#[test]
fn prepared_relative_path_retains_absolute_origin_through_an_edit() {
    let cwd = std::env::current_dir().expect("working directory");
    let directory = tempfile::tempdir_in(&cwd).expect("relative-origin directory");
    for eager in [false, true] {
        let origin = directory.path().join(format!("prepared-{eager}.docx"));
        std::fs::write(&origin, fixture()).expect("write prepared source");
        let relative = origin.strip_prefix(&cwd).expect("relative origin");
        let detection = litchi_ops::detect::detect_and_prepare(
            relative,
            &litchi_ops::detect::SourceLimits::default(),
        )
        .expect("prepare path");
        let options = if eager {
            OpenOptions::new().source_mode(SourceMode::Eager)
        } else {
            OpenOptions::new()
        };
        let original =
            litchi_ops::open_prepared_with(detection, &options).expect("adopt prepared path");
        let edited = original
            .set("/body/p[1]", props! { "text" => "Prepared edit" })
            .expect("edit prepared source");
        let saved = edited.save().expect("save to prepared origin");
        assert_eq!(saved.path, origin);
        assert!(saved.path.is_absolute());
        assert_eq!(
            node_text(&litchi_ops::open(&origin).expect("reopen"), "/body/p[1]"),
            "Prepared edit"
        );
        assert_eq!(node_text(&original, "/body/p[1]"), "Original");
    }
}

#[test]
fn set_and_structural_edits_publish_from_both_owners_without_mutating_clones() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let original = open_bytes_owner(bytes.clone(), eager);
        let clone = original.clone();
        let edited = original
            .set("/body/p[1]", props! { "text" => "Changed" })
            .unwrap_or_else(|error| panic!("{owner} set: {error}"));
        let added = edited
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Added")
                    .at(Placement::End),
            )
            .unwrap_or_else(|error| panic!("{owner} add: {error}"))
            .into_file();

        let directory = tempfile::tempdir().expect("edit save directory");
        let destination = directory.path().join(format!("{owner}-edited.docx"));
        let first = added
            .save_as(&destination)
            .unwrap_or_else(|error| panic!("{owner} edited save: {error}"));
        let first_bytes = std::fs::read(&destination).expect("read edited save");
        assert_eq!(first.bytes, first_bytes.len() as u64);
        assert_eq!(
            member(&first_bytes, OPAQUE_MEMBER),
            member(&bytes, OPAQUE_MEMBER)
        );

        let repeated_destination = directory.path().join(format!("{owner}-repeat.docx"));
        added
            .save_as(&repeated_destination)
            .unwrap_or_else(|error| panic!("{owner} repeated edited save: {error}"));
        assert_bytes(&repeated_destination, &first_bytes);

        let reopened = litchi_ops::open(&destination).expect("reopen edited DOCX");
        assert_eq!(
            reopened.text().expect("edited text"),
            "Changed\nSecond\nAdded"
        );
        assert_eq!(original.to_bytes().expect("original bytes"), bytes);
        assert_eq!(clone.to_bytes().expect("clone bytes"), bytes);
    }
}

#[test]
fn default_and_explicit_durability_options_publish_identical_bytes() {
    let bytes = fixture();
    let file = open_bytes_owner(bytes.clone(), false);
    let directory = tempfile::tempdir().expect("durability directory");

    let default_destination = directory.path().join("default.docx");
    let default_options = SaveOptions::new();
    let saved = file
        .save_with(SaveTo::Path(&default_destination), &default_options)
        .expect("default full durability save");
    assert_eq!(saved.path, saved_path(&default_destination));
    assert_bytes(&default_destination, &bytes);

    for durability in [Durability::Full, Durability::FileOnly, Durability::NoSync] {
        let destination = directory
            .path()
            .join(format!("{}.docx", durability.as_str()));
        let options = SaveOptions::new().durability(durability);
        file.save_with(SaveTo::Path(&destination), &options)
            .unwrap_or_else(|error| panic!("{durability:?} save: {error}"));
        assert_bytes(&destination, &bytes);
    }
    assert_eq!(
        entries(directory.path()),
        [
            "default.docx",
            "file-only.docx",
            "full.docx",
            "no-sync.docx"
        ]
    );
}

#[test]
fn overwrite_refuse_preserves_existing_destination_and_new_targets_work() {
    let bytes = fixture();
    let file = open_bytes_owner(bytes.clone(), false);
    let directory = tempfile::tempdir().expect("overwrite directory");
    let existing = directory.path().join("existing.docx");
    std::fs::write(&existing, b"keep this destination").expect("write existing destination");
    let refuse = SaveOptions::new().overwrite(Overwrite::Refuse);
    let error = file
        .save_with(SaveTo::Path(&existing), &refuse)
        .expect_err("overwrite refusal");
    assert_code(error, Code::FileExists);
    assert_bytes(&existing, b"keep this destination");
    assert_eq!(entries(directory.path()), vec!["existing.docx"]);

    let fresh = directory.path().join("fresh.docx");
    file.save_with(SaveTo::Path(&fresh), &refuse)
        .expect("refuse option permits a new target");
    assert_bytes(&fresh, &bytes);
    file.save_as(&existing).expect("default replace option");
    assert_bytes(&existing, &bytes);
    assert_eq!(
        entries(directory.path()),
        vec!["existing.docx", "fresh.docx"]
    );
}

#[test]
fn cross_family_extensions_are_refused_without_creating_files() {
    let bytes = fixture();
    let file = open_bytes_owner(bytes.clone(), false);
    let directory = tempfile::tempdir().expect("extension directory");
    for extension in ["xlsx", "XLS", "pptx", "doc", "odt", "rtf"] {
        let destination = directory.path().join(format!("wrong.{extension}"));
        let error = file
            .save_as(&destination)
            .expect_err("cross-family extension");
        assert_code(error, Code::InvalidArgument);
        assert!(!destination.exists(), "{extension} destination was created");
    }

    for name in [
        "same.docx",
        "same.DOCM",
        "same.dotx",
        "same.bin",
        "no_extension",
    ] {
        let destination = directory.path().join(name);
        file.save_as(&destination)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_bytes(&destination, &bytes);
    }
}

#[test]
fn symlink_directory_and_missing_parent_destinations_fail_safely() {
    let bytes = fixture();
    let file = open_bytes_owner(bytes, false);
    let directory = tempfile::tempdir().expect("destination directory");

    let folder = directory.path().join("folder.docx");
    std::fs::create_dir(&folder).expect("create destination directory");
    assert_code(
        file.save_as(&folder).expect_err("directory is not a file"),
        Code::IoError,
    );
    assert!(folder.is_dir());
    assert!(entries(&folder).is_empty());

    let missing = directory.path().join("missing").join("out.docx");
    assert_code(
        file.save_as(&missing)
            .expect_err("missing parent is not created implicitly"),
        Code::FileNotFound,
    );
    assert!(!missing.parent().expect("missing parent").exists());

    #[cfg(unix)]
    {
        let target = directory.path().join("target.docx");
        let link = directory.path().join("link.docx");
        std::fs::write(&target, b"link target remains").expect("write link target");
        std::os::unix::fs::symlink(&target, &link).expect("create destination symlink");
        assert_code(
            file.save_as(&link)
                .expect_err("symlink destination is refused"),
            Code::IoError,
        );
        assert!(
            std::fs::symlink_metadata(&link)
                .expect("link metadata")
                .file_type()
                .is_symlink()
        );
        assert_bytes(&target, b"link target remains");
    }
}

#[test]
fn changed_path_source_refuses_before_replacing_destination() {
    let bytes = fixture();
    for (owner, eager) in [("source", false), ("eager", true)] {
        let directory = tempfile::tempdir().expect("stale source directory");
        let origin = directory.path().join(format!("{owner}-source.docx"));
        let destination = directory.path().join(format!("{owner}-destination.docx"));
        std::fs::write(&origin, &bytes).expect("write source");
        std::fs::write(&destination, b"destination before stale save").expect("write destination");
        let file = open_path_owner(&origin, eager);
        std::fs::write(&origin, b"source changed after open").expect("change source");

        let result = file.save_as(&destination);
        if eager {
            // An eager owner is detached from the path's later contents.  It
            // publishes its captured snapshot even after the origin changes.
            result.expect("eager snapshot remains publishable after source change");
            assert_bytes(&destination, &bytes);
        } else {
            let error = result.expect_err("stale source save");
            assert_code(error, Code::IoError);
            assert_bytes(&destination, b"destination before stale save");
        }
        assert_eq!(
            entries(directory.path()),
            vec![
                format!("{owner}-destination.docx"),
                format!("{owner}-source.docx")
            ]
        );
    }
}

#[test]
fn signed_noop_preserves_signature_and_explicit_unsign_publishes_unsigned_docx() {
    let signed_fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/xmldsign/hello-world-signed.docx");
    let original = std::fs::read(&signed_fixture).expect("read signed fixture");

    for (owner, eager) in [("source", false), ("eager", true)] {
        let directory = tempfile::tempdir().expect("signed directory");
        let origin = directory.path().join(format!("{owner}-signed.docx"));
        std::fs::write(&origin, &original).expect("write signed source");
        let signed = open_path_owner(&origin, eager);
        assert!(signed.is_signed(), "{owner} signed route");

        let preserved = directory.path().join(format!("{owner}-preserved.docx"));
        signed
            .save_as(&preserved)
            .unwrap_or_else(|error| panic!("{owner} signed save: {error}"));
        assert_bytes(&preserved, &original);
        assert!(
            litchi_ops::open(&preserved)
                .expect("reopen signed copy")
                .is_signed()
        );

        let unsigned = match signed {
            File::Signed(signed) => signed.unsign(),
            other => panic!("expected signed DOCX, got {other:?}"),
        };
        assert!(!unsigned.is_signed());
        // Explicit consumption takes effect at publication even without an
        // intervening edit. Repeated saves share the successful removal
        // snapshot while leaving the input content unchanged.
        let unedited_output = directory
            .path()
            .join(format!("{owner}-unsigned-noedit.docx"));
        unsigned
            .save_as(&unedited_output)
            .expect("publish signature removal");
        let first_unsigned = std::fs::read(&unedited_output).expect("read unsigned copy");
        unsigned
            .save_as(&unedited_output)
            .expect("repeat signature removal publication");
        assert_eq!(
            std::fs::read(&unedited_output).expect("repeat bytes"),
            first_unsigned
        );
        assert!(
            !litchi_ops::open_bytes(first_unsigned)
                .expect("reopen unsigned copy")
                .is_signed()
        );
        let edited = unsigned
            .set("/body/p[1]", props! { "text" => "Explicitly unsigned" })
            .unwrap_or_else(|error| panic!("{owner} explicit unsign edit: {error}"));
        let output = directory.path().join(format!("{owner}-unsigned.docx"));
        edited
            .save_as(&output)
            .unwrap_or_else(|error| panic!("{owner} unsigned save: {error}"));
        let output_bytes = std::fs::read(&output).expect("read unsigned output");
        let archive = zip::ZipArchive::new(Cursor::new(&output_bytes)).expect("open unsigned ZIP");
        assert!(
            archive.file_names().all(|name| !name
                .expect("valid ZIP member name")
                .starts_with("_xmlsignatures/")),
            "signature members remain after explicit unsign"
        );
        let reopened = litchi_ops::open_bytes(output_bytes).expect("reopen unsigned output");
        assert!(!reopened.is_signed());
        assert_eq!(node_text(&reopened, "/body/p[1]"), "Explicitly unsigned");
    }
}

#[cfg(feature = "encryption")]
#[test]
fn locked_save_refuses_before_destination_policy_validation() {
    let encrypted = litchi_crypto::ooxml::encrypt(
        fixture(),
        "locked-save",
        litchi_crypto::ooxml::Mode::Agile,
        &litchi_crypto::ooxml::Limits::default(),
    )
    .expect("encrypt DOCX fixture");
    let file = litchi_ops::open_bytes(encrypted).expect("open encrypted DOCX");
    assert!(matches!(file, File::Locked(_)));

    let directory = tempfile::tempdir().expect("locked save directory");
    let wrong_family = directory.path().join("locked.xlsx");
    assert_code(
        file.save_as(&wrong_family)
            .expect_err("locked save refuses before extension validation"),
        Code::DocumentProtected,
    );
    assert!(!wrong_family.exists());

    let existing = directory.path().join("locked.docx");
    std::fs::write(&existing, b"keep locked destination").expect("write locked destination");
    assert_code(
        file.save_with(
            SaveTo::Path(&existing),
            &SaveOptions::new().overwrite(Overwrite::Refuse),
        )
        .expect_err("locked save refuses before overwrite validation"),
        Code::DocumentProtected,
    );
    assert_bytes(&existing, b"keep locked destination");
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_docx_save_requires_reencryption_and_preserves_origin() {
    let password = "save-password";
    let encrypted = litchi_crypto::ooxml::encrypt(
        fixture(),
        password,
        litchi_crypto::ooxml::Mode::Agile,
        &litchi_crypto::ooxml::Limits::default(),
    )
    .expect("encrypt DOCX fixture");
    let directory = tempfile::tempdir().expect("encrypted save directory");
    let origin = directory.path().join("encrypted-origin.docx");
    std::fs::write(&origin, &encrypted).expect("write encrypted origin");

    let locked = litchi_ops::open(&origin).expect("open encrypted origin");
    let locked = match locked {
        File::Locked(locked) => locked,
        other => panic!("expected locked DOCX, got {other:?}"),
    };
    let unlocked = locked
        .unlock(litchi_ops::Password::new(password.to_owned()))
        .expect("unlock encrypted origin");
    assert_code(
        unlocked
            .save()
            .expect_err("origin is retained but encrypted"),
        Code::DocumentProtected,
    );

    let default_destination = directory.path().join("default.docx");
    std::fs::write(&default_destination, b"default destination remains")
        .expect("write default destination");
    assert_code(
        unlocked
            .save_as(&default_destination)
            .expect_err("encrypted publication needs an explicit password"),
        Code::DocumentProtected,
    );
    assert_bytes(&default_destination, b"default destination remains");
    assert_bytes(&origin, &encrypted);

    let reencrypted = directory.path().join("reencrypted.docx");
    let saved = unlocked
        .save_with(
            SaveTo::Path(&reencrypted),
            &SaveOptions::new().reencrypt(litchi_ops::Password::new(password.to_owned())),
        )
        .expect("reencrypt encrypted DOCX");
    assert_eq!(saved.path, saved_path(&reencrypted));
    assert_ne!(
        std::fs::read(&reencrypted).expect("read reencrypted output"),
        fixture()
    );
    let reopened = litchi_ops::open(&reencrypted).expect("reopen reencrypted output");
    let File::Locked(locked_again) = reopened else {
        panic!("reencrypted output must remain locked");
    };
    let round_trip = locked_again
        .unlock(litchi_ops::Password::new(password.to_owned()))
        .expect("unlock reencrypted output");
    assert_eq!(
        round_trip.text().expect("round-trip text"),
        "Original\nSecond"
    );

    let weak_destination = directory.path().join("weak.docx");
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
    assert_bytes(&weak_destination, b"weak destination remains");

    let plain_destination = directory.path().join("plain-reencrypt.docx");
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
        .expect("reencrypt at retained origin");
    assert_eq!(saved_origin.path, saved_path(&origin));
    assert!(matches!(
        litchi_ops::open(&origin).expect("reopen encrypted origin"),
        File::Locked(_)
    ));
}
