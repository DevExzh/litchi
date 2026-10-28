//! The signature policy audit (record 0013): every signed input a `File`
//! opens follows ADR 0006's explicit policy, on every open route.
//!
//! For each signed fixture and each route:
//!
//! 1. opening yields [`File::Signed`];
//! 2. the editor builder of the wrapped file refuses with
//!    `document_protected` and a suggestion naming `Signed::unsign`;
//! 3. after [`Signed::unsign`](crate::Signed::unsign), a changed commit
//!    through the editor the builder returns publishes the edit, the output
//!    carries no signature part, and the clone that was not unsigned still
//!    refuses;
//! 4. exact no-ops created through both the signed leaf and the unified
//!    facade replay only onto an independently opened recipient with the same
//!    published bytes; an authorized unsigned no-op stays unsigned.
//!
//! The DOCX, PPTX and XLSX fixtures are producer-signed (Apache POI's
//! `xmldsign` corpus). No signed XLSB or compound-file fixture exists, so
//! those inputs are signed here: an OPC signature origin, signature part and
//! relationships for XLSB, as `litchi-xlsb`'s own tests build them, and both
//! a legacy `_signatures` stream and an `_xmlsignatures` storage for DOC, PPT
//! and XLS, as S6's security tests mark them.

use std::io::Cursor;
use std::sync::Arc;

#[cfg(any(feature = "doc", feature = "docx", feature = "ppt", feature = "xls"))]
use litchi_core::Position;
use litchi_core::{FileFormat, OwnedSource, ReadAt};

use super::test_data;
use crate::{Code, Error, File, OpenOptions, SignatureOrigin, SourceMode};

/// The text a changed OOXML commit writes, so the reopened output can be
/// checked. The compound-file editors only replace text of the same width.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
const MARKER: &str = "EditedAfterUnsign";

/// One signed input.
struct Input {
    label: String,
    format: FileFormat,
    origin: SignatureOrigin,
    bytes: Vec<u8>,
    /// The file-name suffix of the path routes.
    suffix: &'static str,
}

fn read(relative: &str) -> Vec<u8> {
    std::fs::read(test_data(relative)).unwrap_or_else(|error| panic!("{relative}: {error}"))
}

/// Every signed input a `File` can open in this build.
fn inputs() -> Vec<Input> {
    let mut inputs = Vec::new();
    let xmldsign = |format: FileFormat, name: &str, suffix: &'static str| Input {
        label: name.to_owned(),
        format,
        origin: SignatureOrigin::Opc,
        bytes: read(&format!("poi/test-data/xmldsign/{name}")),
        suffix,
    };
    #[cfg(feature = "docx")]
    for name in [
        "hello-world-signed.docx",
        "hello-world-signed-twice.docx",
        "ms-office-2010-signed.docx",
    ] {
        inputs.push(xmldsign(FileFormat::Docx, name, ".docx"));
    }
    #[cfg(feature = "pptx")]
    for name in ["hello-world-signed.pptx", "ms-office-2010-signed.pptx"] {
        inputs.push(xmldsign(FileFormat::Pptx, name, ".pptx"));
    }
    #[cfg(feature = "xlsx")]
    for name in ["hello-world-signed.xlsx", "ms-office-2010-signed.xlsx"] {
        inputs.push(xmldsign(FileFormat::Xlsx, name, ".xlsx"));
    }
    #[cfg(feature = "xlsb")]
    inputs.push(Input {
        label: "Simple.xlsb, signed in-test".to_owned(),
        format: FileFormat::Xlsb,
        origin: SignatureOrigin::Opc,
        bytes: signed_opc(&read("ooxml/xlsb/Simple.xlsb")),
        suffix: ".xlsb",
    });
    let _ = xmldsign;
    // The compound-file inputs are fixtures the narrow source-backed editors
    // admit; most corpus files carry records they refuse whether or not
    // they are signed.
    #[cfg(feature = "doc")]
    inputs.push(Input {
        label: "documentProperties.doc, signed in-test".to_owned(),
        format: FileFormat::Doc,
        origin: SignatureOrigin::Cfb,
        bytes: signed_compound_file(read("ole/doc/documentProperties.doc")),
        suffix: "",
    });
    #[cfg(feature = "ppt")]
    inputs.push(Input {
        label: "authored text boxes, signed in-test".to_owned(),
        format: FileFormat::Ppt,
        origin: SignatureOrigin::Cfb,
        bytes: signed_compound_file(editable_presentation()),
        suffix: "",
    });
    #[cfg(feature = "xls")]
    inputs.push(Input {
        label: "FormulaSheetRange.xls, signed in-test".to_owned(),
        format: FileFormat::Xls,
        origin: SignatureOrigin::Cfb,
        bytes: signed_compound_file(read("ole/xls/FormulaSheetRange.xls")),
        suffix: "",
    });
    inputs
}

/// A presentation whose text boxes the source-backed text editor admits, as
/// the `Ppt` handle's own tests author it.
#[cfg(feature = "ppt")]
fn editable_presentation() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    for slide_index in 0..2 {
        let slide = writer.add_slide().unwrap();
        for shape_index in 0..3 {
            let x = 10 + shape_index * 260;
            let text = format!("slide-{slide_index}-shape-{shape_index}");
            writer.add_textbox(slide, x, 10, 240, 40, &text).unwrap();
        }
    }
    let mut output = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

/// `plain` with an OPC signature graph (an origin part, one XML signature
/// part and their relationships), written out so that the signature is an
/// ingressed source signature, as in a file a producer signed.
#[cfg(feature = "xlsb")]
fn signed_opc(plain: &[u8]) -> Vec<u8> {
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};

    let source = OpcPackage::from_bytes(plain, litchi_opc::SourceLimits::default()).unwrap();
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
            .unwrap();
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).unwrap();
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
                .unwrap();
        }
        package.try_add_part(Box::new(copy)).unwrap();
    }
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
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
        .unwrap();
    package.try_add_part(Box::new(origin_part)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .unwrap();
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

/// `plain` rewritten with a legacy `_signatures` stream and an
/// `_xmlsignatures` storage holding one XML signature stream.
#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
fn signed_compound_file(plain: Vec<u8>) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source =
        OleFile::open(Cursor::new(plain), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let path = path.iter().map(String::as_str).collect::<Vec<_>>();
        for depth in 1..path.len() {
            writer.create_storage(&path[..depth]).unwrap();
        }
        let data = source.open_stream(&path).unwrap();
        writer.create_stream(&path, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"legacy signature")
        .unwrap();
    writer.create_storage(&["_xmlsignatures"]).unwrap();
    writer
        .create_stream(&["_xmlsignatures", "1"], b"<Signature/>")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// The file every route opens `input` as, with the temporary files the path
/// routes read kept alive alongside.
#[allow(
    clippy::type_complexity,
    reason = "a test helper returning labelled files and their backing files"
)]
fn routes(input: &Input) -> (Vec<(&'static str, File)>, Vec<tempfile::NamedTempFile>) {
    let eager = OpenOptions::new().source_mode(SourceMode::Eager);
    let mut files = vec![
        ("bytes", crate::open_bytes(input.bytes.clone()).unwrap()),
        (
            "bytes, eager",
            crate::open_bytes_with(input.bytes.clone(), &eager).unwrap(),
        ),
    ];
    let mut temporaries = Vec::new();
    #[cfg(any(unix, windows))]
    {
        let temporary = super::temporary_with(&input.bytes, input.suffix);
        files.push(("path", crate::open(temporary.path()).unwrap()));
        files.push((
            "path, eager",
            crate::open_with(temporary.path(), &eager).unwrap(),
        ));
        temporaries.push(temporary);
    }
    (files, temporaries)
}

/// The refusal of the editor builder of a tier 1 or tier 2 `file`.
fn builder_refusal(file: &File) -> Error {
    fn refused<T>(result: crate::Result<T>, format: &str) -> Error {
        match result {
            Ok(_) => panic!("a signed {format} built an editor"),
            Err(error) => error,
        }
    }
    match file {
        #[cfg(feature = "docx")]
        File::Docx(handle) => refused(handle.edit(), "DOCX"),
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => refused(handle.editable(), "PPTX"),
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => refused(handle.editor(), "XLSX"),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => refused(handle.editor(), "XLSB"),
        #[cfg(feature = "doc")]
        File::Doc(handle) => refused(handle.body_text_editor(), "DOC"),
        #[cfg(feature = "xls")]
        File::Xls(handle) => refused(handle.cell_values_editor(), "XLS"),
        #[cfg(feature = "ppt")]
        File::Ppt(handle) => refused(handle.editable_package(), "PPT"),
        other => panic!("{other:?} has no tier 1 or tier 2 editor"),
    }
}

#[track_caller]
fn assert_signed_refusal(error: &Error, format: FileFormat, what: &str) {
    assert_eq!(error.code(), Code::DocumentProtected, "{what}: {error:?}");
    assert_eq!(error.format(), Some(format), "{what}: {error:?}");
    let suggestion = error
        .suggestion()
        .expect("a signed refusal suggests a way out");
    assert!(
        suggestion.contains("Signed::unsign()"),
        "{what}: {suggestion}"
    );
}

/// `text` with every ASCII letter and digit changed, keeping its UTF-16
/// width and encoded length; `None` when nothing would change.
#[cfg(any(feature = "doc", feature = "ppt"))]
fn same_width(text: &str) -> Option<String> {
    let replacement: String = text
        .chars()
        .map(|character| match character {
            'x' => 'y',
            character if character.is_ascii_alphanumeric() => 'x',
            character => character,
        })
        .collect();
    (replacement != text).then_some(replacement)
}

/// Drive one changed commit through the editor the builder of the unsigned
/// `file` returns. Returns the published bytes and a text the reopened
/// output must contain.
fn changed_commit(file: &File) -> (Vec<u8>, String) {
    match file {
        #[cfg(feature = "docx")]
        File::Docx(handle) => {
            let Ok(crate::format::docx::DocxEditor::Unsigned(package)) = handle.edit() else {
                panic!("an unsigned DOCX builds an owning unsigned copy");
            };
            // Append a paragraph: Office's signed fixture keeps its only
            // paragraph in structure the text transaction does not rewrite.
            let mut edit = package.edit().unwrap();
            let count = edit.source().unwrap().paragraph_count();
            edit.insert_paragraph(Position::new(count), MARKER).unwrap();
            let commit = edit.commit().unwrap();
            let package = package.apply(commit.patch()).unwrap().into_package();
            let mut output = Vec::new();
            package.write_to(&mut output).unwrap();
            (output, MARKER.to_owned())
        },
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => {
            let Ok(crate::format::pptx::Editable::Eager(package)) = handle.editable() else {
                panic!("an unsigned PPTX builds an owning unsigned copy");
            };
            let Ok(crate::format::pptx::Editable::Eager(repeated)) = handle.editable() else {
                panic!("repeated unsigned promotion");
            };
            assert!(Arc::ptr_eq(&package, &repeated));
            let mut transaction = package.edit().unwrap();
            transaction
                .add_text_box(0_usize, MARKER, (0, 0, 914_400, 914_400))
                .unwrap();
            let commit = transaction.commit().unwrap();
            let package = commit.into_package();
            (package.to_bytes().unwrap(), MARKER.to_owned())
        },
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => {
            let Ok(workbook) = handle.editor() else {
                panic!("an unsigned XLSX builds an owning unsigned copy");
            };
            let repeated = handle.editor().expect("reuse unsigned XLSX owner");
            assert!(
                workbook
                    .workbook_ref()
                    .sheet(0_usize)
                    .unwrap()
                    .unwrap()
                    .same_workbook(&repeated.workbook_ref().sheet(0_usize).unwrap().unwrap()),
                "signature-authorized promotion is shared"
            );
            let mut edit = workbook.edit().unwrap();
            edit.sheet(0_usize)
                .unwrap()
                .expect("a first worksheet")
                .set("A1", MARKER)
                .unwrap();
            let output = edit.commit().unwrap().package().to_bytes().unwrap();
            (output, MARKER.to_owned())
        },
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => {
            use litchi_xlsb::cell_values::{Reference, StyleIndex, Value};

            let Ok(crate::format::xlsb::XlsbEditor::Materialized(package)) = handle.editor() else {
                panic!("an unsigned XLSB builds an owning unsigned copy");
            };
            let mut edit = package
                .edit_cell_values(0)
                .unwrap()
                .resolve()
                .unwrap()
                .expect("worksheet selector resolves");
            edit.insert(
                Reference::new(100, 0).unwrap(),
                StyleIndex::new(0).unwrap(),
                Value::InlineString(MARKER.to_owned()),
            )
            .unwrap();
            let commit = edit.commit().unwrap();
            let package = commit.into_package();
            let mut output = Cursor::new(Vec::new());
            package.write_to(&mut output).unwrap();
            (output.into_inner(), MARKER.to_owned())
        },
        #[cfg(feature = "doc")]
        File::Doc(handle) => {
            let snapshot = handle
                .body_text_editor()
                .unwrap()
                .expect("a retained DOC source");
            let (commit, text) = (0..256)
                .find_map(|index| {
                    let position = Position::new(index);
                    let text = same_width(snapshot.paragraph(position).ok()?.text())?;
                    let mut transaction = snapshot.edit_paragraph(position).ok()?;
                    transaction.set_text(text.clone()).ok()?;
                    Some((transaction.commit().ok()?, text))
                })
                .expect("an editable DOC paragraph");
            let mut output = Vec::new();
            commit.package().write_to(&mut output).unwrap();
            (output, text.trim().to_owned())
        },
        #[cfg(feature = "xls")]
        File::Xls(handle) => {
            use litchi_xls::cell_values::{Selector, Value};

            let package = handle.cell_values_editor().unwrap();
            let snapshot = package.cell_values().unwrap();
            let (sheet, reference) = snapshot
                .worksheets()
                .find_map(|sheet| {
                    sheet
                        .cells()
                        .find(|cell| matches!(cell.value(), Value::Number(_)))
                        .map(|cell| (sheet.position(), cell.reference()))
                })
                .expect("a numeric XLS cell");
            let mut edit = package.edit().unwrap();
            edit.set_value(
                Selector::Position(Position::new(sheet)),
                reference,
                Value::Number(4242.0),
            )
            .unwrap();
            let commit = edit.commit().unwrap();
            (commit.package().to_bytes().unwrap(), "4242".to_owned())
        },
        #[cfg(feature = "ppt")]
        File::Ppt(handle) => {
            use litchi_ppt::package::edit::Target;

            let package = handle.editable_package().unwrap();
            let presentation = package.presentation().unwrap();
            for (slide, shapes) in presentation.slides().unwrap().iter().enumerate() {
                for (shape, item) in shapes.shapes().unwrap().iter().enumerate() {
                    let Some(text) = item.text().ok().and_then(|text| same_width(&text)) else {
                        continue;
                    };
                    let target = Target::new(Position::new(slide), Position::new(shape));
                    let mut edit = package.edit().unwrap();
                    if edit.set_shape_text(target, text.clone()).is_err() {
                        continue;
                    }
                    let Ok(commit) = edit.commit() else {
                        continue;
                    };
                    let output = commit.package().to_bytes().unwrap();
                    return (output, text.trim().to_owned());
                }
            }
            panic!("no editable PPT text atom");
        },
        other => panic!("{other:?} has no tier 1 or tier 2 editor"),
    }
}

/// The published output carries no signature part, opens unsigned, and
/// shows the edit.
#[track_caller]
fn assert_unsigned_output(format: FileFormat, output: &[u8], expected: &str, what: &str) {
    match format {
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        FileFormat::Docx | FileFormat::Pptx | FileFormat::Xlsx | FileFormat::Xlsb => {
            let archive = zip::ZipArchive::new(Cursor::new(output)).unwrap();
            let stale: Vec<_> = archive
                .file_names()
                .filter(|name| name.to_ascii_lowercase().starts_with("_xmlsignatures"))
                .collect();
            assert!(
                stale.is_empty(),
                "{what}: stale signature members {stale:?}"
            );
            let package =
                litchi_opc::OpcPackage::from_bytes(output, litchi_opc::SourceLimits::default())
                    .unwrap();
            assert!(!package.is_signed(), "{what}: the output graph is signed");
        },
        #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
        FileFormat::Doc | FileFormat::Ppt | FileFormat::Xls => {
            let ole = litchi_cfb::OleFile::open(
                Cursor::new(output),
                litchi_cfb::OleFileLimits::default(),
            )
            .unwrap();
            assert!(!ole.exists(&["_signatures"]), "{what}: legacy signature");
            assert!(
                !ole.directory_exists(&["_xmlsignatures"]),
                "{what}: XML signature storage"
            );
        },
        other => panic!("{what}: {other:?} is not signed by this test"),
    }
    let reopened = crate::open_bytes(output.to_vec()).unwrap();
    assert!(!reopened.is_signed(), "{what}: reopened as signed");
    assert_eq!(reopened.format(), Some(format), "{what}");
    let text = reopened.text().unwrap();
    assert!(
        text.contains(expected),
        "{what}: {expected:?} not in {text:?}"
    );
}

/// An exact no-op through the leaf entry the editor builder reaches, on the
/// signed input, keeps it byte for byte. Changed operations remain covered by
/// each format's signed-input policy tests below.
#[track_caller]
fn assert_exact_noop(input: &Input) {
    let label = &input.label;
    let bytes = &input.bytes;
    let source = || -> Arc<dyn ReadAt> { Arc::new(OwnedSource::new(bytes.clone())) };
    let _ = &source;
    match input.format {
        #[cfg(feature = "docx")]
        FileFormat::Docx => {
            let package = litchi_docx::source_backed::Package::from_read_at(source()).unwrap();
            let commit = package.edit().unwrap().commit().unwrap();
            let mut output = Vec::new();
            commit.package().write_to(&mut output).unwrap();
            assert_eq!(&output, bytes, "{label}: source-backed no-op");

            let package = litchi_docx::Package::from_reader(Cursor::new(bytes)).unwrap();
            let commit = package.edit().unwrap().commit().unwrap();
            let package = package.apply(commit.patch()).unwrap().into_package();
            let mut output = Vec::new();
            package.write_to(&mut output).unwrap();
            assert_eq!(&output, bytes, "{label}: eager no-op");
        },
        #[cfg(feature = "pptx")]
        FileFormat::Pptx => {
            let presentation = litchi_pptx::source::Package::from_read_at(source()).unwrap();
            let commit = presentation.edit().unwrap().commit().unwrap();
            let mut output = Vec::new();
            commit.package().write_to(&mut output).unwrap();
            assert_eq!(&output, bytes, "{label}: source-backed no-op");

            let mut package = litchi_pptx::Package::from_reader(&bytes[..]).unwrap();
            let commit = package.edit().unwrap().commit().unwrap();
            assert!(!commit.diagnostics().changed(), "{label}");
            package = commit.into_package();
            assert_eq!(&package.to_bytes().unwrap(), bytes, "{label}: eager no-op");
        },
        #[cfg(feature = "xlsx")]
        FileFormat::Xlsx => {
            // Every source-backed XLSX family publishes through the same OPC
            // overlay; the workbook-level defined-names family reaches it on
            // any workbook.
            let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(source()).unwrap();
            let commit = workbook.edit_defined_names().unwrap().commit().unwrap();
            let mut output = Vec::new();
            commit.package().write_to(&mut output).unwrap();
            assert_eq!(&output, bytes, "{label}: source-backed no-op");

            // The eager semantic transaction returns a signed workbook
            // unchanged for an exact no-op (ADR 0006) and refuses a change:
            // it never strips the signature.
            let workbook = litchi_xlsx::Package::from_bytes(bytes).unwrap();
            let commit = workbook.edit().unwrap().commit().unwrap();
            assert!(!commit.diagnostics().changed(), "{label}");
            assert_eq!(
                &commit.package().to_bytes().unwrap(),
                bytes,
                "{label}: eager no-op"
            );
            let mut edit = workbook.edit().unwrap();
            edit.sheet(0_usize)
                .unwrap()
                .expect("a first worksheet")
                .set("A1", "changed")
                .unwrap();
            assert!(
                matches!(edit.commit(), Err(litchi_xlsx::Error::Signed)),
                "{label}: eager XLSX edit"
            );
        },
        #[cfg(feature = "xlsb")]
        FileFormat::Xlsb => {
            let package = litchi_xlsb::Package::from_bytes(bytes).unwrap();
            let commit = package
                .edit_cell_values(0)
                .unwrap()
                .resolve()
                .unwrap()
                .expect("worksheet selector resolves")
                .commit()
                .unwrap();
            let mut output = Cursor::new(Vec::new());
            commit.into_package().write_to(&mut output).unwrap();
            assert_eq!(&output.into_inner(), bytes, "{label}: eager no-op");
        },
        #[cfg(feature = "doc")]
        FileFormat::Doc => assert!(
            litchi_doc::source_backed::Package::from_bytes(bytes.clone()).is_err(),
            "{label}: the DOC editor admitted a signed container"
        ),
        #[cfg(feature = "xls")]
        FileFormat::Xls => {
            let package = litchi_xls::Package::from_bytes(bytes).unwrap();
            let commit = package.edit().unwrap().commit().unwrap();
            assert!(!commit.diagnostics().changed(), "{label}");
            assert!(commit.patch().is_noop(), "{label}: eager no-op patch");
            assert_eq!(
                &commit.package().to_bytes().unwrap(),
                bytes,
                "{label}: eager no-op"
            );
        },
        #[cfg(feature = "ppt")]
        FileFormat::Ppt => assert!(
            litchi_ppt::text_edit::SourceSnapshot::open(source()).is_err(),
            "{label}: the PPT editor admitted a signed container"
        ),
        other => panic!("{label}: {other:?} is not signed by this test"),
    }
}

/// Return the readable leaf below the signed facade. The leaf keeps its
/// format signature disposition, so direct setter tests exercise the same
/// source owner and publication path exposed by the signed facade.
fn signed_leaf(file: &File) -> File {
    match file {
        File::Signed(signed) => signed.file().clone(),
        other => panic!("expected a signed facade, got {other:?}"),
    }
}

#[cfg(any(feature = "xls", feature = "xlsb", feature = "xlsx"))]
fn signed_name(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .find(|property| property.key() == "name")
        .map(|property| property.value().to_string())
        .unwrap_or_else(|| panic!("{path} has no name property"))
}

/// Find a small plain-text cell for legacy workbook cancellation coverage.
/// The signed XLS fixture intentionally contains formulas and packed numeric
/// records, so the test must choose an authored scalar rather than assuming
/// that A1 is writable by the coalesced value family.
#[cfg(feature = "xls")]
fn signed_text_cell(file: &File) -> Option<(String, String)> {
    for row in 0..12 {
        for column in 0..8 {
            let mut label = String::new();
            let mut value = column + 1;
            while value > 0 {
                let digit = (value - 1) % 26;
                label.push(char::from(b'A' + u8::try_from(digit).ok()?));
                value = (value - 1) / 26;
            }
            let column = label.chars().rev().collect::<String>();
            let path = format!("/sheet[1]/{column}{}", row + 1);
            let Ok(node) = file.get(path.as_str()) else {
                continue;
            };
            let mut type_name = None;
            let mut formula = false;
            let Ok(format) = node.view().format() else {
                continue;
            };
            for property in format {
                if property.key().eq_ignore_ascii_case("type") {
                    type_name = Some(property.value().to_string());
                }
                formula |= property.key().eq_ignore_ascii_case("formula");
            }
            if formula
                || !type_name.as_deref().is_some_and(|kind| {
                    kind.eq_ignore_ascii_case("string")
                        || kind.eq_ignore_ascii_case("sharedstring")
                        || kind.eq_ignore_ascii_case("text")
                })
            {
                continue;
            }
            let Ok(text) = node.view().text() else {
                continue;
            };
            let text = text.unwrap_or_default().into_owned();
            return Some((path, text));
        }
    }
    None
}

/// Find a coordinate with no stored cell record.  The signed spreadsheet
/// tests use it to distinguish an absent cell from an authored cell whose
/// value happens to equal the native default.
#[cfg(any(feature = "xls", feature = "xlsb", feature = "xlsx"))]
fn signed_missing_cell(file: &File) -> Option<String> {
    for row in 0..32 {
        for column in 0..16 {
            let mut label = String::new();
            let mut value = column + 1;
            while value > 0 {
                let digit = (value - 1) % 26;
                label.push(char::from(b'A' + u8::try_from(digit).ok()?));
                value = (value - 1) / 26;
            }
            let column = label.chars().rev().collect::<String>();
            let path = format!("/sheet[1]/{column}{}", row + 1);
            let Ok(node) = file.get(path.as_str()) else {
                continue;
            };
            let Ok(text) = node.view().text() else {
                continue;
            };
            let missing_text = text
                .as_deref()
                .is_none_or(|text| text.is_empty() || text == "(empty)");
            let Ok(format) = node.view().format() else {
                continue;
            };
            if missing_text && format.into_iter().next().is_none() {
                return Some(path);
            }
        }
    }
    None
}

#[cfg(feature = "pptx")]
#[test]
fn signed_pptx_shape_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Pptx)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/slide[1]/shape[1]";
        let text = file
            .get(path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .unwrap()
            .into_owned();
        let props = crate::props! { "text" => text.as_str() };
        let leaf = signed_leaf(&file);

        let leaf_set = leaf
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed leaf no-op setter: {error:?}", input.label));
        assert!(leaf_set.is_signed(), "{}: leaf lost signature", input.label);
        assert_eq!(
            leaf_set.to_bytes().unwrap(),
            input.bytes,
            "{}: leaf bytes",
            input.label
        );

        let facade_set = file.set(path, props.clone()).unwrap_or_else(|error| {
            panic!("{}: signed facade no-op setter: {error:?}", input.label)
        });
        assert!(
            matches!(facade_set, File::Signed(_)),
            "{}: facade unwrapped",
            input.label
        );
        assert_eq!(
            facade_set.to_bytes().unwrap(),
            input.bytes,
            "{}: facade bytes",
            input.label
        );

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): signed setter edit: {error:?}", input.label)
                });
            assert!(!commit.diagnostics().changed(), "{} ({route})", input.label);
            assert!(
                commit.file().is_signed(),
                "{} ({route}): lost signature",
                input.label
            );
            assert_eq!(
                commit.file().to_bytes().unwrap(),
                input.bytes,
                "{} ({route})",
                input.label
            );
        }

        let changed = format!("{text} changed");
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancelling edit: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }
        // A cancellation followed by a real change must remain a signed
        // refusal. The final plan, rather than the first changed operation,
        // decides whether the original signed bytes may be retained.
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let error = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())?;
                    edit.set(path, crate::props! { "text" => changed.as_str() })
                })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Pptx,
                &format!("{} ({route}) mixed signed edit", input.label),
            );
        }
        for (route, candidate) in [("leaf", leaf), ("facade", file)] {
            let error = candidate
                .set(path, crate::props! { "text" => changed.as_str() })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Pptx,
                &format!("{} ({route})", input.label),
            );
        }
    }
}

#[cfg(feature = "docx")]
#[test]
fn signed_docx_paragraph_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Docx)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/body/p[1]";
        let text = file
            .get(path)
            .unwrap_or_else(|error| panic!("{}: resolve DOCX paragraph: {error:?}", input.label))
            .view()
            .text()
            .unwrap_or_else(|error| panic!("{}: read DOCX paragraph: {error:?}", input.label))
            .unwrap_or_default()
            .into_owned();
        let props = crate::props! { "text" => text.as_str() };
        let leaf = signed_leaf(&file);

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let updated = candidate.set(path, props.clone()).unwrap_or_else(|error| {
                panic!(
                    "{} ({route}): DOCX signed no-op setter: {error:?}",
                    input.label
                )
            });
            assert!(
                updated.is_signed(),
                "{} ({route}): lost signature",
                input.label
            );
            assert_eq!(updated.to_bytes().unwrap(), input.bytes);

            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): DOCX signed edit: {error:?}", input.label)
                });
            assert!(!commit.diagnostics().changed(), "{} ({route})", input.label);
            assert!(commit.file().is_signed(), "{} ({route})", input.label);
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
            let restored = commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): DOCX signed no-op inverse replay: {error:?}",
                        input.label
                    )
                });
            assert!(restored.is_signed(), "{} ({route})", input.label);
            assert_eq!(restored.to_bytes().unwrap(), input.bytes);
        }

        let changed = format!("{text} changed");
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): DOCX signed cancellation: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed(), "{} ({route})", input.label);
            assert!(commit.file().is_signed(), "{} ({route})", input.label);
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
            let restored = commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): DOCX signed cancellation inverse replay: {error:?}",
                        input.label
                    )
                });
            assert_eq!(restored.to_bytes().unwrap(), input.bytes);

            let error = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())?;
                    edit.set(path, crate::props! { "text" => changed.as_str() })
                })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Docx,
                &format!("{} ({route}) mixed signed DOCX edit", input.label),
            );
            assert_eq!(
                candidate.to_bytes().unwrap(),
                input.bytes,
                "{} ({route}): changed signed DOCX edit published bytes",
                input.label
            );
        }
    }
}

#[cfg(feature = "xlsx")]
#[test]
fn signed_xlsx_sheet_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Xlsx)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/sheet[1]";
        let name = signed_name(&file, path);
        let props = crate::props! { "name" => name.as_str() };
        let leaf = signed_leaf(&file);

        let leaf_set = leaf
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed leaf no-op setter: {error:?}", input.label));
        assert!(leaf_set.is_signed(), "{}: leaf lost signature", input.label);
        assert_eq!(
            leaf_set.to_bytes().unwrap(),
            input.bytes,
            "{}: leaf bytes",
            input.label
        );

        let facade_set = file.set(path, props.clone()).unwrap_or_else(|error| {
            panic!("{}: signed facade no-op setter: {error:?}", input.label)
        });
        assert!(
            matches!(facade_set, File::Signed(_)),
            "{}: facade unwrapped",
            input.label
        );
        assert_eq!(
            facade_set.to_bytes().unwrap(),
            input.bytes,
            "{}: facade bytes",
            input.label
        );

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): signed setter edit: {error:?}", input.label)
                });
            assert!(!commit.diagnostics().changed(), "{} ({route})", input.label);
            assert!(
                commit.file().is_signed(),
                "{} ({route}): lost signature",
                input.label
            );
            assert_eq!(
                commit.file().to_bytes().unwrap(),
                input.bytes,
                "{} ({route})",
                input.label
            );
        }

        let changed = format!("{name} changed");
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let cell_path = "/sheet[1]/A1";
            let cell_text = candidate
                .get(cell_path)
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): resolve cell: {error:?}", input.label)
                })
                .view()
                .text()
                .unwrap_or_else(|error| panic!("{} ({route}): read cell: {error:?}", input.label))
                .unwrap_or_default()
                .into_owned();
            let cell_props = crate::props! { "value" => cell_text.as_str() };
            let changed = format!("{cell_text} changed");
            let commit = candidate
                .edit(|edit| {
                    edit.set(cell_path, crate::props! { "value" => changed.as_str() })?;
                    edit.set(cell_path, cell_props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancelling edit: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }
        // Keep a signed cancellation run on the native reservation path even
        // after it returns to the immutable source. A tight budget makes the
        // following read/write exercise the worksheet admission again rather
        // than relying on a stale coalescer cache.
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let cell_path = "/sheet[1]/A1";
            let cell_text = candidate
                .get(cell_path)
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): resolve cell: {error:?}", input.label)
                })
                .view()
                .text()
                .unwrap_or_else(|error| panic!("{} ({route}): read cell: {error:?}", input.label))
                .unwrap_or_default()
                .into_owned();
            let original = crate::props! { "value" => cell_text.as_str() };
            let changed = format!("{cell_text} changed again");
            let probe = candidate
                .edit(|edit| {
                    edit.set(cell_path, crate::props! { "value" => changed.as_str() })?;
                    edit.set(cell_path, original.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancellation probe: {error:?}",
                        input.label
                    )
                });
            let snapshot_budget = probe.diagnostics().snapshot_bytes();
            assert!(
                snapshot_budget > 0,
                "{} ({route}): empty signed capture",
                input.label
            );
            let limits = crate::transaction::SemanticLimits::builder()
                .max_snapshot_bytes(snapshot_budget)
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): snapshot bound: {error:?}", input.label)
                })
                .build()
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): snapshot profile: {error:?}", input.label)
                });
            let repeated = candidate
                .edit_with_options(crate::EditOptions::new().semantic_limits(limits), |edit| {
                    edit.set(cell_path, crate::props! { "value" => changed.as_str() })?;
                    edit.set(cell_path, original.clone())?;
                    edit.set(cell_path, crate::props! { "value" => changed.as_str() })?;
                    edit.set(cell_path, original.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancellation under tight budget: {error:?}",
                        input.label
                    )
                });
            assert!(!repeated.diagnostics().changed());
            assert!(repeated.file().is_signed());
            assert_eq!(repeated.file().to_bytes().unwrap(), input.bytes);
        }
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let cell_path = "/sheet[1]/A1";
            let cell_text = candidate
                .get(cell_path)
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): resolve cell: {error:?}", input.label)
                })
                .view()
                .text()
                .unwrap_or_else(|error| panic!("{} ({route}): read cell: {error:?}", input.label))
                .unwrap_or_default()
                .into_owned();
            let cell_props = crate::props! { "value" => cell_text.as_str() };
            let cell_changed = format!("{cell_text} changed");
            let error = candidate
                .edit(|edit| {
                    edit.set(
                        cell_path,
                        crate::props! { "value" => cell_changed.as_str() },
                    )?;
                    edit.set(cell_path, cell_props.clone())?;
                    edit.set(
                        cell_path,
                        crate::props! { "value" => cell_changed.as_str() },
                    )
                })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xlsx,
                &format!("{} ({route}) mixed signed cell edit", input.label),
            );
        }
        let missing_path = signed_missing_cell(&file)
            .unwrap_or_else(|| panic!("{}: signed XLSX fixture has no absent cell", input.label));
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let error = candidate
                .set(missing_path.as_str(), crate::props! { "value" => 0 })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xlsx,
                &format!("{} ({route}) absent cell", input.label),
            );
        }
        for (route, candidate) in [("leaf", leaf), ("facade", file)] {
            let error = candidate
                .set(path, crate::props! { "name" => changed.as_str() })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xlsx,
                &format!("{} ({route})", input.label),
            );
        }
    }
}

#[cfg(feature = "xlsb")]
#[test]
fn signed_xlsb_sheet_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Xlsb)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/sheet[1]";
        let name = signed_name(&file, path);
        let props = crate::props! { "name" => name.as_str() };
        let leaf = signed_leaf(&file);

        let leaf_set = leaf
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed leaf no-op setter: {error:?}", input.label));
        assert!(leaf_set.is_signed(), "{}: leaf lost signature", input.label);
        assert_eq!(
            leaf_set.to_bytes().unwrap(),
            input.bytes,
            "{}: leaf bytes",
            input.label
        );

        let facade_set = file.set(path, props.clone()).unwrap_or_else(|error| {
            panic!("{}: signed facade no-op setter: {error:?}", input.label)
        });
        assert!(
            matches!(facade_set, File::Signed(_)),
            "{}: facade unwrapped",
            input.label
        );
        assert_eq!(
            facade_set.to_bytes().unwrap(),
            input.bytes,
            "{}: facade bytes",
            input.label
        );

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): signed setter edit: {error:?}", input.label)
                });
            assert!(!commit.diagnostics().changed(), "{} ({route})", input.label);
            assert!(
                commit.file().is_signed(),
                "{} ({route}): lost signature",
                input.label
            );
            assert_eq!(
                commit.file().to_bytes().unwrap(),
                input.bytes,
                "{} ({route})",
                input.label
            );
        }

        let changed = format!("{name} changed");
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let cell_path = "/sheet[1]/A1";
            let cell_text = candidate
                .get(cell_path)
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): resolve cell: {error:?}", input.label)
                })
                .view()
                .text()
                .unwrap_or_else(|error| panic!("{} ({route}): read cell: {error:?}", input.label))
                .unwrap_or_default()
                .into_owned();
            let cell_props = crate::props! { "value" => cell_text.as_str() };
            let changed = format!("{cell_text} changed");
            let commit = candidate
                .edit(|edit| {
                    edit.set(cell_path, crate::props! { "value" => changed.as_str() })?;
                    edit.set(cell_path, cell_props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancelling edit: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let cell_path = "/sheet[1]/A1";
            let cell_text = candidate
                .get(cell_path)
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): resolve cell: {error:?}", input.label)
                })
                .view()
                .text()
                .unwrap_or_else(|error| panic!("{} ({route}): read cell: {error:?}", input.label))
                .unwrap_or_default()
                .into_owned();
            let cell_props = crate::props! { "value" => cell_text.as_str() };
            let cell_changed = format!("{cell_text} changed");
            let error = candidate
                .edit(|edit| {
                    edit.set(
                        cell_path,
                        crate::props! { "value" => cell_changed.as_str() },
                    )?;
                    edit.set(cell_path, cell_props.clone())?;
                    edit.set(
                        cell_path,
                        crate::props! { "value" => cell_changed.as_str() },
                    )
                })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xlsb,
                &format!("{} ({route}) mixed signed cell edit", input.label),
            );
        }
        let missing_path = signed_missing_cell(&file)
            .unwrap_or_else(|| panic!("{}: signed XLSB fixture has no absent cell", input.label));
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let error = candidate
                .set(missing_path.as_str(), crate::props! { "value" => 0 })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xlsb,
                &format!("{} ({route}) absent cell", input.label),
            );
        }
        for (route, candidate) in [("leaf", leaf), ("facade", file)] {
            let error = candidate
                .set(path, crate::props! { "name" => changed.as_str() })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xlsb,
                &format!("{} ({route})", input.label),
            );
        }
    }
}

#[cfg(feature = "doc")]
#[test]
fn signed_doc_paragraph_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Doc)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/body/p[1]";
        let text = file
            .get(path)
            .unwrap_or_else(|error| panic!("{}: resolve DOC paragraph: {error:?}", input.label))
            .view()
            .text()
            .unwrap_or_else(|error| panic!("{}: read DOC paragraph: {error:?}", input.label))
            .unwrap_or_default()
            .into_owned();
        let props = crate::props! { "text" => text.as_str() };
        let leaf = signed_leaf(&file);

        let leaf_set = leaf
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed DOC leaf no-op: {error:?}", input.label));
        assert!(
            leaf_set.is_signed(),
            "{}: DOC leaf lost signature",
            input.label
        );
        assert_eq!(leaf_set.to_bytes().unwrap(), input.bytes);

        let facade_set = file
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed DOC facade no-op: {error:?}", input.label));
        assert!(matches!(facade_set, File::Signed(_)));
        assert_eq!(facade_set.to_bytes().unwrap(), input.bytes);

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| panic!("{} ({route}): DOC edit: {error:?}", input.label));
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }

        let changed = same_width(&text).unwrap_or_else(|| format!("{text} changed"));
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancelling edit: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let error = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())?;
                    edit.set(path, crate::props! { "text" => changed.as_str() })
                })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Doc,
                &format!("{} ({route}) mixed signed edit", input.label),
            );
        }
        for (route, candidate) in [("leaf", leaf), ("facade", file)] {
            let error = candidate
                .set(path, crate::props! { "text" => changed.as_str() })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Doc,
                &format!("{} ({route})", input.label),
            );
        }
    }
}

#[cfg(feature = "ppt")]
#[test]
fn signed_ppt_shape_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Ppt)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/slide[1]/shape[1]";
        let text = file
            .get(path)
            .unwrap_or_else(|error| panic!("{}: resolve PPT shape: {error:?}", input.label))
            .view()
            .text()
            .unwrap_or_else(|error| panic!("{}: read PPT shape: {error:?}", input.label))
            .unwrap_or_default()
            .into_owned();
        let props = crate::props! { "text" => text.as_str() };
        let leaf = signed_leaf(&file);

        let leaf_set = leaf
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed PPT leaf no-op: {error:?}", input.label));
        assert!(
            leaf_set.is_signed(),
            "{}: PPT leaf lost signature",
            input.label
        );
        assert_eq!(leaf_set.to_bytes().unwrap(), input.bytes);

        let facade_set = file
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed PPT facade no-op: {error:?}", input.label));
        assert!(matches!(facade_set, File::Signed(_)));
        assert_eq!(facade_set.to_bytes().unwrap(), input.bytes);

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| panic!("{} ({route}): PPT edit: {error:?}", input.label));
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }

        let changed = same_width(&text).unwrap_or_else(|| format!("{text} changed"));
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancelling edit: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let error = candidate
                .edit(|edit| {
                    edit.set(path, crate::props! { "text" => changed.as_str() })?;
                    edit.set(path, props.clone())?;
                    edit.set(path, crate::props! { "text" => changed.as_str() })
                })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Ppt,
                &format!("{} ({route}) mixed signed edit", input.label),
            );
        }
        for (route, candidate) in [("leaf", leaf), ("facade", file)] {
            let error = candidate
                .set(path, crate::props! { "text" => changed.as_str() })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Ppt,
                &format!("{} ({route})", input.label),
            );
        }
    }
}

#[cfg(feature = "xls")]
#[test]
fn signed_xls_sheet_setter_preserves_exact_noops_and_refuses_changes() {
    for input in inputs()
        .into_iter()
        .filter(|input| input.format == FileFormat::Xls)
    {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        let path = "/sheet[1]";
        let name = signed_name(&file, path);
        let props = crate::props! { "name" => name.as_str() };
        let leaf = signed_leaf(&file);

        let leaf_set = leaf
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed XLS leaf no-op: {error:?}", input.label));
        assert!(
            leaf_set.is_signed(),
            "{}: XLS leaf lost signature",
            input.label
        );
        assert_eq!(leaf_set.to_bytes().unwrap(), input.bytes);

        let facade_set = file
            .set(path, props.clone())
            .unwrap_or_else(|error| panic!("{}: signed XLS facade no-op: {error:?}", input.label));
        assert!(matches!(facade_set, File::Signed(_)));
        assert_eq!(facade_set.to_bytes().unwrap(), input.bytes);

        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let commit = candidate
                .edit(|edit| edit.set(path, props.clone()))
                .unwrap_or_else(|error| panic!("{} ({route}): XLS edit: {error:?}", input.label));
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
        }

        let changed = format!("{name} changed");
        let (cell_path, _) = signed_text_cell(&file).unwrap_or_else(|| {
            panic!("{}: signed XLS fixture has no plain text cell", input.label)
        });
        for (route, candidate) in [("leaf", leaf.clone()), ("facade", file.clone())] {
            let cell_text = candidate
                .get(cell_path.as_str())
                .unwrap_or_else(|error| {
                    panic!("{} ({route}): resolve cell: {error:?}", input.label)
                })
                .view()
                .text()
                .unwrap_or_else(|error| panic!("{} ({route}): read cell: {error:?}", input.label))
                .unwrap_or_default()
                .into_owned();
            let cell_props = crate::props! { "value" => cell_text.as_str() };
            let cell_changed = format!("{cell_text} changed");
            let commit = candidate
                .edit(|edit| {
                    edit.set(
                        cell_path.as_str(),
                        crate::props! { "value" => cell_changed.as_str() },
                    )?;
                    edit.set(cell_path.as_str(), cell_props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed cancelling edit: {error:?}",
                        input.label
                    )
                });
            assert!(!commit.diagnostics().changed());
            assert!(commit.file().is_signed());
            assert_eq!(commit.file().to_bytes().unwrap(), input.bytes);
            let remove_reinsert = candidate
                .edit(|edit| {
                    edit.remove(cell_path.as_str())?;
                    edit.set(cell_path.as_str(), cell_props.clone())
                })
                .unwrap_or_else(|error| {
                    panic!(
                        "{} ({route}): signed remove/reinsert cancellation: {error:?}",
                        input.label
                    )
                });
            assert!(!remove_reinsert.diagnostics().changed());
            assert!(remove_reinsert.file().is_signed());
            assert_eq!(remove_reinsert.file().to_bytes().unwrap(), input.bytes);
            let mixed_error = candidate
                .edit(|edit| {
                    edit.set(
                        cell_path.as_str(),
                        crate::props! { "value" => cell_changed.as_str() },
                    )?;
                    edit.set(cell_path.as_str(), cell_props.clone())?;
                    edit.set(
                        cell_path.as_str(),
                        crate::props! { "value" => cell_changed.as_str() },
                    )
                })
                .unwrap_err();
            assert_signed_refusal(
                &mixed_error,
                FileFormat::Xls,
                &format!("{} ({route}) mixed signed cell edit", input.label),
            );
        }
        let missing_path = signed_missing_cell(&file)
            .unwrap_or_else(|| panic!("{}: signed XLS fixture has no absent cell", input.label));
        for (route, candidate) in [("leaf", signed_leaf(&file)), ("facade", file.clone())] {
            let error = candidate
                .set(missing_path.as_str(), crate::props! { "value" => 0 })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xls,
                &format!("{} ({route}) absent cell", input.label),
            );
        }
        for (route, candidate) in [("leaf", leaf), ("facade", file)] {
            let error = candidate
                .set(path, crate::props! { "name" => changed.as_str() })
                .unwrap_err();
            assert_signed_refusal(
                &error,
                FileFormat::Xls,
                &format!("{} ({route})", input.label),
            );
        }
    }
}

#[test]
fn every_signed_input_follows_the_explicit_signature_policy() {
    let inputs = inputs();
    assert!(!inputs.is_empty());
    for input in &inputs {
        let (files, _temporaries) = routes(input);
        for (route, file) in files {
            let what = format!("{} ({route})", input.label);

            // 1. Opening yields `Signed`, and every read works.
            let File::Signed(signed) = &file else {
                panic!("{what}: expected a signed file, got {file:?}");
            };
            assert!(file.is_signed(), "{what}");
            assert_eq!(signed.format(), input.format, "{what}");
            assert_eq!(signed.origin(), input.origin, "{what}");
            let text = file.text().unwrap();

            // 2. The editor builder refuses while the file is signed, and the
            // wrapped file reports it.
            assert!(signed.file().is_signed(), "{what}");
            assert_signed_refusal(&builder_refusal(signed.file()), input.format, &what);

            // 3. A signed no-op is created from the signed source itself and
            // replays onto an independently opened signed recipient without
            // changing its wrapper or publication bytes.
            let signed_noop = file
                .edit(|_| Ok(()))
                .unwrap_or_else(|error| panic!("{what}: unified signed no-op: {error:?}"));
            assert!(signed_noop.patch().is_noop(), "{what}");
            let reopened_signed = crate::open_bytes(file.to_bytes().unwrap())
                .unwrap_or_else(|error| panic!("{what}: reopen signed: {error:?}"));
            let replayed = signed_noop
                .patch()
                .apply(&reopened_signed)
                .unwrap_or_else(|error| panic!("{what}: unified signed replay: {error:?}"));
            assert!(replayed.is_signed(), "{what}: replay lost signature");
            assert_eq!(
                replayed.to_bytes().unwrap(),
                file.to_bytes().unwrap(),
                "{what}: unified replay changed signed bytes"
            );
            let leaf_noop = signed
                .file()
                .edit(|_| Ok(()))
                .unwrap_or_else(|error| panic!("{what}: leaf signed no-op: {error:?}"));
            let reopened_leaf = match crate::open_bytes(file.to_bytes().unwrap()) {
                Ok(File::Signed(reopened)) => reopened.file().clone(),
                Ok(other) => panic!("{what}: expected signed leaf, got {other:?}"),
                Err(error) => panic!("{what}: reopen signed leaf: {error:?}"),
            };
            let replayed_leaf = leaf_noop
                .patch()
                .apply(&reopened_leaf)
                .unwrap_or_else(|error| panic!("{what}: leaf signed replay: {error:?}"));
            assert!(
                replayed_leaf.is_signed(),
                "{what}: leaf replay lost signature"
            );
            assert_eq!(
                replayed_leaf.to_bytes().unwrap(),
                file.to_bytes().unwrap(),
                "{what}: leaf replay changed signed bytes"
            );

            // 4. `unsign` authorizes a changed commit for the returned file
            // only; its output carries no signature part.
            let unsigned = signed.clone().unsign();
            assert!(!unsigned.is_signed(), "{what}");
            assert_eq!(unsigned.format(), Some(input.format), "{what}");
            assert_eq!(
                unsigned.text().unwrap(),
                text,
                "{what}: unsign rewrote nothing"
            );
            let unsigned_bytes = unsigned.to_bytes().unwrap();
            let reopened_unsigned = crate::open_bytes(unsigned_bytes.clone())
                .unwrap_or_else(|error| panic!("{what}: reopen unsigned: {error:?}"));
            assert!(!reopened_unsigned.is_signed(), "{what}");
            let noop = unsigned
                .edit(|_| Ok(()))
                .unwrap_or_else(|error| panic!("{what}: authorized unsigned no-op: {error:?}"));
            assert!(noop.patch().is_noop(), "{what}");
            let replayed = noop
                .patch()
                .apply(&reopened_unsigned)
                .unwrap_or_else(|error| panic!("{what}: unsigned replay: {error:?}"));
            assert!(
                !replayed.is_signed(),
                "{what}: unsigned replay gained signature"
            );
            assert_eq!(
                replayed.to_bytes().unwrap(),
                unsigned_bytes,
                "{what}: unsigned replay changed bytes"
            );
            assert_eq!(
                noop.patch().apply(&file).unwrap_err().code(),
                Code::InvalidInput,
                "{what}: unsigned replay matched signed bytes"
            );
            assert_signed_refusal(&builder_refusal(signed.file()), input.format, &what);
            let (output, expected) = changed_commit(&unsigned);
            assert_unsigned_output(input.format, &output, &expected, &what);
            assert_signed_refusal(&builder_refusal(signed.file()), input.format, &what);
            assert_eq!(
                file.text().unwrap(),
                text,
                "{what}: signed snapshot changed"
            );
        }

        // 5. The per-format leaf transaction still keeps a signed exact no-op.
        assert_exact_noop(input);
    }
}

#[test]
fn invalid_intermediate_signed_edits_return_their_typed_error_without_publication() {
    for input in inputs() {
        let file = crate::open_bytes(input.bytes.clone()).unwrap();
        for (route, candidate) in [("leaf", signed_leaf(&file)), ("facade", file.clone())] {
            let before = candidate.to_bytes().unwrap();
            let error = candidate
                .edit(|edit| {
                    edit.set(
                        "/__signed_noop_missing__",
                        crate::props! { "text" => "invalid" },
                    )
                })
                .unwrap_err();
            assert_ne!(
                error.code(),
                Code::DocumentProtected,
                "{} ({route}): invalid admission was hidden by signature refusal: {error:?}",
                input.label
            );
            assert!(
                matches!(
                    error.code(),
                    Code::PathNotFound
                        | Code::ZeroMatches
                        | Code::NotFound
                        | Code::InvalidPath
                        | Code::UnsupportedPath
                ),
                "{} ({route}): unexpected invalid intermediate code: {error:?}",
                input.label
            );
            assert_eq!(
                candidate.to_bytes().unwrap(),
                before,
                "{} ({route}): failed intermediate edit published bytes",
                input.label
            );
        }
    }
}

/// Encrypted packages keep their encryption profile through `unsign`: the
/// unsigned copy would publish them in the clear, so the builder refuses, as
/// every edit of an unlocked encrypted package is refused.
#[cfg(all(feature = "encryption", feature = "docx"))]
#[test]
fn an_unlocked_signed_package_is_not_unsigned_into_the_clear() {
    let plain = read("poi/test-data/xmldsign/hello-world-signed.docx");
    let encrypted = litchi_crypto::ooxml::encrypt(
        plain,
        "signed",
        litchi_crypto::ooxml::Mode::Agile,
        &litchi_crypto::ooxml::Limits::default(),
    )
    .unwrap();
    let File::Locked(locked) = crate::open_bytes(encrypted).unwrap() else {
        panic!("an encrypted package opens locked");
    };
    let file = locked
        .unlock(crate::Password::new("signed".to_owned()))
        .unwrap();
    let File::Signed(signed) = &file else {
        panic!("the unlocked package is signed");
    };
    let signed_noop = file.edit(|_| Ok(())).unwrap();
    let error = signed_noop.patch().apply(&file).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
    assert_signed_refusal(
        &builder_refusal(signed.file()),
        FileFormat::Docx,
        "unlocked",
    );
    let error = builder_refusal(&signed.clone().unsign());
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
}
