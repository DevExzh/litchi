//! Opt-in correctness checks over checked-in files emitted by real producers.
//!
//! This corpus is deliberately outside the performance selector and every test
//! is ignored by default. Run it explicitly with:
//!
//! ```text
//! cargo test --manifest-path tools/perf-baseline/Cargo.toml --lib security_corpus -- --ignored
//! ```
//!
//! The assertions cover bounded ingress/publication and security semantics,
//! not latency. Fixture paths and source SHA-256 values are fixed so a changed
//! producer artifact is a visible manifest change. The macro slice is limited
//! to inert discovery, exact CFB stream identity, and the existing typed
//! source-backed edit refusal; this harness never runs or authors VBA. The
//! managed budget assertion checks the retained (RAII) memory/object/output
//! dimensions after dropping the package; input and work are cumulative
//! counters by design and are not expected to return to zero.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the opt-in corpus is an assertion-oriented integration harness"
)]

use std::{
    collections::BTreeMap,
    error::Error,
    fs,
    io::Cursor,
    num::{NonZeroU64, NonZeroUsize},
    path::{Path, PathBuf},
    sync::Arc,
};

use litchi_cfb::OleFile;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource,
};
use litchi_doc::{Error as DocError, OpenOptions, Package as DocPackage};
use litchi_docx::source_backed::Package as DocxSourcePackage;
use litchi_opc::{OpcError, OpcPackage, ReadLimits, ReadResource, SourceBackedPackage, TargetMode};
use litchi_sign::{Policy, Status};
use litchi_xls::{Workbook as XlsWorkbook, cell_values::Snapshot as XlsSnapshot};
use sha2::{Digest, Sha256};

const SIGNED_DOCX_SHA256: &str = "bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f";
const SIGNED_XLSX_SHA256: &str = "4cbd8cbe613f036b7a0c779ffaaec7c5838710896c6ef26b3f27410d25d5ce45";
const SIGNED_PPTX_SHA256: &str = "4d925d282dcca86e62b6716647a458246f8b9ea0eae0ec6664bbbf5a3f91bce1";
const PROTECTED_DOCX_SHA256: &str =
    "5d4c919f2e06b84fbe35cfaaa4012e8f469b811e1f643deb3e660b798bfe4544";
const EXTERNAL_XLSX_SHA256: &str =
    "e06155747da482bfb7c1ac5f0ab3a80cbe5b510e664926709c356ba6b59e9bc4";
const MACRO_XLS_SHA256: &str = "0e92c9bb018abd8a5f9121d65827c9e3bd280777219cb77a2efd70635143c00a";
const ENCRYPTED_CRYPTOAPI_DOC_SHA256: &str =
    "f2d0dc59ad7ec2356695ad5dc550057052a4017d5f1eb46e887297f5089896fb";
const ENCRYPTED_BINARY_RC4_DOC_SHA256: &str =
    "9231e724bb17a2e5f74815728d90b06e15684cf5fb2443a6fa24deebd33be952";
const CRYPTOAPI_SEMANTIC_SHA256: &str =
    "6dd4273bea0a8f70f4b6d8448e0ea1cb22b54713ad78d177394bd8b496e0aea6";
const BINARY_RC4_SEMANTIC_SHA256: &str =
    "5c5c945257fcd1569b5161722e15a9d73283daf786aca1daf04d0029d3736b78";
const EXTERNAL_INVENTORY_SHA256: &str =
    "16a2466d394b25d4a465c4db740fca842b9627e98a91c37163312b9585b80beb";

const SIGNED_DOCX_PATH: &str = "../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx";
const SIGNED_XLSX_PATH: &str = "../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx";
const SIGNED_PPTX_PATH: &str = "../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx";
const PROTECTED_DOCX_PATH: &str =
    "../../test-data/ooxml/docx/documentProtection_readonly_no_password.docx";
const EXTERNAL_XLSX_PATH: &str = "../../test-data/ooxml/xlsx/external-link-path-startup.xlsx";
const MACRO_XLS_PATH: &str = "../../test-data/poi/test-data/spreadsheet/SimpleMacro.xls";

#[derive(Clone, Copy)]
struct Fixture {
    id: &'static str,
    relative_path: &'static str,
    sha256: &'static str,
}

const SIGNED_FIXTURES: &[Fixture] = &[
    Fixture {
        id: "poi-office-2010-signed-docx",
        relative_path: SIGNED_DOCX_PATH,
        sha256: SIGNED_DOCX_SHA256,
    },
    Fixture {
        id: "poi-office-2010-signed-xlsx",
        relative_path: SIGNED_XLSX_PATH,
        sha256: SIGNED_XLSX_SHA256,
    },
    Fixture {
        id: "poi-office-2010-signed-pptx",
        relative_path: SIGNED_PPTX_PATH,
        sha256: SIGNED_PPTX_SHA256,
    },
];

const ENCRYPTED_FIXTURES: &[(&str, &str, &str)] = &[
    (
        "poi-password-password-cryptoapi",
        "../../test-data/poi/test-data/document/password_password_cryptoapi.doc",
        ENCRYPTED_CRYPTOAPI_DOC_SHA256,
    ),
    (
        "poi-password-tika-binaryrc4",
        "../../test-data/poi/test-data/document/password_tika_binaryrc4.doc",
        ENCRYPTED_BINARY_RC4_DOC_SHA256,
    ),
];

#[test]
#[ignore = "opt-in correctness-only real-producer security corpus"]
fn real_producer_security_corpus_is_bounded_and_deterministic() -> Result<(), Box<dyn Error>> {
    signed_ooxml_security()?;
    protected_docx_security()?;
    encrypted_ole_security()?;
    inert_macro_security()?;
    external_target_inventory()?;
    bounded_ingress_and_publication()?;
    Ok(())
}

fn signed_ooxml_security() -> Result<(), Box<dyn Error>> {
    for fixture in SIGNED_FIXTURES {
        let bytes = load_fixture(*fixture)?;
        let package = OpcPackage::from_bytes(&bytes)?;
        assert!(
            package.is_signed(),
            "{} lost its signature graph",
            fixture.id
        );
        let reports = package.signatures_with(&Policy::compatible())?;
        assert!(
            !reports.is_empty(),
            "{} has no signature report",
            fixture.id
        );
        assert!(reports.iter().all(|report| {
            report.integrity() == Status::Valid && report.signature() == Status::Valid
        }));

        let source = SourceBackedPackage::from_vec(bytes.clone())?;
        let mut no_op = Vec::new();
        source.write_part_overlays_to_stream(&mut no_op, Vec::new())?;
        assert_eq!(no_op, bytes, "{} changed on exact no-op", fixture.id);

        let source = SourceBackedPackage::from_vec(bytes)?;
        let (target, replacement) = {
            let main = source.main_document_part()?;
            let target = main.partname().clone();
            let mut replacement = main.data()?.as_bytes().to_vec();
            replacement.push(b' ');
            (target, replacement)
        };
        let mut output = Vec::new();
        let error = source
            .write_part_overlay_to_stream(&mut output, &target, replacement)
            .unwrap_err();
        assert!(matches!(
            error,
            OpcError::SignedSourceRequiresExplicitPolicy
        ));
        assert!(output.is_empty(), "{} wrote before refusal", fixture.id);
    }
    Ok(())
}

fn protected_docx_security() -> Result<(), Box<dyn Error>> {
    let fixture = Fixture {
        id: "ooxml-document-protection-readonly-no-password",
        relative_path: PROTECTED_DOCX_PATH,
        sha256: PROTECTED_DOCX_SHA256,
    };
    let bytes = load_fixture(fixture)?;

    let package = SourceBackedPackage::from_vec(bytes.clone())?;
    let docx = DocxSourcePackage::from_source_backed_package(package)?;
    let no_op = docx.edit_document_variables()?.commit()?;
    assert!(!no_op.changed());
    let mut no_op_output = Vec::new();
    docx.publish_document_variables_commit_to_stream(&mut no_op_output, &no_op)?;
    assert_eq!(no_op_output, bytes);

    let package = SourceBackedPackage::from_vec(bytes)?;
    let docx = DocxSourcePackage::from_source_backed_package(package)?;
    let mut edit = docx.edit_document_variables()?;
    edit.set_variable("security_matrix", "must-refuse")?;
    let commit = edit.commit()?;
    assert!(commit.changed());
    let mut output = Vec::new();
    let error = match docx.publish_document_variables_commit_to_stream(&mut output, &commit) {
        Ok(_) => return Err("protected DOCX publication unexpectedly succeeded".into()),
        Err(error) => error,
    };
    assert!(matches!(error, litchi_docx::Error::UnsafeEdit { .. }));
    assert!(output.is_empty());
    Ok(())
}

fn encrypted_ole_security() -> Result<(), Box<dyn Error>> {
    for &(id, relative_path, sha256) in ENCRYPTED_FIXTURES {
        let bytes = load_fixture(Fixture {
            id,
            relative_path,
            sha256,
        })?;

        let mut package = DocPackage::from_reader(Cursor::new(bytes.clone()))?;
        assert!(matches!(
            package.document(),
            Err(DocError::PasswordRequired)
        ));

        let mut package = DocPackage::from_reader(Cursor::new(bytes.clone()))?;
        assert!(matches!(
            package.document_with_options(
                OpenOptions::default().with_password("wrong".to_owned().into())
            ),
            Err(DocError::InvalidPassword)
        ));

        let password = if id == "poi-password-password-cryptoapi" {
            "password"
        } else {
            "tika"
        };
        let mut package = DocPackage::from_reader(Cursor::new(bytes))?;
        let document = package.document_with_options(
            OpenOptions::default().with_password(password.to_owned().into()),
        )?;
        let text = document.text()?;
        let semantic_digest = sha256_hex(text.as_bytes());
        assert!(!text.trim().is_empty(), "{id} has no semantic text");
        let expected = if id == "poi-password-password-cryptoapi" {
            CRYPTOAPI_SEMANTIC_SHA256
        } else {
            BINARY_RC4_SEMANTIC_SHA256
        };
        assert_eq!(semantic_digest, expected, "{id} semantic digest changed");

        let mut package = DocPackage::from_reader(Cursor::new(load_fixture(Fixture {
            id,
            relative_path,
            sha256,
        })?))?;
        let document = package.document_with_options(
            OpenOptions::default().with_password(password.to_owned().into()),
        )?;
        assert_eq!(semantic_digest, sha256_hex(document.text()?.as_bytes()));
    }
    Ok(())
}

fn inert_macro_security() -> Result<(), Box<dyn Error>> {
    let bytes = load_fixture(Fixture {
        id: "poi-simple-macro-xls",
        relative_path: MACRO_XLS_PATH,
        sha256: MACRO_XLS_SHA256,
    })?;
    let before_macro_digest = macro_stream_digest(&bytes)?;

    let mut workbook = XlsWorkbook::new(Cursor::new(bytes.clone()))?;
    let metadata = workbook.vba_metadata();
    assert!(metadata.has_project_marker());
    assert!(metadata.has_project_storage());
    assert!(metadata.may_contain_executable_code());
    let storage = workbook
        .vba_project_storage()
        .ok_or("macro fixture has no VBA project storage")?;
    assert!(storage.is_structurally_complete());
    assert!(storage.may_contain_macro_code());
    let project = workbook.vba()?.ok_or("macro fixture has no VBA project")?;
    assert!(
        project
            .modules()
            .iter()
            .any(|module| { module.source().text().contains("Sub ") })
    );
    let source = XlsSnapshot::from_bytes(bytes.clone())?;
    let no_op = source.edit().commit()?;
    assert_eq!(no_op.snapshot().bytes(), bytes.as_slice());
    assert_eq!(
        before_macro_digest,
        macro_stream_digest(no_op.snapshot().bytes())?
    );
    let mut edit = source.edit();
    edit.insert_rows("Sheet1".into(), 1, 1)?;
    let error = edit.commit_source_backed().unwrap_err();
    assert!(matches!(error, litchi_xls::Error::UnsupportedFeature(_)));
    assert_eq!(source.bytes(), bytes.as_slice());
    assert_eq!(before_macro_digest, macro_stream_digest(&bytes)?);
    Ok(())
}

fn external_target_inventory() -> Result<(), Box<dyn Error>> {
    let bytes = load_fixture(Fixture {
        id: "ooxml-external-link-path-startup",
        relative_path: EXTERNAL_XLSX_PATH,
        sha256: EXTERNAL_XLSX_SHA256,
    })?;
    let package = SourceBackedPackage::from_vec(bytes)?;
    let inventory = collect_external_targets(&package);
    assert_eq!(
        inventory,
        vec![
            "part=/xl/externalLinks/externalLink1.xml|rId1|http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup|personal.xls|external"
            .to_owned()
        ]
    );
    assert_eq!(
        inventory_digest(&inventory),
        EXTERNAL_INVENTORY_SHA256,
        "external target inventory changed"
    );
    assert!(
        package
            .physical_member_names()
            .all(|member| member != "personal.xls")
    );
    Ok(())
}

fn bounded_ingress_and_publication() -> Result<(), Box<dyn Error>> {
    let bytes = load_fixture(Fixture {
        id: "ooxml-external-link-path-startup",
        relative_path: EXTERNAL_XLSX_PATH,
        sha256: EXTERNAL_XLSX_SHA256,
    })?;
    let one_under = ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len().saturating_sub(1))?)?
        .build()?;
    let error = match SourceBackedPackage::from_vec_with_limits(bytes.clone(), one_under) {
        Ok(_) => return Err("one-under input limit unexpectedly opened".into()),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::InputBytes,
            ..
        }
    ));

    let budget = Budget::root(
        "real-producer-security-publication",
        Limits::new(64 * 1024 * 1024, u64::MAX, 0, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(64 * 1024 * 1024).unwrap(),
        0,
    )?;
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.clone()));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )?;
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    let error = artifact.write_to_stream(&mut output).unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(ref limit))
            if limit.resource == Resource::OutputBytes
    ));
    assert!(output.is_empty());
    drop(artifact);
    drop(package);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
    assert_eq!(budget.used(Resource::OutputBytes), 0);
    Ok(())
}

fn load_fixture(fixture: Fixture) -> Result<Vec<u8>, Box<dyn Error>> {
    let path = fixture_path(fixture.relative_path);
    let bytes = fs::read(&path)?;
    let actual = sha256_hex(&bytes);
    if actual != fixture.sha256 {
        return Err(format!(
            "fixture {} changed: {} has SHA-256 {}, expected {}",
            fixture.id,
            path.display(),
            actual,
            fixture.sha256
        )
        .into());
    }
    Ok(bytes)
}

fn fixture_path(relative_path: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative_path)
}

fn collect_external_targets(package: &SourceBackedPackage) -> Vec<String> {
    let mut inventory = Vec::new();
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| relationship.is_external())
    {
        inventory.push(format_relationship("package", relationship));
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| relationship.is_external())
        {
            inventory.push(format_relationship(part.partname().as_str(), relationship));
        }
    }
    inventory.sort();
    inventory
}

fn format_relationship(owner: &str, relationship: &litchi_opc::Relationship) -> String {
    let mode = match relationship.target_mode() {
        TargetMode::Internal => "internal",
        TargetMode::External => "external",
    };
    format!(
        "part={owner}|{}|{}|{}|{mode}",
        relationship.r_id(),
        relationship.reltype(),
        relationship.target_ref()
    )
}

fn macro_stream_digest(bytes: &[u8]) -> Result<String, Box<dyn Error>> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    let mut streams = BTreeMap::new();
    for path in ole.list_streams() {
        if !path
            .iter()
            .any(|component| component.eq_ignore_ascii_case("_VBA_PROJECT_CUR"))
        {
            continue;
        }
        let references = path.iter().map(String::as_str).collect::<Vec<_>>();
        streams.insert(path.join("/"), ole.open_stream(&references)?);
    }
    if streams.is_empty() {
        return Err("macro fixture has no _VBA_PROJECT_CUR streams".into());
    }
    let mut hasher = Sha256::new();
    for (path, payload) in streams {
        hasher.update((path.len() as u64).to_le_bytes());
        hasher.update(path.as_bytes());
        hasher.update((payload.len() as u64).to_le_bytes());
        hasher.update(payload);
    }
    Ok(hex_digest(hasher.finalize()))
}

fn inventory_digest(inventory: &[String]) -> String {
    let mut hasher = Sha256::new();
    for item in inventory {
        hasher.update((item.len() as u64).to_le_bytes());
        hasher.update(item.as_bytes());
    }
    hex_digest(hasher.finalize())
}

fn sha256_hex(bytes: &[u8]) -> String {
    hex_digest(Sha256::digest(bytes))
}

fn hex_digest<D>(digest: D) -> String
where
    D: AsRef<[u8]>,
{
    let digest = digest.as_ref();
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}
