//! Reproducible downstream probe for the bounded Theme-family hardening batch.
//!
//! This binary uses only public DrawingML/XLSB APIs.  It keeps valid scalar
//! edit outputs for offline XSD validation and keeps malformed/opaque cases as
//! diagnostics only; arbitrary opaque extension payload is not claimed to be
//! schema-valid.

use std::{
    error::Error,
    fs,
    io::{Cursor, Write},
    path::{Path, PathBuf},
};

use litchi_drawingml::theme::family::{self, part};
use litchi_xlsb::Workbook;

const FAMILY_NAMESPACE: &str = "http://schemas.microsoft.com/office/thememl/2012/main";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const OFFICE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const OFFICE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

struct Case {
    name: String,
    detail: String,
}

fn main() -> Result<(), Box<dyn Error>> {
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let root = manifest.join("../../../../../");
    let output_dir = manifest.join("outputs");
    fs::create_dir_all(&output_dir)?;

    let date_path = root.join("test-data/ooxml/xlsb/date.xlsb");
    let hyperlink_path = root.join("test-data/ooxml/xlsb/hyperlink.xlsb");
    let date_package = fs::read(&date_path)?;
    let hyperlink_package = fs::read(&hyperlink_path)?;
    let date_workbook = Workbook::new(Cursor::new(date_package))?;
    let hyperlink_workbook = Workbook::new(Cursor::new(hyperlink_package))?;
    let date_theme = date_workbook.theme()?.ok_or("date.xlsb has no Theme")?;
    let hyperlink_theme = hyperlink_workbook
        .theme()?
        .ok_or("hyperlink.xlsb has no Theme")?;
    let native = date_theme.source_xml().to_vec();
    assert_eq!(native, hyperlink_theme.source_xml());
    assert_eq!(native.len(), 8_390, "native Theme fixture size changed");
    write_output(&output_dir, "native.xml", &native)?;

    let native_fragment = native_family_fragment(&native);
    let parsed = part::read(&native)?;
    let native_family = parsed.family().ok_or("native family missing")?.clone();
    assert_eq!(
        parsed.family_profile(),
        Some(part::ExtensionProfile::NativeDiscriminator)
    );

    let mut cases = Vec::new();
    cases.push(Case {
        name: "native_fixture_theme_source".to_owned(),
        detail: format!(
            "bytes={}; date_and_hyperlink_theme_equal=true",
            native.len()
        ),
    });

    let mut changed = native_family.clone();
    changed.set_name("Audit Edited Theme")?;
    let updated = parsed.replace_family(&changed)?;
    assert_eq!(
        part::read(&updated)?.family().map(family::Family::name),
        Some("Audit Edited Theme")
    );
    let changed_fragment = replace_once(
        &native_fragment,
        b"name=\"Office Theme\"",
        b"name=\"Audit Edited Theme\"",
    );
    let expected_update = replace_once(&native, &native_fragment, &changed_fragment);
    assert_eq!(
        updated, expected_update,
        "scalar edit changed unrelated bytes"
    );
    write_output(&output_dir, "updated.xml", &updated)?;
    cases.push(Case {
        name: "native_scalar_name_edit".to_owned(),
        detail: format!("bytes={}", updated.len()),
    });

    let removed = parsed.remove_family()?;
    assert!(
        !removed
            .windows(b"<a:extLst".len())
            .any(|window| window == b"<a:extLst")
    );
    assert!(
        !removed
            .windows(b"<a:ext ".len())
            .any(|window| window == b"<a:ext ")
            && !removed
                .windows(b"<a:ext>".len())
                .any(|window| window == b"<a:ext>")
            && !removed
                .windows(b"<a:ext/".len())
                .any(|window| window == b"<a:ext/")
    );
    assert!(part::read(&removed)?.family().is_none());
    write_output(&output_dir, "removed.xml", &removed)?;
    cases.push(Case {
        name: "native_removal_closes_semantically_empty_wrappers".to_owned(),
        detail: format!("bytes={}; extLst=false; ext=false", removed.len()),
    });

    let added = part::add_family(&removed, &native_family)?;
    let added_snapshot = part::read(&added)?;
    assert_eq!(
        added_snapshot.family_profile(),
        Some(part::ExtensionProfile::Normative)
    );
    assert_eq!(
        added_snapshot.family().map(family::Family::name),
        Some("Office Theme")
    );
    write_output(&output_dir, "added.xml", &added)?;
    cases.push(Case {
        name: "normative_add_after_removal".to_owned(),
        detail: format!("bytes={}; profile=normative", added.len()),
    });

    let mut bom = b"\xEF\xBB\xBF".to_vec();
    bom.extend_from_slice(&native_fragment);
    let bom_family = family::read(&bom)?;
    let bom_added = part::add_family_with_uri(&removed, &bom_family, part::NATIVE_EXTENSION_URI)?;
    assert!(!bom_added.windows(3).any(|window| window == b"\xEF\xBB\xBF"));
    assert!(part::read(&bom_added)?.family().is_some());
    cases.push(Case {
        name: "standalone_bom_is_stripped_before_embedding".to_owned(),
        detail: format!("bytes={}; embedded_bom_count=0", bom_added.len()),
    });

    let declaration = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>{}",
        String::from_utf8_lossy(&native_fragment)
    );
    let declaration_family = family::read(declaration.as_bytes())?;
    assert!(part::add_family(&removed, &declaration_family).is_err());
    let pi = format!(
        "<?future data?>{}",
        String::from_utf8_lossy(&native_fragment)
    );
    assert!(family::read(pi.as_bytes()).is_err());
    cases.push(Case {
        name: "standalone_declaration_and_processing_instruction_policy".to_owned(),
        detail: "declaration_parses_but_embedding_refuses; actual_pi_refused".to_owned(),
    });

    let opaque = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:opaque" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:future><!-- opaque <?x --><![CDATA[<?x]]></x:future></thm15:themeFamily>"#
    )
    .into_bytes();
    let opaque_theme = replace_once(&native, &native_fragment, &opaque);
    let opaque_snapshot = part::read(&opaque_theme)?;
    let opaque_replacement = family::Family::new("Opaque Edited", OFFICE_ID, OFFICE_VID)?;
    let opaque_edited = opaque_snapshot.replace_family(&opaque_replacement)?;
    assert!(
        opaque_edited
            .windows(b"<!-- opaque <?x -->".len())
            .any(|window| window == b"<!-- opaque <?x -->")
    );
    assert!(
        opaque_edited
            .windows(b"<![CDATA[<?x]]>".len())
            .any(|window| window == b"<![CDATA[<?x]]>")
    );
    assert_eq!(
        part::read(&opaque_edited)?
            .family()
            .map(family::Family::name),
        Some("Opaque Edited")
    );
    write_output(
        &output_dir,
        "opaque-preserved-diagnostic.xml",
        &opaque_edited,
    )?;
    cases.push(Case {
        name: "opaque_comment_and_cdata_processing_tokens".to_owned(),
        detail: "preserved; diagnostic output intentionally excluded from XSD claims".to_owned(),
    });

    let empty_prefix_family = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="" x:y="z" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#
    );
    assert!(family::read(empty_prefix_family.as_bytes()).is_err());
    assert!(
        part::read(&replace_once(
            &native,
            &native_fragment,
            empty_prefix_family.as_bytes()
        ))
        .is_err()
    );
    let default_xml_family = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns="{XML_NAMESPACE}" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#
    );
    assert!(family::read(default_xml_family.as_bytes()).is_err());
    assert!(
        part::read(&replace_once(
            &native,
            &native_fragment,
            default_xml_family.as_bytes()
        ))
        .is_err()
    );
    let invalid_root = replace_once(
        &native,
        b"<a:theme xmlns:a=",
        format!(r#"<a:theme xmlns="{XML_NAMESPACE}" xmlns:a="#).as_bytes(),
    );
    assert!(part::read(&invalid_root).is_err());
    let oversized_prefix = format!("p{}", "x".repeat(family::MAX_NAMESPACE_BYTES + 1));
    let oversized_prefix_family = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:{oversized_prefix}="urn:litchi:oversized-prefix" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#
    );
    assert_namespace_prefix_limit(
        family::read(oversized_prefix_family.as_bytes()),
        "standalone oversized namespace prefix",
    );
    assert_namespace_prefix_limit(
        part::read(&replace_once(
            &native,
            &native_fragment,
            oversized_prefix_family.as_bytes(),
        )),
        "complete-part oversized namespace prefix",
    );
    let exact_prefix = "p".repeat(family::MAX_NAMESPACE_BYTES);
    let exact_prefix_family = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:{exact_prefix}="urn:litchi:exact-prefix" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#
    );
    assert!(family::read(exact_prefix_family.as_bytes()).is_ok());
    assert!(
        part::read(&replace_once(
            &native,
            &native_fragment,
            exact_prefix_family.as_bytes(),
        ))
        .is_ok()
    );
    cases.push(Case {
        name: "reserved_and_empty_namespace_bindings".to_owned(),
        detail:
            "standalone and complete-part readers reject invalid bindings and oversized prefixes"
                .to_owned(),
    });

    let invalid_known = vec![
        (
            "ext-text",
            format!(r#"<a:ext uri="{}">bad</a:ext>"#, part::NATIVE_EXTENSION_URI),
        ),
        (
            "ext-reference",
            format!(
                r#"<a:ext uri="{}">&#65;</a:ext>"#,
                part::NATIVE_EXTENSION_URI
            ),
        ),
        (
            "ext-cdata",
            format!(
                r#"<a:ext uri="{}"><![CDATA[bad]]></a:ext>"#,
                part::NATIVE_EXTENSION_URI
            ),
        ),
        ("extlst-text", "<a:extLst>bad</a:extLst>".to_owned()),
        ("extlst-reference", "<a:extLst>&#65;</a:extLst>".to_owned()),
        (
            "extlst-cdata",
            "<a:extLst><![CDATA[bad]]></a:extLst>".to_owned(),
        ),
        (
            "extlst-foreign-child",
            r#"<a:extLst><x:future xmlns:x="urn:litchi:foreign"/></a:extLst>"#.to_owned(),
        ),
    ];
    for (label, container) in &invalid_known {
        let invalid = if label.starts_with("extlst") {
            insert_before_root_close(&removed, container.as_bytes())
        } else {
            insert_before_root_close(
                &removed,
                format!("<a:extLst>{container}</a:extLst>").as_bytes(),
            )
        };
        assert!(part::read(&invalid).is_err(), "accepted malformed {label}");
        assert!(
            part::add_family(&invalid, &native_family).is_err(),
            "add accepted malformed {label}"
        );
        assert!(
            part::replace_family(&invalid, &changed).is_err(),
            "replace accepted malformed {label}"
        );
        cases.push(Case {
            name: format!("rejects_{label}"),
            detail: "read/add/replace all returned Err".to_owned(),
        });
    }

    let mce = format!(
        r#"<mc:AlternateContent xmlns:mc="{MCE_NAMESPACE}"><mc:Choice Requires="a"><a:extLst><a:ext uri="{}">{}</a:ext></a:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        part::NATIVE_EXTENSION_URI,
        String::from_utf8_lossy(&native_fragment)
    );
    let mce_source = insert_before_root_close(&removed, mce.as_bytes());
    let mce_snapshot = part::read(&mce_source)?;
    assert!(mce_snapshot.family().is_none());
    assert!(
        mce_snapshot
            .add_family_with_uri(&native_family, part::NATIVE_EXTENSION_URI)
            .is_err()
    );
    assert!(mce_snapshot.replace_family(&changed).is_err());
    assert!(mce_snapshot.remove_family().is_err());
    assert_eq!(mce_snapshot.xml_bytes(), mce_source.as_slice());
    cases.push(Case {
        name: "mce_ignored_owner_refuses_all_mutations".to_owned(),
        detail: "read-only projection opaque; add/replace/remove refused".to_owned(),
    });

    let replacement_under_limit = updated.len() - 1;
    assert!(
        parsed
            .replace_family_with_limit(&changed, replacement_under_limit)
            .is_err()
    );
    assert!(
        parsed
            .replace_family_with_limit(&changed, updated.len())
            .is_ok()
    );
    assert!(part::remove_family_with_limit(&native, removed.len() - 1).is_err());
    assert!(part::remove_family_with_limit(&native, removed.len()).is_ok());
    assert!(part::add_family_with_limit(&removed, &native_family, added.len() - 1).is_err());
    assert!(part::add_family_with_limit(&removed, &native_family, added.len()).is_ok());
    cases.push(Case {
        name: "caller_specific_output_caps".to_owned(),
        detail: format!(
            "replace_limit={}; remove_limit={}; add_limit={}",
            updated.len(),
            removed.len(),
            added.len()
        ),
    });

    let mut receipt = String::from("case\tdetail\n");
    for case in &cases {
        receipt.push_str(&case.name);
        receipt.push('\t');
        receipt.push_str(&case.detail.replace('\n', "\\n"));
        receipt.push('\n');
    }
    receipt.push_str("all_cases_passed\ttrue\n");
    let mut file = fs::File::create(output_dir.join("probe-results.tsv"))?;
    file.write_all(receipt.as_bytes())?;
    println!("theme-family hardening probe passed {} cases", cases.len());
    Ok(())
}

fn write_output(directory: &Path, name: &str, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    fs::write(directory.join(name), bytes)?;
    Ok(())
}

fn native_family_fragment(source: &[u8]) -> Vec<u8> {
    let start = source
        .windows(b"<thm15:themeFamily".len())
        .position(|window| window == b"<thm15:themeFamily")
        .expect("native family root");
    let end = source[start..]
        .windows(2)
        .position(|window| window == b"/>")
        .map(|offset| start + offset + 2)
        .expect("native family close");
    source[start..end].to_vec()
}

fn replace_once(source: &[u8], old: &[u8], replacement: &[u8]) -> Vec<u8> {
    let start = source
        .windows(old.len())
        .position(|window| window == old)
        .expect("replacement marker");
    let mut output = Vec::with_capacity(source.len() + replacement.len());
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[start + old.len()..]);
    output
}

fn insert_before_root_close(source: &[u8], insertion: &[u8]) -> Vec<u8> {
    let marker = b"</a:theme>";
    let start = source
        .windows(marker.len())
        .rposition(|window| window == marker)
        .expect("Theme root close");
    let mut output = Vec::with_capacity(source.len() + insertion.len());
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(insertion);
    output.extend_from_slice(&source[start..]);
    output
}

fn assert_namespace_prefix_limit<T>(result: litchi_drawingml::Result<T>, label: &str) {
    match result {
        Err(litchi_drawingml::Error::Limit { resource, .. })
            if resource.contains("namespace prefix") => {},
        Err(error) => panic!("{label} returned the wrong error: {error:?}"),
        Ok(_) => panic!("{label} was accepted"),
    }
}
