#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "MCE ownership fixtures are deliberately checked and fail fast"
)]

//! Mutation-boundary coverage for DrawingML Theme Family ownership through
//! markup-compatibility wrappers.
//!
//! A `themeFamily` below an `mc:AlternateContent` branch is opaque to the
//! direct Theme owner.  The owner must refuse every shared mutation when that
//! hidden branch could be an effective owner, while preserving the exact
//! source.  Foreign wrappers and unrecognized extension URIs are separate
//! opaque payloads: they remain editable around a newly admitted direct owner.

use litchi_drawingml::theme::family::{self, part};

const NATIVE_THEME: &[u8] = include_bytes!("fixtures/theme-part-native.xml");
const OFFICE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const OFFICE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn native_family_fragment() -> &'static [u8] {
    let start = NATIVE_THEME
        .windows(b"<thm15:themeFamily".len())
        .position(|window| window == b"<thm15:themeFamily")
        .expect("native family root");
    let end = NATIVE_THEME[start..]
        .windows(2)
        .position(|window| window == b"/>")
        .map(|offset| start + offset + 2)
        .expect("native family close");
    &NATIVE_THEME[start..end]
}

fn native_extension_list() -> &'static [u8] {
    let start = NATIVE_THEME
        .windows(b"<a:extLst>".len())
        .position(|window| window == b"<a:extLst>")
        .expect("native extension list");
    let end = NATIVE_THEME[start..]
        .windows(b"</a:extLst>".len())
        .position(|window| window == b"</a:extLst>")
        .map(|offset| start + offset + b"</a:extLst>".len())
        .expect("native extension list close");
    &NATIVE_THEME[start..end]
}

fn replace_bytes(source: &[u8], old: &[u8], new: &[u8]) -> Vec<u8> {
    let start = source
        .windows(old.len())
        .position(|window| window == old)
        .expect("source marker");
    let mut output = Vec::with_capacity(source.len() + new.len() - old.len());
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(new);
    output.extend_from_slice(&source[start + old.len()..]);
    output
}

fn enclosed_bytes(source: &[u8], open: &[u8], close: &[u8]) -> Vec<u8> {
    let start = source
        .windows(open.len())
        .position(|window| window == open)
        .expect("enclosed opening marker");
    let end = source[start..]
        .windows(close.len())
        .position(|window| window == close)
        .map(|offset| start + offset + close.len())
        .expect("enclosed closing marker");
    source[start..end].to_vec()
}

fn with_root_namespaces(source: &[u8], declarations: &str) -> Vec<u8> {
    let marker = b"<a:theme ";
    let start = source
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("Theme root");
    let insertion = start + marker.len();
    let mut output = Vec::with_capacity(source.len() + declarations.len());
    output.extend_from_slice(&source[..insertion]);
    output.extend_from_slice(declarations.as_bytes());
    output.extend_from_slice(&source[insertion..]);
    output
}

fn with_mc_namespace(source: &[u8]) -> Vec<u8> {
    with_root_namespaces(source, &format!(r#"xmlns:mc="{MC_NAMESPACE}" "#))
}

fn mce(choice: &str, fallback: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="a">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    )
}

fn empty_mce() -> String {
    mce("", "")
}

fn family_extension() -> String {
    format!(
        r#"<a:ext uri="{}">{}</a:ext>"#,
        part::NATIVE_EXTENSION_URI,
        std::str::from_utf8(native_family_fragment()).expect("family UTF-8")
    )
}

fn empty_family_extension() -> String {
    format!(r#"<a:ext uri="{}"></a:ext>"#, part::NATIVE_EXTENSION_URI)
}

fn extension_list(content: &str) -> String {
    format!("<a:extLst>{content}</a:extLst>")
}

fn replace_native_extension_list(replacement: &str) -> Vec<u8> {
    with_mc_namespace(&replace_bytes(
        NATIVE_THEME,
        native_extension_list(),
        replacement.as_bytes(),
    ))
}

fn root_alternate_content(choice: &str, fallback: &str) -> Vec<u8> {
    replace_native_extension_list(&mce(choice, fallback))
}

fn extension_list_alternate_content(choice: &str, fallback: &str) -> Vec<u8> {
    replace_native_extension_list(&extension_list(&mce(choice, fallback)))
}

fn alternate_content_inside_extension(choice: &str, fallback: &str) -> Vec<u8> {
    let extension = format!(
        r#"<a:ext uri="{}">{}</a:ext>"#,
        part::NATIVE_EXTENSION_URI,
        mce(choice, fallback)
    );
    replace_native_extension_list(&extension_list(&extension))
}

fn alternate_content_inside_empty_extension() -> Vec<u8> {
    let extension = format!(
        r#"<a:ext uri="{}">{}</a:ext>"#,
        part::NATIVE_EXTENSION_URI,
        empty_mce()
    );
    replace_native_extension_list(&extension_list(&extension))
}

fn hidden_mce_cases() -> Vec<(&'static str, Vec<u8>)> {
    let family = std::str::from_utf8(native_family_fragment()).expect("family UTF-8");
    let ext = family_extension();
    let empty_ext = empty_family_extension();
    let self_closing_ext = format!(r#"<a:ext uri="{}"/>"#, part::NATIVE_EXTENSION_URI);
    let self_closing_mce_ext = format!(
        r#"<a:ext uri="{}"><mc:AlternateContent/></a:ext>"#,
        part::NATIVE_EXTENSION_URI
    );
    vec![
        (
            "root Choice self-closing extLst",
            root_alternate_content("<a:extLst/>", ""),
        ),
        (
            "list Choice self-closing admitted ext",
            extension_list_alternate_content(&self_closing_ext, ""),
        ),
        (
            "admitted ext self-closing AlternateContent",
            replace_native_extension_list(&extension_list(&self_closing_mce_ext)),
        ),
        (
            "direct AlternateContent wrapped extLst",
            replace_native_extension_list("<mc:AlternateContent><a:extLst/></mc:AlternateContent>"),
        ),
        (
            "root AlternateContent Choice extLst family",
            root_alternate_content(&extension_list(&ext), ""),
        ),
        (
            "root AlternateContent paired Choice and Fallback",
            root_alternate_content(&extension_list(&ext), &extension_list(&empty_ext)),
        ),
        (
            "root AlternateContent empty branches",
            root_alternate_content(&extension_list(""), &extension_list(&empty_ext)),
        ),
        (
            "direct extLst AlternateContent Choice family",
            extension_list_alternate_content(&ext, ""),
        ),
        (
            "direct extLst AlternateContent paired branches",
            extension_list_alternate_content(&ext, &empty_ext),
        ),
        (
            "direct extLst AlternateContent empty branches",
            extension_list_alternate_content("", &empty_ext),
        ),
        (
            "recognized ext AlternateContent family",
            alternate_content_inside_extension(family, ""),
        ),
        (
            "recognized ext AlternateContent paired branches",
            alternate_content_inside_extension(family, family),
        ),
        (
            "recognized ext AlternateContent empty branches",
            alternate_content_inside_empty_extension(),
        ),
    ]
}

#[test]
fn mce_owners_are_opaque_and_all_shared_mutations_refuse_without_source_change() {
    let incoming = family::Family::new("Fresh MCE Family", OFFICE_ID, OFFICE_VID)
        .expect("valid incoming family");

    for (label, source) in hidden_mce_cases() {
        let snapshot = part::read(&source).expect(label);
        assert!(
            snapshot.family().is_none(),
            "{label}: hidden family projected"
        );
        assert!(
            part::read_family(&source)
                .expect("read optional hidden projection")
                .is_none(),
            "{label}: hidden family projected by borrowed helper"
        );
        assert_eq!(snapshot.xml_bytes(), source.as_slice(), "{label}: source");

        assert!(
            snapshot
                .add_family_with_uri(&incoming, part::NATIVE_EXTENSION_URI)
                .is_err(),
            "{label}: snapshot add must refuse"
        );
        assert!(
            snapshot.replace_family(&incoming).is_err(),
            "{label}: snapshot replace must refuse"
        );
        assert!(
            snapshot.remove_family().is_err(),
            "{label}: snapshot remove must refuse"
        );
        assert!(
            part::add_family_with_uri(&source, &incoming, part::NATIVE_EXTENSION_URI).is_err(),
            "{label}: borrowed add must refuse"
        );
        assert!(
            part::replace_family(&source, &incoming).is_err(),
            "{label}: borrowed replace must refuse"
        );
        assert!(
            part::remove_family(&source).is_err(),
            "{label}: borrowed remove must refuse"
        );
        assert_eq!(
            snapshot.xml_bytes(),
            source.as_slice(),
            "{label}: source changed"
        );
    }
}

fn foreign_wrapper_mce_theme() -> Vec<u8> {
    let family = std::str::from_utf8(native_family_fragment()).expect("family UTF-8");
    let hidden = format!(
        r#"<x:wrapper><mc:AlternateContent><mc:Choice Requires="a"><a:extLst><a:ext uri="{uri}">{family}</a:ext></a:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent></x:wrapper>"#,
        uri = part::NATIVE_EXTENSION_URI,
    );
    with_root_namespaces(
        replace_bytes(NATIVE_THEME, native_extension_list(), hidden.as_bytes()).as_slice(),
        &format!(r#"xmlns:mc="{MC_NAMESPACE}" xmlns:x="urn:litchi:foreign-wrapper" "#),
    )
}

#[test]
fn foreign_wrapper_mce_family_stays_opaque_while_direct_owner_can_be_added() {
    let source = foreign_wrapper_mce_theme();
    let opaque_wrapper = enclosed_bytes(&source, b"<x:wrapper>", b"</x:wrapper>");
    let snapshot = part::read(&source).expect("foreign wrapper Theme");
    assert!(snapshot.family().is_none());

    let incoming = family::Family::new("Fresh Direct Family", OFFICE_ID, OFFICE_VID)
        .expect("valid incoming family");
    let added = snapshot
        .add_family_with_uri(&incoming, part::NATIVE_EXTENSION_URI)
        .expect("foreign wrapper must not block direct owner");
    let added_snapshot = part::read(&added).expect("read direct owner after insertion");
    assert_eq!(
        added_snapshot.family().expect("direct owner").name(),
        "Fresh Direct Family"
    );
    assert!(
        added
            .windows(opaque_wrapper.len())
            .any(|window| window == opaque_wrapper.as_slice()),
        "opaque foreign wrapper bytes must remain exact"
    );

    let removed = part::remove_family(&added).expect("remove direct owner");
    assert_eq!(
        removed, source,
        "removing direct owner must restore opaque source"
    );
    assert!(
        part::read(&removed)
            .expect("read restored source")
            .family()
            .is_none()
    );
}

fn foreign_uri_mce_theme() -> (Vec<u8>, Vec<u8>) {
    let family = std::str::from_utf8(native_family_fragment()).expect("family UTF-8");
    let payload = format!(
        r#"<a:ext uri="urn:litchi:foreign-mce"><mc:AlternateContent><mc:Choice Requires="a">{family}</mc:Choice><mc:Fallback><x:themeFamily xmlns:x="urn:litchi:lookalike" marker="keep"/></mc:Fallback></mc:AlternateContent></a:ext>"#
    );
    let source = with_root_namespaces(
        replace_bytes(
            NATIVE_THEME,
            native_extension_list(),
            extension_list(&payload).as_bytes(),
        )
        .as_slice(),
        &format!(r#"xmlns:mc="{MC_NAMESPACE}" "#),
    );
    (source, payload.into_bytes())
}

#[test]
fn foreign_uri_mce_family_and_lookalike_remain_opaque_around_direct_owner() {
    let (source, foreign_payload) = foreign_uri_mce_theme();
    let snapshot = part::read(&source).expect("foreign URI Theme");
    assert!(snapshot.family().is_none());

    let incoming = family::Family::new("Fresh Direct Family", OFFICE_ID, OFFICE_VID)
        .expect("valid incoming family");
    let added = part::add_family_with_uri(&source, &incoming, part::NATIVE_EXTENSION_URI)
        .expect("foreign URI must not block direct owner");
    let added_snapshot = part::read(&added).expect("read direct owner");
    assert_eq!(
        added_snapshot.family().expect("direct owner").name(),
        "Fresh Direct Family"
    );
    assert!(
        added
            .windows(foreign_payload.len())
            .any(|window| window == foreign_payload.as_slice()),
        "foreign URI payload must remain exact"
    );

    let removed = part::remove_family(&added).expect("remove direct owner");
    assert_eq!(
        removed, source,
        "foreign URI payload must survive inverse edit"
    );
    assert!(
        part::read(&removed)
            .expect("read foreign source")
            .family()
            .is_none()
    );
}
