#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "theme-part fixtures are deliberately checked and fail fast"
)]

//! Independent coverage for a complete DrawingML Theme part carrying the
//! shared DrawingML 2012 `themeFamily` fragment.
//!
//! The host package owns the extension path; this file keeps the shared crate
//! responsible for two narrower contracts: a complete `a:theme` remains a
//! valid typed theme when the extension is present, and a parsed family edit
//! keeps its exact source and opaque family markup.

use litchi_drawingml::theme::family::part;
use litchi_drawingml::theme::{codec, family};

const NATIVE_THEME: &[u8] = include_bytes!("fixtures/theme-part-native.xml");
const FAMILY_NAMESPACE: &str = "http://schemas.microsoft.com/office/thememl/2012/main";
const OFFICE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const OFFICE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

fn native_family_fragment() -> &'static [u8] {
    let start = NATIVE_THEME
        .windows(b"<thm15:themeFamily".len())
        .position(|window| window == b"<thm15:themeFamily")
        .expect("native family root");
    let end = NATIVE_THEME[start..]
        .windows(b"/>".len())
        .position(|window| window == b"/>")
        .map(|offset| start + offset + 2)
        .expect("native family close");
    &NATIVE_THEME[start..end]
}

fn native_family_range() -> std::ops::Range<usize> {
    let start = NATIVE_THEME
        .windows(b"<thm15:themeFamily".len())
        .position(|window| window == b"<thm15:themeFamily")
        .expect("native family root");
    start..start + native_family_fragment().len()
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

fn replace_native_family(replacement: &[u8]) -> Vec<u8> {
    replace_bytes(NATIVE_THEME, native_family_fragment(), replacement)
}

fn native_family_free_theme() -> Vec<u8> {
    part::remove_family(NATIVE_THEME).expect("remove native family for fixture mutation")
}

fn native_empty_extension() -> Vec<u8> {
    format!(r#"<a:ext uri="{}"></a:ext>"#, part::NATIVE_EXTENSION_URI).into_bytes()
}

fn oversized_family_theme() -> Vec<u8> {
    let open = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:oversized" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:payload>"#
    );
    let close = b"</x:payload></thm15:themeFamily>";
    let target = family::MAX_XML_BYTES + 1;
    let mut family = Vec::with_capacity(target);
    family.extend_from_slice(open.as_bytes());
    family.resize(target - close.len(), b'x');
    family.extend_from_slice(close);
    assert!(family.len() > family::MAX_XML_BYTES);
    let theme = replace_native_family(&family);
    assert!(theme.len() < codec::MAX_XML_BYTES);
    theme
}

fn inherited_opaque_family_theme() -> Vec<u8> {
    let original = std::str::from_utf8(native_family_fragment()).expect("family UTF-8");
    let inherited = original
        .replace(
            &format!(" vid=\"{OFFICE_VID}\"/>"),
            &format!(" vid=\"{OFFICE_VID}\" x:marker=\"opaque\"><x:child/></thm15:themeFamily>"),
        )
        .replace(&format!(" xmlns:thm15=\"{FAMILY_NAMESPACE}\""), "");
    String::from_utf8(NATIVE_THEME.to_vec())
        .expect("native Theme is UTF-8")
        .replace(
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
            &format!(
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:thm15=\"{FAMILY_NAMESPACE}\" xmlns:x=\"urn:litchi:inherited-opaque\""
            ),
        )
        .replace(original, &inherited)
        .into_bytes()
}

fn opaque_family() -> Vec<u8> {
    format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:theme-part" x:marker="keep" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:future value="keep"/></thm15:themeFamily>"#
    )
    .into_bytes()
}

fn opaque_markup_with_processing_token() -> Vec<u8> {
    format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:processing-token" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:future><!-- opaque <? marker --><![CDATA[opaque <? marker]]></x:future></thm15:themeFamily>"#
    )
    .into_bytes()
}

fn namespace_heavy_foreign_family() -> Vec<u8> {
    const LEVELS: usize = 64;
    let mut source = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}">"#
    )
    .into_bytes();
    for index in 0..LEVELS {
        source.extend_from_slice(
            format!(
                r#"<p{index}:node xmlns:p{index}="urn:litchi:namespace:{index}" p{index}:marker="{index}">"#
            )
            .as_bytes(),
        );
    }
    source.extend_from_slice(b"<!-- namespace-heavy leaf -->");
    for index in (0..LEVELS).rev() {
        source.extend_from_slice(format!("</p{index}:node>").as_bytes());
    }
    source.extend_from_slice(b"</thm15:themeFamily>");
    source
}

fn mce_hidden_extlst_theme() -> Vec<u8> {
    let source = NATIVE_THEME;
    let list_start = source
        .windows(b"<a:extLst>".len())
        .position(|window| window == b"<a:extLst>")
        .expect("native extLst");
    let list_end = source[list_start..]
        .windows(b"</a:extLst>".len())
        .position(|window| window == b"</a:extLst>")
        .map(|offset| list_start + offset + b"</a:extLst>".len())
        .expect("native extLst close");
    let family = std::str::from_utf8(native_family_fragment()).expect("family UTF-8");
    let hidden = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="x"><a:extLst><a:ext uri="{native}">{family}</a:ext></a:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        native = part::NATIVE_EXTENSION_URI,
    );
    let mut replaced = Vec::with_capacity(source.len() + hidden.len());
    replaced.extend_from_slice(&source[..list_start]);
    replaced.extend_from_slice(hidden.as_bytes());
    replaced.extend_from_slice(&source[list_end..]);
    let replaced = String::from_utf8(replaced)
        .expect("hidden-family Theme is UTF-8")
        .replacen(
            "<a:theme ",
            "<a:theme xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:litchi:mce\" ",
            1,
        );
    replaced.into_bytes()
}

fn mce_selected_extlst_theme() -> Vec<u8> {
    let source = NATIVE_THEME;
    let list_start = source
        .windows(b"<a:extLst>".len())
        .position(|window| window == b"<a:extLst>")
        .expect("native extLst");
    let list_end = source[list_start..]
        .windows(b"</a:extLst>".len())
        .position(|window| window == b"</a:extLst>")
        .map(|offset| list_start + offset + b"</a:extLst>".len())
        .expect("native extLst close");
    let old = &source[list_start..list_end];
    let old_text = std::str::from_utf8(old).expect("extLst UTF-8");
    let replacement = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="a">{old_text}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#
    );
    let wrapped = replace_bytes(source, old, replacement.as_bytes());
    String::from_utf8(wrapped)
        .expect("selected MCE Theme is UTF-8")
        .replacen(
            "<a:theme ",
            "<a:theme xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
            1,
        )
        .into_bytes()
}

#[test]
fn native_complete_theme_part_keeps_core_and_family_projections_distinct() {
    let parsed = codec::read(NATIVE_THEME).expect("native complete Theme part");
    assert_eq!(parsed.name, "Office Theme");
    assert_eq!(parsed.colors.name(), "Office");
    assert_eq!(parsed.fonts.major().latin, "Calibri Light");

    let family = family::read(native_family_fragment()).expect("native family extension");
    assert_eq!(family.name(), "Office Theme");
    assert_eq!(family.id().as_str(), OFFICE_ID);
    assert_eq!(family.variant_id().as_str(), OFFICE_VID);
    assert_eq!(family.source(), Some(native_family_fragment()));
    assert_eq!(
        family::write(&family).expect("exact family replay"),
        native_family_fragment()
    );
}

#[test]
fn shared_theme_part_owner_reads_profiles_and_replays_exact_source() {
    let snapshot = part::read(NATIVE_THEME).expect("native complete Theme owner");
    assert_eq!(snapshot.xml_bytes(), NATIVE_THEME);
    assert_eq!(
        snapshot.family_profile(),
        Some(part::ExtensionProfile::NativeDiscriminator)
    );
    assert_eq!(
        snapshot.family_extension_uri(),
        Some(part::NATIVE_EXTENSION_URI)
    );
    assert_eq!(snapshot.family_range(), Some(native_family_range()));
    assert_eq!(
        &snapshot.xml_bytes()[snapshot.family_range().expect("native family range")],
        native_family_fragment()
    );

    let detached =
        family::Family::new("Edited Office", OFFICE_ID, OFFICE_VID).expect("detached family");
    let replaced = snapshot
        .replace_family(&detached)
        .expect("replace source-backed family");
    let changed = part::read(&replaced).expect("read replaced Theme owner");
    assert_eq!(
        changed.family().expect("changed family").name(),
        "Edited Office"
    );
    assert_eq!(
        changed.family_profile(),
        Some(part::ExtensionProfile::NativeDiscriminator)
    );

    let original = family::Family::new("Office Theme", OFFICE_ID, OFFICE_VID)
        .expect("original detached family");
    let restored = changed
        .replace_family(&original)
        .expect("restore source-backed family");
    assert_eq!(restored, NATIVE_THEME);
}

#[test]
fn shared_theme_part_owner_add_remove_is_exact_and_absent_remove_is_noop() {
    let removed = part::remove_family(NATIVE_THEME).expect("remove native family");
    let no_op = part::remove_family(&removed).expect("remove absent family");
    assert_eq!(no_op, removed);
    assert!(
        part::read(&removed)
            .expect("read family-free Theme")
            .family()
            .is_none()
    );
    assert_eq!(
        part::family_profile(&removed).expect("family-free profile"),
        None
    );

    let detached =
        family::Family::new("Office Theme", OFFICE_ID, OFFICE_VID).expect("detached family");
    let restored = part::add_family_with_uri(&removed, &detached, part::NATIVE_EXTENSION_URI)
        .expect("add native family");
    let restored_snapshot = part::read(&restored).expect("read restored native family");
    assert_eq!(
        restored_snapshot.family_profile(),
        Some(part::ExtensionProfile::NativeDiscriminator)
    );
    assert_eq!(restored_snapshot.family(), Some(&detached));

    let normative = String::from_utf8(removed)
        .expect("family-free Theme is UTF-8")
        .replace(part::NATIVE_EXTENSION_URI, part::EXTENSION_URI)
        .into_bytes();
    let added = part::add_family(&normative, &detached).expect("add normative family");
    let parsed = part::read(&added).expect("read normative family");
    assert_eq!(
        parsed.family_profile(),
        Some(part::ExtensionProfile::Normative)
    );
    assert_eq!(parsed.family_extension_uri(), Some(part::EXTENSION_URI));
    assert_eq!(
        parsed.family().expect("normative family").name(),
        "Office Theme"
    );
}

#[test]
fn bom_prefixed_family_add_strips_only_the_leading_bom_before_insertion() {
    let mut bom_family = b"\xEF\xBB\xBF".to_vec();
    bom_family.extend_from_slice(native_family_fragment());
    let parsed = family::read(&bom_family).expect("read BOM-prefixed standalone family");
    let family_free = native_family_free_theme();
    let added = part::add_family_with_uri(&family_free, &parsed, part::NATIVE_EXTENSION_URI)
        .expect("standalone BOM is stripped before embedding");
    assert!(
        !added.windows(3).any(|window| window == b"\xEF\xBB\xBF"),
        "a standalone BOM must not be embedded inside a complete Theme ext"
    );
    assert!(
        part::read(&added).is_ok(),
        "BOM-free insertion remains readable"
    );
}

#[test]
fn admitted_extension_context_rejects_nonwhitespace_text_but_allows_whitespace_and_foreign_children()
 {
    // Keep the direct empty ext/extLst wrappers here: this test targets their
    // element-only content grammar independently of removal closure.
    let family_free = replace_native_family(b"");
    let empty_ext = native_empty_extension();

    let invalid_ext_text = replace_bytes(
        &family_free,
        &empty_ext,
        format!(r#"<a:ext uri="{}">bad</a:ext>"#, part::NATIVE_EXTENSION_URI).as_bytes(),
    );
    let invalid_ext_cdata = replace_bytes(
        &family_free,
        &empty_ext,
        format!(
            r#"<a:ext uri="{}"><![CDATA[bad]]></a:ext>"#,
            part::NATIVE_EXTENSION_URI
        )
        .as_bytes(),
    );
    let invalid_ext_reference = replace_bytes(
        &family_free,
        &empty_ext,
        format!(
            r#"<a:ext uri="{}">&#65;</a:ext>"#,
            part::NATIVE_EXTENSION_URI
        )
        .as_bytes(),
    );
    for (label, source) in [
        ("a:ext text", invalid_ext_text),
        ("a:ext CDATA", invalid_ext_cdata),
        ("a:ext general reference", invalid_ext_reference),
    ] {
        assert!(
            part::read(&source).is_err(),
            "accepted nonwhitespace {label}"
        );
    }

    let invalid_extlst_text = replace_bytes(&family_free, b"<a:extLst>", b"<a:extLst>bad");
    let invalid_extlst_cdata =
        replace_bytes(&family_free, b"<a:extLst>", b"<a:extLst><![CDATA[bad]]>");
    let invalid_extlst_reference = replace_bytes(&family_free, b"<a:extLst>", b"<a:extLst>&#65;");
    for (label, source) in [
        ("a:extLst text", invalid_extlst_text),
        ("a:extLst CDATA", invalid_extlst_cdata),
        ("a:extLst general reference", invalid_extlst_reference),
    ] {
        assert!(
            part::read(&source).is_err(),
            "accepted nonwhitespace {label}"
        );
    }

    let valid_ext_whitespace = replace_bytes(
        &family_free,
        &empty_ext,
        format!(
            "<a:ext uri=\"{}\">\r\n \t</a:ext>",
            part::NATIVE_EXTENSION_URI
        )
        .as_bytes(),
    );
    assert!(part::read(&valid_ext_whitespace).is_ok());

    let valid_extlst_whitespace = replace_bytes(&family_free, b"<a:extLst>", b"<a:extLst>\r\n \t");
    assert!(part::read(&valid_extlst_whitespace).is_ok());

    let valid_foreign_child = replace_bytes(
        &family_free,
        &empty_ext,
        format!(
            r#"<a:ext uri="{}"><x:future xmlns:x="urn:litchi:foreign"><x:child/></x:future></a:ext>"#,
            part::NATIVE_EXTENSION_URI
        )
        .as_bytes(),
    );
    assert!(part::read(&valid_foreign_child).is_ok());
}

#[test]
fn inherited_family_projection_is_reusable_and_insertable() {
    let source = inherited_opaque_family_theme();
    let projected = part::read_family(&source)
        .expect("read inherited family projection")
        .expect("inherited family owner");
    let written = family::write(&projected).expect("write inherited family projection");
    let reread = family::read(&written).expect("reopen written inherited family");
    assert_eq!(reread, projected);

    let family_free = part::remove_family(NATIVE_THEME).expect("remove native family");
    let inserted = part::add_family_with_uri(&family_free, &projected, part::NATIVE_EXTENSION_URI)
        .expect("insert inherited family projection");
    assert_eq!(
        part::read(&inserted)
            .expect("read inserted family")
            .family()
            .expect("inserted family"),
        &projected
    );
}

#[test]
fn hidden_mce_family_is_ignored_and_all_mutations_refuse_without_editing_subtree() {
    let source = mce_hidden_extlst_theme();
    let snapshot = part::read(&source).expect("read MCE-wrapped Theme");
    assert!(snapshot.family().is_none());

    let detached =
        family::Family::new("Fresh Office", OFFICE_ID, OFFICE_VID).expect("detached family");
    assert!(
        snapshot
            .add_family_with_uri(&detached, part::NATIVE_EXTENSION_URI)
            .is_err()
    );
    assert!(snapshot.replace_family(&detached).is_err());
    assert!(snapshot.remove_family().is_err());
    assert!(part::add_family_with_uri(&source, &detached, part::NATIVE_EXTENSION_URI).is_err());
    assert!(part::replace_family(&source, &detached).is_err());
    assert!(part::remove_family(&source).is_err());
    assert_eq!(snapshot.xml_bytes(), source.as_slice());
}

#[test]
fn selected_mce_family_owner_refuses_a_second_direct_owner() {
    let source = mce_selected_extlst_theme();
    let Ok(snapshot) = part::read(&source) else {
        // Refusing an unsupported selected MCE owner while reading is a safe
        // outcome: callers must not receive a snapshot that can create a
        // second effective owner through direct insertion.
        return;
    };
    assert!(snapshot.family().is_none());
    let detached =
        family::Family::new("Second Direct Family", OFFICE_ID, OFFICE_VID).expect("family");
    assert!(
        snapshot
            .add_family_with_uri(&detached, part::NATIVE_EXTENSION_URI)
            .is_err(),
        "adding a direct owner must refuse a selected MCE owner"
    );
}

#[test]
fn shared_theme_part_rejects_raw_xml_delimiters_quick_xml_tokenizes() {
    let text = replace_bytes(NATIVE_THEME, b"</a:theme>", b"]]></a:theme>");
    let family_text = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:raw-text" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child>]]></x:child></thm15:themeFamily>"#
    );
    let root_attribute = replace_bytes(
        NATIVE_THEME,
        b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"",
        b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office <Theme\"",
    );
    let family_attribute = replace_bytes(
        NATIVE_THEME,
        b"name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\"",
        b"name=\"Office <Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\"",
    );
    let rejected = (
        part::read(&text).is_err(),
        part::read(&replace_native_family(family_text.as_bytes())).is_err(),
        part::read(&root_attribute).is_err(),
        part::read(&family_attribute).is_err(),
    );
    assert_eq!(
        rejected,
        (true, true, true, true),
        "raw XML delimiter acceptance (Theme text, family text, root attribute, family attribute)"
    );
}

#[test]
fn shared_theme_part_allows_delimiters_in_comments_and_attributes() {
    let theme_comment = replace_bytes(
        NATIVE_THEME,
        b"</a:theme>",
        b"<!-- a comment may contain ]]> --></a:theme>",
    );
    let family_comment = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:comment" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child><!-- a comment may contain ]]> --></x:child></thm15:themeFamily>"#
    );
    let raw_attribute = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:raw-attribute" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child marker="raw ]]>"></x:child></thm15:themeFamily>"#
    );
    let escaped_attribute = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:escaped-attribute" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child marker="escaped ]]&gt;"></x:child></thm15:themeFamily>"#
    );

    let accepted = (
        part::read(&theme_comment).is_ok(),
        part::read(&replace_native_family(family_comment.as_bytes())).is_ok(),
        part::read(&replace_native_family(raw_attribute.as_bytes())).is_ok(),
        part::read(&replace_native_family(escaped_attribute.as_bytes())).is_ok(),
    );
    assert_eq!(
        accepted,
        (true, true, true, true),
        "XML delimiter scope (Theme comment, family comment, raw attribute, escaped attribute)"
    );
}

#[test]
fn add_reuses_an_extension_list_prefix_when_the_container_is_prefixed() {
    let retained = format!(
        "{}<!-- retain extension wrapper -->",
        std::str::from_utf8(native_family_fragment()).expect("family UTF-8")
    );
    let source = replace_native_family(retained.as_bytes());
    let removed = part::remove_family(&source).expect("remove native family");
    let empty = String::from_utf8(removed)
        .expect("family-free Theme is UTF-8")
        .replace(
            &format!(
                "<a:ext uri=\"{}\"><!-- retain extension wrapper --></a:ext>",
                part::NATIVE_EXTENSION_URI
            ),
            "",
        )
        .replace(
            "<a:extLst>",
            &format!("<p:extLst xmlns:p=\"{}\">", family::DRAWINGML_NAMESPACE),
        )
        .replace("</a:extLst>", "</p:extLst>")
        .into_bytes();
    let detached =
        family::Family::new("Prefixed Family", OFFICE_ID, OFFICE_VID).expect("detached family");
    let added = part::add_family_with_uri(&empty, &detached, part::NATIVE_EXTENSION_URI)
        .expect("add family to prefixed extLst");
    let marker = format!("<p:ext uri=\"{}\">", part::NATIVE_EXTENSION_URI).into_bytes();
    assert!(
        added
            .windows(marker.len())
            .any(|window| window == marker.as_slice())
    );
    assert_eq!(
        part::read(&added)
            .expect("read prefixed extension list")
            .family()
            .expect("prefixed family")
            .name(),
        "Prefixed Family"
    );
}

#[test]
fn complete_theme_codec_accepts_family_extension_with_inherited_namespace() {
    let source = String::from_utf8(NATIVE_THEME.to_vec()).expect("native UTF-8");
    let family = native_family_fragment();
    let family_without_local_declaration =
        String::from_utf8(family.iter().copied().filter(|_| true).collect::<Vec<_>>())
            .expect("family UTF-8")
            .replace(&format!(" xmlns:thm15=\"{FAMILY_NAMESPACE}\""), "");
    let source = source
        .replace(
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
            &format!(
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:thm15=\"{FAMILY_NAMESPACE}\""
            ),
        )
        .replace(
            std::str::from_utf8(family).expect("family UTF-8"),
            &family_without_local_declaration,
        );

    let parsed = codec::read(source.as_bytes()).expect("inherited family does not alter Theme");
    assert_eq!(parsed.name, "Office Theme");
}

#[test]
fn family_edit_preserves_unknown_attributes_children_and_exact_inverse() {
    let source = opaque_family();
    let snapshot = family::Snapshot::from_xml(source.clone()).expect("opaque family snapshot");
    assert_eq!(snapshot.xml_bytes(), source.as_slice());

    let mut edit = snapshot.edit();
    edit.set_name("Edited Office").expect("set family name");
    let commit = edit.commit().expect("commit family edit");
    assert!(commit.changed());
    assert_eq!(commit.snapshot().value().name(), "Edited Office");
    assert!(
        commit
            .snapshot()
            .xml_bytes()
            .windows(b"x:marker=\"keep\"".len())
            .any(|window| window == b"x:marker=\"keep\"")
    );
    assert!(
        commit
            .snapshot()
            .xml_bytes()
            .windows(b"<x:future value=\"keep\"/>".len())
            .any(|window| window == b"<x:future value=\"keep\"/>")
    );

    let restored = commit
        .patch()
        .clone()
        .inverse()
        .apply(commit.snapshot())
        .expect("exact inverse family patch");
    assert_eq!(restored.xml_bytes(), source.as_slice());
}

#[test]
fn scalar_family_update_preserves_processing_tokens_in_opaque_markup() {
    let source = opaque_markup_with_processing_token();
    let snapshot = family::Snapshot::from_xml(source.clone()).expect("opaque family snapshot");
    let mut edit = snapshot.edit();
    edit.set_name("Edited Office").expect("set family name");
    let commit = edit
        .commit()
        .expect("scalar update preserves opaque comment and CDATA");
    assert_eq!(commit.snapshot().value().name(), "Edited Office");
    assert!(
        commit
            .snapshot()
            .xml_bytes()
            .windows(b"<!-- opaque <? marker -->".len())
            .any(|window| window == b"<!-- opaque <? marker -->")
    );
    assert!(
        commit
            .snapshot()
            .xml_bytes()
            .windows(b"<![CDATA[opaque <? marker]]>".len())
            .any(|window| window == b"<![CDATA[opaque <? marker]]>")
    );
    assert_eq!(
        family::read(commit.snapshot().xml_bytes())
            .expect("reopen scalar-updated family")
            .name(),
        "Edited Office"
    );

    let complete = replace_native_family(&source);
    let snapshot = part::read(&complete).expect("complete Theme with processing token");
    let detached = family::Family::new("Edited Complete Office", OFFICE_ID, OFFICE_VID)
        .expect("detached complete replacement family");
    let changed = snapshot
        .replace_family(&detached)
        .expect("complete Theme scalar update preserves processing token");
    assert!(
        changed
            .windows(b"<!-- opaque <? marker -->".len())
            .any(|window| { window == b"<!-- opaque <? marker -->" })
    );
    assert!(
        changed
            .windows(b"<![CDATA[opaque <? marker]]>".len())
            .any(|window| { window == b"<![CDATA[opaque <? marker]]>" })
    );
}

#[test]
fn namespace_validation_rejects_empty_prefixed_and_xmlns_bindings() {
    let empty_prefix = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"><payload x:y="1"/></thm15:themeFamily>"#
    );
    let xmlns_uri = family::XMLNS_NAMESPACE;
    let xmlns_default = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns="{xmlns_uri}" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#
    );
    let xml_uri_default = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns="{xml_uri}" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#,
        xml_uri = family::XML_NAMESPACE,
    );
    let xml_uri_prefixed = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="{xml_uri}" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child/></thm15:themeFamily>"#,
        xml_uri = family::XML_NAMESPACE,
    );
    let xml_uri_escaped = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child/></thm15:themeFamily>"#
    );
    assert!(family::read(empty_prefix.as_bytes()).is_err());
    assert!(family::read(xmlns_default.as_bytes()).is_err());
    assert!(family::read(xml_uri_default.as_bytes()).is_err());
    assert!(family::read(xml_uri_prefixed.as_bytes()).is_err());
    assert!(family::read(xml_uri_escaped.as_bytes()).is_err());
    assert!(part::read(&replace_native_family(empty_prefix.as_bytes())).is_err());
    assert!(part::read(&replace_native_family(xmlns_default.as_bytes())).is_err());
    assert!(part::read(&replace_native_family(xml_uri_default.as_bytes())).is_err());
    assert!(part::read(&replace_native_family(xml_uri_prefixed.as_bytes())).is_err());
    assert!(part::read(&replace_native_family(xml_uri_escaped.as_bytes())).is_err());
}

#[test]
fn legal_default_namespace_reset_and_namespace_heavy_foreign_children_round_trip() {
    let default_reset = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns="urn:litchi:default-parent" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"><payload xmlns=""/></thm15:themeFamily>"#
    )
    .into_bytes();
    let reset = family::read(&default_reset).expect("legal default namespace reset");
    assert_eq!(
        family::write(&reset).expect("write default namespace reset"),
        default_reset
    );
    let complete_reset = replace_native_family(&default_reset);
    assert_eq!(
        part::read(&complete_reset)
            .expect("complete Theme with legal default namespace reset")
            .family()
            .expect("complete reset family")
            .name(),
        "Office"
    );

    let heavy = namespace_heavy_foreign_family();
    assert!(heavy.len() < family::MAX_XML_BYTES);
    let source = replace_native_family(&heavy);
    let snapshot = part::read(&source).expect("bounded namespace-heavy family");
    assert_eq!(
        snapshot.family().expect("namespace-heavy family").name(),
        "Office"
    );
    let detached = family::Family::new("Edited Office", OFFICE_ID, OFFICE_VID)
        .expect("detached replacement family");
    let changed = snapshot
        .replace_family(&detached)
        .expect("scalar update namespace-heavy family");
    let reopened = part::read(&changed).expect("reopen namespace-heavy family");
    assert_eq!(
        reopened.family().expect("edited family").name(),
        "Edited Office"
    );
    assert!(
        changed
            .windows(b"p0:marker=\"0\"".len())
            .any(|window| window == b"p0:marker=\"0\"")
    );
    assert!(
        changed
            .windows(b"p63:marker=\"63\"".len())
            .any(|window| window == b"p63:marker=\"63\"")
    );
}

#[test]
fn family_noop_and_stale_source_are_source_checked() {
    let source = opaque_family();
    let snapshot = family::Snapshot::from_xml(source.clone()).expect("family snapshot");
    let noop = snapshot.edit().commit().expect("family no-op");
    assert!(!noop.changed());
    assert_eq!(noop.snapshot().xml_bytes(), source.as_slice());
    assert_eq!(noop.patch().before_xml(), noop.patch().after_xml());

    let mut changed = snapshot.edit();
    changed
        .set_id("{11111111-2222-3333-4444-555555555555}")
        .expect("change family id");
    let changed = changed.commit().expect("changed family");
    let other =
        family::Snapshot::from_xml(source.iter().copied().chain(Some(b' ')).collect::<Vec<_>>())
            .expect("different source snapshot");
    assert!(changed.patch().apply(&other).is_err());
}

#[test]
fn complete_theme_limits_and_malformed_family_are_bounded() {
    assert!(codec::read(&NATIVE_THEME[..NATIVE_THEME.len() - 1]).is_err());

    let mut malformed = NATIVE_THEME.to_vec();
    let family = native_family_fragment();
    let replacement = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" name="Office" id="{OFFICE_ID}"/>"#
    );
    let start = malformed
        .windows(family.len())
        .position(|window| window == family)
        .expect("family in native source");
    malformed.splice(start..start + family.len(), replacement.bytes());
    assert!(
        family::read(
            malformed
                .windows(replacement.len())
                .find(|window| window.starts_with(b"<thm15:themeFamily"))
                .unwrap_or_default()
        )
        .is_err()
    );

    let over = vec![b' '; family::MAX_XML_BYTES + 1];
    assert!(family::read(&over).is_err());
    assert!(codec::read(&vec![b' '; codec::MAX_XML_BYTES + 1]).is_err());
}

#[test]
fn shared_theme_part_validates_xml_declarations_and_root_boundaries() {
    let declaration = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    let mut bom = b"\xEF\xBB\xBF".to_vec();
    bom.extend_from_slice(NATIVE_THEME);
    let _ = part::read(&bom).unwrap_or_else(|error| panic!("UTF-8 BOM was rejected: {error:?}"));

    let mut duplicate_declaration = declaration.to_vec();
    duplicate_declaration.extend_from_slice(declaration);
    let duplicate = replace_bytes(NATIVE_THEME, declaration, &duplicate_declaration);
    let version = replace_bytes(NATIVE_THEME, b"version=\"1.0\"", b"version=\"1.1\"");
    let encoding = replace_bytes(
        NATIVE_THEME,
        b"encoding=\"UTF-8\"",
        b"encoding=\"ISO-8859-1\"",
    );
    let standalone = replace_bytes(NATIVE_THEME, b"standalone=\"yes\"", b"standalone=\"maybe\"");
    let unsupported_attribute = replace_bytes(
        NATIVE_THEME,
        b"standalone=\"yes\"",
        b"standalone=\"yes\" foo=\"bar\"",
    );
    for (label, source) in [
        ("duplicate declaration", duplicate),
        ("XML 1.1", version),
        ("non-UTF-8 encoding", encoding),
        ("invalid standalone", standalone),
        ("unsupported declaration attribute", unsupported_attribute),
    ] {
        assert!(part::read(&source).is_err(), "accepted {label}");
    }

    let mut declaration_after_whitespace = b"\n".to_vec();
    declaration_after_whitespace.extend_from_slice(NATIVE_THEME);
    assert!(part::read(&declaration_after_whitespace).is_err());
    let mut declaration_after_comment = b"<!-- before declaration -->".to_vec();
    declaration_after_comment.extend_from_slice(NATIVE_THEME);
    assert!(part::read(&declaration_after_comment).is_err());

    let without_declaration = &NATIVE_THEME[declaration.len()..];
    let mut leading_text = b"outside".to_vec();
    leading_text.extend_from_slice(without_declaration);
    assert!(part::read(&leading_text).is_err());
    let mut leading_cdata = b"<![CDATA[outside]]>".to_vec();
    leading_cdata.extend_from_slice(without_declaration);
    assert!(part::read(&leading_cdata).is_err());

    let mut trailing_text = NATIVE_THEME.to_vec();
    trailing_text.extend_from_slice(b"outside");
    assert!(part::read(&trailing_text).is_err());
    let mut trailing_cdata = NATIVE_THEME.to_vec();
    trailing_cdata.extend_from_slice(b"<![CDATA[outside]]>");
    assert!(part::read(&trailing_cdata).is_err());
    let mut trailing_whitespace = NATIVE_THEME.to_vec();
    trailing_whitespace.extend_from_slice(b"\r\n \t");
    assert!(part::read(&trailing_whitespace).is_ok());
}

#[test]
fn shared_theme_part_rejects_unknown_references_and_forbidden_controls() {
    let standard_references = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:refs" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child value="&amp;&#10;">&amp;&#10;</x:child></thm15:themeFamily>"#
    );
    assert!(
        part::read(&replace_native_family(standard_references.as_bytes())).is_ok(),
        "predefined and character references in opaque family markup are valid XML"
    );

    let unknown_reference = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" xmlns:x="urn:litchi:bad" name="Office Theme" id="{OFFICE_ID}" vid="{OFFICE_VID}"><x:child>&foo;</x:child></thm15:themeFamily>"#
    );
    assert!(part::read(&replace_native_family(unknown_reference.as_bytes())).is_err());

    let mut forbidden_control = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" name="Office" id="{OFFICE_ID}" vid="{OFFICE_VID}"/>"#
    )
    .into_bytes();
    let marker = b"name=\"Office\"";
    let start = forbidden_control
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("family name marker");
    forbidden_control.splice(
        start..start + marker.len(),
        b"name=\"bad\0name\"".iter().copied(),
    );
    assert!(part::read(&replace_native_family(&forbidden_control)).is_err());
}

#[test]
fn removing_a_malformed_recognized_family_refuses_to_sanitize_it() {
    let malformed = format!(
        r#"<thm15:themeFamily xmlns:thm15="{FAMILY_NAMESPACE}" name="Office Theme" id="{OFFICE_ID}"/>"#
    );
    let source = replace_native_family(malformed.as_bytes());
    assert!(part::read(&source).is_err());
    assert!(part::remove_family(&source).is_err());
}

#[test]
fn oversized_family_is_rejected_inside_an_under_limit_theme_part() {
    let source = oversized_family_theme();
    assert!(source.len() > family::MAX_XML_BYTES);
    assert!(source.len() < codec::MAX_XML_BYTES);
    assert!(part::read(&source).is_err());
    assert!(part::remove_family(&source).is_err());
}
