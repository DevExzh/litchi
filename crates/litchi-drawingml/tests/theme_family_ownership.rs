#![allow(clippy::expect_used, reason = "regression fixtures must fail loudly")]

//! Direct complete-Theme scanner regressions, independent of XSD validation.

use litchi_drawingml::theme::family::{self, Family, part};

const NATIVE: &str = include_str!("fixtures/theme-part-native.xml");
const ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

fn fragment() -> &'static str {
    let start = NATIVE.find("<thm15:themeFamily").expect("native family");
    let end = start + NATIVE[start..].find("/>").expect("family end") + 2;
    &NATIVE[start..end]
}

fn with_extensions(contents: &str) -> String {
    let start = NATIVE.rfind("<a:extLst>").expect("root extension list");
    let end = NATIVE.rfind("</a:extLst>").expect("list end") + "</a:extLst>".len();
    format!(
        "{}<a:extLst>{contents}</a:extLst>{}",
        &NATIVE[..start],
        &NATIVE[end..]
    )
}

fn extension(uri: &str, contents: &str) -> String {
    format!("<a:ext uri=\"{uri}\">{contents}</a:ext>")
}

#[test]
fn spaced_uri_diagnostics_are_normalized_while_read_add_and_edit_keep_source_spelling() {
    for profile in [
        part::ExtensionProfile::Normative,
        part::ExtensionProfile::NativeDiscriminator,
    ] {
        let uri = profile.uri();
        for lexical in [format!(" \t{uri}\r\n "), format!("&#x20;{uri}&#9;")] {
            let attribute = format!("uri=\"{lexical}\"");
            let source = with_extensions(&extension(&lexical, fragment()));
            let snapshot = part::read(source.as_bytes()).expect("spaced recognized URI");
            assert_eq!(snapshot.family_extension_uri(), Some(uri));
            assert_eq!(snapshot.family_profile(), Some(profile));
            assert_eq!(snapshot.xml_bytes(), source.as_bytes());

            let removed = snapshot
                .remove_family()
                .expect("retain spaced empty extension");
            let empty = part::read(&removed).expect("read retained empty extension");
            assert_eq!(empty.family_extension_uri(), None);
            assert_eq!(empty.family_profile(), None);
            let detached = Family::new("Added", ID, VID).expect("new family");
            let added = empty
                .add_family(&detached)
                .expect("reuse existing spaced URI");
            let added_snapshot = part::read(&added).expect("read added family");
            assert_eq!(added_snapshot.family_extension_uri(), Some(uri));
            assert_eq!(added_snapshot.family_profile(), Some(profile));
            assert_eq!(added_snapshot.family(), Some(&detached));

            let replacement = Family::new("Replaced", ID, VID).expect("replacement family");
            let changed = added_snapshot
                .replace_family(&replacement)
                .expect("update with spaced URI");
            let changed_snapshot = part::read(&changed).expect("read replacement");
            assert_eq!(changed_snapshot.family_extension_uri(), Some(uri));
            assert_eq!(changed_snapshot.family_profile(), Some(profile));
            assert_eq!(changed_snapshot.family(), Some(&replacement));
            for bytes in [&removed, &added, &changed] {
                assert!(
                    std::str::from_utf8(bytes)
                        .expect("Theme UTF-8")
                        .contains(&attribute)
                );
            }
        }
    }
}

#[test]
fn duplicate_supported_extensions_and_families_refuse_every_scanner_edit() {
    let native = extension(part::NATIVE_EXTENSION_URI, fragment());
    let normative = extension(part::EXTENSION_URI, fragment());
    let duplicate_family = extension(
        part::NATIVE_EXTENSION_URI,
        &format!("{}{}", fragment(), fragment()),
    );
    let detached = Family::new("Changed", ID, VID).expect("detached family");
    for contents in [
        format!("{native}{native}"),
        format!("{native}{normative}"),
        format!(
            "{}{}",
            extension(part::NATIVE_EXTENSION_URI, ""),
            extension(part::EXTENSION_URI, "")
        ),
        duplicate_family,
    ] {
        let source = with_extensions(&contents);
        assert!(part::read(source.as_bytes()).is_err());
        assert!(part::read_family(source.as_bytes()).is_err());
        assert!(part::family_range(source.as_bytes()).is_err());
        assert!(part::add_family(source.as_bytes(), &detached).is_err());
        assert!(part::replace_family(source.as_bytes(), &detached).is_err());
        assert!(part::remove_family(source.as_bytes()).is_err());
    }
}

#[test]
fn duplicate_family_lookalikes_under_foreign_uris_remain_opaque() {
    let foreign = extension("urn:unrecognized:family", &fragment().repeat(2));
    let source = with_extensions(&foreign.repeat(2));
    let snapshot = part::read(source.as_bytes()).expect("opaque foreign extensions");
    assert!(snapshot.family().is_none());
    assert_eq!(
        snapshot.remove_family().expect("absent no-op"),
        source.as_bytes()
    );

    let detached = Family::new("Added", ID, VID).expect("detached family");
    let added = snapshot
        .add_family(&detached)
        .expect("add direct normative owner");
    let added_text = std::str::from_utf8(&added).expect("Theme UTF-8");
    assert!(added_text.contains(&foreign.repeat(2)));
    let reopened = part::read(&added).expect("reopen direct owner");
    assert_eq!(reopened.family(), Some(&detached));
    assert_eq!(
        reopened.family_profile(),
        Some(part::ExtensionProfile::Normative)
    );
}

#[test]
fn inherited_bindings_used_only_by_opaque_descendants_survive_detached_replacement() {
    let opaque = format!(
        "{}<x:child y:flag=\"keep\" hint=\"z:Type\"/></thm15:themeFamily>",
        fragment()
            .strip_suffix("/>")
            .expect("empty native family")
            .to_owned()
            + ">"
    );
    let source = with_extensions(&extension(part::NATIVE_EXTENSION_URI, &opaque)).replacen(
        "<a:theme ",
        "<a:theme xmlns:x=\"urn:opaque:element\" xmlns:y=\"urn:opaque:attribute\" xmlns:z=\"urn:opaque:qname\" ",
        1,
    );
    let snapshot = part::read(source.as_bytes()).expect("inherited descendant bindings");
    let projected = snapshot.family().expect("family owner");
    let standalone = family::write(projected).expect("namespace-complete projection");
    let standalone_text = std::str::from_utf8(&standalone).expect("family UTF-8");
    for declaration in [
        "xmlns:x=\"urn:opaque:element\"",
        "xmlns:y=\"urn:opaque:attribute\"",
        "xmlns:z=\"urn:opaque:qname\"",
    ] {
        assert!(standalone_text.contains(declaration));
    }
    assert_eq!(
        family::read(&standalone).expect("standalone readback"),
        *projected
    );
    let detached = Family::new("Changed", ID, VID).expect("fresh scalar values");
    let changed = snapshot
        .replace_family(&detached)
        .expect("replace detached scalars");
    assert_eq!(
        changed,
        source
            .replacen(
                &opaque,
                &opaque.replacen("name=\"Office Theme\"", "name=\"Changed\"", 1),
                1,
            )
            .as_bytes()
    );
    assert_eq!(
        part::read(&changed).expect("changed readback").family(),
        Some(&detached)
    );
}

#[test]
fn removal_retains_empty_wrappers_and_exact_whitespace_comments_and_payload() {
    for (list_prefix, ext_prefix, ext_suffix, list_suffix) in [
        ("", "", "", ""),
        (" \r\n", "\t ", "\r\n ", "\t"),
        (
            "<!-- list-before -->",
            "<!-- ext-before -->",
            "<!-- ext-after -->",
            "<!-- list-after -->",
        ),
        (
            "",
            "<v:opaque xmlns:v=\"urn:vendor\">keep</v:opaque>",
            "",
            "",
        ),
    ] {
        let body = format!("{ext_prefix}{}{ext_suffix}", fragment());
        let source = with_extensions(&format!(
            "{list_prefix}{}{list_suffix}",
            extension(part::NATIVE_EXTENSION_URI, &body)
        ));
        let snapshot = part::read(source.as_bytes()).expect("removal source");
        let removed = snapshot.remove_family().expect("remove selected element");
        assert_eq!(removed, source.replacen(fragment(), "", 1).as_bytes());
        let reopened = part::read(&removed).expect("retained wrappers remain readable");
        assert!(reopened.family().is_none());
        assert_eq!(
            reopened.remove_family().expect("second removal no-op"),
            removed
        );
    }
}
