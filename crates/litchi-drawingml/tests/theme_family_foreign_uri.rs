#![allow(clippy::expect_used, reason = "regression fixture assertions")]

//! Foreign URI ownership must be checked before rejecting QName lookalikes.

use litchi_drawingml::theme::family::{Family, part};

const NATIVE: &str = include_str!("fixtures/theme-part-native.xml");

#[test]
fn foreign_namespace_lookalikes_are_opaque_under_unknown_extension_uris() {
    let start = NATIVE.find("<thm15:themeFamily").expect("native family");
    let end = start + NATIVE[start..].find("/>").expect("family end") + 2;
    let detached = Family::new(
        "Added",
        "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
        "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
    )
    .expect("detached family");
    for payload in [
        r#"<x:themeFamily xmlns:x="urn:vendor" marker="keep"/>"#,
        r#"<x:themeFamily xmlns:x="urn:vendor"><x:payload>keep</x:payload></x:themeFamily>"#,
    ] {
        let template = format!("{}{payload}{}", &NATIVE[..start], &NATIVE[end..]);
        for uri in [part::EXTENSION_URI, part::NATIVE_EXTENSION_URI] {
            let recognized = template.replace(part::NATIVE_EXTENSION_URI, uri);
            assert!(part::read(recognized.as_bytes()).is_err());
        }

        let source = template.replace(part::NATIVE_EXTENSION_URI, "urn:vendor:unknown-extension");
        let snapshot = part::read(source.as_bytes()).expect("foreign lookalike is opaque");
        assert!(snapshot.family().is_none());
        assert_eq!(snapshot.xml_bytes(), source.as_bytes());
        assert!(
            part::read_family(source.as_bytes())
                .expect("borrowed projection")
                .is_none()
        );
        assert_eq!(
            snapshot.remove_family().expect("absent no-op"),
            source.as_bytes()
        );

        let added = snapshot
            .add_family(&detached)
            .expect("add direct normative owner");
        assert!(
            std::str::from_utf8(&added)
                .expect("Theme UTF-8")
                .contains(payload)
        );
        let reopened = part::read(&added).expect("read added owner");
        assert_eq!(reopened.family(), Some(&detached));
        assert_eq!(
            reopened.family_profile(),
            Some(part::ExtensionProfile::Normative)
        );
        assert_eq!(
            reopened.remove_family().expect("remove only typed owner"),
            source.as_bytes()
        );
    }
}
