//! Focused regression tests for the bounded DDE owner.

use super::{OFFICE, TEXT};
use crate::variable_declaration::{Part, parse_parts};

const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:body><o:text>"#;
const SUFFIX: &str = "</o:text></o:body></o:document-content>";

#[test]
fn retains_inert_declarations_and_references() {
    let xml = format!(
        r#"{PREFIX}<t:dde-connection-decls>
                <t:dde-connection-decl o:name="Prices" o:dde-application="soffice"
                    o:dde-topic="file:///tmp/prices.ods" o:dde-item="Sheet1.A1"
                    o:automatic-update="true"/>
            </t:dde-connection-decls><t:p><t:dde-connection t:connection-name="Prices"/></t:p>{SUFFIX}"#
    );
    let inventory = parse_parts(&[(xml.as_str(), Part::Content)]).unwrap();
    assert_eq!(inventory.dde_connections.len(), 1);
    assert_eq!(inventory.dde_connection_uses.len(), 1);
    let declaration = &inventory.dde_connections[0];
    assert_eq!(declaration.name, "Prices");
    assert_eq!(declaration.topic, "file:///tmp/prices.ods");
    assert!(declaration.effective_automatic_update());
    assert_eq!(inventory.dde_connection_uses[0].connection_name, "Prices");
}

#[test]
fn accepts_declaration_from_another_part_and_default_update() {
    let styles = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:master-styles/></o:document-styles>"#
    );
    let content = format!(
        r#"{PREFIX}<t:dde-connection-decls><t:dde-connection-decl o:name="Feed"
                o:dde-application="app" o:dde-topic="topic" o:dde-item="item"/>
                </t:dde-connection-decls><t:dde-connection t:connection-name="Feed"/>{SUFFIX}"#
    );
    let parsed = parse_parts(&[
        (styles.as_str(), Part::Styles),
        (content.as_str(), Part::Content),
    ])
    .unwrap();
    assert!(!parsed.dde_connections[0].effective_automatic_update());
}

#[test]
fn rejects_malformed_active_and_unresolved_connections() {
    let bodies = [
        r#"<t:dde-connection-decls><t:dde-connection-decl o:name="x" o:dde-application="a" o:dde-topic="t"/></t:dde-connection-decls>"#,
        r#"<t:dde-connection-decls><t:dde-connection-decl o:name="x" o:dde-application="a" o:dde-topic="t" o:dde-item="i" o:automatic-update="yes"/></t:dde-connection-decls>"#,
        r#"<t:dde-connection-decls><t:dde-connection-decl o:name="x" o:dde-application="a" o:dde-topic="t" o:dde-item="i">payload</t:dde-connection-decl></t:dde-connection-decls>"#,
        r#"<t:dde-connection t:connection-name="missing"/>"#,
        r"<t:dde-connection-decls><t:p/></t:dde-connection-decls>",
    ];
    for body in bodies {
        let xml = format!("{PREFIX}{body}{SUFFIX}");
        assert!(
            parse_parts(&[(xml.as_str(), Part::Content)]).is_err(),
            "accepted {body}"
        );
    }
}
