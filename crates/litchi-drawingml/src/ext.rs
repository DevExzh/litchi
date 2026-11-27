//! `DrawingML` extension-list primitives.

use std::{fmt, fmt::Write as _};

/// Write the Office 2014 creation-ID extension with escaped content.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_creation_id(xml: &mut String, creation_id: &str) -> fmt::Result {
    let creation_id = quick_xml::escape::escape(creation_id);
    write!(
        xml,
        r#"<a:extLst><a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{creation_id}"/></a:ext></a:extLst>"#,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn creation_ids_are_xml_escaped() {
        let mut xml = String::new();
        write_creation_id(&mut xml, "<&").expect("String formatting is infallible");
        assert!(xml.contains("id=\"&lt;&amp;\""));
    }
}
