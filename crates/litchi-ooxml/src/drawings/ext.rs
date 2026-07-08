use std::fmt;
use std::fmt::Write as _;

pub fn write_a16_creation_id_extlst(xml: &mut String, creation_id: &str) -> fmt::Result {
    write!(
        xml,
        r#"<a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{}\"/></a:ext></a:extLst>"#,
        creation_id
    )
}
