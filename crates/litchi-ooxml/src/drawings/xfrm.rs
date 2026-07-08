use std::fmt;
use std::fmt::Write as _;

pub fn write_a_xfrm_off_ext(xml: &mut String, x: i64, y: i64, cx: i64, cy: i64) -> fmt::Result {
    xml.push_str("<a:xfrm>");
    write!(xml, r#"<a:off x=\"{}\" y=\"{}\"/>"#, x, y)?;
    write!(xml, r#"<a:ext cx=\"{}\" cy=\"{}\"/>"#, cx, cy)?;
    xml.push_str("</a:xfrm>");
    Ok(())
}
