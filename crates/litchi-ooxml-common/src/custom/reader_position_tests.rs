//! Where the custom-properties reader accepts each kind of XML content, and
//! what a strict and a tolerant read make of it.

use super::codec::{decode, decode_tolerant};
use super::model::Props;
use super::schema::FORMAT_ID;
use crate::Result;

const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const ROOT_OPEN: &str = concat!(
    r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" "#,
    r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#
);
const ROOT_CLOSE: &str = "</Properties>";
const EMPTY_ROOT: &str = concat!(
    r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" "#,
    r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>"#
);

fn open(pid: i32, name: &str) -> String {
    format!(r#"<property fmtid="{FORMAT_ID}" pid="{pid}" name="{name}">"#)
}

fn property(pid: i32, name: &str, content: &str) -> String {
    format!("{}{content}</property>", open(pid, name))
}

fn good(pid: i32, name: &str) -> String {
    property(pid, name, "<vt:i4>7</vt:i4>")
}

/// A document whose root holds `body`.
fn rooted(body: &str) -> String {
    format!("{DECLARATION}{ROOT_OPEN}{body}{ROOT_CLOSE}")
}

fn summarize(result: Result<Props>) -> String {
    let props = match result {
        Ok(props) => props,
        Err(error) => return format!("Err({error:?})").replace('"', "'"),
    };
    let mut values: Vec<String> = props
        .iter()
        .map(|(name, value)| format!("{name}={value:?}"))
        .collect();
    values.sort();
    let defects: Vec<String> = props
        .tolerance_report()
        .iter()
        .map(|tolerated| {
            format!(
                "{:?}({:?}, {:?})",
                tolerated.defect(),
                tolerated.item().unwrap_or("-"),
                tolerated.value().unwrap_or("-")
            )
        })
        .collect();
    format!("Ok[{}] repaired[{}]", values.join(", "), defects.join(", ")).replace('"', "'")
}

/// The strict and the tolerant outcome of reading `xml`.
fn outcomes(xml: &str) -> (String, String) {
    (
        summarize(decode(xml.as_bytes())),
        summarize(decode_tolerant(xml.as_bytes(), "/docProps/custom.xml")),
    )
}

fn cases() -> Vec<String> {
    let after = good(9, "After");
    let unknown_child = r#"<x:ext xmlns:x="urn:x">"#;
    vec![
        // Prolog, root and epilog.
        rooted(&good(2, "A")),
        format!("{ROOT_OPEN}{}{ROOT_CLOSE}", good(2, "A")),
        format!("{DECLARATION}{EMPTY_ROOT}"),
        EMPTY_ROOT.to_owned(),
        rooted(" \n<!--c-->\n "),
        format!("<!--c-->\n{DECLARATION}{EMPTY_ROOT}"),
        format!(" {DECLARATION}{EMPTY_ROOT}"),
        format!("{DECLARATION}{DECLARATION}{EMPTY_ROOT}"),
        format!("{DECLARATION}\n<!--c-->\n{EMPTY_ROOT}\n<!--c-->\n"),
        format!("{EMPTY_ROOT}{DECLARATION}"),
        format!("{ROOT_OPEN}{ROOT_CLOSE}{DECLARATION}"),
        rooted(DECLARATION),
        rooted(&property(2, "A", DECLARATION)),
        rooted(&property(2, "A", &format!("<vt:i4>{DECLARATION}7</vt:i4>"))),
        String::new(),
        DECLARATION.to_owned(),
        "<!--c-->".to_owned(),
        format!("{EMPTY_ROOT}{EMPTY_ROOT}"),
        format!("{ROOT_OPEN}{ROOT_CLOSE}{ROOT_OPEN}{ROOT_CLOSE}"),
        format!("{ROOT_OPEN}{ROOT_CLOSE}{EMPTY_ROOT}"),
        format!("{EMPTY_ROOT}{ROOT_OPEN}{ROOT_CLOSE}"),
        r#"<Properties xmlns="urn:wrong"/>"#.to_owned(),
        r#"<Other xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"></Other>"#.to_owned(),
        ROOT_OPEN.to_owned(),
        format!("{ROOT_OPEN}{}", open(2, "A")),
        format!("{ROOT_OPEN}{}<vt:i4>7", open(2, "A")),
        format!("{ROOT_OPEN}{unknown_child}"),
        format!("{ROOT_OPEN}{unknown_child}<x:y>"),
        ROOT_CLOSE.to_owned(),
        // Character content at each level.
        format!("text{EMPTY_ROOT}"),
        format!("{EMPTY_ROOT}text"),
        format!("{EMPTY_ROOT}\n"),
        rooted(&format!("text{after}")),
        rooted(&format!("{}{after}", property(2, "A", "text<vt:i4>7</vt:i4>"))),
        rooted(&format!("{}{after}", property(2, "A", "<vt:i4>7</vt:i4>text"))),
        rooted(&format!("{}{after}", property(2, "A", " <vt:i4>7</vt:i4> "))),
        rooted(&property(2, "A", "<vt:lpwstr> a <!--c--> b </vt:lpwstr>")),
        format!("<![CDATA[c]]>{EMPTY_ROOT}"),
        format!("{EMPTY_ROOT}<![CDATA[c]]>"),
        rooted(&format!("<![CDATA[c]]>{after}")),
        rooted(&format!("{}{after}", property(2, "A", "<![CDATA[c]]><vt:i4>7</vt:i4>"))),
        rooted(&property(2, "A", "<vt:lpwstr>a<![CDATA[<b>]]>c</vt:lpwstr>")),
        format!("&amp;{EMPTY_ROOT}"),
        format!("{EMPTY_ROOT}&amp;"),
        rooted(&format!("&amp;{after}")),
        rooted(&format!("{}{after}", property(2, "A", "&amp;<vt:i4>7</vt:i4>"))),
        rooted(&property(2, "A", "<vt:lpwstr>a&amp;&#66;&lt;</vt:lpwstr>")),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:lpwstr>a&nbsp;b</vt:lpwstr>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:lpwstr>a&nbsp;b</vt:lpwstr><vt:i4>7</vt:i4>")
        )),
        // Elements below the root.
        rooted(&format!("{unknown_child}</x:ext>{after}")),
        rooted(&format!("{unknown_child}<x:y>t&amp;<![CDATA[c]]><x:z/><!--c--></x:y></x:ext>{after}")),
        rooted(&format!("{unknown_child}<x:y><x:z></x:z></x:y></x:ext>{after}")),
        rooted(&format!("{unknown_child}{DECLARATION}</x:ext>{after}")),
        rooted(&format!("{unknown_child}<?pi x?></x:ext>{after}")),
        rooted(&format!("{unknown_child}<!DOCTYPE x></x:ext>{after}")),
        rooted(&format!(r#"<x:ext xmlns:x="urn:x"/>{after}"#)),
        rooted(&format!(
            r#"<property fmtid="{FORMAT_ID}" pid="2" name="Empty"/>{after}"#
        )),
        rooted(&format!(r#"<property pid="2"/>{after}"#)),
        rooted(&format!("{}{after}", good(9, "SamePid"))),
        rooted(&format!("{}{after}", good(3, "after"))),
        rooted(&format!(
            r#"<property pid="2" name="NoFmtid"><vt:i4>7</vt:i4>text</property>{after}"#
        )),
        rooted(&format!("{}{after}", property(2, "A", &good(3, "Nested")))),
        rooted(&format!(
            "{}{after}",
            property(2, "A", &format!("<vt:i4>7</vt:i4>{}", good(3, "Nested")))
        )),
        // Values.
        rooted(&format!("{}{after}", property(2, "A", ""))),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:i4>7</vt:i4><vt:i4>8</vt:i4>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:i4>7</vt:i4><vt:lpwstr/>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:lpwstr/><vt:i4>7</vt:i4>")
        )),
        rooted(&format!("{}{after}", property(2, "A", "<vt:lpwstr/>"))),
        rooted(&format!("{}{after}", property(2, "A", "<vt:i4/>"))),
        rooted(&format!("{}{after}", property(2, "A", "<vt:date/>"))),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:date>2026-01-01</vt:date>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", r#"<vt:i4 a="b">7</vt:i4>"#)
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", r#"<x:i4 xmlns:x="urn:x">7</x:i4>"#)
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:bool>maybe</vt:bool>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:bool>maybe</vt:bool><vt:i4>7</vt:i4>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:bool>maybe</vt:bool>text")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:lpwstr><x/></vt:lpwstr>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:lpwstr><x></x></vt:lpwstr>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:date><x/></vt:date>")
        )),
        rooted(&format!(
            "{}{after}",
            property(2, "A", "<vt:date><x></x></vt:date>")
        )),
        // Processing instructions and DTDs.
        format!("<?pi x?>{EMPTY_ROOT}"),
        format!("{EMPTY_ROOT}<?pi x?>"),
        rooted(&format!("<?pi x?>{after}")),
        rooted(&format!("{}{after}", property(2, "A", "<?pi x?><vt:i4>7</vt:i4>"))),
        rooted(&format!("{}{after}", property(2, "A", "<vt:i4>7<?pi x?></vt:i4>"))),
        format!("<!DOCTYPE Properties>{EMPTY_ROOT}"),
        rooted(&format!("<!DOCTYPE x>{after}")),
    ]
}

/// Writes the current outcomes for a maintainer who extends [`cases`].
fn render(found: &[(String, String)]) -> String {
    let mut text = String::new();
    for (strict, tolerant) in found {
        text.push_str(&format!(
            "    (\n        r#\"{strict}\"#,\n        r#\"{tolerant}\"#,\n    ),\n"
        ));
    }
    text
}

#[test]
fn reader_outcomes_by_document_position() {
    let found: Vec<(String, String)> = cases().iter().map(|xml| outcomes(xml)).collect();
    let expected: Vec<(String, String)> = EXPECTED
        .iter()
        .map(|(strict, tolerant)| ((*strict).to_owned(), (*tolerant).to_owned()))
        .collect();
    assert!(
        found == expected,
        "outcomes changed; found:\n{}",
        render(&found)
    );
}

const EXPECTED: &[(&str, &str)] = &[
    (r#"Ok[A=I32(7)] repaired[]"#, r#"Ok[A=I32(7)] repaired[]"#),
    (r#"Ok[A=I32(7)] repaired[]"#, r#"Ok[A=I32(7)] repaired[]"#),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
    ),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
    ),
    (
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
    ),
    (
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
    ),
    (
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
    ),
    (
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
        r#"Err(Invalid('XML declaration must occur once before the root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
        r#"Err(Invalid('custom-properties XML must contain exactly one root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties root must be Properties in namespace 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties''))"#,
        r#"Err(Invalid('custom-properties root must be Properties in namespace 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties''))"#,
    ),
    (
        r#"Err(Invalid('custom-properties root must be Properties in namespace 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties''))"#,
        r#"Err(Invalid('custom-properties root must be Properties in namespace 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties''))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Err(Invalid('custom-properties XML must contain one complete Properties root'))"#,
    ),
    (
        r#"Err(Xml('ill-formed document: close tag `</Properties>` does not match any open tag'))"#,
        r#"Err(Xml('ill-formed document: close tag `</Properties>` does not match any open tag'))"#,
    ),
    (
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
    ),
    (
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
    ),
    (r#"Ok[] repaired[]"#, r#"Ok[] repaired[]"#),
    (
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
    ),
    (
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('non-whitespace text is not allowed outside a property value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Ok[A=I32(7), After=I32(7)] repaired[]"#,
        r#"Ok[A=I32(7), After=I32(7)] repaired[]"#,
    ),
    (
        r#"Ok[A=Text(' a  b ')] repaired[]"#,
        r#"Ok[A=Text(' a  b ')] repaired[]"#,
    ),
    (
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
    ),
    (
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
    ),
    (
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
    ),
    (
        r#"Err(Invalid('CDATA is only allowed inside a property value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Ok[A=Text('a<b>c')] repaired[]"#,
        r#"Ok[A=Text('a<b>c')] repaired[]"#,
    ),
    (
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
    ),
    (
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
    ),
    (
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
    ),
    (
        r#"Err(Invalid('entity references are only allowed inside a property value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Ok[A=Text('a&B<')] repaired[]"#,
        r#"Ok[A=Text('a&B<')] repaired[]"#,
    ),
    (
        r#"Err(Invalid('unsupported entity reference '&nbsp;''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidValue)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('unsupported entity reference '&nbsp;''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidValue)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnknownElement)('x:ext', '-')]"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnknownElement)('x:ext', '-')]"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnknownElement)('x:ext', '-')]"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnknownElement)('x:ext', '-')]"#,
    ),
    (
        r#"Err(Invalid('unexpected element 'x:ext' below custom-properties root'))"#,
        r#"Err(Invalid('DTD declarations are forbidden in custom-properties XML'))"#,
    ),
    (
        r#"Err(Invalid('custom property must contain exactly one typed value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnknownElement)('x:ext', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom property must contain exactly one typed value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('Empty', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom property must contain exactly one typed value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('-', '-')]"#,
    ),
    (
        r#"Err(Invalid('duplicate custom property PID 9'))"#,
        r#"Ok[SamePid=I32(7)] repaired[CustomProperties(DuplicateProperty)('After', '9')]"#,
    ),
    (
        r#"Err(Invalid('duplicate custom property name 'After' (names are case-insensitive)'))"#,
        r#"Ok[after=I32(7)] repaired[CustomProperties(DuplicateProperty)('After', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom property 'NoFmtid' is missing its format ID'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidIdentity)('NoFmtid', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom-property value 'property' is not in the variant-types namespace'))"#,
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
    ),
    (
        r#"Err(Invalid('custom property 'A' must contain exactly one value'))"#,
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
    ),
    (
        r#"Err(Invalid('custom property 'A' must contain exactly one value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom property 'A' must contain exactly one value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom property 'A' must contain exactly one value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom property 'A' must contain exactly one value'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Ok[A=Text(''), After=I32(7)] repaired[]"#,
        r#"Ok[A=Text(''), After=I32(7)] repaired[]"#,
    ),
    (
        r#"Err(Invalid('invalid vt:i4 value '': cannot parse integer from empty string'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidValue)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('unsupported custom-property value type 'date''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnsupportedValueType)('A', 'vt:date')]"#,
    ),
    (
        r#"Err(Invalid('unsupported custom-property value type 'date''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnsupportedValueType)('A', 'vt:date')]"#,
    ),
    (
        r#"Err(Invalid('custom-property value has unexpected attribute 'a''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(MalformedProperty)('A', '-')]"#,
    ),
    (
        r#"Err(Invalid('custom-property value 'x:i4' is not in the variant-types namespace'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnsupportedValueType)('A', 'x:i4')]"#,
    ),
    (
        r#"Err(Invalid('invalid vt:bool value 'maybe''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidValue)('A', 'maybe')]"#,
    ),
    (
        r#"Err(Invalid('invalid vt:bool value 'maybe''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidValue)('A', 'maybe')]"#,
    ),
    (
        r#"Err(Invalid('invalid vt:bool value 'maybe''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(InvalidValue)('A', 'maybe')]"#,
    ),
    (
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
    ),
    (
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
    ),
    (
        r#"Err(Invalid('unsupported custom-property value type 'date''))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(UnsupportedValueType)('A', 'vt:date')]"#,
    ),
    (
        r#"Err(Invalid('unsupported custom-property value type 'date''))"#,
        r#"Err(Limit { kind: CustomPropertiesXmlDepth, observed: 4, maximum: 3 })"#,
    ),
    (
        r#"Err(Invalid('processing instructions are forbidden in custom-properties XML'))"#,
        r#"Ok[] repaired[CustomProperties(ProcessingInstruction)('-', '-')]"#,
    ),
    (
        r#"Err(Invalid('processing instructions are forbidden in custom-properties XML'))"#,
        r#"Ok[] repaired[CustomProperties(ProcessingInstruction)('-', '-')]"#,
    ),
    (
        r#"Err(Invalid('processing instructions are forbidden in custom-properties XML'))"#,
        r#"Ok[After=I32(7)] repaired[CustomProperties(ProcessingInstruction)('-', '-')]"#,
    ),
    (
        r#"Err(Invalid('processing instructions are forbidden in custom-properties XML'))"#,
        r#"Ok[A=I32(7), After=I32(7)] repaired[CustomProperties(ProcessingInstruction)('-', '-')]"#,
    ),
    (
        r#"Err(Invalid('processing instructions are forbidden in custom-properties XML'))"#,
        r#"Ok[A=I32(7), After=I32(7)] repaired[CustomProperties(ProcessingInstruction)('-', '-')]"#,
    ),
    (
        r#"Err(Invalid('DTD declarations are forbidden in custom-properties XML'))"#,
        r#"Err(Invalid('DTD declarations are forbidden in custom-properties XML'))"#,
    ),
    (
        r#"Err(Invalid('DTD declarations are forbidden in custom-properties XML'))"#,
        r#"Err(Invalid('DTD declarations are forbidden in custom-properties XML'))"#,
    ),
];
