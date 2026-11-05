//! ODP character formatting through derived automatic styles (record 0308,
//! G-ODP-1): bold, italic, underline, size, colour and font on runs and
//! paragraphs. A set never changes a style other markup may share; it
//! copies the node's automatic style under a new name, merges the
//! properties, and points the node at the copy. Every other byte of
//! `content.xml` keeps its source.

#![cfg(feature = "odp")]

use litchi_ops::{Code, File, props};

fn open(relative: &str) -> File {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative);
    litchi_ops::open_bytes(std::fs::read(path).unwrap()).unwrap()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn content(file: &File) -> String {
    let package = litchi_odp::Package::from_bytes(&file.to_bytes().unwrap()).unwrap();
    package.presentation().unwrap().content_xml().to_owned()
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

const RUN: &str = "/slide[1]/shape[2]/p[1]/r[1]";

#[test]
fn run_formatting_derives_a_style_and_keeps_the_rest() {
    let file = open("odf/corpus/impress-basic.odp");
    // `T1` sets 24pt and nothing else.
    assert_eq!(prop(&file, RUN, "size").as_deref(), Some("24pt"));
    assert_eq!(prop(&file, RUN, "bold"), None);
    let source = content(&file);
    let changed = file
        .set(
            RUN,
            props! { "bold" => true, "italic" => true, "color" => "C00000", "font" => "Liberation Serif" },
        )
        .unwrap();
    let changed = reopen(&changed);
    assert_eq!(prop(&changed, RUN, "bold").as_deref(), Some("true"));
    assert_eq!(prop(&changed, RUN, "italic").as_deref(), Some("true"));
    assert_eq!(prop(&changed, RUN, "size").as_deref(), Some("24pt"));
    assert_eq!(prop(&changed, RUN, "color").as_deref(), Some("#C00000"));
    assert_eq!(
        prop(&changed, RUN, "font").as_deref(),
        Some("Liberation Serif")
    );
    let style = prop(&changed, RUN, "style").unwrap();
    assert_ne!(style, "T1");
    // The other runs that share T1 keep it, and their formatting.
    assert_eq!(
        prop(&changed, "/slide[1]/shape[2]/p[2]/r[1]", "style").as_deref(),
        Some("T1")
    );
    assert_eq!(prop(&changed, "/slide[1]/shape[2]/p[2]/r[1]", "bold"), None);
    // Only the span's start tag and the new style differ.
    let after = content(&changed);
    let definition = after
        .find(&format!("<style:style style:name=\"{style}\""))
        .unwrap();
    let end = definition + after[definition..].find("</style:style>").unwrap() + 14;
    let without = format!("{}{}", &after[..definition], &after[end..]);
    let without = without.replacen(
        &format!("text:style-name=\"{style}\">Eins"),
        "text:style-name=\"T1\">Eins",
        1,
    );
    assert_eq!(without, source);
}

#[test]
fn paragraph_formatting_and_no_ops() {
    let file = open("odf/corpus/impress-basic.odp");
    let title = "/slide[1]/shape[1]/p[1]";
    let changed = file
        .set(title, props! { "underline" => true, "size" => "40pt" })
        .unwrap();
    assert_eq!(prop(&changed, title, "underline").as_deref(), Some("true"));
    assert_eq!(prop(&changed, title, "size").as_deref(), Some("40pt"));
    let style = prop(&changed, title, "style").unwrap();
    assert_ne!(style, "P1");
    // The derived paragraph style keeps P1's paragraph properties.
    let after = content(&changed);
    let at = after
        .find(&format!("<style:style style:name=\"{style}\""))
        .unwrap();
    assert!(after[at..].starts_with(&format!(
        "<style:style style:name=\"{style}\" style:family=\"paragraph\"><style:paragraph-properties fo:margin-left=\"0cm\""
    )));
    // A value the style already has is an exact no-op.
    let same = file.set(RUN, props! { "size" => "24pt" }).unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
    // A set of the text and the formatting of one run in one call.
    let both = file
        .set(RUN, props! { "text" => "One", "bold" => true })
        .unwrap();
    assert_eq!(prop(&both, RUN, "bold").as_deref(), Some("true"));
    assert_eq!(
        both.get(RUN).unwrap().view().text().unwrap().as_deref(),
        Some("One")
    );
}

#[test]
fn batches_and_refusals() {
    let file = open("odf/corpus/impress-basic.odp");
    let commit = file
        .edit(|tx| {
            for paragraph in 1..=4 {
                tx.set(
                    format!("/slide[1]/shape[2]/p[{paragraph}]/r[1]").as_str(),
                    props! { "bold" => true },
                )?;
            }
            tx.set(RUN, props! { "italic" => true })?;
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 1);
    for paragraph in 1..=4 {
        let path = format!("/slide[1]/shape[2]/p[{paragraph}]/r[1]");
        assert_eq!(prop(commit.file(), &path, "bold").as_deref(), Some("true"));
    }
    assert_eq!(prop(commit.file(), RUN, "italic").as_deref(), Some("true"));
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());

    // A text set followed by formatting of the same story resolves against
    // the text set's result.
    let commit = file
        .edit(|tx| {
            tx.set("/slide[1]/shape[2]/p[1]", props! { "text" => "Plain" })?;
            tx.set("/slide[1]/shape[2]/p[1]/r[1]", props! { "bold" => true })?;
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(
        prop(commit.file(), "/slide[1]/shape[2]/p[1]/r[1]", "bold").as_deref(),
        Some("true")
    );

    let error = file
        .set(RUN, props! { "color" => "accent1" })
        .map(drop)
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
    let error = file
        .set("/slide[1]/shape[1]", props! { "strike" => true })
        .map(drop)
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
}

/// Record 0316: a shape, a table cell and the notes take formatting as a
/// whole; every run then reads it, and the inverse restores the bytes.
#[test]
fn shapes_cells_and_notes_take_formatting_as_a_whole() {
    let file = open("odf/corpus/impress-basic.odp")
        .set("/slide[1]/notes", props! { "text" => "Speaker notes" })
        .unwrap();
    let file = reopen(&file);
    let shape = "/slide[1]/shape[2]";
    assert_eq!(prop(&file, shape, "size"), None, "the levels differ");
    let commit = file
        .edit(|tx| {
            tx.set(shape, props! { "bold" => true, "size" => "20pt" })?;
            tx.set("/slide[1]/notes", props! { "italic" => true })?;
            Ok(())
        })
        .unwrap();
    let changed = reopen(commit.file());
    assert_eq!(prop(&changed, shape, "bold").as_deref(), Some("true"));
    assert_eq!(prop(&changed, shape, "size").as_deref(), Some("20pt"));
    // The spans' T1..T4 set the size, so they take the new one too.
    for path in [
        "/slide[1]/shape[2]/p[1]/r[1]",
        "/slide[1]/shape[2]/p[3]/r[1]",
    ] {
        assert_eq!(
            prop(&changed, path, "size").as_deref(),
            Some("20pt"),
            "{path}"
        );
        assert_eq!(
            prop(&changed, path, "bold").as_deref(),
            Some("true"),
            "{path}"
        );
    }
    assert_eq!(
        prop(&changed, "/slide[1]/notes", "italic").as_deref(),
        Some("true")
    );
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    // The same values again: an exact no-op.
    let again = commit.file().set(shape, props! { "bold" => true }).unwrap();
    assert_eq!(again.to_bytes().unwrap(), commit.file().to_bytes().unwrap());

    let table = open("odf/odp/cellspan.odp");
    let cell = "/slide[1]/table[1]/tr[1]/tc[2]";
    let changed = table
        .set(cell, props! { "underline" => true, "color" => "1F4E79" })
        .unwrap();
    let changed = reopen(&changed);
    assert_eq!(prop(&changed, cell, "underline").as_deref(), Some("true"));
    assert_eq!(prop(&changed, cell, "color").as_deref(), Some("#1F4E79"));
    assert_eq!(
        prop(&changed, "/slide[1]/table[1]/tr[1]/tc[1]", "underline"),
        None
    );
}

/// A flat presentation with a common paragraph style, a link and a field.
const FLAT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
    r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" "#,
    r#"office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">"#,
    r#"<office:styles>"#,
    r#"<style:style style:name="Emphasis" style:family="paragraph"><style:text-properties fo:font-weight="bold" fo:font-size="28pt"/></style:style>"#,
    r#"</office:styles>"#,
    r#"<office:automatic-styles>"#,
    r#"<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Emphasis"><style:paragraph-properties fo:text-align="center"/></style:style>"#,
    r#"</office:automatic-styles>"#,
    r#"<office:body><office:presentation><draw:page draw:name="One">"#,
    r#"<draw:frame draw:name="Box" svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="3cm"><draw:text-box>"#,
    r#"<text:p text:style-name="P1">See <text:a xlink:href="https://example.com">the site</text:a> on page <text:page-number>1</text:page-number></text:p>"#,
    r#"</draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document>"#,
);

#[test]
fn links_fields_and_common_styles() {
    let file = litchi_ops::open_bytes(FLAT.as_bytes().to_vec()).unwrap();
    let paragraph = "/slide[1]/shape[1]/p[1]";
    // P1 sets no character property; its parent common style does.
    assert_eq!(prop(&file, paragraph, "bold").as_deref(), Some("true"));
    assert_eq!(prop(&file, paragraph, "size").as_deref(), Some("28pt"));
    assert_eq!(
        prop(&file, "/slide[1]/shape[1]", "bold").as_deref(),
        Some("true")
    );
    let link = "/slide[1]/shape[1]/p[1]/r[2]";
    let field = "/slide[1]/shape[1]/p[1]/r[4]";
    let changed = file
        .edit(|tx| {
            tx.set(link, props! { "italic" => true })?;
            tx.set(field, props! { "color" => "C00000" })?;
            Ok(())
        })
        .unwrap()
        .into_file();
    let changed = litchi_ops::open_bytes(changed.to_bytes().unwrap()).unwrap();
    assert_eq!(prop(&changed, link, "italic").as_deref(), Some("true"));
    assert_eq!(
        prop(&changed, link, "link").as_deref(),
        Some("https://example.com")
    );
    assert_eq!(prop(&changed, field, "color").as_deref(), Some("#C00000"));
    assert_eq!(
        changed
            .get(field)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("1")
    );
    // The field still refuses a text set, inside its formatting span.
    let error = changed
        .set(field, props! { "text" => "2" })
        .map(drop)
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
}

#[test]
fn a_created_deck_takes_run_formatting() {
    let deck = litchi_ops::create(litchi_core::FileFormat::Odp).unwrap();
    let deck = deck
        .add("/", litchi_ops::NewNode::new("slide"))
        .unwrap()
        .file;
    let deck = deck
        .add(
            "/slide[1]",
            litchi_ops::NewNode::new("shape")
                .prop("text", "Hello")
                .prop("width", "8cm")
                .prop("height", "2cm"),
        )
        .unwrap()
        .file;
    let path = "/slide[1]/shape[1]/paragraph[1]/run[1]";
    let changed = deck
        .set(path, props! { "bold" => true, "size" => "24pt" })
        .unwrap_or_else(|error| {
            panic!(
                "{error}: {:?}",
                std::error::Error::source(&error).map(ToString::to_string)
            )
        });
    assert_eq!(prop(&changed, path, "bold").as_deref(), Some("true"));
}
