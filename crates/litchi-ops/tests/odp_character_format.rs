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
    assert_eq!(commit.diagnostics.commits(), 1);
    for paragraph in 1..=4 {
        let path = format!("/slide[1]/shape[2]/p[{paragraph}]/r[1]");
        assert_eq!(prop(&commit.file, &path, "bold").as_deref(), Some("true"));
    }
    assert_eq!(prop(&commit.file, RUN, "italic").as_deref(), Some("true"));
    let restored = commit.patch.inverse().apply(&commit.file).unwrap();
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
    assert_eq!(commit.diagnostics.commits(), 2);
    assert_eq!(
        prop(&commit.file, "/slide[1]/shape[2]/p[1]/r[1]", "bold").as_deref(),
        Some("true")
    );

    let error = file
        .set(RUN, props! { "color" => "accent1" })
        .map(drop)
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
    let error = file
        .set("/slide[1]/shape[1]", props! { "bold" => true })
        .map(drop)
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
}
