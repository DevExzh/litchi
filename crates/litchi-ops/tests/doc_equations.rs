//! DOC equations through the unified API: Word-written MathType and
//! Equation Editor objects read as `equation` nodes, and added, replaced and
//! removed for source and eager owners, with read-back and the other
//! equations kept.
#![cfg(feature = "doc")]

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode};
#[cfg(feature = "equation")]
use litchi_ops::{Placement, props};

const TRIGONOMETRIC_RATIOS: &[u8] =
    include_bytes!("../../../test-data/ole/doc/trigonometric_ratios.doc");

/// Both owners of `bytes`: source-backed and eager.
fn owners(bytes: &[u8]) -> [File; 2] {
    [SourceMode::Auto, SourceMode::Eager].map(|mode| {
        litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode)).unwrap()
    })
}

#[cfg(feature = "equation")]
fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

#[cfg(feature = "equation")]
fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[cfg(feature = "equation")]
fn equation_paths(file: &File) -> Vec<String> {
    file.query("equation")
        .unwrap()
        .map(|node| node.unwrap().view().path().to_string())
        .collect()
}

#[cfg(feature = "equation")]
fn latexes(file: &File) -> Vec<Option<String>> {
    equation_paths(file)
        .iter()
        .map(|path| prop(file, path, "latex"))
        .collect()
}

#[cfg(feature = "equation")]
fn mtef(file: &File, path: &str) -> Vec<u8> {
    use base64::Engine as _;
    base64::engine::general_purpose::STANDARD
        .decode(prop(file, path, "mtef").unwrap())
        .unwrap()
}

#[cfg(not(feature = "equation"))]
#[test]
fn equations_need_the_equation_feature() {
    for file in owners(TRIGONOMETRIC_RATIOS) {
        let error = file
            .add("/body/p[1]", NewNode::new("equation").prop("latex", "x"))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

#[cfg(feature = "equation")]
#[test]
fn word_equations_are_listed_with_latex_and_mtef() {
    let file = litchi_ops::open_bytes(TRIGONOMETRIC_RATIOS.to_vec()).unwrap();
    let paths = equation_paths(&file);
    assert_eq!(paths.len(), 671);
    // The first equation is in the eighth body paragraph.
    let first = &paths[0];
    assert!(first.ends_with("/equation[1]"), "{first}");
    assert_eq!(
        prop(&file, first, "latex").as_deref(),
        Some(r"P(r\cos \alpha,r\sin \alpha)")
    );
    assert_eq!(prop(&file, first, "mode").as_deref(), Some("inline"));
    let native = mtef(&file, first);
    assert_eq!(native[28], 3, "Equation Editor 3");
    assert_eq!(
        litchi_equation::mtef_to_latex(&native, &litchi_equation::Limits::DEFAULT).unwrap(),
        r"P(r\cos \alpha,r\sin \alpha)"
    );
    assert_eq!(mtef(&file, &paths[1])[28], 5, "MathType");
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_is_added_set_and_removed() {
    for file in owners(TRIGONOMETRIC_RATIOS) {
        let before = file.to_bytes().unwrap();
        let old = latexes(&file);
        let first = equation_paths(&file)[0].clone();
        let paragraph = first.rsplit_once('/').unwrap().0.to_owned();

        // Before the paragraph's first run: the new equation comes first.
        let added = file
            .add(
                &paragraph,
                NewNode::new("equation")
                    .prop("latex", r"\frac{1}{2}")
                    .at(Placement::before(&format!("{paragraph}/r[1]")).unwrap()),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), format!("{paragraph}/equation[1]"));
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
        let read = reopen(&added.file);
        let mut expected = old.clone();
        expected.insert(0, Some(r"\frac{1}{2}".to_owned()));
        assert_eq!(latexes(&read), expected);
        assert_eq!(mtef(&read, &added.path.to_string())[28], 5, "MathType");

        // An Equation Editor 3 equation stays MTEF 3.
        let second = format!("{paragraph}/equation[2]");
        let set = read.set(&second, props! { "latex" => "y^{2}" }).unwrap();
        let read = reopen(&set);
        assert_eq!(prop(&read, &second, "latex").as_deref(), Some("y^{2}"));
        assert_eq!(mtef(&read, &second)[28], 3);
        expected[1] = Some("y^{2}".to_owned());
        assert_eq!(latexes(&read), expected);

        // Raw MTEF is written as given.
        let raw = mtef(&file, &equation_paths(&file)[2]);
        let raw_text = prop(&file, &equation_paths(&file)[2], "mtef").unwrap();
        let set = read
            .set(
                &added.path.to_string(),
                props! { "mtef" => raw_text.as_str() },
            )
            .unwrap();
        let read = reopen(&set);
        assert_eq!(mtef(&read, &added.path.to_string()), raw);

        let removed = read.remove(&added.path.to_string()).unwrap();
        let read = reopen(&removed);
        expected.remove(0);
        assert_eq!(latexes(&read), expected);
    }
}

#[cfg(feature = "equation")]
#[test]
fn a_removed_equation_leaves_no_storage_or_preview() {
    let streams = |file: &File| {
        let ole = litchi_cfb::OleFile::open(
            std::io::Cursor::new(file.to_bytes().unwrap()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        ole.list_streams()
    };
    let data = |file: &File| {
        let mut ole = litchi_cfb::OleFile::open(
            std::io::Cursor::new(file.to_bytes().unwrap()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        ole.open_stream(&["Data"]).unwrap_or_default()
    };
    let created = litchi_ops::create(litchi_core::FileFormat::Doc).unwrap();
    let mut file = created.clone();
    // Twice, reopening between: the second equation reuses the first one's
    // storage number and Data offset.
    for _ in 0..2 {
        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("equation").prop("latex", r"\frac{a}{b}"),
            )
            .unwrap();
        let removed = reopen(&added.file).remove(&added.path.to_string()).unwrap();
        file = reopen(&removed);
        assert!(
            streams(&file)
                .iter()
                .all(|path| path.first().map(String::as_str) != Some("ObjectPool")),
            "{:?}",
            streams(&file)
        );
        assert_eq!(data(&file), data(&created));
    }
}

#[cfg(feature = "equation")]
#[test]
fn equations_are_added_to_a_new_document() {
    let file = litchi_ops::create(litchi_core::FileFormat::Doc).unwrap();
    let file = file
        .set("/body/p[1]", props! { "text" => "Area: " })
        .unwrap();
    let added = file
        .add(
            "/body/p[1]",
            NewNode::new("equation").prop("latex", r"\pi r^{2}"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[1]/equation[1]");
    let read = reopen(&added.file);
    assert_eq!(latexes(&read), vec![Some(r"\pi r^{2}".to_owned())]);
    let omml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>q</m:t></m:r></m:oMath>"#;
    let added = read
        .add("/body/p[1]", NewNode::new("math").prop("omml", omml))
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[1]/equation[2]");
    assert_eq!(
        latexes(&reopen(&added.file)),
        vec![Some(r"\pi r^{2}".to_owned()), Some("q".to_owned())]
    );
}

#[cfg(feature = "equation")]
#[test]
fn bad_equation_writes_are_typed_refusals() {
    let file = litchi_ops::open_bytes(TRIGONOMETRIC_RATIOS.to_vec()).unwrap();
    let first = equation_paths(&file)[0].clone();
    let error = file
        .set(&first, props! { "latex" => r"\frac{" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set(&first, props! { "mtef" => "not base64!" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set(&first, props! { "latex" => "x", "omml" => "<m:oMath/>" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file.set(&first, props! { "mode" => "inline" }).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    // MathType MTEF does not replace an Equation Editor 3 object.
    let math_type = prop(&file, &equation_paths(&file)[1], "mtef").unwrap();
    let error = file
        .set(&first, props! { "mtef" => math_type.as_str() })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    assert!(error.to_string().contains("Equation.3"), "{error}");
    // A conversion failure says what the converter found.
    let error = file
        .set(&first, props! { "latex" => r"\frac{" })
        .unwrap_err();
    assert!(error.to_string().contains("does not convert"), "{error}");
}

#[cfg(feature = "equation")]
#[test]
fn added_mtef_keeps_the_kind_its_version_names() {
    let file = litchi_ops::open_bytes(TRIGONOMETRIC_RATIOS.to_vec()).unwrap();
    let paths = equation_paths(&file);
    // The first equation is Equation Editor 3, the second MathType.
    for (source, version) in [(&paths[0], 3), (&paths[1], 5)] {
        let raw = prop(&file, source, "mtef").unwrap();
        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("equation").prop("mtef", raw.as_str()),
            )
            .unwrap();
        let read = reopen(&added.file);
        let path = added.path.to_string();
        assert_eq!(mtef(&read, &path)[28], version);
        assert_eq!(mtef(&read, &path), mtef(&file, source));
        assert_eq!(prop(&read, &path, "latex"), prop(&file, source, "latex"));
    }
}
