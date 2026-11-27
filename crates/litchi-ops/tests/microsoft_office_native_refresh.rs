//! Native cache refresh is bounded and distinct from application layout or external execution.
#![cfg(all(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "eval"
))]
use litchi_ops::refresh::{Clock, Options, PageNumbers, Status};
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};
use std::{
    collections::BTreeMap,
    io::{Cursor, Write},
    path::Path,
};
fn native(relative: &str, mode: SourceMode) -> File {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office")
        .join(relative);
    let options = OpenOptions::new().source_mode(mode);
    if mode == SourceMode::Auto {
        litchi_ops::open_with(&path, &options).unwrap()
    } else {
        litchi_ops::open_bytes_with(std::fs::read(path).unwrap(), &options).unwrap()
    }
}
#[test]
fn native_blank_string_formula_cache_is_present_and_does_not_trigger_refresh() {
    use litchi_ops::calc::{Origin, Value};
    for ext in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(&format!("excel/excel-values-formulas-chart.{ext}"), mode);
            let before = f.to_bytes().unwrap();
            let cached = f
                .evaluate("Sheet1!B18", &litchi_ops::calc::Options::new())
                .unwrap();
            assert_eq!(
                cached.origin(),
                Origin::Cached,
                "{ext} {mode:?}: blank string is a saved value"
            );
            assert_eq!(cached.value(), &Value::Text(String::new()));
            let c = f.refresh(&Options::new().write_cache(true)).unwrap();
            assert!(
                c.patch().is_noop(),
                "{ext} {mode:?}: a current cache was rewritten"
            );
            assert!(c.file().to_bytes().unwrap() == before);
        }
    }
}

#[test]
fn native_chart_tabs_keep_cross_sheet_calculation_and_refresh_positions() {
    use litchi_ops::calc::Value;
    for ext in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(&format!("excel/excel-values-formulas-chart.{ext}"), mode);
            let v = f
                .evaluate(
                    "Sheet2!B3",
                    &litchi_ops::calc::Options::new().recalculate(true),
                )
                .unwrap();
            assert_eq!(v.value(), &Value::Number(49.75), "{ext} {mode:?}");
            let source = f.to_bytes().unwrap();
            let report = f.recalculate(&litchi_ops::calc::Options::new()).unwrap();
            assert_eq!(report.formulas().len(), 16);
            assert_eq!(
                report.formulas().iter().filter(|f| f.sheet() == 1).count(),
                14
            );
            assert_eq!(
                report.formulas().iter().filter(|f| f.sheet() == 2).count(),
                2
            );
            let error = f
                .evaluate("/sheet[1]/A1", &litchi_ops::calc::Options::new())
                .unwrap_err();
            assert_eq!(
                error.code(),
                if ext == "xls" {
                    Code::UnsupportedCapability
                } else {
                    Code::UnsupportedType
                },
                "{ext} {mode:?}"
            );
            if ext == "xlsx" {
                let name = f.workbook().unwrap().worksheet_names().unwrap()[0].clone();
                let invalid = f
                    .set(
                        "Sheet1!Z100",
                        props! {"formula" => format!("'{}'!A1", name.replace("'", "''"))},
                    )
                    .unwrap();
                assert_eq!(
                    invalid
                        .evaluate(
                            "Sheet1!Z100",
                            &litchi_ops::calc::Options::new().recalculate(true)
                        )
                        .unwrap_err()
                        .code(),
                    Code::UnsupportedCapability
                );
            }
            assert!(f.to_bytes().unwrap() == source);
            let c = f.refresh(&Options::new()).unwrap();
            assert!(
                c.value()
                    .fields()
                    .iter()
                    .any(|field| field.location() == Some("Sheet2!B3")),
                "{ext}: trailing worksheet omitted or mislabeled"
            );
        }
    }
}

fn source_bytes(relative: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(relative),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    p.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = p.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn xml_derivative(bytes: &[u8], name: &str, change: impl FnOnce(&str) -> String) -> Vec<u8> {
    let mut parts = members(bytes);
    let value = change(std::str::from_utf8(&parts[name]).unwrap());
    parts.insert(name.into(), value.into_bytes());
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in &parts {
        writer
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(data).unwrap();
    }
    let bytes = writer.finish().unwrap().into_inner();
    assert!(members(&bytes) == parts);
    bytes
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|n| {
            let b = ole
                .open_stream(&n.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (n, b)
        })
        .collect()
}
fn preserve_except(a: &[u8], b: &[u8], allowed: &[&str]) {
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if !allowed.contains(&name.as_str()) {
            assert!(b[&name] == bytes, "changed {name}");
        }
    }
}
fn open_derivative(bytes: &[u8], ext: &str, mode: SourceMode) -> (tempfile::TempDir, File) {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join(format!("derivative.{ext}"));
    std::fs::write(&path, bytes).unwrap();
    let options = OpenOptions::new().source_mode(mode);
    let file = if mode == SourceMode::Auto {
        litchi_ops::open_with(&path, &options).unwrap()
    } else {
        litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap()
    };
    (dir, file)
}
fn check_commit(
    base: &File,
    c: &litchi_ops::transaction::Commit<litchi_ops::refresh::Summary>,
    mode: SourceMode,
    ext: &str,
    check: impl FnOnce(&File),
) {
    let before = base.to_bytes().unwrap();
    let after = c.file().to_bytes().unwrap();
    assert_eq!(c.diagnostics().operations(), 1);
    assert_eq!(c.diagnostics().commits(), 1);
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join(format!("refreshed.{ext}"));
    std::fs::write(&path, &after).unwrap();
    let r = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(mode)).unwrap();
    check(&r);
    assert!(c.patch().apply(base).unwrap().to_bytes().unwrap() == after);
    assert!(c.patch().inverse().apply(&r).unwrap().to_bytes().unwrap() == before);
    assert!(base.to_bytes().unwrap() == before);
}
fn exact_statuses(
    c: &litchi_ops::transaction::Commit<litchi_ops::refresh::Summary>,
    expected: &[(&str, Status)],
) {
    assert_eq!(
        c.value()
            .fields()
            .iter()
            .map(|f| (f.kind(), f.status()))
            .collect::<Vec<_>>(),
        expected
    );
}

#[test]
fn native_caption_caches_refresh_in_doc_docx_and_templates_without_layout() {
    for ext in ["doc", "dot", "docx", "dotx"] {
        let relative = format!("word/word-layout-extensions.{ext}");
        let original = source_bytes(&relative);
        let stale = if matches!(ext, "doc" | "dot") {
            let p = litchi_doc::Package::from_bytes(&original).unwrap();
            let mut e = p.edit().unwrap();
            e.replace_text(
                litchi_doc::body_text::TextTarget::FieldResult(litchi_core::Position::new(0)),
                "9",
            )
            .unwrap();
            e.commit().unwrap().package().to_bytes().unwrap()
        } else {
            xml_derivative(&original, "word/document.xml", |xml| {
                assert_eq!(xml.matches("<w:t>1</w:t>").count(), 2);
                xml.replace("<w:t>1</w:t>", "<w:t>9</w:t>")
            })
        };
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let baseline = native(&relative, mode);
            let noop = baseline.refresh(&Options::new()).unwrap();
            exact_statuses(
                &noop,
                &[("SEQ", Status::Unchanged), ("SEQ", Status::Unchanged)],
            );
            assert!(noop.patch().is_noop());
            let (_input, f) = open_derivative(&stale, ext, mode);
            let c = f.refresh(&Options::new()).unwrap();
            assert_eq!(
                c.value().refreshed(),
                if matches!(ext, "doc" | "dot") { 1 } else { 2 }
            );
            let after = c.file().to_bytes().unwrap();
            if matches!(ext, "doc" | "dot") {
                let a = streams(&stale);
                let b = streams(&after);
                assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
                for (name, data) in a {
                    if name != ["WordDocument"] && name != ["1Table"] {
                        assert!(b[&name] == data, "changed {name:?}");
                    }
                }
            } else {
                preserve_except(&stale, &after, &["word/document.xml"]);
                assert!(
                    members(&after)["word/document.xml"] == members(&original)["word/document.xml"]
                );
            }
            check_commit(&f, &c, mode, ext, |r| {
                let current = r.refresh(&Options::new()).unwrap();
                exact_statuses(
                    &current,
                    &[("SEQ", Status::Unchanged), ("SEQ", Status::Unchanged)],
                );
                assert!(current.patch().is_noop());
            });
        }
        assert!(source_bytes(&relative) == original);
    }
}

fn without_toc(xml: &[u8]) -> Vec<u8> {
    let s = std::str::from_utf8(xml).unwrap();
    assert_eq!(s.matches("<w:sdt>").count(), 1);
    let start = s.find("<w:sdt>").unwrap();
    let end = s.find("</w:sdt>").unwrap() + 8;
    [&xml[..start], &xml[end..]].concat()
}
#[test]
fn native_toc_refresh_reuses_bookmarks_and_saved_pages_and_refuses_unknown_pages() {
    let relative = "word/word-math-navigation.docx";
    let original = source_bytes(relative);
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let base = native(relative, mode);
        let noop = base.refresh(&Options::new()).unwrap();
        exact_statuses(
            &noop,
            &[("TOC", Status::Unchanged), ("PAGEREF", Status::NeedsLayout)],
        );
        assert!(noop.patch().is_noop());
        let f = base
            .set(
                "/body/p[6]",
                props! {"find"=>"Equations","replace"=>"Reviewed equations"},
            )
            .unwrap();
        let before = f.to_bytes().unwrap();
        let c = f
            .refresh(&Options::new().page_numbers(PageNumbers::Require))
            .unwrap();
        exact_statuses(
            &c,
            &[("TOC", Status::Refreshed), ("PAGEREF", Status::NeedsLayout)],
        );
        assert_eq!(
            (c.value().bookmarks(), c.value().missing_page_numbers()),
            (0, 0)
        );
        let after = c.file().to_bytes().unwrap();
        preserve_except(&before, &after, &["word/document.xml"]);
        assert_eq!(
            without_toc(&members(&before)["word/document.xml"]),
            without_toc(&members(&after)["word/document.xml"])
        );
        let xml = String::from_utf8(members(&after)["word/document.xml"].clone()).unwrap();
        assert_eq!(xml.matches("Reviewed equations").count(), 2);
        assert!(xml.contains("_Toc242967666") && xml.contains("_Toc242967667"));
        check_commit(&f, &c, mode, "docx", |r| {
            assert!(r.refresh(&Options::new()).unwrap().patch().is_noop())
        });
        let extended = f
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "NEW_HEADING_059")
                    .prop("style", "Heading1"),
            )
            .unwrap()
            .file;
        let bytes = extended.to_bytes().unwrap();
        let e = extended
            .refresh(&Options::new().page_numbers(PageNumbers::Require))
            .unwrap_err();
        assert_eq!(e.code(), Code::RefreshFailed);
        assert!(extended.to_bytes().unwrap() == bytes);
        let partial = extended.refresh(&Options::new()).unwrap();
        assert_eq!(
            (
                partial.value().bookmarks(),
                partial.value().missing_page_numbers()
            ),
            (1, 1)
        );
        assert!(
            String::from_utf8(
                members(&partial.file().to_bytes().unwrap())["word/document.xml"].clone()
            )
            .unwrap()
            .contains("w:dirty=\"true\"")
        );
        assert!(
            partial
                .patch()
                .inverse()
                .apply(partial.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
    }
    assert!(source_bytes(relative) == original);
}

#[test]
fn native_workbook_cache_refresh_writes_only_stale_results_including_after_chart_tabs() {
    use litchi_ops::calc::{Origin, Value};
    for ext in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(&format!("excel/excel-values-formulas-chart.{ext}"), mode)
                .set("Sheet1!B2", props! {"value"=>4})
                .unwrap();
            let before = f.to_bytes().unwrap();
            let pending = f.refresh(&Options::new()).unwrap();
            let writes = if ext == "xls" { 6 } else { 5 };
            assert_eq!(pending.value().count(Status::Stale), writes);
            assert!(pending.patch().is_noop());
            let c = f.refresh(&Options::new().write_cache(true)).unwrap();
            assert_eq!(c.value().refreshed(), writes);
            assert_eq!(
                c.value().count(Status::Unsupported),
                if ext == "xls" { 1 } else { 2 }
            );
            let changed = c
                .value()
                .fields()
                .iter()
                .filter(|r| r.status() == Status::Refreshed)
                .map(|r| r.location().unwrap())
                .collect::<Vec<_>>();
            let expected = if ext == "xls" {
                vec![
                    "Sheet1!D2",
                    "Sheet1!B6",
                    "Sheet1!D6",
                    "Sheet1!B23",
                    "Sheet2!B3",
                    "Sheet2!B4",
                ]
            } else {
                vec![
                    "Sheet1!D2",
                    "Sheet1!B6",
                    "Sheet1!D6",
                    "Sheet1!B23",
                    "Sheet2!B3",
                ]
            };
            assert_eq!(changed, expected);
            let after = c.file().to_bytes().unwrap();
            if ext == "xls" {
                let a = streams(&before);
                let b = streams(&after);
                assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
                for (name, bytes) in &a {
                    if name.as_slice() != ["Workbook"] {
                        assert!(*bytes == b[name]);
                    }
                }
                let a = &a[&vec!["Workbook".into()]];
                let b = &b[&vec!["Workbook".into()]];
                assert_eq!(a.len(), b.len());
                let mut at = 0;
                let mut updated = 0;
                while at < a.len() {
                    let kind = u16::from_le_bytes(a[at..at + 2].try_into().unwrap());
                    let len =
                        usize::from(u16::from_le_bytes(a[at + 2..at + 4].try_into().unwrap()));
                    let end = at + 4 + len;
                    assert_eq!(&a[at..at + 4], &b[at..at + 4]);
                    if a[at..end] != b[at..end] {
                        assert_eq!(kind, 6);
                        assert_eq!(&a[at + 4..at + 10], &b[at + 4..at + 10]);
                        assert_eq!(&a[at + 18..end], &b[at + 18..end]);
                        updated += 1;
                    }
                    at = end;
                }
                assert_eq!(updated, writes);
            } else {
                preserve_except(
                    &before,
                    &after,
                    &[
                        "xl/worksheets/sheet1.xml",
                        "xl/worksheets/sheet2.xml",
                        "xl/worksheets/sheet1.bin",
                        "xl/worksheets/sheet2.bin",
                    ],
                );
            }
            check_commit(&f, &c, mode, ext, |r| {
                for (path, expected) in [
                    ("Sheet1!D2", 5.0),
                    ("Sheet1!B6", 15.0),
                    ("Sheet1!D6", 51.0),
                    ("Sheet1!B23", 3.75),
                    ("Sheet2!B3", 51.0),
                    ("Sheet2!B4", if ext == "xls" { 15.0 } else { 14.0 }),
                ] {
                    let v = r.evaluate(path, &litchi_ops::calc::Options::new()).unwrap();
                    assert_eq!(
                        (v.value(), v.origin()),
                        (&Value::Number(expected), Origin::Cached)
                    );
                }
                assert!(
                    r.refresh(&Options::new().write_cache(true))
                        .unwrap()
                        .patch()
                        .is_noop()
                );
            });
            let limits = litchi_ops::transaction::SemanticLimits::builder()
                .max_expanded_targets(1)
                .unwrap()
                .build()
                .unwrap();
            let e = f
                .refresh_with_options(
                    &Options::new().write_cache(true),
                    litchi_ops::transaction::EditOptions::new().semantic_limits(limits),
                )
                .unwrap_err();
            assert_eq!(e.code(), Code::LimitExceeded);
            assert!(matches!(
                e.leaf(),
                litchi_ops::error::Source::Transaction(litchi_ops::transaction::Error::Limit {
                    kind: litchi_ops::transaction::LimitKind::ExpandedTargets,
                    observed: 2,
                    maximum: 1
                })
            ));
            assert!(f.to_bytes().unwrap() == before);
        }
    }
}

#[test]
fn native_refresh_keeps_external_and_mail_merge_inputs_inert_and_refuses_missing_adapters() {
    for ext in ["doc", "docx"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(&format!("word/word-mailmerge-fields.{ext}"), mode);
            let before = f.to_bytes().unwrap();
            let c = f.refresh(&Options::new()).unwrap();
            assert_eq!(c.value().count(Status::Unsupported), 7);
            assert!(c.patch().is_noop());
            assert!(c.file().to_bytes().unwrap() == before);
        }
    }
    for ext in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(&format!("excel/excel-links-objects.{ext}"), mode);
            let before = f.to_bytes().unwrap();
            let c = f.refresh(&Options::new().write_cache(true)).unwrap();
            assert_eq!(c.value().count(Status::Unsupported), 3);
            assert!(c.patch().is_noop());
            assert!(c.file().to_bytes().unwrap() == before);
        }
    }
    for ext in ["xls", "xlsm", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(&format!("excel/excel-legacy-sheets.{ext}"), mode);
            let before = f.to_bytes().unwrap();
            assert_eq!(f.query("sheet[name=Macro_016]").unwrap().count(), 1);
            let c = f.refresh(&Options::new().write_cache(true)).unwrap();
            assert!(c.value().fields().is_empty());
            assert!(c.patch().is_noop());
            assert!(c.file().to_bytes().unwrap() == before);
            let error = f
                .evaluate(
                    "Macro_016!A1",
                    &litchi_ops::calc::Options::new().recalculate(true),
                )
                .unwrap_err();
            assert_eq!(
                error.code(),
                if ext == "xls" {
                    Code::UnsupportedCapability
                } else {
                    Code::UnsupportedType
                }
            );
        }
    }
    for relative in [
        "word/word-mailmerge-fields.rtf",
        "word/word-layout-extensions.rtf",
        "powerpoint/powerpoint-master-motion.ppt",
    ] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native(relative, mode);
            let before = f.to_bytes().unwrap();
            let e = f.refresh(&Options::new()).unwrap_err();
            assert_eq!(e.code(), Code::UnsupportedCapability);
            assert_eq!(e.message(), "this format has no field refresh adapter");
            assert!(f.to_bytes().unwrap() == before);
        }
    }
}

#[test]
fn native_slide_number_and_date_derivatives_refresh_with_an_explicit_clock() {
    let relative = "powerpoint/powerpoint-master-motion.pptx";
    let original = source_bytes(relative);
    let dated = xml_derivative(&original, "ppt/slides/slide4.xml", |xml| {
        assert_eq!(xml.matches("type=\"slidenum\"").count(), 1);
        xml.replace("type=\"slidenum\"", "type=\"datetime1\"")
    });
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let base = native(relative, mode);
        assert!(base.refresh(&Options::new()).unwrap().patch().is_noop());
        let mce = base.swap("/slide[1]", "/slide[2]").unwrap();
        let mce_source = mce.to_bytes().unwrap();
        let error = mce.refresh(&Options::new()).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(
            matches!(error.leaf(), litchi_ops::error::Source::Pptx(litchi_pptx::Error::Refused(r))
            if r.kind() == litchi_core::edit::RefusalKind::UnsupportedContent && r.reason().contains("markup-compatibility branch selection"))
        );
        assert!(mce.to_bytes().unwrap() == mce_source);
        let swapped = base.swap("/slide[4]", "/slide[5]").unwrap();
        let before = swapped.to_bytes().unwrap();
        let c = swapped.refresh(&Options::new()).unwrap();
        assert_eq!(c.value().refreshed(), 2);
        preserve_except(
            &before,
            &c.file().to_bytes().unwrap(),
            &["ppt/slides/slide4.xml", "ppt/slides/slide5.xml"],
        );
        check_commit(&swapped, &c, mode, "pptx", |r| {
            assert!(r.refresh(&Options::new()).unwrap().patch().is_noop())
        });
        let (_input, f) = open_derivative(&dated, "pptx", mode);
        let c = f.refresh(&Options::new()).unwrap();
        assert_eq!(c.value().count(Status::NeedsClock), 1);
        assert!(c.patch().is_noop());
        let options = Options::new().clock(Clock::parse("2026-11-07T09:30:00").unwrap());
        let c = f.refresh(&options).unwrap();
        assert_eq!(c.value().refreshed(), 1);
        let after = c.file().to_bytes().unwrap();
        preserve_except(&dated, &after, &["ppt/slides/slide4.xml"]);
        assert!(
            String::from_utf8(members(&after)["ppt/slides/slide4.xml"].clone())
                .unwrap()
                .contains("11/7/2026")
        );
        check_commit(&f, &c, mode, "pptx", |r| {
            assert!(r.refresh(&options).unwrap().patch().is_noop())
        });
    }
    assert!(source_bytes(relative) == original);
}
