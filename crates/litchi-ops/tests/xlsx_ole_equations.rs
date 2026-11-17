//! XLSX OLE equation objects (MathType and Equation Editor 3.0) through the
//! unified API, beside the Office Math equations of `xlsx_equations`: add,
//! get, query, set and remove on both owners, with read-back through the
//! leaf reader and MTEF, untouched objects and parts kept byte for byte, a
//! sheet holding both kinds, and typed refusals that leave the input
//! intact.
#![cfg(feature = "xlsx")]

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode};

/// An ONLYOFFICE-written workbook whose first sheet holds one Office Math
/// equation (see `test-data/poi/PROVENANCE.md`).
#[cfg(feature = "equation")]
const ONLYOFFICE: &[u8] = include_bytes!(
    "../../../test-data/poi/test-data/spreadsheet/LIBRE_OFFICE-128382-0-small-media.xlsx"
);

const EXCEL: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/excel/excel-links-objects.xlsx");

/// Both owners of `bytes`: source-backed and eager.
fn owners(bytes: &[u8]) -> [File; 2] {
    [SourceMode::Auto, SourceMode::Eager].map(|mode| {
        litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode)).unwrap()
    })
}

#[cfg(not(feature = "equation"))]
#[test]
fn equation_objects_need_the_equation_feature() {
    for file in owners(EXCEL) {
        assert_eq!(file.query("equation").unwrap().count(), 0);
        let error = file
            .add(
                "/Sheet1",
                NewNode::new("equation")
                    .prop("progId", "Equation.DSMT4")
                    .prop("latex", "x"),
            )
            .unwrap_err();
        assert_ne!(error.code(), Code::InternalError);
    }
}

#[cfg(feature = "equation")]
mod with_equation {
    use super::*;
    use litchi_core::FileFormat;
    use litchi_ops::props;
    use std::collections::BTreeMap;
    use std::io::Read as _;

    fn blank() -> Vec<u8> {
        litchi_ops::create(FileFormat::Xlsx)
            .unwrap()
            .to_bytes()
            .unwrap()
    }

    /// An Equation Editor 3.0 `Equation Native` stream holding the digit
    /// 2, base64-encoded.
    const EQUATION3: &str = "HAAAAAIAbcENAAAAAAAAAAAAAAAAAAAAAAAAAAMBAQMACgECiDIAAAA=";

    fn prop(file: &File, path: &str, key: &str) -> Option<String> {
        file.get(path)
            .unwrap()
            .view()
            .format()
            .unwrap()
            .find(|prop| prop.key() == key)
            .map(|prop| prop.value().to_string())
    }

    fn compact(text: Option<String>) -> Option<String> {
        text.map(|text| text.replace([' ', '{', '}'], ""))
    }

    /// Every ZIP entry of a saved workbook.
    fn entries(file: &File) -> BTreeMap<String, Vec<u8>> {
        let bytes = file.to_bytes().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        (0..archive.len())
            .map(|index| {
                let mut entry = archive.by_index(index).unwrap();
                let mut data = Vec::new();
                entry.read_to_end(&mut data).unwrap();
                (entry.name().unwrap().into_owned(), data)
            })
            .collect()
    }

    /// The leaf's equation objects of the first worksheet of `file`.
    fn leaf(file: &File) -> Vec<litchi_xlsx::equation_objects::EquationObject> {
        litchi_xlsx::Package::from_bytes(&file.to_bytes().unwrap())
            .unwrap()
            .equation_objects(0)
            .unwrap()
    }

    #[test]
    fn capabilities_declare_the_equation_object_surface() {
        let capabilities = litchi_ops::capabilities::Capabilities::of(FileFormat::Xlsx);
        let element = capabilities.element("equation").unwrap();
        let access = |name: &str| {
            let property = element.property(name).unwrap();
            (property.get(), property.set(), property.add())
        };
        assert_eq!(access("latex"), (true, true, true));
        assert_eq!(access("omml"), (true, true, true));
        assert_eq!(access("mtef"), (true, true, true));
        assert_eq!(access("progId"), (true, false, true));
        assert_eq!(access("anchor"), (true, false, true));
        assert_eq!(access("error"), (true, false, false));
        assert_eq!(access("name"), (true, false, false));
        assert_eq!(access("size"), (false, false, true));
        let json = serde_json::to_value(capabilities).unwrap();
        let equation = json["elements"]
            .as_array()
            .unwrap()
            .iter()
            .find(|element| element["element"] == "equation")
            .unwrap();
        for verb in ["get", "query", "set", "add", "remove"] {
            assert_eq!(equation["operations"][verb], true, "{verb}");
        }
    }

    #[test]
    fn office_math_equations_come_first_and_each_kind_keeps_its_verbs() {
        for file in owners(ONLYOFFICE) {
            let sheet = "/Basic features";
            let omml = prop(&file, &format!("{sheet}/equation[1]"), "omml");
            assert!(omml.is_some());
            let added = file
                .add(
                    sheet,
                    NewNode::new("equation")
                        .prop("mtef", EQUATION3)
                        .prop("anchor", "H3"),
                )
                .unwrap();
            assert_eq!(added.path.to_string(), "/sheet[1]/equation[2]");
            let file = added.file;
            for read in owners(&file.to_bytes().unwrap()) {
                let paths = read
                    .query("equation")
                    .unwrap()
                    .map(|node| node.unwrap().view().path().to_string())
                    .collect::<Vec<_>>();
                assert_eq!(paths, ["/sheet[1]/equation[1]", "/sheet[1]/equation[2]"]);
                let first = format!("{sheet}/equation[1]");
                let second = format!("{sheet}/equation[2]");
                assert_eq!(prop(&read, &first, "omml"), omml);
                assert_eq!(prop(&read, &first, "progId"), None);
                assert_eq!(prop(&read, &first, "mtef"), None);
                assert_eq!(
                    prop(&read, &second, "progId").as_deref(),
                    Some("Equation.3")
                );
                assert_eq!(prop(&read, &second, "omml"), None);
                assert_eq!(prop(&read, &second, "anchor").as_deref(), Some("H3"));
            }

            // `mtef` is refused on the Office Math equation; the object
            // takes `latex` and keeps its kind.
            let error = file
                .set(
                    &format!("{sheet}/equation[1]"),
                    props! { "mtef" => EQUATION3 },
                )
                .unwrap_err();
            assert_eq!(error.code(), Code::InvalidValue);
            let set = file
                .set(&format!("{sheet}/equation[2]"), props! { "latex" => "y" })
                .unwrap();
            assert_eq!(
                prop(&set, &format!("{sheet}/equation[2]"), "progId").as_deref(),
                Some("Equation.3")
            );
            assert_eq!(prop(&set, &format!("{sheet}/equation[1]"), "omml"), omml);

            // A new Office Math equation numbers before the object.
            let more = set
                .add(
                    sheet,
                    NewNode::new("equation")
                        .prop("latex", "z")
                        .prop("anchor", "A40"),
                )
                .unwrap()
                .file;
            assert_eq!(more.query("equation").unwrap().count(), 3);
            assert_eq!(
                prop(&more, &format!("{sheet}/equation[3]"), "progId").as_deref(),
                Some("Equation.3")
            );

            // Removing the object leaves both Office Math equations.
            let removed = more.remove(&format!("{sheet}/equation[3]")).unwrap();
            assert_eq!(removed.query("equation").unwrap().count(), 2);
            assert_eq!(
                prop(&removed, &format!("{sheet}/equation[1]"), "omml"),
                omml
            );
            assert!(removed.query("equation").unwrap().all(|node| {
                prop(&removed, &node.unwrap().view().path().to_string(), "progId").is_none()
            }));
        }
    }

    #[test]
    fn contains_matches_an_equation_object_by_its_latex() {
        let added = litchi_ops::open_bytes(blank())
            .unwrap()
            .add(
                "/Sheet1",
                NewNode::new("equation")
                    .prop("progId", "Equation.DSMT4")
                    .prop("latex", r"\frac{a}{b}"),
            )
            .unwrap()
            .file
            .add("/Sheet1", NewNode::new("equation").prop("mtef", EQUATION3))
            .unwrap()
            .file;
        for file in owners(&added.to_bytes().unwrap()) {
            let paths = |query: &str| {
                file.query(query)
                    .unwrap()
                    .map(|node| node.unwrap().view().path().to_string())
                    .collect::<Vec<_>>()
            };
            assert_eq!(paths("equation:contains(frac)"), ["/sheet[1]/equation[1]"]);
            assert_eq!(paths("equation:contains(2)"), ["/sheet[1]/equation[2]"]);
            assert!(paths("equation:contains(sqrt)").is_empty());
        }
    }

    #[test]
    fn excel_embedded_objects_are_not_equations() {
        for file in owners(EXCEL) {
            assert_eq!(file.query("equation").unwrap().count(), 0);
            assert!(file.get("/Sheet1/equation[1]").is_err());
        }
    }

    #[test]
    fn equation_objects_are_added_read_set_and_removed() {
        for file in owners(&blank()) {
            let before = file.to_bytes().unwrap();
            let added = file
                .add(
                    "/Sheet1",
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("latex", r"\frac{a}{b}")
                        .prop("anchor", "C3"),
                )
                .unwrap();
            assert_eq!(added.path.to_string(), "/sheet[1]/equation[1]");
            assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
            for read in owners(&added.file.to_bytes().unwrap()) {
                let path = "/Sheet1/equation[1]";
                assert_eq!(prop(&read, path, "latex").as_deref(), Some(r"\frac{a}{b}"));
                assert_eq!(
                    prop(&read, path, "progId").as_deref(),
                    Some("Equation.DSMT4")
                );
                assert_eq!(prop(&read, path, "anchor").as_deref(), Some("C3"));
                assert_eq!(prop(&read, path, "error"), None);
                assert!(prop(&read, path, "mtef").is_some());
                assert_eq!(read.query("equation").unwrap().count(), 1);
                assert_eq!(read.query("math").unwrap().count(), 1);
                assert_eq!(read.query("picture").unwrap().count(), 0);
                assert_eq!(read.query("shape").unwrap().count(), 0);
            }
            // The leaf reader and the MTEF reader agree.
            let objects = leaf(&added.file);
            assert_eq!(objects.len(), 1);
            let mtef = litchi_equation::read_mtef(
                objects[0].native().unwrap(),
                &litchi_equation::Limits::DEFAULT,
            )
            .unwrap();
            assert_eq!(mtef.latex(), r"\frac{a}{b}");
            assert_eq!(mtef.info().unwrap().version, 5);

            // A second object from MTEF keeps Equation Editor 3.0, under a
            // cell parent.
            let second = added
                .file
                .add("/Sheet1/B10", NewNode::new("math").prop("mtef", EQUATION3))
                .unwrap();
            assert_eq!(second.path.to_string(), "/sheet[1]/equation[2]");
            let file = second.file;
            assert_eq!(
                prop(&file, "/Sheet1/equation[2]", "progId").as_deref(),
                Some("Equation.3")
            );
            assert_eq!(
                prop(&file, "/Sheet1/equation[2]", "anchor").as_deref(),
                Some("B10")
            );
            assert_eq!(
                prop(&file, "/Sheet1/equation[2]", "mtef").as_deref(),
                Some(EQUATION3)
            );

            // Set keeps the kind and every byte of the other object.
            let parts = entries(&file);
            let set = file
                .set("/Sheet1/equation[2]", props! { "latex" => r"\sqrt{x}" })
                .unwrap();
            assert_eq!(
                prop(&set, "/Sheet1/equation[2]", "latex").as_deref(),
                Some(r"\sqrt{x}")
            );
            assert_eq!(
                prop(&set, "/Sheet1/equation[2]", "progId").as_deref(),
                Some("Equation.3")
            );
            let after = entries(&set);
            for name in ["xl/embeddings/oleObject1.bin", "xl/media/image1.png"] {
                assert_eq!(parts[name], after[name], "{name}");
            }
            assert_ne!(
                parts["xl/embeddings/oleObject2.bin"],
                after["xl/embeddings/oleObject2.bin"]
            );
            let objects = leaf(&set);
            let mtef = litchi_equation::read_mtef(
                objects[1].native().unwrap(),
                &litchi_equation::Limits::DEFAULT,
            )
            .unwrap();
            assert_eq!(mtef.info().unwrap().version, 3);

            let set = set
                .set(
                    "/Sheet1/equation[1]",
                    props! { "omml" => "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>y</m:t></m:r></m:oMath>" },
                )
                .unwrap();
            assert_eq!(
                compact(prop(&set, "/Sheet1/equation[1]", "latex")).as_deref(),
                Some("y")
            );
            let set = set
                .set("/Sheet1/equation[2]", props! { "mtef" => EQUATION3 })
                .unwrap();
            assert_eq!(
                prop(&set, "/Sheet1/equation[2]", "latex").as_deref(),
                Some("2")
            );

            // Remove one, then the other: the workbook is blank again.
            let removed = set.remove("/Sheet1/equation[1]").unwrap();
            assert_eq!(removed.query("equation").unwrap().count(), 1);
            assert_eq!(
                prop(&removed, "/Sheet1/equation[1]", "progId").as_deref(),
                Some("Equation.3")
            );
            let removed = removed.remove("/Sheet1/equation[1]").unwrap();
            assert_eq!(removed.query("equation").unwrap().count(), 0);
            assert_eq!(entries(&removed), entries(&owners(&before)[0]));
        }
    }

    #[test]
    fn equation_objects_join_and_leave_excel_embedded_objects() {
        for file in owners(EXCEL) {
            let before = entries(&file);
            let pictures = file.query("picture").unwrap().count();
            let added = file
                .add(
                    "/Sheet1",
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("latex", "E=mc^2")
                        .prop("anchor", "D30"),
                )
                .unwrap();
            for read in owners(&added.file.to_bytes().unwrap()) {
                assert_eq!(read.query("picture").unwrap().count(), pictures);
                assert_eq!(read.query("equation").unwrap().count(), 1);
                let latex = compact(prop(&read, "/Sheet1/equation[1]", "latex"));
                assert_eq!(latex.as_deref(), Some("E=mc^2"));
            }
            let after = entries(&added.file);
            for (name, bytes) in &before {
                if !name.starts_with("xl/worksheets/")
                    && !name.starts_with("xl/drawings/")
                    && name != "[Content_Types].xml"
                {
                    assert_eq!(&after[name], bytes, "{name}");
                }
            }
            let removed = added.file.remove("/Sheet1/equation[1]").unwrap();
            let after = entries(&removed);
            for (name, bytes) in &before {
                if name != "[Content_Types].xml" {
                    assert_eq!(&after[name], bytes, "{name}");
                }
            }
            assert_eq!(after.len(), before.len());
        }
    }

    #[test]
    fn refusals_leave_the_workbook_intact() {
        for file in owners(&blank()) {
            let before = file.to_bytes().unwrap();
            let refuse = |node: NewNode<'_>| file.add("/Sheet1", node).unwrap_err().code();
            assert_eq!(
                refuse(
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("latex", "x")
                        .prop("size", "12pt")
                ),
                Code::InvalidValue
            );
            assert_eq!(
                refuse(NewNode::new("equation").prop("progId", "Equation.DSMT4")),
                Code::MissingProperty
            );
            assert_eq!(
                refuse(
                    NewNode::new("equation")
                        .prop("progId", "Equation.3")
                        .prop("latex", "x")
                ),
                Code::InvalidValue
            );
            assert_eq!(
                refuse(
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("mtef", EQUATION3)
                ),
                Code::InvalidValue
            );
            assert_eq!(
                refuse(
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("latex", "x")
                        .prop("omml", "<m:oMath/>")
                ),
                Code::InvalidArgument
            );
            assert_eq!(
                refuse(
                    NewNode::new("equation")
                        .prop("progId", "Equation.DSMT4")
                        .prop("latex", r"\frac{")
                ),
                Code::InvalidValue
            );
            assert_eq!(
                refuse(NewNode::new("equation").prop("mtef", "not base64!")),
                Code::InvalidValue
            );
            assert_eq!(
                refuse(
                    NewNode::new("equation")
                        .prop("mtef", EQUATION3)
                        .prop("error", "x")
                ),
                Code::UnsupportedProperty
            );
            assert_eq!(file.query("equation").unwrap().count(), 0);
            assert_eq!(file.to_bytes().unwrap(), before);

            let added = file
                .add("/Sheet1", NewNode::new("equation").prop("mtef", EQUATION3))
                .unwrap()
                .file;
            let staged = added.to_bytes().unwrap();
            let error = added
                .set(
                    "/Sheet1/equation[1]",
                    props! { "progId" => "Equation.DSMT4" },
                )
                .unwrap_err();
            assert_eq!(error.code(), Code::UnsupportedProperty);
            // MathType MTEF 5 does not go into an Equation Editor 3.0 object.
            let native = litchi_equation::latex_to_equation_native(
                "z",
                &litchi_equation::MtefTarget::MathType,
                &litchi_equation::Limits::DEFAULT,
            )
            .unwrap();
            let encoded = {
                use base64::Engine as _;
                base64::engine::general_purpose::STANDARD.encode(native.bytes())
            };
            let error = added
                .set("/Sheet1/equation[1]", props! { "mtef" => encoded.as_str() })
                .unwrap_err();
            assert_eq!(error.code(), Code::InvalidValue);
            assert!(added.get("/Sheet1/equation[2]").is_err());
            assert_eq!(added.to_bytes().unwrap(), staged);
        }
    }
}
