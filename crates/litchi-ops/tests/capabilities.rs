//! Reports are generated from the same tables as production reads.
use litchi_core::FileFormat;
#[cfg(any(feature = "docx", feature = "ppt", feature = "pptx"))]
use litchi_ops::Code;
use litchi_ops::capabilities::Capabilities;
#[cfg(any(
    feature = "docx",
    feature = "ppt",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "xls"
))]
use litchi_ops::capabilities::Verb;

#[test]
fn capability_goldens_follow_enabled_formats() {
    for (format, name, enabled) in [
        (FileFormat::Doc, "doc", cfg!(feature = "doc")),
        (FileFormat::Xls, "xls", cfg!(feature = "xls")),
        (FileFormat::Docx, "docx", cfg!(feature = "docx")),
        (FileFormat::Xlsx, "xlsx", cfg!(feature = "xlsx")),
        (FileFormat::Xlsb, "xlsb", cfg!(feature = "xlsb")),
        (FileFormat::Ppt, "ppt", cfg!(feature = "ppt")),
        (FileFormat::Pptx, "pptx", cfg!(feature = "pptx")),
        (FileFormat::Rtf, "rtf", cfg!(feature = "rtf")),
        (FileFormat::Ods, "ods", cfg!(feature = "ods")),
        (FileFormat::Odt, "odt", cfg!(feature = "odt")),
        (FileFormat::Odp, "odp", cfg!(feature = "odp")),
        (FileFormat::Oth, "oth", cfg!(feature = "oth")),
        (FileFormat::Odm, "odm", cfg!(feature = "odm")),
        (FileFormat::Odg, "odg", cfg!(feature = "odg")),
        (FileFormat::Odc, "odc", cfg!(feature = "odc")),
        (FileFormat::Odi, "odi", cfg!(feature = "odi")),
        (FileFormat::Odb, "odb", cfg!(feature = "odb")),
        (FileFormat::Odf, "odf", cfg!(feature = "odf-formula")),
        (FileFormat::Pages, "pages", cfg!(feature = "pages")),
        (FileFormat::Keynote, "keynote", cfg!(feature = "keynote")),
        (FileFormat::Numbers, "numbers", cfg!(feature = "numbers")),
    ] {
        if !enabled {
            continue;
        }
        let actual = format!(
            "{}\n",
            serde_json::to_string_pretty(&Capabilities::of(format)).unwrap()
        );
        let golden_name = match (name, cfg!(feature = "encryption")) {
            ("docx", true) => "docx-encryption",
            ("xlsx", true) => "xlsx-encryption",
            _ => name,
        };
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/golden/capabilities")
            .join(format!("{golden_name}.json"));
        if std::env::var_os("LITCHI_BLESS").is_some() {
            std::fs::create_dir_all(path.parent().unwrap()).unwrap();
            std::fs::write(&path, &actual).unwrap();
        }
        assert_eq!(
            std::fs::read_to_string(path).unwrap(),
            actual,
            "{golden_name}"
        );
    }
}

#[cfg(feature = "xls")]
#[test]
fn xls_capabilities_expose_write_surface() {
    use litchi_ops::capabilities::Subject;

    let caps = Capabilities::of(FileFormat::Xls);
    for verb in [
        Verb::Set,
        Verb::Add,
        Verb::Remove,
        Verb::Move,
        Verb::Swap,
        Verb::Edit,
        Verb::Batch,
    ] {
        assert!(caps.verb(verb).is_supported(), "XLS {verb:?}");
    }

    let cell = caps.element("cell").expect("XLS cell capability");
    assert!(cell.operations().set);
    assert!(cell.operations().remove);
    for key in ["value", "clear"] {
        let property = cell.property(key).unwrap_or_else(|| panic!("{key}"));
        assert!(!property.get(), "XLS {key} is write-only");
        assert!(property.set(), "XLS {key} is settable");
        assert!(!property.add(), "XLS {key} is not an add payload");
    }
    let sheet = caps.element("sheet").expect("XLS sheet capability");
    assert!(sheet.operations().set);
    assert!(sheet.operations().add);
    assert!(sheet.operations().remove);
    assert!(sheet.property("name").unwrap().set());
    assert!(sheet.property("name").unwrap().add());
    assert!(!sheet.property("hidden").unwrap().set());
    assert!(!sheet.property("visibility").unwrap().set());

    for (verb, element, context) in [
        ("set", "cell", "non-scalar-properties"),
        ("set", "cell", "existing-formula-cells"),
        ("set", "cell", "native-dependency-closure"),
        ("set", "sheet", "tab-state"),
        ("remove", "sheet", "native-dependency-closure"),
        ("move", "sheet", "native-dependency-closure"),
        ("add", "row", "native-dependency-closure"),
    ] {
        assert!(
            caps.gaps().any(|gap| {
                gap.subject()
                    == Subject::Operation {
                        verb,
                        element,
                        context,
                    }
            }),
            "missing XLS operation gap {verb} {element} {context}"
        );
    }
}

#[cfg(feature = "docx")]
#[test]
fn ordinary_reports_hide_native_selection_fields() {
    let caps = Capabilities::of(FileFormat::Docx);
    assert!(caps.verb(Verb::Set).is_supported());
    for verb in [Verb::Add, Verb::Remove, Verb::Move, Verb::Swap] {
        assert!(caps.verb(verb).is_supported(), "{verb:?}");
    }
    assert!(caps.verb(Verb::Save).is_supported());
    assert!(caps.verb(Verb::Create).is_supported());
    for (element, native) in [("paragraph", "paraId"), ("sdt", "sdtId")] {
        let element = caps.element(element).unwrap();
        assert!(element.operations().get);
        assert!(element.property(native).is_none());
        assert!(element.paths().all(|(path, _)| !path.contains(native)));
    }
    let serialized = serde_json::to_string(&caps).unwrap();
    assert!(!serialized.contains("paraId"));
    assert!(!serialized.contains("textId"));
    assert!(!serialized.contains("sdtId"));
    assert!(!serialized.contains("[@id="));
    assert!(!serialized.contains("<id>"));
    for element in caps.elements() {
        let writable_element = matches!(element.name(), "paragraph" | "run");
        let operations = element.operations();
        // Record 0251a: cells and hyperlinks take `text`; tables and rows
        // are added and removed; cells, hyperlinks and content controls are
        // removed.
        let (set, add, remove) = match element.name() {
            "paragraph" | "run" | "hyperlink" => (true, true, true),
            "cell" => (true, false, true),
            // Record 0251c: comments are added and removed.
            "table" | "row" | "comment" => (false, true, true),
            // Record 0251d: pictures are added; record 0295: SmartArt too.
            "picture" | "smartart" => (false, true, false),
            // Record 0283: charts are added, set and removed.
            "chart" => (true, true, true),
            "sdt" => (false, false, true),
            _ => (false, false, false),
        };
        assert_eq!(operations.set, set, "{} set", element.name());
        assert_eq!(operations.add, add, "{} add", element.name());
        assert_eq!(operations.remove, remove, "{} remove", element.name());
        for property in element.properties() {
            if writable_element {
                let paragraph_initial_run = element.name() == "paragraph"
                    && matches!(
                        property.name(),
                        "font.latin"
                            | "font.ascii"
                            | "font.hAnsi"
                            | "font.ea"
                            | "font.asciiTheme"
                            | "font.hAnsiTheme"
                            | "font.eaTheme"
                            | "font.csTheme"
                            | "font.hint"
                            | "lang.latin"
                            | "lang.ea"
                            | "lang.cs"
                            | "size"
                            | "bold"
                            | "italic"
                            | "font.cs"
                            | "size.cs"
                            | "bold.cs"
                            | "italic.cs"
                            | "color"
                            | "underline"
                            | "underline.color"
                            | "strike"
                            | "highlight"
                            | "caps"
                            | "smallcaps"
                            | "dstrike"
                            | "vanish"
                            | "outline"
                            | "shadow"
                            | "emboss"
                            | "imprint"
                            | "noproof"
                            | "superscript"
                            | "subscript"
                            | "position"
                            | "charSpacing"
                            | "rStyle"
                    );
                if paragraph_initial_run {
                    assert!(!property.get(), "{}", property.name());
                    assert!(!property.set(), "{}", property.name());
                    assert!(property.add(), "{}", property.name());
                } else if property.name() == "text" {
                    assert!(!property.get());
                    assert!(property.set());
                    assert!(property.add());
                } else if element.name() == "paragraph"
                    && matches!(property.name(), "listStyle" | "start")
                {
                    // Record 0251e: list kinds are set through the
                    // numbering part and not read back yet.
                    assert!(!property.get(), "{}", property.name());
                    assert!(property.set(), "{}", property.name());
                    assert!(!property.add(), "{}", property.name());
                } else if property.name() != "styleId" {
                    assert!(property.get(), "{}", property.name());
                    assert!(property.set(), "{}", property.name());
                    assert!(property.add(), "{}", property.name());
                } else {
                    assert!(property.get(), "{}", property.name());
                    assert!(!property.set(), "{}", property.name());
                    assert!(!property.add(), "{}", property.name());
                }
            } else if element.name() == "hyperlink"
                && matches!(
                    property.name(),
                    "url"
                        | "anchor"
                        | "font"
                        | "size"
                        | "bold"
                        | "italic"
                        | "color"
                        | "underline"
                        | "rStyle"
                )
            {
                // Record 0251b: the target and the run formatting seed a new
                // hyperlink; only `url` changes an existing one.
                assert!(property.get(), "hyperlink {}", property.name());
                assert!(property.add(), "hyperlink {}", property.name());
                assert_eq!(property.set(), property.name() == "url");
            } else if element.name() == "picture" && property.name() != "relId" {
                // Record 0251d: the extent, name and alternative text seed a
                // new picture; `src` is add-only.
                assert_eq!(
                    property.get(),
                    property.name() != "src",
                    "picture {}",
                    property.name()
                );
                assert!(!property.set(), "picture {}", property.name());
                assert!(property.add(), "picture {}", property.name());
            } else if element.name() == "comment" && property.name() != "id" {
                // Record 0251c: author, initials, date and text seed a new
                // comment.
                assert!(property.get(), "comment {}", property.name());
                assert!(!property.set(), "comment {}", property.name());
                assert!(property.add(), "comment {}", property.name());
            } else if element.name() == "smartart" {
                // Record 0295: layout, text, name, alternative text and
                // extent seed a new graphic; the projection reads layout,
                // its identifier, the node count and the outline.
                assert!(!property.set(), "smartart {}", property.name());
                assert_eq!(
                    property.add(),
                    !matches!(property.name(), "layoutId" | "nodeCount"),
                    "smartart {}",
                    property.name()
                );
                assert_eq!(
                    property.get(),
                    matches!(
                        property.name(),
                        "layout" | "layoutId" | "nodeCount" | "text"
                    ),
                    "smartart {}",
                    property.name()
                );
            } else if element.name() == "chart" {
                // Record 0283: chart keys seed a new chart and most are
                // also settable; the registry declares each.
                assert!(
                    property.set() || property.add() || property.get(),
                    "chart {}",
                    property.name()
                );
            } else if property.name() == "text" {
                // `text` of a hyperlink or cell is write-only.
                assert!(!property.get(), "{} text", element.name());
                assert!(property.set(), "{} text", element.name());
            } else if matches!(element.name(), "table" | "row")
                && matches!(
                    property.name(),
                    "cols" | "rows" | "style" | "colWidths" | "data"
                )
            {
                // Seeds of a new table or row.
                assert!(!property.set(), "{} {}", element.name(), property.name());
                assert!(property.add(), "{} {}", element.name(), property.name());
            } else {
                assert!(property.get());
                assert!(!property.set());
                assert!(!property.add());
            }
        }
    }
}

#[cfg(feature = "docx")]
#[test]
fn docx_write_context_gaps_are_explicit() {
    use litchi_ops::capabilities::Subject;

    let caps = Capabilities::of(FileFormat::Docx);
    let operation_gaps: Vec<_> = caps
        .gaps()
        .filter(|gap| matches!(gap.subject(), Subject::Operation { .. }))
        .collect();
    assert_eq!(
        operation_gaps.len(),
        27 + if cfg!(feature = "encryption") { 8 } else { 0 }
    );

    let subjects: Vec<_> = operation_gaps.iter().map(|gap| gap.subject()).collect();
    for expected in [
        Subject::Operation {
            verb: "set",
            element: "paragraph",
            context: "wrapped-or-unproven-selection",
        },
        Subject::Operation {
            verb: "set",
            element: "run",
            context: "wrapped-or-unproven-selection",
        },
        Subject::Operation {
            verb: "set",
            element: "paragraph/run",
            context: "mixed-paragraph-run-fanout",
        },
        Subject::Operation {
            verb: "set",
            element: "paragraph/run",
            context: "managed-transaction",
        },
        Subject::Operation {
            verb: "set",
            element: "paragraph/run",
            context: "complex-leaf-content",
        },
        Subject::Operation {
            verb: "add",
            element: "paragraph/run",
            context: "managed-transaction",
        },
        Subject::Operation {
            verb: "add",
            element: "paragraph",
            context: "wrapped-or-unproven-parent",
        },
        Subject::Operation {
            verb: "add",
            element: "run",
            context: "wrapped-or-unproven-parent",
        },
        Subject::Operation {
            verb: "remove",
            element: "paragraph/run",
            context: "managed-transaction",
        },
        Subject::Operation {
            verb: "remove",
            element: "paragraph/run",
            context: "wrapped-or-unproven-selection",
        },
        Subject::Operation {
            verb: "remove",
            element: "paragraph/run",
            context: "complex-leaf-content",
        },
        Subject::Operation {
            verb: "move",
            element: "paragraph/run",
            context: "managed-transaction",
        },
        Subject::Operation {
            verb: "move",
            element: "paragraph/run",
            context: "wrapped-or-unproven-selection",
        },
        Subject::Operation {
            verb: "move",
            element: "paragraph/run",
            context: "complex-leaf-content",
        },
        Subject::Operation {
            verb: "swap",
            element: "paragraph/run",
            context: "managed-transaction",
        },
        Subject::Operation {
            verb: "swap",
            element: "paragraph/run",
            context: "nested-or-unproven-selection",
        },
        Subject::Operation {
            verb: "swap",
            element: "paragraph/run",
            context: "complex-leaf-content",
        },
        Subject::Operation {
            verb: "create",
            element: "document",
            context: "strict-conformance",
        },
        Subject::Operation {
            verb: "create",
            element: "document",
            context: "noncore-initial-metadata",
        },
    ] {
        assert!(subjects.contains(&expected), "missing {expected:?}");
    }
    #[cfg(feature = "encryption")]
    assert!(subjects.contains(&Subject::Operation {
        verb: "save",
        element: "document",
        context: "encrypted-without-reencrypt",
    }));

    let json = serde_json::to_value(caps).unwrap();
    let json_gaps = json["gaps"].as_array().unwrap();
    assert!(json_gaps.iter().any(|gap| {
        gap["subject"] == "operation"
            && gap["verb"] == "set"
            && gap["element"] == "paragraph/run"
            && gap["context"] == "mixed-paragraph-run-fanout"
    }));
    for (verb, context) in [
        ("create", "strict-conformance"),
        ("create", "noncore-initial-metadata"),
    ] {
        assert!(
            json_gaps.iter().any(|gap| {
                gap["subject"] == "operation"
                    && gap["verb"] == verb
                    && gap["element"] == "document"
                    && gap["context"] == context
                    && gap["code"] == "unsupported_capability"
            }),
            "missing serialized create gap {context}"
        );
    }
    #[cfg(feature = "encryption")]
    {
        let encrypted_gap = operation_gaps
            .iter()
            .find(|gap| {
                gap.subject()
                    == Subject::Operation {
                        verb: "save",
                        element: "document",
                        context: "encrypted-without-reencrypt",
                    }
            })
            .expect("missing encrypted publication gap");
        assert_eq!(encrypted_gap.code(), Code::DocumentProtected);
        assert!(json_gaps.iter().any(|gap| {
            gap["subject"] == "operation"
                && gap["verb"] == "save"
                && gap["element"] == "document"
                && gap["context"] == "encrypted-without-reencrypt"
                && gap["code"] == "document_protected"
        }));
    }
    assert!(caps.element("paragraph").unwrap().operations().set);
    assert!(caps.element("run").unwrap().operations().set);
    assert!(caps.element("paragraph").unwrap().operations().add);
    assert!(caps.element("run").unwrap().operations().add);
    assert!(caps.element("paragraph").unwrap().operations().remove);
    assert!(caps.element("run").unwrap().operations().remove);
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_reports_keep_selection_semantic_and_declare_cell_and_sheet_writes() {
    let caps = Capabilities::of(FileFormat::Xlsx);
    assert!(caps.verb(Verb::Get).is_supported());
    assert!(caps.element("cell").is_some());
    assert!(caps.element("table").is_some());
    for element in caps.elements() {
        assert!(element.paths().all(|(path, _)| {
            !path.contains("<id>") && !path.contains("[@id=") && !path.contains("[@rId=")
        }));
        // Chart writes have their own report test (`tests/charts.rs`), and
        // worksheet features theirs (`tests/xlsx_features.rs`, record 0320).
        if matches!(
            element.name(),
            "chart"
                | "series"
                | "axis"
                | "comment"
                | "hyperlink"
                | "picture"
                | "shape"
                | "conditionalFormatting"
                | "dataValidation"
                | "autofilter"
                | "rowbreak"
                | "colbreak"
        ) {
            continue;
        }
        for property in element.properties() {
            // Cell hyperlinks and the sheet auto-filter are write-only
            // setters (record 0320).
            let link_write = (element.name() == "cell"
                && matches!(property.name(), "link" | "tooltip" | "display"))
                || (element.name() == "sheet" && property.name() == "autoFilter");
            let cell_write = matches!(element.name(), "cell" | "range")
                && matches!(
                    property.name(),
                    "value"
                        | "formula"
                        | "clear"
                        | "font.bold"
                        | "font.italic"
                        | "strike"
                        | "underline"
                        | "font.name"
                        | "font.size"
                        | "font.color"
                        | "font.outline"
                        | "font.shadow"
                        | "font.condense"
                        | "font.extend"
                        | "font.charset"
                        | "font.family"
                        | "font.scheme"
                        | "alignment.horizontal"
                        | "alignment.vertical"
                        | "alignment.wrapText"
                        | "alignment.shrinkToFit"
                        | "alignment.textRotation"
                        | "alignment.indent"
                        | "alignment.readingOrder"
                        | "alignment.relativeIndent"
                        | "alignment.justifyLastLine"
                        | "protection.locked"
                        | "protection.hidden"
                        | "quotePrefix"
                        | "numberformat"
                        | "fill"
                        | "fillPattern"
                        | "fillBg"
                        | "border.all"
                        | "border"
                        | "border.all.color"
                        | "border.color"
                        | "border.left"
                        | "border.left.color"
                        | "border.right"
                        | "border.right.color"
                        | "border.top"
                        | "border.top.color"
                        | "border.bottom"
                        | "border.bottom.color"
                        | "border.diagonal"
                        | "border.diagonal.color"
                        | "border.diagonalUp"
                        | "border.diagonalDown"
                        | "superscript"
                        | "subscript"
                        | "merge"
                );
            let write_only = (cell_write
                && ((element.name() == "range" && property.name() != "merge")
                    || matches!(
                        property.name(),
                        "value"
                            | "clear"
                            | "border"
                            | "border.all"
                            | "border.color"
                            | "border.all.color"
                    )))
                || (matches!(element.name(), "row" | "column") && property.name() == "outline")
                || (element.name() == "sheet" && property.name() == "merge")
                || link_write;
            assert_eq!(property.get(), !write_only, "{}", property.name());
            let sheet_write =
                element.name() == "sheet" && matches!(property.name(), "name" | "hidden");
            let dimension_write = matches!(element.name(), "row" | "column")
                && matches!(
                    property.name(),
                    "height" | "width" | "hidden" | "outline" | "outlineLevel" | "collapsed"
                );
            assert_eq!(
                property.set(),
                cell_write
                    || sheet_write
                    || dimension_write
                    || link_write
                    || (element.name() == "sheet" && property.name() == "merge"),
                "{}",
                property.name()
            );
            assert_eq!(property.add(), sheet_write, "{}", property.name());
        }
    }
}

#[cfg(feature = "pptx")]
/// Rich text setters of record 0252, and the shape fill, geometry and slide
/// transition of record 0269.
fn pptx_text_format_set(element: &str, property: &str) -> bool {
    match element {
        "shape" | "placeholder" => matches!(
            property,
            "bold"
                | "italic"
                | "size"
                | "color"
                | "font"
                | "font.latin"
                | "align"
                | "fill"
                | "geometry"
        ),
        "paragraph" => matches!(
            property,
            "align" | "size" | "text" | "bold" | "italic" | "color" | "font"
        ),
        "run" => matches!(
            property,
            "font.latin" | "size" | "bold" | "italic" | "color" | "text" | "font"
        ),
        "tc" => matches!(
            property,
            "fill" | "text" | "bold" | "italic" | "size" | "color" | "font" | "align"
        ),
        "notes" => property == "text",
        "slide" => matches!(property, "notes" | "transition"),
        "comment" => matches!(property, "text" | "x" | "y"),
        _ => false,
    }
}

#[cfg(feature = "pptx")]
/// Text-box formatting a shape takes on `add` (record 0269).
fn pptx_shape_add(property: &str) -> bool {
    matches!(property, "text" | "x" | "y" | "width" | "height")
        || pptx_text_format_set("shape", property)
}

#[cfg(feature = "pptx")]
/// Write-only rich text entries (no ordinary read of the same key).
fn pptx_write_only(element: &str, property: &str) -> bool {
    match element {
        "paragraph" => matches!(property, "text" | "bold" | "italic" | "color" | "font"),
        "run" => matches!(property, "text" | "font"),
        "tc" => property != "fill",
        "slide" => matches!(property, "notes" | "transition"),
        "shape" | "placeholder" => matches!(property, "fill" | "geometry"),
        "picture" => property == "src",
        _ => false,
    }
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_reports_selection_and_structural_writes() {
    let caps = Capabilities::of(FileFormat::Pptx);
    assert!(caps.verb(Verb::Get).is_supported());
    assert!(caps.verb(Verb::Add).is_supported());
    for verb in [Verb::Move, Verb::Swap] {
        assert!(caps.verb(verb).is_supported(), "{verb:?}");
    }
    assert!(caps.verb(Verb::Remove).is_supported());
    assert!(caps.element("slide").is_some());
    assert!(caps.element("shape").is_some());
    for element in caps.elements() {
        assert!(element.paths().all(|(path, _)| {
            !path.contains("<id>") && !path.contains("[@id=") && !path.contains("[@rId=")
        }));
        // Chart (record 0270) and SmartArt (record 0284) writes have their own
        // report tests (`tests/charts.rs`, `tests/smartart.rs`).
        if matches!(
            element.name(),
            "chart" | "series" | "axis" | "smartart" | "node"
        ) {
            continue;
        }
        for property in element.properties() {
            let text_write =
                matches!(element.name(), "shape" | "placeholder") && property.name() == "text";
            assert_eq!(
                property.get(),
                !text_write && !pptx_write_only(element.name(), property.name()),
                "{}",
                property.name()
            );
            let shape_geometry_write = matches!(element.name(), "shape" | "placeholder")
                && matches!(property.name(), "x" | "y" | "width" | "height");
            assert_eq!(
                property.set(),
                text_write
                    || shape_geometry_write
                    || pptx_text_format_set(element.name(), property.name())
                    || (element.name() == "slide" && matches!(property.name(), "name" | "hidden")),
                "{}",
                property.name()
            );
            let slide_add = element.name() == "slide"
                && matches!(property.name(), "layout" | "name" | "hidden");
            let shape_add = element.name() == "shape" && pptx_shape_add(property.name());
            let picture_add = element.name() == "picture"
                && matches!(
                    property.name(),
                    "src" | "name" | "x" | "y" | "width" | "height"
                );
            let comment_add = element.name() == "comment"
                && matches!(
                    property.name(),
                    "text" | "author" | "initials" | "date" | "x" | "y"
                );
            assert_eq!(
                property.add(),
                slide_add || shape_add || picture_add || comment_add,
                "{} / {}",
                element.name(),
                property.name()
            );
        }
    }
    for kind in [
        "picture",
        "table",
        "chart",
        "group",
        "connector",
        "smartart",
        "ole",
    ] {
        let operations = caps.element(kind).unwrap().operations();
        assert!(operations.remove, "{kind}");
        // Charts (record 0270) and SmartArt (record 0284) are settable.
        assert_eq!(
            operations.set,
            matches!(kind, "chart" | "smartart"),
            "{kind}"
        );
    }
    let slide = caps.element("slide").unwrap();
    assert!(slide.operations().set);
    assert!(slide.operations().add);
    assert!(slide.operations().remove);
    assert!(slide.property("layout").unwrap().add());
    for property in slide.properties() {
        assert_eq!(
            property.add(),
            matches!(property.name(), "layout" | "name" | "hidden"),
            "{}",
            property.name()
        );
    }
    let shape = caps.element("shape").unwrap();
    assert!(shape.operations().add);
    assert!(shape.operations().remove);
    for property in shape.properties() {
        let add = pptx_shape_add(property.name());
        assert_eq!(property.add(), add, "{}", property.name());
    }
    let placeholder = caps.element("placeholder").unwrap();
    assert!(!placeholder.operations().add);
    assert!(placeholder.operations().remove);
    for property in placeholder.properties() {
        assert!(!property.add(), "{}", property.name());
        assert_eq!(
            property.set(),
            matches!(property.name(), "text" | "x" | "y" | "width" | "height")
                || pptx_text_format_set("placeholder", property.name()),
            "{}",
            property.name()
        );
    }
    assert!(
        caps.gaps()
            .find(|gap| {
                gap.subject()
                    == litchi_ops::capabilities::Subject::Operation {
                        verb: "add",
                        element: "slide",
                        context: "source-owner",
                    }
            })
            .is_none()
    );
    assert!(
        caps.gaps()
            .find(|gap| {
                gap.subject()
                    == litchi_ops::capabilities::Subject::Operation {
                        verb: "add",
                        element: "shape",
                        context: "source-owner",
                    }
            })
            .is_none()
    );
    let placement_gap = caps
        .gaps()
        .find(|gap| {
            gap.subject()
                == litchi_ops::capabilities::Subject::Operation {
                    verb: "add",
                    element: "shape",
                    context: "non-append-placement",
                }
        })
        .expect("non-append shape placement remains an explicit capability gap");
    assert_eq!(placement_gap.code(), Code::UnsupportedCapability);
    for (verb, context) in [("move", "positional-owner"), ("swap", "positional-owner")] {
        let gap = caps
            .gaps()
            .find(|gap| {
                gap.subject()
                    == litchi_ops::capabilities::Subject::Operation {
                        verb,
                        element: "slide",
                        context,
                    }
            })
            .unwrap_or_else(|| panic!("source-backed slide {verb} remains an explicit gap"));
        assert_eq!(gap.code(), Code::UnsupportedCapability);
        assert_eq!(gap.follow_up(), "M2 PPTX positional-owner slide ordering");
        assert_eq!(
            gap.reason(),
            "source-backed PPTX slide ordering maintains p14 sections but still refuses outline slide lists, numeric presentation ranges, and unknown or opaque metadata extensions"
        );
    }
    assert!(
        caps.gaps()
            .find(|gap| {
                gap.subject()
                    == litchi_ops::capabilities::Subject::Operation {
                        verb: "remove",
                        element: "slide",
                        context: "source-owner",
                    }
            })
            .is_none()
    );
    for (context, follow_up, reason) in [
        (
            "dependency-topology",
            "M2 PPTX dependency-aware slide removal",
            "native PPTX slide removal removes the notes, comments, pictures, media and other Parts the slide owns exclusively and keeps shared ones; a slide another slide jumps to, or whose relationships name a missing Part, is still refused",
        ),
        (
            "final-slide",
            "M2 PPTX final-slide removal policy",
            "native PPTX slide removal currently preserves the final slide and cannot remove the last slide from a presentation",
        ),
        (
            "canonical-target",
            "M2 PPTX equivalent relationship-target spellings",
            "source-backed PPTX slide removal requires canonical owner-relative relationship targets so inverse replay can restore the original OPC spelling",
        ),
    ] {
        let gap = caps
            .gaps()
            .find(|gap| {
                gap.subject()
                    == litchi_ops::capabilities::Subject::Operation {
                        verb: "remove",
                        element: "slide",
                        context,
                    }
            })
            .unwrap_or_else(|| panic!("slide removal gap {context} is missing"));
        assert_eq!(
            gap.code(),
            if context == "final-slide" {
                Code::InvalidArgument
            } else {
                Code::UnsupportedCapability
            }
        );
        assert_eq!(gap.follow_up(), follow_up);
        assert_eq!(gap.reason(), reason);
    }
    for (context, follow_up, reason) in [
        (
            "dependency-topology",
            "M2 PPTX dependency-aware shape removal",
            "shape removal refuses retained connector or animation references to removed IDs; source-backed removal preserves unused relationships and parts, while eager removal prunes unreachable relationship closure",
        ),
        (
            "non-top-level",
            "M2 PPTX grouped/master/layout shape removal",
            "native PPTX shape removal currently requires a top-level slide shape; grouped children and master or layout shapes remain outside the admitted topology",
        ),
    ] {
        let gap = caps
            .gaps()
            .find(|gap| {
                gap.subject()
                    == litchi_ops::capabilities::Subject::Operation {
                        verb: "remove",
                        element: "shape",
                        context,
                    }
            })
            .unwrap_or_else(|| panic!("shape removal gap {context} is missing"));
        assert_eq!(gap.code(), Code::UnsupportedCapability);
        assert_eq!(gap.follow_up(), follow_up);
        assert_eq!(gap.reason(), reason);
    }
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_reports_unified_reads_and_keeps_writes_granular() {
    let format = Capabilities::of(FileFormat::Ppt);
    for verb in [
        Verb::Get,
        Verb::Find,
        Verb::Query,
        Verb::View,
        Verb::Set,
        Verb::Edit,
        Verb::Batch,
    ] {
        assert!(format.verb(verb).is_supported(), "PPT {verb:?}");
    }
    assert!(format.verb(Verb::Save).is_supported());
    assert!(format.verb(Verb::Create).is_supported());
    for verb in [Verb::Add, Verb::Remove, Verb::Move, Verb::Swap] {
        assert!(format.verb(verb).is_supported(), "PPT {verb:?}");
    }
    for name in ["notes", "paragraph"] {
        let element = format.element(name).expect("text element");
        assert!(
            element.operations().get && element.operations().set,
            "PPT {name}"
        );
        assert!(
            element.property("text").expect("text").set(),
            "PPT {name}.text"
        );
    }
    for name in ["shape", "placeholder"] {
        let element = format.element(name).expect("shape element");
        for key in ["text", "x", "y", "width", "height"] {
            assert!(
                element.property(key).expect("setter property").set(),
                "PPT {name}.{key}"
            );
        }
        assert!(
            !element
                .property("rotation")
                .expect("read-only property")
                .set()
        );
    }
    assert!(
        format
            .element("slide")
            .expect("slide")
            .property("hidden")
            .expect("hidden")
            .set()
    );
    for name in ["presentation", "slide", "shape"] {
        let element = format
            .element(name)
            .unwrap_or_else(|| panic!("missing PPT {name}"));
        let operations = element.operations();
        assert!(operations.get, "PPT {name} is readable");
        assert_eq!(operations.set, name != "presentation", "PPT {name} setter");
        assert_eq!(operations.add, name == "slide", "PPT {name} insertion");
        assert_eq!(operations.remove, name == "slide", "PPT {name} removal");
    }

    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ole/ppt/SampleShow.ppt"
    ))
    .unwrap();
    let released = litchi_ops::open_bytes_with(
        bytes,
        &litchi_ops::OpenOptions::default().retain_source(false),
    )
    .unwrap();
    assert_eq!(released.capabilities().retains_source(), Some(false));
    for verb in [Verb::Get, Verb::Find, Verb::Query, Verb::View] {
        assert!(
            released.capabilities().verb(verb).is_supported(),
            "PPT {verb:?}"
        );
    }
    assert_eq!(
        released.capabilities().verb(Verb::Save).code(),
        Some(Code::UnsupportedCapability)
    );
    for verb in [
        Verb::Set,
        Verb::Add,
        Verb::Remove,
        Verb::Move,
        Verb::Swap,
        Verb::Edit,
        Verb::Batch,
    ] {
        assert_eq!(
            released.capabilities().verb(verb).code(),
            Some(Code::UnsupportedCapability)
        );
    }
    assert!(
        !released
            .capabilities()
            .element("slide")
            .unwrap()
            .operations()
            .remove
    );
    for name in ["slide", "shape", "placeholder"] {
        assert!(
            !released
                .capabilities()
                .element(name)
                .expect("element")
                .operations()
                .set
        );
        assert!(
            released
                .capabilities()
                .element(name)
                .expect("element")
                .properties()
                .all(|property| !property.set())
        );
    }
    assert!(released.capabilities().verb(Verb::Create).is_supported());
    assert!(released.root().is_ok());
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_structural_capabilities_follow_owner_and_retention_policy() {
    use litchi_ops::{OpenOptions, SourceMode};
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/pptx/sample.pptx"
    ))
    .unwrap();

    let source = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Auto),
    )
    .unwrap();
    let source_caps = source.capabilities();
    assert!(source_caps.verb(Verb::Add).is_supported());
    for verb in [Verb::Move, Verb::Swap] {
        assert!(source_caps.verb(verb).is_supported(), "{verb:?}");
    }
    assert!(source_caps.verb(Verb::Remove).is_supported());
    assert!(source_caps.element("slide").unwrap().operations().add);
    assert!(source_caps.element("slide").unwrap().operations().remove);
    assert!(
        source_caps
            .element("slide")
            .unwrap()
            .property("layout")
            .unwrap()
            .add()
    );
    assert!(source_caps.element("shape").unwrap().operations().add);
    assert!(source_caps.element("shape").unwrap().operations().remove);
    for property in source_caps.element("shape").unwrap().properties() {
        let add = pptx_shape_add(property.name());
        assert_eq!(property.add(), add, "{}", property.name());
        let set = matches!(property.name(), "text" | "x" | "y" | "width" | "height")
            || pptx_text_format_set("shape", property.name());
        assert_eq!(property.set(), set, "{}", property.name());
    }
    assert!(!source_caps.element("placeholder").unwrap().operations().add);
    assert!(
        source_caps
            .element("placeholder")
            .unwrap()
            .operations()
            .remove
    );
    for property in source_caps.element("placeholder").unwrap().properties() {
        assert!(!property.add(), "{}", property.name());
        assert_eq!(
            property.set(),
            matches!(property.name(), "text" | "x" | "y" | "width" | "height")
                || pptx_text_format_set("placeholder", property.name()),
            "{}",
            property.name()
        );
    }
    for element in source_caps.elements() {
        assert_eq!(
            element.operations().remove,
            matches!(
                element.name(),
                "slide"
                    | "shape"
                    | "placeholder"
                    | "picture"
                    | "table"
                    | "chart"
                    | "group"
                    | "connector"
                    | "smartart"
                    | "node"
                    | "ole"
            ),
            "{}",
            element.name()
        );
    }

    let eager = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .unwrap();
    let eager_caps = eager.capabilities();
    assert!(eager_caps.verb(Verb::Add).is_supported());
    assert!(eager_caps.verb(Verb::Remove).is_supported());
    assert!(eager_caps.verb(Verb::Move).is_supported());
    assert!(eager_caps.verb(Verb::Swap).is_supported());
    assert!(eager_caps.element("slide").unwrap().operations().add);
    assert!(eager_caps.element("slide").unwrap().operations().remove);
    for element in eager_caps.elements() {
        assert_eq!(
            element.operations().remove,
            matches!(
                element.name(),
                "slide"
                    | "shape"
                    | "placeholder"
                    | "picture"
                    | "table"
                    | "chart"
                    | "group"
                    | "connector"
                    | "smartart"
                    | "node"
                    | "ole"
            ),
            "{}",
            element.name()
        );
    }
    assert!(
        eager_caps
            .element("slide")
            .unwrap()
            .property("layout")
            .unwrap()
            .add()
    );
    assert!(eager_caps.element("shape").unwrap().operations().add);
    for property in eager_caps.element("shape").unwrap().properties() {
        let add = pptx_shape_add(property.name());
        assert_eq!(property.add(), add, "{}", property.name());
        let set = matches!(property.name(), "text" | "x" | "y" | "width" | "height")
            || pptx_text_format_set("shape", property.name());
        assert_eq!(property.set(), set, "{}", property.name());
    }
    assert!(eager_caps.element("shape").unwrap().operations().remove);
    assert!(!eager_caps.element("placeholder").unwrap().operations().add);
    assert!(
        eager_caps
            .element("placeholder")
            .unwrap()
            .operations()
            .remove
    );
    for property in eager_caps.element("placeholder").unwrap().properties() {
        assert!(!property.add(), "{}", property.name());
        assert_eq!(
            property.set(),
            matches!(property.name(), "text" | "x" | "y" | "width" | "height")
                || pptx_text_format_set("placeholder", property.name()),
            "{}",
            property.name()
        );
    }

    let released = litchi_ops::open_bytes_with(
        bytes,
        &OpenOptions::new()
            .retain_source(false)
            .source_mode(SourceMode::Eager),
    )
    .unwrap();
    assert_eq!(
        released.capabilities().verb(Verb::Add).code(),
        Some(Code::UnsupportedCapability)
    );
    assert_eq!(
        released.capabilities().verb(Verb::Remove).code(),
        Some(Code::UnsupportedCapability)
    );
    for verb in [Verb::Move, Verb::Swap] {
        assert_eq!(
            released.capabilities().verb(verb).code(),
            Some(Code::UnsupportedCapability),
            "{verb:?}"
        );
    }
    assert!(
        !released
            .capabilities()
            .element("slide")
            .unwrap()
            .operations()
            .add
    );
    assert!(
        !released
            .capabilities()
            .element("slide")
            .unwrap()
            .operations()
            .remove
    );
    assert!(
        !released
            .capabilities()
            .element("shape")
            .unwrap()
            .operations()
            .add
    );
    assert!(
        !released
            .capabilities()
            .element("shape")
            .unwrap()
            .operations()
            .remove
    );
    for property in released
        .capabilities()
        .element("shape")
        .unwrap()
        .properties()
    {
        assert!(!property.add(), "{}", property.name());
    }
    assert!(
        !released
            .capabilities()
            .element("placeholder")
            .unwrap()
            .operations()
            .add
    );
    assert!(
        !released
            .capabilities()
            .element("placeholder")
            .unwrap()
            .operations()
            .remove
    );
    for property in released
        .capabilities()
        .element("placeholder")
        .unwrap()
        .properties()
    {
        assert!(!property.add(), "{}", property.name());
    }
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_reports_keep_selection_semantic_and_declare_sheet_and_cell_writes() {
    let caps = Capabilities::of(FileFormat::Xlsb);
    assert!(caps.element("sheet").unwrap().operations().add);
    assert!(caps.element("sheet").unwrap().operations().remove);
    assert!(!caps.element("workbook").unwrap().operations().add);
    assert!(caps.verb(Verb::Get).is_supported());
    assert!(caps.verb(Verb::Set).is_supported());
    assert!(caps.verb(Verb::Add).is_supported());
    assert!(caps.verb(Verb::Remove).is_supported());
    assert!(caps.verb(Verb::Save).is_supported());
    assert!(caps.element("cell").is_some());
    let computed_value = caps
        .gaps()
        .find(|gap| {
            matches!(
                gap.subject(),
                litchi_ops::capabilities::Subject::Property {
                    element: "cell",
                    key: "computedValue"
                }
            )
        })
        .expect("computed values in cell views remain an explicit XLSB capability gap");
    assert_eq!(
        computed_value.follow_up(),
        "M4 follow-up: project computed values in cell views"
    );
    // Record 0266: import and evaluation are worksheet verbs.
    assert!(caps.verb(Verb::Import).is_supported());
    assert!(caps.verb(Verb::Evaluate).is_supported());
    for context in [
        "order-opaque-dependencies",
        "removed-scoped-names",
        "textual-references",
    ] {
        assert!(
            caps.gaps().any(|gap| {
                matches!(
                    gap.subject(),
                    litchi_ops::capabilities::Subject::Operation {
                        verb: "remove",
                        element: "sheet",
                        context: actual,
                    } if actual == context
                )
            }),
            "missing XLSB sheet-removal gap: {context}"
        );
    }
    let named_range = caps.element("namedrange").unwrap();
    assert_eq!(
        named_range
            .properties()
            .map(|property| property.name())
            .collect::<Vec<_>>(),
        vec!["name", "ref", "scope"]
    );
    let cell = caps.element("cell").unwrap();
    for unsupported in ["protection.locked", "numFmtId", "font.scheme"] {
        assert!(
            cell.property(unsupported).is_none(),
            "XLSB must not advertise {unsupported}"
        );
    }
    // Record 0261: formatting reads and writes.
    for supported in ["font.bold", "fill", "border.left", "numberformat"] {
        let property = cell.property(supported).expect("advertised");
        assert!(property.get() && property.set(), "{supported}");
    }
    for element in caps.elements() {
        assert!(element.paths().all(|(path, _)| {
            !path.contains("<id>") && !path.contains("[@id=") && !path.contains("[@rId=")
        }));
        // Record 0261: the cell setters (also accepted by ranges) and `merge`.
        const CELL_WRITES: &[&str] = &[
            "value",
            "formula",
            "clear",
            "font.bold",
            "font.italic",
            "strike",
            "underline",
            "font.name",
            "font.size",
            "font.color",
            "fill",
            "border.all",
            "border.all.color",
            "border.left",
            "border.left.color",
            "border.right",
            "border.right.color",
            "border.top",
            "border.top.color",
            "border.bottom",
            "border.bottom.color",
            "numberformat",
            "alignment.horizontal",
            "alignment.vertical",
            "alignment.wrapText",
            "alignment.textRotation",
            "alignment.indent",
            "alignment.shrinkToFit",
            "alignment.readingOrder",
        ];
        // Record 0322: the worksheet features use the XLSX feature tables.
        if [
            "comment",
            "hyperlink",
            "picture",
            "shape",
            "conditionalFormatting",
            "dataValidation",
            "autofilter",
            "rowbreak",
            "colbreak",
        ]
        .contains(&element.name())
        {
            let operations = element.operations();
            assert!(
                operations.get && operations.set && operations.add && operations.remove,
                "{}",
                element.name()
            );
            continue;
        }
        for property in element.properties() {
            let name = property.name();
            let payload = matches!(element.name(), "cell" | "range")
                && matches!(name, "value" | "clear" | "border.all" | "border.all.color");
            let range_write_only = element.name() == "range" && name != "merge";
            assert_eq!(
                property.get(),
                !(payload || range_write_only),
                "{}.{name}",
                element.name()
            );
            assert_eq!(
                property.set(),
                (element.name() == "sheet" && matches!(name, "name" | "hidden" | "visibility"))
                    || (element.name() == "namedrange" && name == "ref")
                    || (element.name() == "cell" && CELL_WRITES.contains(&name))
                    || (element.name() == "range"
                        && (CELL_WRITES.contains(&name) || name == "merge")),
                "{}.{}",
                element.name(),
                property.name()
            );
            let sheet_add = (element.name() == "sheet"
                && matches!(property.name(), "name" | "hidden" | "visibility"))
                || (element.name() == "namedrange"
                    && matches!(property.name(), "name" | "ref" | "scope"));
            assert_eq!(
                property.add(),
                sheet_add,
                "{}.{}",
                element.name(),
                property.name()
            );
        }
    }
}

#[cfg(all(feature = "docx", feature = "xlsx"))]
#[test]
fn available_and_pending_reads_have_distinct_reports() {
    assert!(
        Capabilities::of(FileFormat::Docx)
            .verb(Verb::Get)
            .is_supported()
    );
    // RTF is a read adapter when its feature is enabled. Focused builds keep
    // the whole-format refusal as the authoritative answer.
    let rtf = Capabilities::of(FileFormat::Rtf);
    if cfg!(feature = "rtf") {
        assert!(rtf.verb(Verb::Get).is_supported());
        assert!(rtf.element("document").unwrap().operations().get);
    } else {
        assert!(!rtf.verb(Verb::Get).is_supported());
    }
    #[cfg(feature = "pptx")]
    {
        let file = litchi_ops::open(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/pptx/smartart-cycle-matrix.pptx"
        ))
        .unwrap();
        assert_eq!(file.root().unwrap().kind(), "presentation");
    }
}

#[cfg(feature = "docx")]
#[test]
fn file_state_retention_and_staleness_do_not_decode_capabilities() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/docx/drawing.docx"
    ))
    .unwrap();
    let file = litchi_ops::open_bytes_with(
        bytes.clone(),
        &litchi_ops::OpenOptions::default().retain_source(false),
    )
    .unwrap();
    let unretained_caps = file.capabilities();
    assert_eq!(unretained_caps.retains_source(), Some(false));
    assert_eq!(
        unretained_caps.verb(Verb::Set).code(),
        Some(Code::UnsupportedCapability)
    );
    for verb in [Verb::Add, Verb::Remove, Verb::Move, Verb::Swap] {
        assert_eq!(
            unretained_caps.verb(verb).code(),
            Some(Code::UnsupportedCapability),
            "{verb:?}"
        );
    }
    assert!(unretained_caps.verb(Verb::Save).is_supported());
    assert!(unretained_caps.verb(Verb::Create).is_supported());
    assert!(
        !unretained_caps
            .element("paragraph")
            .unwrap()
            .operations()
            .set
    );
    assert!(
        !unretained_caps
            .element("paragraph")
            .unwrap()
            .operations()
            .add
    );
    assert!(
        !unretained_caps
            .element("paragraph")
            .unwrap()
            .operations()
            .remove
    );
    assert!(
        unretained_caps
            .element("paragraph")
            .unwrap()
            .properties()
            .all(|property| !property.set() && !property.add())
    );
    let retained = litchi_ops::open_bytes(bytes.clone()).unwrap();
    assert!(retained.capabilities().verb(Verb::Set).is_supported());
    let created = litchi_ops::create(FileFormat::Docx).expect("create DOCX");
    assert!(created.capabilities().verb(Verb::Save).is_supported());
    assert!(created.capabilities().verb(Verb::Create).is_supported());
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("revision.docx");
    std::fs::write(&path, bytes).unwrap();
    let file = litchi_ops::open(&path).unwrap();
    std::fs::write(&path, b"source changed").unwrap();
    assert!(file.capabilities().verb(Verb::Get).is_supported());
    assert_eq!(file.root().unwrap_err().code(), Code::IoError);
    let signed = litchi_ops::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    ))
    .unwrap();
    let caps = signed.capabilities();
    assert_eq!(caps.state(), litchi_ops::capabilities::State::Signed);
    assert!(caps.verb(Verb::Get).is_supported());
    assert_eq!(caps.verb(Verb::Set).code(), Some(Code::DocumentProtected));
    assert!(caps.verb(Verb::Save).is_supported());
    assert_eq!(
        caps.verb(Verb::Create).code(),
        Some(Code::DocumentProtected)
    );
}
