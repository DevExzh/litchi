//! Reports are generated from the same tables as production reads.
use litchi_core::FileFormat;
#[cfg(any(feature = "docx", feature = "pptx"))]
use litchi_ops::Code;
use litchi_ops::capabilities::Capabilities;
#[cfg(all(feature = "docx", feature = "xlsx"))]
use litchi_ops::capabilities::Support;
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
use litchi_ops::capabilities::Verb;

#[test]
fn capability_goldens_follow_enabled_formats() {
    for (format, name, enabled) in [
        (FileFormat::Docx, "docx", cfg!(feature = "docx")),
        (FileFormat::Xlsx, "xlsx", cfg!(feature = "xlsx")),
        (FileFormat::Xlsb, "xlsb", cfg!(feature = "xlsb")),
        (FileFormat::Pptx, "pptx", cfg!(feature = "pptx")),
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
        if writable_element {
            assert!(element.operations().set);
            assert!(element.operations().add);
            assert!(element.operations().remove);
        } else {
            assert!(!element.operations().set);
            assert!(!element.operations().add);
            assert!(!element.operations().remove);
        }
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
                } else if property.name() != "styleId" {
                    assert!(property.get(), "{}", property.name());
                    assert!(property.set(), "{}", property.name());
                    assert!(property.add(), "{}", property.name());
                } else {
                    assert!(property.get(), "{}", property.name());
                    assert!(!property.set(), "{}", property.name());
                    assert!(!property.add(), "{}", property.name());
                }
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
        25 + if cfg!(feature = "encryption") { 8 } else { 0 }
    );

    let subjects: Vec<_> = operation_gaps.iter().map(|gap| gap.subject()).collect();
    for expected in [
        Subject::Operation {
            verb: "set",
            element: "paragraph",
            context: "nested-or-unproven-selection",
        },
        Subject::Operation {
            verb: "set",
            element: "run",
            context: "nested-or-unproven-selection",
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
            context: "nested-or-unproven-parent",
        },
        Subject::Operation {
            verb: "add",
            element: "run",
            context: "nested-or-unproven-parent",
        },
        Subject::Operation {
            verb: "remove",
            element: "paragraph/run",
            context: "managed-transaction",
        },
        Subject::Operation {
            verb: "remove",
            element: "paragraph/run",
            context: "nested-or-unproven-selection",
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
            context: "nested-or-unproven-selection",
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
        for property in element.properties() {
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
                || (element.name() == "sheet" && property.name() == "merge");
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
                    || (element.name() == "sheet" && property.name() == "merge"),
                "{}",
                property.name()
            );
            assert_eq!(property.add(), sheet_write, "{}", property.name());
        }
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
        for property in element.properties() {
            let text_write =
                matches!(element.name(), "shape" | "placeholder") && property.name() == "text";
            assert_eq!(property.get(), !text_write, "{}", property.name());
            let shape_geometry_write = matches!(element.name(), "shape" | "placeholder")
                && matches!(property.name(), "x" | "y" | "width" | "height");
            assert_eq!(
                property.set(),
                text_write
                    || shape_geometry_write
                    || (element.name() == "slide" && matches!(property.name(), "name" | "hidden")),
                "{}",
                property.name()
            );
            let slide_add = element.name() == "slide"
                && matches!(property.name(), "layout" | "name" | "hidden");
            let shape_add = element.name() == "shape"
                && matches!(property.name(), "text" | "x" | "y" | "width" | "height");
            assert_eq!(
                property.add(),
                slide_add || shape_add,
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
        assert!(!operations.set, "{kind}");
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
        let add = matches!(property.name(), "text" | "x" | "y" | "width" | "height");
        assert_eq!(property.add(), add, "{}", property.name());
    }
    let placeholder = caps.element("placeholder").unwrap();
    assert!(!placeholder.operations().add);
    assert!(placeholder.operations().remove);
    for property in placeholder.properties() {
        assert!(!property.add(), "{}", property.name());
        assert_eq!(
            property.set(),
            matches!(property.name(), "text" | "x" | "y" | "width" | "height"),
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
            "source-backed PPTX slide ordering still refuses sections, outline slide lists, numeric presentation ranges, and unknown or opaque metadata extensions"
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
            "native PPTX slide removal currently admits only a dependency-free slide with exactly one layout relationship; notes, comments, images, media, and external or internal relationship closure remain pending",
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
        let add = matches!(property.name(), "text" | "x" | "y" | "width" | "height");
        assert_eq!(property.add(), add, "{}", property.name());
        let set = matches!(property.name(), "text" | "x" | "y" | "width" | "height");
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
            matches!(property.name(), "text" | "x" | "y" | "width" | "height"),
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
        let add = matches!(property.name(), "text" | "x" | "y" | "width" | "height");
        assert_eq!(property.add(), add, "{}", property.name());
        let set = matches!(property.name(), "text" | "x" | "y" | "width" | "height");
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
            matches!(property.name(), "text" | "x" | "y" | "width" | "height"),
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
fn xlsb_reports_keep_selection_semantic_and_read_only() {
    let caps = Capabilities::of(FileFormat::Xlsb);
    assert!(caps.verb(Verb::Get).is_supported());
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
        .expect("formula evaluation remains an explicit XLSB capability gap");
    assert_eq!(computed_value.follow_up(), "M4 (formula evaluation)");
    let named_range = caps.element("namedrange").unwrap();
    assert_eq!(
        named_range
            .properties()
            .map(|property| property.name())
            .collect::<Vec<_>>(),
        vec!["name"]
    );
    let cell = caps.element("cell").unwrap();
    for unsupported in [
        "font.bold",
        "fill",
        "border.left",
        "protection.locked",
        "numFmtId",
        "numberformat",
    ] {
        assert!(
            cell.property(unsupported).is_none(),
            "XLSB must not advertise {unsupported}"
        );
    }
    for element in caps.elements() {
        assert!(element.paths().all(|(path, _)| {
            !path.contains("<id>") && !path.contains("[@id=") && !path.contains("[@rId=")
        }));
        for property in element.properties() {
            assert!(property.get(), "{}", property.name());
            assert!(!property.set(), "{}", property.name());
            assert!(!property.add(), "{}", property.name());
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
    // XLSB now has a unified read adapter. Keep this check on a legacy format
    // whose adapter remains pending; its feature may be disabled in a focused
    // build, so the report's declared refusal is authoritative.
    let pending = Capabilities::of(FileFormat::Xls);
    assert!(!pending.verb(Verb::Get).is_supported());
    #[cfg(feature = "pptx")]
    {
        let file = litchi_ops::open(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/pptx/smartart-cycle-matrix.pptx"
        ))
        .unwrap();
        assert_eq!(file.root().unwrap().kind(), "presentation");
    }
    assert!(matches!(
        pending.verb(Verb::Get),
        Support::Unsupported {
            code: Code::UnsupportedCapability | Code::UnsupportedFormat,
            ..
        }
    ));
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
