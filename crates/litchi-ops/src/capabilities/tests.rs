//! Unit tests of the capability report: its generation rules, and the
//! agreement of every adapter's spec table with the resolver's vocabulary.
//! The golden snapshots and the liveness tests over the fixture corpus are
//! integration tests (`tests/capabilities.rs`).

use litchi_core::FileFormat;

use super::{
    CREATE_PENDING, Capabilities, EDITS, IWORK_STRUCTURE, SAVE_PENDING, State, Support, Verb,
    path_segment, positional, refused,
};
use crate::Code;
use crate::view::Mode;

/// Every format the detector knows, for the per-format rules.
const FORMATS: [FileFormat; 21] = [
    FileFormat::Doc,
    FileFormat::Docx,
    FileFormat::Ppt,
    FileFormat::Pptx,
    FileFormat::Xls,
    FileFormat::Xlsx,
    FileFormat::Xlsb,
    FileFormat::Rtf,
    FileFormat::Pages,
    FileFormat::Keynote,
    FileFormat::Numbers,
    FileFormat::Odt,
    FileFormat::Ods,
    FileFormat::Odp,
    FileFormat::Odg,
    FileFormat::Odc,
    FileFormat::Odf,
    FileFormat::Odi,
    FileFormat::Odm,
    FileFormat::Oth,
    FileFormat::Odb,
];

#[test]
fn templates_are_classified_by_their_ordinal_steps() {
    assert!(positional("/<sheet>/row[N]"));
    assert!(positional("/<sheet>/chart[N]/series[K]"));
    assert!(!positional("/"));
    assert!(!positional("/<sheet>/<A1>"));
    assert!(!positional("/<sheet>/col[<letters>]"));
    assert!(!positional("/<sheet>/table[@name=<name>]"));
    assert!(!positional("<sheet>!N:N"));
    assert_eq!(path_segment(&["/<sheet>/row[N]"]), Some("row"));
    assert_eq!(
        path_segment(&["<sheet>!<A1>", "/<sheet>/<A1>"]),
        Some("<A1>")
    );
    assert_eq!(
        path_segment(&["/<sheet>/chart[N]/axis[@role=<role>]"]),
        Some("axis")
    );
    assert_eq!(path_segment(&["/"]), None);
    assert_eq!(path_segment(&["Sheet!A1"]), None);
}

#[test]
fn every_format_has_one_answer_for_its_reads() {
    for format in FORMATS {
        let capabilities = Capabilities::of(format);
        assert_eq!(capabilities.format(), Some(format));
        assert_eq!(capabilities.state(), State::Open);
        assert_eq!(capabilities.retains_source(), None);
        let get = capabilities.verb(Verb::Get);
        // The reads agree with each other and with the view modes.
        for verb in [Verb::Find, Verb::Query] {
            assert_eq!(capabilities.verb(verb), get, "{format:?} {verb:?}");
        }
        match get {
            Support::Supported => {
                assert!(capabilities.elements().next().is_some(), "{format:?}");
                assert!(capabilities.view(Mode::Text).is_supported(), "{format:?}");
            },
            Support::Unsupported { code, .. } => {
                assert_eq!(capabilities.elements().count(), 0, "{format:?}");
                let creation_gaps = if (format == FileFormat::Ppt && cfg!(feature = "ppt"))
                    || (format == FileFormat::Rtf && cfg!(feature = "rtf"))
                {
                    3
                } else {
                    0
                };
                assert_eq!(capabilities.gaps().count(), creation_gaps, "{format:?}");
                for mode in Mode::ALL {
                    assert_eq!(capabilities.view(*mode), get, "{format:?} {mode:?}");
                }
                assert!(
                    matches!(code, Code::UnsupportedCapability | Code::UnsupportedFormat),
                    "{format:?}: {code:?}"
                );
                assert_eq!(
                    capabilities.refusal().code(),
                    code,
                    "the read entry points refuse as the report says"
                );
            },
        }
        // Import is a worksheet verb with its own answer (`m4_verbs_...`);
        // merge is a template verb (`merge_is_a_template_verb`); refresh
        // has its own answer (`refresh_is_supported_where_fields_are_computed`);
        // the raw-layer writes answer by package kind (record 0336,
        // `raw::tests::raw_layer_verbs_answer_by_package`).
        for verb in Verb::ALL.into_iter().filter(|verb| {
            verb.is_edit()
                && !matches!(
                    verb,
                    Verb::Import
                        | Verb::Merge
                        | Verb::Replace
                        | Verb::Refresh
                        | Verb::AddPart
                        | Verb::RawSet
                )
        }) {
            #[cfg(feature = "pptx")]
            if format == FileFormat::Pptx
                && matches!(verb, Verb::Add | Verb::Remove | Verb::Move | Verb::Swap)
            {
                let expected = match verb {
                    Verb::Add => Support::Supported,
                    Verb::Move | Verb::Swap => Support::Supported,
                    Verb::Remove => Support::Supported,
                    _ => unreachable!("the PPTX contextual check only contains structural edits"),
                };
                assert_eq!(capabilities.verb(verb), expected);
                continue;
            }
            let docx_edit = format == FileFormat::Docx
                && cfg!(feature = "docx")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let rtf_edit = format == FileFormat::Rtf
                && cfg!(feature = "rtf")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let doc_edit = format == FileFormat::Doc
                && cfg!(feature = "doc")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let xlsx_edit = format == FileFormat::Xlsx
                && cfg!(feature = "xlsx")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let ppt_edit = format == FileFormat::Ppt
                && cfg!(feature = "ppt")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let pptx_edit = format == FileFormat::Pptx
                && cfg!(feature = "pptx")
                && matches!(verb, Verb::Set | Verb::Edit | Verb::Batch);
            let xls_edit = format == FileFormat::Xls
                && cfg!(feature = "xls")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let ods_edit = format == FileFormat::Ods
                && cfg!(feature = "ods")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            let odt_edit = format == FileFormat::Odt
                && cfg!(feature = "odt")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            // OTH and ODM share the ODT adapter (record 0282).
            let text_family_edit = matches!(format, FileFormat::Oth | FileFormat::Odm)
                && crate::read::Adapter::of(format).is_some()
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            // The object formats edit through their leaves; what a leaf
            // lacks is an element-level gap, never a verb refusal (record
            // 0292).
            #[cfg(odf_objects)]
            let object_edit = crate::read::Adapter::of(format)
                .and_then(crate::format::odf_tree::writes)
                .is_some()
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            #[cfg(not(odf_objects))]
            let object_edit = false;
            #[cfg(feature = "iwork-adapter")]
            let iwork_edit = crate::read::Adapter::of(format)
                .and_then(crate::format::iwork::writes)
                .is_some_and(|writes| match verb {
                    Verb::Set | Verb::Edit | Verb::Batch => true,
                    Verb::Add => !writes.addable.is_empty(),
                    Verb::Remove => !writes.removable.is_empty(),
                    // Pages exchanges paragraphs (record 0299); Keynote
                    // slides, paragraphs and drawables (record 0298);
                    // Numbers sheets and tables (record 0347).
                    Verb::Move | Verb::Swap => !writes.movable.is_empty(),
                    _ => false,
                });
            #[cfg(not(feature = "iwork-adapter"))]
            let iwork_edit = false;
            #[cfg(feature = "iwork-adapter")]
            let iwork_refusal = crate::read::Adapter::of(format)
                .and_then(crate::format::iwork::writes)
                .is_some()
                && !iwork_edit;
            #[cfg(not(feature = "iwork-adapter"))]
            let iwork_refusal = false;
            let odp_edit = format == FileFormat::Odp
                && cfg!(feature = "odp")
                && matches!(
                    verb,
                    Verb::Set
                        | Verb::Add
                        | Verb::Remove
                        | Verb::Move
                        | Verb::Swap
                        | Verb::Edit
                        | Verb::Batch
                );
            assert_eq!(
                capabilities.verb(verb),
                if rtf_edit
                    || odt_edit
                    || odp_edit
                    || doc_edit
                    || docx_edit
                    || xlsx_edit
                    || pptx_edit
                    || ppt_edit
                    || xls_edit
                    || ods_edit
                    || (format == FileFormat::Xlsb
                        && cfg!(feature = "xlsb")
                        && matches!(
                            verb,
                            Verb::Set
                                | Verb::Add
                                | Verb::Remove
                                | Verb::Move
                                | Verb::Swap
                                | Verb::Edit
                                | Verb::Batch
                        ))
                    || text_family_edit
                    || object_edit
                    || iwork_edit
                {
                    Support::Supported
                } else if iwork_refusal {
                    IWORK_STRUCTURE.support()
                } else if get.is_supported() {
                    EDITS.support()
                } else {
                    get
                },
                "{format:?} {verb:?}: whole-format refusals precede M2"
            );
        }

        // Native creators and publishers can precede navigation integration.
        // A format without a read adapter
        // keeps its whole-format refusal as the authoritative answer.
        for verb in [Verb::Save, Verb::Create] {
            let integrated = (format == FileFormat::Ppt && cfg!(feature = "ppt"))
                || (format == FileFormat::Xls && cfg!(feature = "xls"))
                || (format == FileFormat::Doc && cfg!(feature = "doc"))
                || (format == FileFormat::Docx && cfg!(feature = "docx"))
                || (format == FileFormat::Xlsx && cfg!(feature = "xlsx"))
                || (format == FileFormat::Pptx && cfg!(feature = "pptx"))
                || (format == FileFormat::Xlsb && cfg!(feature = "xlsb"))
                || (format == FileFormat::Rtf && cfg!(feature = "rtf"))
                || (format == FileFormat::Ods && cfg!(feature = "ods"))
                || (format == FileFormat::Odt && cfg!(feature = "odt"))
                || (format == FileFormat::Odp && cfg!(feature = "odp"))
                // Record 0297: Numbers builds a blank workbook.
                || (format == FileFormat::Numbers && cfg!(feature = "numbers"))
                || (matches!(
                    format,
                    FileFormat::Oth
                        | FileFormat::Odm
                        | FileFormat::Odg
                        | FileFormat::Odc
                        | FileFormat::Odi
                        | FileFormat::Odb
                        | FileFormat::Odf
                ) && crate::read::Adapter::of(format).is_some());
            let iwork = matches!(
                format,
                FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers
            ) && crate::read::Adapter::of(format).is_some();
            // Pages (record 0299), Numbers (record 0297) and Keynote (record
            // 0298) create their minimal documents.
            let expected = if integrated || iwork {
                Support::Supported
            } else if get.is_supported() {
                match verb {
                    Verb::Save => SAVE_PENDING.support(),
                    Verb::Create => CREATE_PENDING.support(),
                    _ => unreachable!("the W4 check only contains save/create"),
                }
            } else {
                get
            };
            assert_eq!(capabilities.verb(verb), expected, "{format:?} {verb:?}");
        }
    }
}

#[test]
fn publication_and_creation_are_not_existing_document_edits() {
    assert!(Verb::Set.is_edit());
    assert!(!Verb::Save.is_edit());
    assert!(!Verb::Create.is_edit());
}

#[cfg(feature = "xls")]
#[test]
fn xls_adapter_advertises_writes_and_keeps_granular_gaps() {
    let capabilities = Capabilities::of(FileFormat::Xls);
    assert!(capabilities.verb(Verb::Get).is_supported());
    assert!(capabilities.verb(Verb::Find).is_supported());
    assert!(capabilities.verb(Verb::Query).is_supported());
    assert!(capabilities.verb(Verb::View).is_supported());
    assert!(capabilities.verb(Verb::Save).is_supported());
    assert!(capabilities.verb(Verb::Create).is_supported());
    assert!(capabilities.elements().next().is_some());

    for verb in [
        Verb::Set,
        Verb::Add,
        Verb::Remove,
        Verb::Move,
        Verb::Swap,
        Verb::Edit,
        Verb::Batch,
    ] {
        assert!(
            capabilities.verb(verb).is_supported(),
            "XLS {verb:?} is supported for admitted writes"
        );
    }

    let cell = capabilities.element("cell").unwrap();
    assert!(cell.operations().set);
    assert!(cell.operations().remove);
    assert!(cell.property("value").unwrap().set());
    assert!(cell.property("clear").unwrap().set());
    assert!(!cell.property("value").unwrap().get());
    assert!(!cell.property("clear").unwrap().get());
    let sheet = capabilities.element("sheet").unwrap();
    assert!(sheet.operations().set);
    assert!(sheet.operations().add);
    assert!(sheet.operations().remove);
    assert!(sheet.property("name").unwrap().set());
    assert!(sheet.property("name").unwrap().add());
    for axis in ["row", "col"] {
        let element = capabilities.element(axis).unwrap();
        assert!(element.operations().add, "{axis}");
        assert!(element.operations().remove, "{axis}");
    }
    assert!(!sheet.property("hidden").unwrap().set());
    assert!(!sheet.property("visibility").unwrap().set());

    for mode in [Mode::Text, Mode::Annotated, Mode::Outline, Mode::Stats] {
        assert!(capabilities.view(mode).is_supported(), "XLS {mode:?}");
    }
    for mode in [
        Mode::Issues,
        Mode::Forms,
        Mode::Html,
        Mode::Svg,
        Mode::Screenshot,
        Mode::Pdf,
    ] {
        assert_eq!(
            capabilities.view(mode).code(),
            Some(Code::UnsupportedCapability),
            "XLS {mode:?} remains an M1 view gap"
        );
    }

    assert!(
        !capabilities
            .elements()
            .find(|element| element.data().canonical == "row")
            .unwrap()
            .header_columns()
    );

    let creation_gaps = capabilities
        .gaps()
        .filter(|gap| {
            matches!(
                gap.subject(),
                crate::node::spec::Subject::Operation {
                    verb: "create",
                    element: "workbook",
                    ..
                }
            )
        })
        .count();
    assert_eq!(creation_gaps, 3);

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
            capabilities.gaps().any(|gap| {
                gap.subject()
                    == crate::node::spec::Subject::Operation {
                        verb,
                        element,
                        context,
                    }
            }),
            "missing XLS operation gap {verb} {element} {context}"
        );
    }
}

#[test]
fn iwork_formats_without_their_feature_name_it() {
    // Every claimed format has a `File` variant since record 0296; a build
    // without a format's feature names the feature.
    for (format, feature) in [
        (FileFormat::Pages, "enable the `pages` cargo feature"),
        (FileFormat::Keynote, "enable the `keynote` cargo feature"),
        (FileFormat::Numbers, "enable the `numbers` cargo feature"),
    ] {
        let refusal = refused(format);
        assert_eq!(refusal.code, Code::UnsupportedFormat, "{format:?}");
        assert_eq!(refusal.follow_up, feature, "{format:?}");
    }
}

#[cfg(feature = "rtf")]
#[test]
fn rtf_advertises_native_paragraph_structural_writes_and_keeps_closure_gaps() {
    let capabilities = Capabilities::of(FileFormat::Rtf);
    assert!(capabilities.verb(Verb::Add).is_supported());
    assert!(capabilities.verb(Verb::Swap).is_supported());

    // The body takes added paragraphs; it is not itself added (0289).
    let body = capabilities.element("body").expect("RTF body");
    assert!(!body.operations().add);

    let paragraph = capabilities.element("paragraph").expect("RTF paragraph");
    assert!(paragraph.operations().set);
    assert!(paragraph.operations().remove);
    assert!(paragraph.operations().add);
    assert!(paragraph.property("text").expect("paragraph text").add());
    let align = paragraph.property("align").expect("paragraph alignment");
    assert!(align.get());
    assert!(align.set());
    assert!(!align.add());

    let add_gap = capabilities.gaps().find(|gap| {
        gap.subject()
            == crate::node::spec::Subject::Operation {
                verb: "add",
                element: "paragraph",
                context: "native-source-closure",
            }
    });
    assert!(add_gap.is_some(), "missing RTF paragraph add closure gap");
    let swap_gap = capabilities.gaps().find(|gap| {
        gap.subject()
            == crate::node::spec::Subject::Operation {
                verb: "swap",
                element: "body",
                context: "native-source-closure",
            }
    });
    assert!(swap_gap.is_some(), "missing RTF paragraph swap closure gap");
    assert!(!capabilities.gaps().any(|gap| {
        gap.subject()
            == crate::node::spec::Subject::Operation {
                verb: "swap",
                element: "body",
                context: "paragraph-text-only",
            }
    }));
    assert!(!capabilities.gaps().any(|gap| {
        gap.subject()
            == crate::node::spec::Subject::Operation {
                verb: "add",
                element: "body",
                context: "paragraph-text-only",
            }
    }));
}

/// Every adapter's spec table agrees with the vocabulary its resolver
/// reads: each element's identity keys are the ones a path step may use,
/// each alias normalizes to the element's kind, and the report's element
/// lookup finds the element by every name.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "xls",
    feature = "rtf",
    feature = "odt"
))]
#[test]
fn spec_tables_agree_with_their_vocabularies() {
    use crate::node::spec::FormatSpec;
    use crate::path::navigate::ElementKind;

    fn check<K: ElementKind>(spec: &FormatSpec<K>, format: FileFormat) {
        let capabilities = Capabilities::of(format);
        for element in spec.elements {
            let rule = K::VOCABULARY
                .rule(element.kind.name())
                .unwrap_or_else(|| panic!("{format:?}: no rule for {}", element.name));
            // A gap element resolves to its gap whatever identity it is
            // given, so only readable elements must agree.
            if element.gap.is_none() {
                let keys: Vec<&str> = rule
                    .identities
                    .iter()
                    .map(|identity| identity.key)
                    .collect();
                assert_eq!(
                    element.identities, keys,
                    "{format:?} {}: the report's identities are the resolver's",
                    element.name
                );
            }
            for alias in element.aliases {
                if element.kind == K::ROOT {
                    continue;
                }
                assert_eq!(
                    K::VOCABULARY.canonical(alias),
                    Some(element.kind.name()),
                    "{format:?}: {alias}"
                );
            }
            let found = capabilities
                .element(element.name)
                .unwrap_or_else(|| panic!("{format:?}: {} is reported", element.name));
            assert_eq!(found.name(), element.name);
        }
        assert_eq!(capabilities.elements().count(), spec.elements.len());
    }

    #[cfg(feature = "doc")]
    check(&crate::format::doc::SPEC, FileFormat::Doc);
    #[cfg(feature = "docx")]
    check(&crate::format::docx::SPEC, FileFormat::Docx);
    #[cfg(feature = "xlsx")]
    check(&crate::format::xlsx::SPEC, FileFormat::Xlsx);
    #[cfg(feature = "xlsb")]
    check(&crate::format::xlsb::SPEC, FileFormat::Xlsb);
    #[cfg(feature = "xls")]
    check(&crate::format::xls::SPEC, FileFormat::Xls);
    #[cfg(feature = "pptx")]
    check(&crate::format::pptx::SPEC, FileFormat::Pptx);
    #[cfg(feature = "rtf")]
    check(&crate::format::rtf::SPEC, FileFormat::Rtf);
    #[cfg(feature = "odt")]
    check(&crate::format::odt::SPEC, FileFormat::Odt);
}

#[test]
fn locked_files_refuse_every_verb_and_still_list_their_elements() {
    let locked = crate::Locked::new(
        crate::locked::Container::Ooxml,
        Some(FileFormat::Docx),
        crate::EncryptionKind::Agile,
        crate::locked::LockedSource::Owned(vec![1, 2, 3]),
        &crate::OpenOptions::default(),
    );
    let file = crate::File::Locked(locked);
    let capabilities = file.capabilities();
    assert_eq!(capabilities.state(), State::Locked);
    assert_eq!(capabilities.retains_source(), None);
    for verb in Verb::ALL {
        assert_eq!(
            capabilities.verb(verb).code(),
            Some(Code::DocumentProtected),
            "{verb:?}"
        );
    }
    for mode in Mode::ALL {
        assert_eq!(
            capabilities.view(*mode).code(),
            Some(Code::DocumentProtected)
        );
    }
    assert_eq!(
        capabilities.elements().count(),
        Capabilities::of(FileFormat::Docx).elements().count()
    );
    let error = file.get("/").unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(error.suggestion().is_some());
    assert_eq!(
        file.query("cell").unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(
        file.view(Mode::Text).unwrap_err().code(),
        Code::DocumentProtected
    );

    // A locked OOXML package that hides its format lists nothing.
    let hidden = crate::File::Locked(crate::Locked::new(
        crate::locked::Container::Ooxml,
        None,
        crate::EncryptionKind::Standard,
        crate::locked::LockedSource::Owned(vec![1, 2, 3]),
        &crate::OpenOptions::default(),
    ));
    let capabilities = hidden.capabilities();
    assert_eq!(capabilities.format(), None);
    assert_eq!(capabilities.elements().count(), 0);
    assert_eq!(
        capabilities.verb(Verb::Get).code(),
        Some(Code::DocumentProtected)
    );
    let json = serde_json::to_value(capabilities).expect("serializes");
    assert_eq!(json["format"], serde_json::Value::Null);
    assert_eq!(json["state"], "locked");
}

#[cfg(feature = "docx")]
#[test]
fn gap_subjects_are_unique() {
    let mut subjects = std::collections::HashSet::new();
    for gap in Capabilities::of(FileFormat::Docx).gaps() {
        assert!(
            subjects.insert(gap.subject()),
            "duplicate {:?}",
            gap.subject()
        );
    }
}

#[cfg(feature = "doc")]
#[test]
fn doc_enumerates_its_unprojected_element_gaps() {
    let capabilities = Capabilities::of(FileFormat::Doc);
    // Record 0324 projects styles, numbering, hyperlinks, bookmarks,
    // fields, sections, tab stops, text boxes and shapes.
    for element in ["equation", "formfield", "picture", "ole", "revision"] {
        assert!(
            capabilities
                .gaps()
                .any(|gap| gap.subject() == crate::node::spec::Subject::Element(element)),
            "missing DOC element gap {element}"
        );
    }
    let mut subjects = std::collections::HashSet::new();
    for gap in capabilities.gaps() {
        assert!(
            subjects.insert(gap.subject()),
            "duplicate {:?}",
            gap.subject()
        );
    }
}

/// Records 0273 and 0295: find/replace answers for the formats with a find
/// adapter.
#[test]
fn replace_is_supported_where_a_find_adapter_exists() {
    for format in [
        FileFormat::Docx,
        FileFormat::Xlsx,
        FileFormat::Pptx,
        FileFormat::Doc,
        FileFormat::Ppt,
        FileFormat::Xls,
        FileFormat::Xlsb,
        FileFormat::Rtf,
        FileFormat::Odt,
        FileFormat::Ods,
        FileFormat::Odp,
        FileFormat::Odg,
    ] {
        let support = Capabilities::of(format).verb(Verb::Replace);
        let adapter = crate::read::Adapter::of(format).is_some();
        let integrated = adapter && format != FileFormat::Odg;
        assert_eq!(support.is_supported(), integrated, "{format:?}");
        if adapter && !integrated {
            assert_eq!(
                support.code(),
                Some(Code::UnsupportedCapability),
                "{format:?}"
            );
        }
    }
}

/// Record 0266: import and evaluation answer for worksheets only.
#[test]
fn m4_verbs_are_worksheet_verbs() {
    for format in [FileFormat::Xlsx, FileFormat::Xlsb, FileFormat::Xls] {
        let capabilities = Capabilities::of(format);
        let workbook = crate::read::Adapter::of(format).is_some();
        assert_eq!(capabilities.verb(Verb::Import).is_supported(), workbook);
        assert_eq!(
            capabilities.verb(Verb::Evaluate).is_supported(),
            workbook && cfg!(feature = "eval")
        );
    }
    for format in [
        FileFormat::Docx,
        FileFormat::Doc,
        FileFormat::Pptx,
        FileFormat::Rtf,
    ] {
        let capabilities = Capabilities::of(format);
        for verb in [Verb::Import, Verb::Evaluate] {
            // "No worksheets", or the whole-format refusal of a build
            // without the format.
            assert!(
                !capabilities.verb(verb).is_supported(),
                "{format:?} {verb:?}"
            );
        }
    }
}

/// Records 0274 and 0295: merge answers for the text, presentation and
/// workbook templates, not for the ODF object formats.
#[test]
fn merge_is_a_template_verb() {
    for format in [
        FileFormat::Docx,
        FileFormat::Pptx,
        FileFormat::Xlsx,
        FileFormat::Xlsb,
        FileFormat::Doc,
        FileFormat::Ppt,
        FileFormat::Xls,
        FileFormat::Rtf,
        FileFormat::Odt,
        FileFormat::Ods,
        FileFormat::Odp,
        FileFormat::Odg,
    ] {
        let capabilities = Capabilities::of(format);
        let template = format != FileFormat::Odg && crate::read::Adapter::of(format).is_some();
        assert_eq!(
            capabilities.verb(Verb::Merge).is_supported(),
            template,
            "{format:?}"
        );
    }
}

#[test]
fn refresh_is_supported_where_fields_are_computed() {
    for format in [
        FileFormat::Docx,
        FileFormat::Pptx,
        FileFormat::Xlsx,
        FileFormat::Xlsb,
        FileFormat::Xls,
        FileFormat::Doc,
        FileFormat::Ppt,
        FileFormat::Rtf,
        FileFormat::Odt,
    ] {
        let support = Capabilities::of(format).verb(Verb::Refresh);
        let adapter = (format == FileFormat::Docx && cfg!(feature = "docx"))
            || (format == FileFormat::Pptx && cfg!(feature = "pptx"))
            || (format == FileFormat::Doc && cfg!(feature = "doc"))
            || (format == FileFormat::Odt && cfg!(feature = "odt"))
            || (format == FileFormat::Xlsx && cfg!(all(feature = "xlsx", feature = "eval")))
            || (format == FileFormat::Xlsb && cfg!(all(feature = "xlsb", feature = "eval")))
            || (format == FileFormat::Xls && cfg!(all(feature = "xls", feature = "eval")));
        assert_eq!(support.is_supported(), adapter, "{format:?}");
    }
}
