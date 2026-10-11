//! Unit tests of the capability report: its generation rules, and the
//! agreement of every adapter's spec table with the resolver's vocabulary.
//! The golden snapshots and the liveness tests over the fixture corpus are
//! integration tests (`tests/capabilities.rs`).

use litchi_core::FileFormat;

use super::{
    CREATE_PENDING, Capabilities, EDITS, SAVE_PENDING, State, Support, Verb, path_segment,
    positional, refused,
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
                assert_eq!(capabilities.gaps().count(), 0, "{format:?}");
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
        for verb in Verb::ALL.into_iter().filter(|verb| verb.is_edit()) {
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
            let doc_edit = format == FileFormat::Doc
                && cfg!(feature = "doc")
                && matches!(verb, Verb::Set | Verb::Edit | Verb::Batch);
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
            let pptx_edit = format == FileFormat::Pptx
                && cfg!(feature = "pptx")
                && matches!(verb, Verb::Set | Verb::Edit | Verb::Batch);
            assert_eq!(
                capabilities.verb(verb),
                if doc_edit
                    || docx_edit
                    || xlsx_edit
                    || pptx_edit
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
                {
                    Support::Supported
                } else if get.is_supported() {
                    EDITS.support()
                } else {
                    get
                },
                "{format:?} {verb:?}: whole-format refusals precede M2"
            );
        }

        // DOC, DOCX, XLSX, XLSB and PPTX have creators and publishers.
        // A format without a read adapter
        // keeps its whole-format refusal as the authoritative answer.
        for verb in [Verb::Save, Verb::Create] {
            let integrated = (format == FileFormat::Xls && cfg!(feature = "xls"))
                || (format == FileFormat::Doc && cfg!(feature = "doc"))
                || (format == FileFormat::Docx && cfg!(feature = "docx"))
                || (format == FileFormat::Xlsx && cfg!(feature = "xlsx"))
                || (format == FileFormat::Pptx && cfg!(feature = "pptx"))
                || (format == FileFormat::Xlsb && cfg!(feature = "xlsb"));
            let expected = if integrated {
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
fn xls_read_adapter_advertises_m1_reads_and_keeps_mutations_pending() {
    let capabilities = Capabilities::of(FileFormat::Xls);
    assert!(capabilities.verb(Verb::Get).is_supported());
    assert!(capabilities.verb(Verb::Find).is_supported());
    assert!(capabilities.verb(Verb::Query).is_supported());
    assert!(capabilities.verb(Verb::View).is_supported());
    assert!(capabilities.verb(Verb::Save).is_supported());
    assert!(capabilities.verb(Verb::Create).is_supported());
    assert!(capabilities.elements().next().is_some());

    for verb in Verb::ALL.into_iter().filter(|verb| verb.is_edit()) {
        assert_eq!(
            capabilities.verb(verb).code(),
            Some(Code::UnsupportedCapability),
            "XLS {verb:?} remains a pending mutation"
        );
    }

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
}

#[test]
fn formats_without_a_variant_point_at_m6() {
    for format in [
        FileFormat::Pages,
        FileFormat::Keynote,
        FileFormat::Numbers,
        FileFormat::Odg,
        FileFormat::Odc,
        FileFormat::Odf,
        FileFormat::Odi,
        FileFormat::Odm,
        FileFormat::Oth,
        FileFormat::Odb,
    ] {
        let refusal = refused(format);
        assert_eq!(refusal.code, Code::UnsupportedFormat, "{format:?}");
        assert_eq!(refusal.follow_up, "M6", "{format:?}");
    }
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
    feature = "xls"
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
