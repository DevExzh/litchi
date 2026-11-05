#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! The package-level edit families refuse a signed package (ADR 0006;
//! unified-ops record 0014a). A changed edit is refused with
//! `SignedSourceRequiresExplicitPolicy` (or the format-owned `Error::Signed`
//! for the shared VBA owner) and leaves the package untouched; after an
//! explicit `Package::unsign` the same edit publishes a package without
//! signature parts; an exact no-op keeps the signed source byte for byte.

mod in_place;
use in_place::InPlace as _;

use std::path::PathBuf;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};
use litchi_pptx::chart::{Chart, Series, Type as ChartType};
use litchi_pptx::comments::{
    Author as LegacyAuthor, Comment as LegacyComment, Conformance as LegacyConformance,
};
use litchi_pptx::custom::Props as CustomProps;
use litchi_pptx::font::{Data as FontData, Face, Font, Fonts, Style as FontStyle};
use litchi_pptx::modern_comments::{self, Author as ModernAuthor, Comment as ModernComment};
use litchi_pptx::presentation::embedded::ole::{Frame, Kind as OleKind};
#[cfg(feature = "vba-inspection")]
use litchi_pptx::presentation::embedded::vba;
use litchi_pptx::presentation_properties::math::{BinaryBreak, Properties as MathProperties};
use litchi_pptx::presentation_properties::metadata::changes::{self, Info, Part as ChangesPart};
use litchi_pptx::presentation_properties::metadata::custom_show::{self, List as CustomShows};
use litchi_pptx::presentation_properties::metadata::events::{self, Draft};
use litchi_pptx::presentation_properties::metadata::sections::Section;
use litchi_pptx::presentation_properties::metadata::structure;
use litchi_pptx::shape::classification::Outcome;
use litchi_pptx::shape::designer::{DrawingProperties, Tag as DesignerTag, Tags as DesignerTags};
use litchi_pptx::shape::theme::{self, Override};
use litchi_pptx::table::style::Id as StyleId;
use litchi_pptx::tag::{List as TagList, Tag};
use litchi_pptx::time::Offset;
use litchi_pptx::{Error, Package, Result};

/// The producer-signed fixtures: one signature by `xmlsec`, one by Office 2010.
const SIGNED: [&str; 2] = ["hello-world-signed.pptx", "ms-office-2010-signed.pptx"];
const SLIDE: &str = "/ppt/slides/slide1.xml";
const THEME: &str = "/ppt/theme/theme1.xml";
const SECTION: &str = "{33333333-3333-3333-3333-333333333333}";
const OTHER_TABLE_STYLE: &str = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}";
const MODERN_AUTHOR_A: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
const MODERN_AUTHOR_B: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
const MODERN_COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";

fn fixture(name: &str) -> Vec<u8> {
    test_data(&format!("poi/test-data/xmldsign/{name}"))
}

fn test_data(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(path),
    )
    .unwrap_or_else(|error| panic!("test-data/{path}: {error}"))
}

fn slide() -> PackURI {
    PackURI::new(SLIDE).unwrap()
}

#[cfg(feature = "vba-inspection")]
fn presentation() -> PackURI {
    PackURI::new("/ppt/presentation.xml").unwrap()
}

fn custom_props() -> CustomProps {
    let mut props = CustomProps::new();
    props.insert("Project", "Litchi").unwrap();
    props
}

/// Whether `bytes` still carries any OPC signature infrastructure.
fn carries_signature(bytes: &[u8]) -> bool {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    package.is_signed()
        || package
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
}

fn is_refusal(error: &Error) -> bool {
    matches!(
        error,
        Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
    )
}

#[cfg(feature = "vba-inspection")]
fn is_vba_refusal(error: &Error) -> bool {
    matches!(error, Error::Signed)
}

fn is_family_refusal(family_name: &'static str, error: &Error) -> bool {
    #[cfg(feature = "vba-inspection")]
    if family_name == "VBA project (shared owner)" {
        return is_vba_refusal(error);
    }
    #[cfg(not(feature = "vba-inspection"))]
    let _ = family_name;
    is_refusal(error)
}

/// `plain` with an OPC signature graph (origin part, one XML signature and
/// their relationships), written out so that reading it back ingresses a
/// signed source, as a file a producer signed does. The signature is not
/// valid; only its presence matters to the edit policy.
fn sign(plain: &[u8]) -> Vec<u8> {
    let source = OpcPackage::from_bytes(plain, litchi_opc::SourceLimits::default()).unwrap();
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .unwrap();
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).unwrap();
        let mut copy = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .unwrap();
        }
        package.try_add_part(Box::new(copy)).unwrap();
    }
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let mut origin_part = BlobPart::new(
        origin.clone(),
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(origin_part)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .unwrap();
    package
        .rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    assert!(
        Package::from_reader(std::io::Cursor::new(bytes.clone()))
            .unwrap()
            .opc()
            .unwrap()
            .is_signed()
    );
    bytes
}

/// The producer-signed `name` with its signature removed explicitly and
/// `prepare` applied, then signed again in-test: a signed package that holds
/// the content an exact no-op needs.
fn signed_with(name: &str, prepare: impl FnOnce(&mut Package)) -> Vec<u8> {
    let mut package = Package::from_reader(std::io::Cursor::new(fixture(name))).unwrap();
    package = package.unsign().unwrap();
    prepare(&mut package);
    sign(&package.to_bytes().unwrap())
}

/// One package-level family: a changed edit with the readback that shows
/// it, where the fixtures hold something to change, and an exact no-op,
/// where the family has one.
struct Family {
    name: &'static str,
    /// Content the producer-signed fixtures lack, added after an explicit
    /// unsign and signed again in-test (`signed_with`).
    prepare: Option<fn(&mut Package)>,
    change: Option<Change>,
    no_op: Option<fn(&mut Package) -> Result<()>>,
}

/// A changed edit and the readback of a reopened package that shows it.
struct Change {
    edit: fn(&mut Package) -> Result<()>,
    shows: fn(&Package) -> Result<bool>,
}

/// The shape-span families (classification, design element and Designer
/// properties) read one shape's XML on its own, so their shape declares its
/// namespaces itself.
const SHAPE: &str = "Self-contained";

fn add_self_contained_shape(package: &mut Package) {
    package
        .edit_opc_in_place(|opc| {
            let part = opc.get_part_mut(&slide())?;
            let xml = String::from_utf8(part.blob().to_vec()).unwrap().replacen(
                "</p:spTree>",
                concat!(
                    r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
                    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
                    r#"<p:nvSpPr><p:cNvPr id="40" name="Self-contained"/><p:cNvSpPr/><p:nvPr/>"#,
                    r#"</p:nvSpPr><p:spPr/></p:sp></p:spTree>"#,
                ),
                1,
            );
            part.set_blob(xml.into_bytes());
            Ok(())
        })
        .unwrap();
}

fn tag_list() -> TagList {
    let mut list = TagList::new();
    list.add(Tag::new("Reviewer", "Signed").unwrap()).unwrap();
    list
}

fn designer_tags() -> DesignerTags {
    let mut tags = DesignerTags::new();
    tags.push(DesignerTag::new("layout", "hero").unwrap())
        .unwrap();
    tags
}

fn math() -> MathProperties {
    let mut value = MathProperties::new();
    value.binary_break = Some(BinaryBreak::After);
    value
}

/// A minimal embedded OpenType program header, as the font unit tests build it.
fn eot(marker: u8) -> Vec<u8> {
    let mut value = vec![0; 96];
    value[0..4].copy_from_slice(&108u32.to_le_bytes());
    value[4..8].copy_from_slice(&12u32.to_le_bytes());
    value[8..12].copy_from_slice(&0x0001_0000u32.to_le_bytes());
    value[16] = marker;
    value[34..36].copy_from_slice(&0x504Cu16.to_le_bytes());
    value.extend_from_slice(&[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    value
}

fn fonts() -> Fonts {
    let mut fonts = Fonts::new();
    fonts
        .add(
            Font::from_face(
                "Signed",
                Face::new(FontStyle::Regular, FontData::powerpoint(eot(7)).unwrap()),
            )
            .unwrap(),
        )
        .unwrap();
    fonts
}

fn chart() -> Chart {
    Chart::new(ChartType::Bar, 914_400, 914_400, 4_572_000, 2_743_200).add_series(
        Series::new("2026")
            .with_categories(vec!["Q1".into(), "Q2".into()])
            .with_values(vec![100.0, 150.0]),
    )
}

fn changes_part() -> ChangesPart {
    ChangesPart {
        relationship_id: "rIdChanges".into(),
        part_name: "/ppt/changesInfo.xml".into(),
        changes_information: Info::parse(
            br#"<pc:chgInfo xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"><pc:docChgLst><pc:chgData name="Ada" userId="ada@example.test"/><pc:docChg xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" chg="addSld"><pc:docMkLst><pc:docMk/></pc:docMkLst><pc:sldChg xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" chg="modSld"><pc:sldMkLst><pc:docMk/><pc:sldMk cId="0" sldId="256"/></pc:sldMkLst><pc226:cmChg xmlns:pc226="http://schemas.microsoft.com/office/powerpoint/2022/06/main/command" xmlns:pc2="http://schemas.microsoft.com/office/powerpoint/2019/9/main/command" chg="mod"><pc2:cmMkLst xmlns:pc2="http://schemas.microsoft.com/office/powerpoint/2019/9/main/command"><pc2:cmMk id="{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}"/></pc2:cmMkLst></pc226:cmChg></pc:sldChg></pc:docChg></pc:docChgLst></pc:chgInfo>"#,
            changes::Limits::DEFAULT,
        )
        .unwrap(),
    }
}

fn first_slide_id(opc: &OpcPackage) -> Result<u32> {
    Ok(structure::load(opc, structure::Limits::DEFAULT)?.slides[0].slide_id)
}

fn legacy_author(id: u32, name: &str) -> LegacyAuthor {
    LegacyAuthor::new(id, name, &name[..1])
}

fn modern_author(id: &str, name: &str) -> ModernAuthor {
    ModernAuthor {
        id: id.into(),
        name: name.into(),
        initials: Some(name[..1].to_owned()),
        user_id: format!("{}@example.test", name.to_ascii_lowercase()),
        provider_id: "local-test".into(),
        namespace_declarations: Vec::new(),
        extension_xml: None,
    }
}

fn modern_comment(id: &str, author_id: &str, title: &str) -> ModernComment {
    ModernComment {
        id: id.into(),
        author_id: author_id.into(),
        status: None,
        created: "2026-07-19T12:00:00Z".into(),
        start_date: None,
        due_date: None,
        assigned_to: None,
        complete: None,
        title: Some(title.into()),
        namespace_declarations: Vec::new(),
        anchors: Vec::new(),
        position: None,
        reply_list_namespace_declarations: Vec::new(),
        replies: Vec::new(),
        reply_list_present: false,
        text_body_xml: None,
        extension_xml: None,
    }
}

#[allow(
    clippy::too_many_lines,
    reason = "one table row per package-level family keeps the families comparable"
)]
fn families() -> Vec<Family> {
    let mut families = vec![
        Family {
            name: "legacy comments",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_comments_in_place(|edit| {
                        edit.add_author(
                            legacy_author(1, "Reviewer"),
                            LegacyConformance::Transitional,
                        )
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(litchi_pptx::find_presentation_comment_author(opc, 1)?.is_some())
                    })
                },
            }),
            no_op: Some(|package| {
                package.edit_comments_in_place(|edit| {
                    edit.remove_author(1, LegacyConformance::Transitional)
                        .map(drop)
                })
            }),
        },
        Family {
            name: "modern comments",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_modern_comments_in_place(|edit| {
                        edit.add_author(modern_author(MODERN_AUTHOR_A, "Ada"))
                            .map(drop)
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(
                            modern_comments::find_modern_comment_author(opc, MODERN_AUTHOR_A)?
                                .is_some(),
                        )
                    })
                },
            }),
            no_op: Some(|package| {
                package.edit_modern_comments_in_place(|edit| {
                    edit.remove_author(MODERN_AUTHOR_A).map(drop)
                })
            }),
        },
        Family {
            name: "notes",
            prepare: None,
            // The fixtures have no notes graph to change; see
            // `notes_refuse_a_signed_package_until_it_is_unsigned`.
            change: None,
            no_op: Some(|package| {
                assert!(!package.remove_notes_in_place(0_usize)?);
                assert_eq!(package.clear_notes_in_place()?, 0);
                Ok(())
            }),
        },
        Family {
            name: "slide tags",
            prepare: None,
            change: Some(Change {
                edit: |package| package.put_tags_in_place(0_usize, tag_list()).map(drop),
                shows: |package| Ok(package.tags(0_usize)?.is_some()),
            }),
            no_op: Some(|package| package.remove_tags_in_place(0_usize).map(drop)),
        },
        Family {
            name: "shape tags",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package
                        .put_shape_tags_in_place(0_usize, 0_usize, tag_list())
                        .map(drop)
                },
                shows: |package| Ok(package.shape_tags(0_usize, 0_usize)?.is_some()),
            }),
            no_op: Some(|package| {
                package
                    .remove_shape_tags_in_place(0_usize, 0_usize)
                    .map(drop)
            }),
        },
        Family {
            name: "change tracking",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    let mut edit = package.edit_change_tracking()?;
                    edit.set_creation_id(0_usize, 7_u32)?;
                    let commit = edit.commit()?;
                    *package = commit.into_package();
                    Ok(())
                },
                shows: |package| {
                    Ok(package.change_tracking(0_usize)?.state().creation_id()
                        == Some(litchi_pptx::change_tracking::Id::new(7)))
                },
            }),
            no_op: Some(|package| {
                package.change_tracking(0_usize)?;
                let commit = package.edit_change_tracking()?.commit()?;
                assert!(commit.patch().is_noop());
                *package = commit.into_package();
                Ok(())
            }),
        },
        Family {
            name: "shape classification",
            prepare: Some(add_self_contained_shape),
            change: Some(Change {
                edit: |package| {
                    package
                        .put_shape_classification_in_place(0_usize, SHAPE, Outcome::Header)
                        .map(drop)
                },
                shows: |package| {
                    Ok(package
                        .shape_classification(0_usize, SHAPE)?
                        .and_then(|snapshot| snapshot.outcome())
                        == Some(Outcome::Header))
                },
            }),
            no_op: Some(|package| {
                package
                    .remove_shape_classification_in_place(0_usize, SHAPE)
                    .map(drop)
            }),
        },
        Family {
            name: "shape design element",
            prepare: Some(add_self_contained_shape),
            change: Some(Change {
                edit: |package| {
                    package
                        .put_shape_design_element_in_place(0_usize, SHAPE, true)
                        .map(drop)
                },
                shows: |package| {
                    Ok(package
                        .shape_design_element(0_usize, SHAPE)?
                        .and_then(|snapshot| snapshot.value())
                        == Some(true))
                },
            }),
            no_op: Some(|package| {
                package
                    .remove_shape_design_element_in_place(0_usize, SHAPE)
                    .map(drop)
            }),
        },
        Family {
            name: "shape Designer properties",
            prepare: Some(add_self_contained_shape),
            change: Some(Change {
                edit: |package| {
                    package
                        .put_shape_designer_properties_in_place(
                            0_usize,
                            SHAPE,
                            DrawingProperties::new().with_editable(Some(true)),
                        )
                        .map(drop)
                },
                shows: |package| {
                    Ok(package
                        .shape_designer_properties(0_usize, SHAPE)?
                        .properties()
                        .is_some())
                },
            }),
            no_op: Some(|package| {
                package
                    .remove_shape_designer_properties_in_place(0_usize, SHAPE)
                    .map(drop)
            }),
        },
        Family {
            name: "zoom and model3d owners",
            prepare: None,
            // Neither fixture holds a zoom or a model3d frame to change; see
            // `model3d_refuses_a_signed_package_until_it_is_unsigned`.
            change: None,
            no_op: Some(|package| {
                assert!(package.remove_zooms_in_place(0_usize)?.is_none());
                assert!(package.remove_model3d_in_place(0_usize, 0_usize)?.is_none());
                Ok(())
            }),
        },
        Family {
            name: "themes",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_themes_in_place(|edit| {
                        let current = edit.theme(THEME)?;
                        edit.add("Signed copy", &current.colors, &current.fonts)
                            .map(drop)
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(opc.contains_part(&PackURI::new("/ppt/theme/theme2.xml").unwrap()))
                    })
                },
            }),
            no_op: Some(|package| {
                package.edit_themes_in_place(|edit| {
                    assert!(!edit.remove_override(0_usize)?);
                    Ok(())
                })
            }),
        },
        Family {
            name: "math properties",
            prepare: None,
            change: Some(Change {
                edit: |package| package.put_math_properties_in_place(math()).map(drop),
                shows: |package| Ok(package.math_properties()? == Some(math())),
            }),
            no_op: Some(|package| package.remove_math_properties_in_place().map(drop)),
        },
        Family {
            name: "presentation structure (sections)",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_structure_in_place(|edit| {
                        edit.add_section(Section::new("Signed", SECTION)).map(drop)
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| Ok(structure::find_section(opc, SECTION)?.is_some()))
                },
            }),
            no_op: Some(|package| {
                package.edit_structure_in_place(|edit| {
                    assert!(!edit.remove_section(SECTION)?);
                    Ok(())
                })
            }),
        },
        Family {
            name: "custom shows",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    let slide_id = package.with_opc(first_slide_id)?;
                    package.edit_custom_shows_in_place(|edit| {
                        let mut shows = CustomShows::new();
                        shows.create("Signed", vec![slide_id]);
                        edit.set(&shows)
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(!custom_show::load(
                            opc,
                            custom_show::Limits::from_source_limits(opc.read_limits()),
                        )?
                        .is_empty())
                    })
                },
            }),
            no_op: Some(|package| {
                package.edit_custom_shows_in_place(|edit| {
                    assert!(edit.remove()?.is_none());
                    Ok(())
                })
            }),
        },
        Family {
            name: "slide-show events",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_show_events_in_place(|edit| {
                        edit.set(0_usize, &[Draft::play(Offset::ms(1), 2)])
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(events::load(opc, &slide(), events::Limits::DEFAULT)?.is_some())
                    })
                },
            }),
            no_op: Some(|package| {
                package.edit_show_events_in_place(|edit| {
                    assert!(edit.remove(0_usize)?.is_none());
                    Ok(())
                })
            }),
        },
        Family {
            name: "extended guides",
            prepare: None,
            // A guide graph is not authored here; removing none is the no-op.
            change: None,
            no_op: Some(|package| package.edit_guides_in_place(|edit| edit.remove().map(drop))),
        },
        Family {
            name: "slide Designer tags",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package
                        .put_slide_designer_tags_in_place(0_usize, designer_tags())
                        .map(drop)
                },
                shows: |package| {
                    Ok(package.slide_designer_tags(0_usize)?.tags()? == Some(&designer_tags()))
                },
            }),
            no_op: Some(|package| {
                package
                    .remove_slide_designer_tags_in_place(0_usize)
                    .map(drop)
            }),
        },
        Family {
            name: "Changes Information",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_changes_info_in_place(|edit| edit.set(&changes_part()))
                },
                shows: |package| {
                    package
                        .with_opc(|opc| Ok(changes::load(opc, changes::Limits::DEFAULT)?.is_some()))
                },
            }),
            no_op: Some(|package| {
                package.edit_changes_info_in_place(|edit| {
                    assert!(edit.remove()?.is_none());
                    Ok(())
                })
            }),
        },
        Family {
            name: "embedded OLE objects",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_ole_objects_in_place(|edit| {
                        edit.add(
                            0_usize,
                            OleKind::Package,
                            None,
                            Some("Signed payload"),
                            Frame::new(0, 0, 914_400, 914_400),
                            b"inert payload",
                        )
                        .map(drop)
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(opc
                            .iter_parts()
                            .any(|part| part.partname().as_str().starts_with("/ppt/embeddings/")))
                    })
                },
            }),
            no_op: Some(|package| {
                package.edit_slide_ole_objects_in_place(|edit| edit.update(0, |_| Ok(())))
            }),
        },
        Family {
            name: "embedded fonts",
            prepare: None,
            change: Some(Change {
                edit: |package| package.put_fonts_in_place(fonts()).map(drop),
                shows: |package| Ok(package.fonts()?.is_some()),
            }),
            no_op: Some(|package| {
                assert!(package.remove_fonts_in_place()?.is_none());
                Ok(())
            }),
        },
        Family {
            name: "table styles",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    let mut styles = package.styles()?.expect("table-style catalog");
                    styles.set_default(StyleId::parse(OTHER_TABLE_STYLE)?);
                    package.put_styles_in_place(styles).map(drop)
                },
                shows: |package| {
                    Ok(package.styles()?.expect("table-style catalog").default()
                        == StyleId::parse(OTHER_TABLE_STYLE)?)
                },
            }),
            no_op: Some(|package| {
                let styles = package.styles()?.expect("table-style catalog");
                assert!(!package.put_styles_in_place(styles)?);
                Ok(())
            }),
        },
        Family {
            name: "slide masters, layouts and placeholders",
            prepare: None,
            change: Some(Change {
                edit: |package| package.add_slide_master_in_place().map(drop),
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(opc.contains_part(
                            &PackURI::new("/ppt/slideMasters/slideMaster2.xml").unwrap(),
                        ))
                    })
                },
            }),
            no_op: Some(|package| {
                package.placeholder_type_extension(0_usize, 0_usize)?;
                let commit = package.edit_placeholder_type_extension()?.commit()?;
                assert!(commit.patch().is_noop());
                *package = commit.into_package();
                Ok(())
            }),
        },
        Family {
            name: "charts",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_charts_in_place(|edit| edit.add(0_usize, &chart()).map(drop))
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(opc.contains_part(&PackURI::new("/ppt/charts/chart1.xml").unwrap()))
                    })
                },
            }),
            // Adding a chart always changes the package.
            no_op: None,
        },
        Family {
            name: "Package::edit_opc",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    package.edit_opc_in_place(|opc| {
                        let part = opc.get_part_mut(&slide())?;
                        let xml = String::from_utf8(part.blob().to_vec())
                            .unwrap()
                            .replacen("Hello", "Signed", 1);
                        part.set_blob(xml.into_bytes());
                        Ok(())
                    })
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(opc
                            .get_part(&slide())?
                            .blob()
                            .windows(b"Signed".len())
                            .any(|window| window == b"Signed"))
                    })
                },
            }),
            no_op: Some(|package| package.edit_opc_in_place(|_| Ok(()))),
        },
        Family {
            name: "opened-presentation transaction",
            prepare: None,
            change: Some(Change {
                edit: |package| {
                    let mut transaction = package.edit()?;
                    transaction.add_comment_author(
                        legacy_author(1, "Reviewer"),
                        LegacyConformance::Transitional,
                    )?;
                    let commit = transaction.commit()?;
                    *package = commit.into_package();
                    Ok(())
                },
                shows: |package| {
                    package.with_opc(|opc| {
                        Ok(litchi_pptx::find_presentation_comment_author(opc, 1)?.is_some())
                    })
                },
            }),
            no_op: Some(|package| {
                let styles = package.styles()?.expect("table-style catalog");
                let mut transaction = package.edit()?;
                assert!(!transaction.put_table_styles(styles)?);
                let commit = transaction.commit()?;
                assert!(!commit.diagnostics().changed());
                *package = commit.into_package();
                Ok(())
            }),
        },
    ];
    // Shared litchi-ooxml-common owners reached through this crate (F1b).
    families.push(Family {
        name: "custom properties (shared owner)",
        prepare: None,
        change: Some(Change {
            edit: |package| package.put_custom_props_in_place(custom_props()),
            shows: |package| Ok(package.custom_props()?.get("Project").is_some()),
        }),
        no_op: Some(|package| package.put_custom_props_in_place(litchi_pptx::custom::Props::new())),
    });
    #[cfg(feature = "vba-inspection")]
    families.push(Family {
        name: "VBA project (shared owner)",
        prepare: None,
        change: Some(Change {
            edit: |package| {
                package.edit_vba_in_place(|edit| edit.set(b"inert project".to_vec()).map(drop))
            },
            shows: |package| {
                package.with_opc(|opc| Ok(vba::discover(opc, &presentation())?.is_some()))
            },
        }),
        no_op: Some(|package| {
            package.edit_vba_in_place(|edit| {
                assert!(!edit.remove()?);
                Ok(())
            })
        }),
    });
    families
}

/// For every producer-signed fixture and every family: the changed edit is
/// refused and leaves the source bytes; the exact no-op keeps them; after
/// `Package::unsign` the changed edit publishes a package without signature
/// parts that shows the change.
#[test]
fn every_family_refuses_a_signed_package_until_it_is_unsigned() {
    for name in SIGNED {
        for family in families() {
            let label = format!("{name}: {}", family.name);
            let source = match family.prepare {
                Some(prepare) => signed_with(name, prepare),
                None => fixture(name),
            };

            if let Some(no_op) = family.no_op {
                let mut package =
                    Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();
                no_op(&mut package).unwrap_or_else(|error| panic!("{label}: no-op: {error:?}"));
                assert!(package.opc().unwrap().is_signed(), "{label}");
                assert_eq!(package.to_bytes().unwrap(), source, "{label}: no-op");
            }

            let Some(change) = family.change else {
                continue;
            };
            let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();
            let error = (change.edit)(&mut package)
                .expect_err(&format!("{label}: a signed package must refuse"));
            assert!(is_family_refusal(family.name, &error), "{label}: {error:?}");
            assert!(package.opc().unwrap().is_signed(), "{label}");
            assert_eq!(package.to_bytes().unwrap(), source, "{label}: refusal");

            package = package.unsign().unwrap();
            (change.edit)(&mut package)
                .unwrap_or_else(|error| panic!("{label}: after unsign: {error:?}"));
            let output = package.to_bytes().unwrap();
            assert!(!carries_signature(&output), "{label}");
            let reopened = Package::from_reader(std::io::Cursor::new(output)).unwrap();
            assert!(
                (change.shows)(&reopened).unwrap(),
                "{label}: the change is not visible"
            );
        }
    }
}

/// Removing the signature inside the raw closure authorizes its change, as
/// `Package::unsign` before the edit does.
#[test]
fn raw_edit_that_unsigns_explicitly_publishes_without_a_signature() {
    for name in SIGNED {
        let mut package = Package::from_reader(std::io::Cursor::new(fixture(name))).unwrap();
        package
            .edit_opc_in_place(|opc| {
                opc.unsign();
                opc.add_part(Box::new(BlobPart::new(
                    PackURI::new("/custom/raw.bin").unwrap(),
                    "application/octet-stream".into(),
                    vec![1, 2, 3],
                )));
                Ok(())
            })
            .unwrap();
        let output = package.to_bytes().unwrap();
        assert!(!carries_signature(&output), "{name}");
    }
}

#[test]
fn notes_refuse_a_signed_package_until_it_is_unsigned() {
    let source = sign(&test_data("poi/test-data/slideshow/prProps.pptx"));
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();
    let graph = package.notes().unwrap().expect("notes graph");

    package.put_notes_in_place(graph).unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    assert!(is_refusal(
        &package.remove_notes_in_place(0_usize).unwrap_err()
    ));
    assert!(is_refusal(&package.clear_notes_in_place().unwrap_err()));
    assert_eq!(package.to_bytes().unwrap(), source);

    package = package.unsign().unwrap();
    assert!(package.remove_notes_in_place(0_usize).unwrap());
    let output = package.to_bytes().unwrap();
    assert!(!carries_signature(&output));
    let reopened = Package::from_reader(std::io::Cursor::new(output)).unwrap();
    assert!(
        reopened
            .notes()
            .unwrap()
            .is_none_or(|graph| graph.slides().is_empty())
    );
}

#[test]
fn model3d_refuses_a_signed_package_until_it_is_unsigned() {
    const MODEL_FRAME: &str = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="42" name="3D Model"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" r:embed="rIdModel"><am3d:spPr/><am3d:camera/><am3d:trans/><am3d:raster rName="Office Renderer" rVer="17.0"><am3d:blip r:embed="rIdPreview"/></am3d:raster><am3d:objViewport/></am3d:model3d></a:graphicData></a:graphic></p:graphicFrame>"#;
    let source = signed_with(SIGNED[0], |package| {
        package
            .edit_opc_in_place(|opc| {
                let part = opc.get_part_mut(&slide())?;
                let xml = String::from_utf8(part.blob().to_vec()).unwrap().replacen(
                    "</p:spTree>",
                    &format!("{MODEL_FRAME}</p:spTree>"),
                    1,
                );
                part.set_blob(xml.into_bytes());
                part.rels_mut().add_relationship(
                    "http://schemas.microsoft.com/office/2017/06/relationships/model3d".to_owned(),
                    "../media/model3d.glb".to_owned(),
                    "rIdModel".to_owned(),
                    false,
                );
                part.rels_mut().add_relationship(
                    rt::IMAGE.to_owned(),
                    "../media/model3d-preview.png".to_owned(),
                    "rIdPreview".to_owned(),
                    false,
                );
                opc.add_part(Box::new(BlobPart::new(
                    PackURI::new("/ppt/media/model3d.glb").unwrap(),
                    "model/gltf-binary".to_owned(),
                    b"glb-v1".to_vec(),
                )));
                opc.add_part(Box::new(BlobPart::new(
                    PackURI::new("/ppt/media/model3d-preview.png").unwrap(),
                    ct::PNG.to_owned(),
                    b"png-v1".to_vec(),
                )));
                Ok(())
            })
            .unwrap();
    });
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();
    let model = package
        .model3d(0_usize, "3D Model")
        .unwrap()
        .expect("model");

    package
        .put_model3d_in_place(0_usize, "3D Model", model.clone())
        .unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    assert!(is_refusal(
        &package
            .remove_model3d_in_place(0_usize, "3D Model")
            .unwrap_err()
    ));
    assert_eq!(package.to_bytes().unwrap(), source);

    package = package.unsign().unwrap();
    assert!(
        package
            .remove_model3d_in_place(0_usize, "3D Model")
            .unwrap()
            .is_some()
    );
    let output = package.to_bytes().unwrap();
    assert!(!carries_signature(&output));
    assert!(matches!(
        Package::from_reader(std::io::Cursor::new(output))
            .unwrap()
            .model3d(0_usize, "3D Model"),
        Ok(None)
    ));
}

/// Legacy-comment updates and reorders that change nothing publish nothing
/// (record 0014a class c): they used to rewrite the part and strip.
#[test]
fn legacy_comment_no_ops_keep_a_signed_package_exact() {
    let conformance = LegacyConformance::Transitional;
    let comment = |author_id, text: &str| LegacyComment {
        author_id,
        date_time: Some("2026-07-19T12:00:00Z".into()),
        index: 1,
        x: 100,
        y: 200,
        text: text.into(),
    };
    let source = signed_with(SIGNED[0], |package| {
        package
            .edit_comments_in_place(|edit| {
                edit.add_author(legacy_author(1, "Ada"), conformance)?;
                edit.add_author(legacy_author(2, "Bob"), conformance)?;
                edit.add(0_usize, comment(1, "first"), conformance)?;
                edit.add(0_usize, comment(2, "second"), conformance)
            })
            .unwrap();
    });
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();

    package
        .edit_comments_in_place(|edit| {
            let author = edit.author(1)?.unwrap();
            edit.update_author(1, author, conformance)?;
            edit.reorder_authors(&[1, 2], conformance)?;
            let first = edit.comment(0_usize, 1, 1)?.unwrap();
            edit.update(0_usize, 1, 1, first, conformance)?;
            edit.reorder(0_usize, &[(1, 1), (2, 1)], conformance)
        })
        .unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    let refused = [
        package.edit_comments_in_place(|edit| edit.reorder_authors(&[2, 1], conformance)),
        package.edit_comments_in_place(|edit| {
            edit.update(0_usize, 1, 1, comment(1, "changed"), conformance)
        }),
        package.edit_comments_in_place(|edit| edit.remove(0_usize, 2, 1, conformance).map(drop)),
    ];
    for result in refused {
        assert!(is_refusal(&result.unwrap_err()));
    }
    assert_eq!(package.to_bytes().unwrap(), source);

    package = package.unsign().unwrap();
    package
        .edit_comments_in_place(|edit| edit.reorder_authors(&[2, 1], conformance))
        .unwrap();
    assert!(!carries_signature(&package.to_bytes().unwrap()));
}

/// Modern-comment and modern-author updates and reorders that change
/// nothing publish nothing (record 0014a class c).
#[test]
fn modern_comment_no_ops_keep_a_signed_package_exact() {
    let source = signed_with(SIGNED[1], |package| {
        package
            .edit_modern_comments_in_place(|edit| {
                edit.add_author(modern_author(MODERN_AUTHOR_A, "Ada"))?;
                edit.add_author(modern_author(MODERN_AUTHOR_B, "Bob"))?;
                edit.add(
                    0_usize,
                    modern_comment(MODERN_COMMENT, MODERN_AUTHOR_A, "Review"),
                )
                .map(drop)
            })
            .unwrap();
    });
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();
    let authors = [MODERN_AUTHOR_A.to_owned(), MODERN_AUTHOR_B.to_owned()];

    package
        .edit_modern_comments_in_place(|edit| {
            assert!(edit.update_author(MODERN_AUTHOR_A, |_| {})?);
            edit.reorder_authors(&authors)?;
            assert!(edit.update(0_usize, MODERN_COMMENT, |_| {})?);
            edit.reorder(0_usize, &[MODERN_COMMENT.to_owned()])
                .map(drop)
        })
        .unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    let reversed = [MODERN_AUTHOR_B.to_owned(), MODERN_AUTHOR_A.to_owned()];
    let refused = [
        package.edit_modern_comments_in_place(|edit| edit.reorder_authors(&reversed).map(drop)),
        package.edit_modern_comments_in_place(|edit| {
            edit.update(0_usize, MODERN_COMMENT, |comment| {
                comment.title = Some("Changed".into());
            })
            .map(drop)
        }),
        package.edit_modern_comments_in_place(|edit| edit.remove_author(MODERN_AUTHOR_B).map(drop)),
    ];
    for result in refused {
        assert!(is_refusal(&result.unwrap_err()));
    }
    assert_eq!(package.to_bytes().unwrap(), source);

    package = package.unsign().unwrap();
    package
        .edit_modern_comments_in_place(|edit| edit.reorder_authors(&reversed).map(drop))
        .unwrap();
    assert!(!carries_signature(&package.to_bytes().unwrap()));
}

/// Modern-comment change commands that a callback leaves unchanged publish
/// nothing (record 0014a class c).
#[test]
fn modern_comment_command_no_op_keeps_a_signed_package_exact() {
    let source = signed_with(SIGNED[0], |package| {
        package
            .edit_changes_info_in_place(|edit| edit.set(&changes_part()))
            .unwrap();
        package
            .edit_modern_comments_in_place(|edit| {
                // Publish the commands in this library's own spelling first,
                // so the no-op below reproduces the stored bytes.
                assert!(edit.update_changes(|_| {})?);
                Ok(())
            })
            .unwrap();
    });
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();
    package
        .edit_modern_comments_in_place(|edit| {
            assert!(edit.update_changes(|_| {})?);
            Ok(())
        })
        .unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    let change = |edit: &mut modern_comments::Edit<'_>| {
        edit.update_changes(|commands| {
            commands[0].value.changes = vec![modern_comments::Change::Add];
        })
        .map(drop)
    };
    assert!(is_refusal(
        &package.edit_modern_comments_in_place(change).unwrap_err()
    ));
    assert_eq!(package.to_bytes().unwrap(), source);

    package = package.unsign().unwrap();
    package.edit_modern_comments_in_place(change).unwrap();
    assert!(!carries_signature(&package.to_bytes().unwrap()));
}

/// Theme scheme and override replacements that encode to the stored bytes
/// publish nothing (record 0014a class c).
#[test]
fn theme_no_ops_keep_a_signed_package_exact() {
    const AUTHORED: &str = "/ppt/theme/theme2.xml";
    let (colors, fonts) = {
        let package = Package::from_reader(std::io::Cursor::new(fixture(SIGNED[1]))).unwrap();
        let current = theme::load(package.opc().unwrap(), THEME).unwrap();
        (current.colors, current.fonts)
    };
    let palette_override = Override {
        colors: Some(colors.clone()),
        fonts: None,
    };
    let source = signed_with(SIGNED[1], |package| {
        package
            .edit_themes_in_place(|edit| {
                edit.add("Signed", &colors, &fonts)?;
                // Replace the schemes once while unsigned, so the theme holds
                // them in the spelling a replacement publishes.
                edit.set_colors(AUTHORED, &colors)?;
                edit.set_fonts(AUTHORED, &fonts)?;
                edit.set_override(0_usize, &palette_override).map(drop)
            })
            .unwrap();
    });
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();

    package
        .edit_themes_in_place(|edit| {
            edit.set_colors(AUTHORED, &colors)?;
            edit.set_fonts(AUTHORED, &fonts)?;
            edit.set_override(0_usize, &palette_override).map(drop)
        })
        .unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    let fonts_override = Override {
        colors: None,
        fonts: Some(fonts.clone()),
    };
    let master = in_place::master_id(
        &package,
        &PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap(),
    )
    .unwrap();
    let refused = [
        package.edit_themes_in_place(|edit| edit.set_override(0_usize, &fonts_override).map(drop)),
        package.edit_themes_in_place(|edit| edit.remove_override(0_usize).map(drop)),
        package.edit_themes_in_place(|edit| edit.attach(master, AUTHORED).map(drop)),
    ];
    for result in refused {
        assert!(is_refusal(&result.unwrap_err()));
    }
    assert_eq!(package.to_bytes().unwrap(), source);

    package = package.unsign().unwrap();
    package
        .edit_themes_in_place(|edit| edit.remove_override(0_usize).map(drop))
        .unwrap();
    assert!(!carries_signature(&package.to_bytes().unwrap()));
}

/// Shape-span puts of the stored value return before the package is
/// borrowed mutably, so a signed package stays exact; before record 0014a
/// that borrow revoked the exact source and the unchanged signed package
/// could no longer be published.
#[test]
fn shape_put_no_ops_keep_a_signed_package_exact() {
    let source = signed_with(SIGNED[0], |package| {
        add_self_contained_shape(package);
        package
            .put_shape_classification_in_place(0_usize, SHAPE, Outcome::Header)
            .unwrap();
        package
            .put_shape_design_element_in_place(0_usize, SHAPE, true)
            .unwrap();
    });
    let mut package = Package::from_reader(std::io::Cursor::new(source.clone())).unwrap();

    package
        .put_shape_classification_in_place(0_usize, SHAPE, Outcome::Header)
        .unwrap();
    package
        .put_shape_design_element_in_place(0_usize, SHAPE, true)
        .unwrap();
    assert_eq!(package.to_bytes().unwrap(), source);

    let refused = [
        package
            .put_shape_classification_in_place(0_usize, SHAPE, Outcome::Footer)
            .map(drop),
        package
            .remove_shape_classification_in_place(0_usize, SHAPE)
            .map(drop),
        package
            .put_shape_design_element_in_place(0_usize, SHAPE, false)
            .map(drop),
        package
            .remove_shape_design_element_in_place(0_usize, SHAPE)
            .map(drop),
    ];
    for result in refused {
        assert!(is_refusal(&result.unwrap_err()));
    }
    assert_eq!(package.to_bytes().unwrap(), source);
}

/// Generic removal of a signed source's signature parts is not an explicit
/// removal (ADR 0006): the OPC policy requirement survives it, and the
/// families keep refusing until `OpcPackage::unsign`.
#[test]
fn generic_removal_of_signature_parts_does_not_authorize_a_change() {
    for name in SIGNED {
        let mut opc =
            OpcPackage::from_vec(fixture(name), litchi_opc::SourceLimits::default()).unwrap();
        let signature_parts: Vec<_> = opc
            .iter_parts()
            .map(|part| part.partname().clone())
            .filter(|part| part.as_str().starts_with("/_xmlsignatures/"))
            .collect();
        assert!(!signature_parts.is_empty(), "{name}");
        for part in &signature_parts {
            assert!(opc.remove_part(part), "{name}");
        }
        let origin: Vec<_> = opc
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in origin {
            opc.rels_mut().remove(&id);
        }
        assert!(!opc.is_signed(), "{name}");
        assert!(opc.requires_signature_edit_policy(), "{name}");

        let package = Package::from_opc_package(opc.clone()).unwrap();
        let mut edit = package.edit_tags().unwrap();
        let error = edit.set(0_usize, tag_list()).unwrap_err();
        assert!(is_refusal(&error), "{name}: {error:?}");

        opc.unsign();
        let package = Package::from_opc_package(opc).unwrap();
        let mut edit = package.edit_tags().unwrap();
        edit.set(0_usize, tag_list()).unwrap();
        let published = edit.commit().unwrap().into_package();
        assert!(!carries_signature(&published.to_bytes().unwrap()), "{name}");
    }
}
