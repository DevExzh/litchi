//! Shared, inert Office web-extension and persisted task-pane metadata.

mod codec;
mod model;
mod package;

use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use std::sync::Arc;

const WEB_EXTENSION_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
const TASK_PANES_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";

/// Low-level OPC constants for callers constructing synthetic or specialized graphs.
pub mod raw {
    /// Package relationship to the persisted task-pane part.
    pub const TASK_PANES_RELATIONSHIP: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes";
    /// Relationship from the task-pane part to one Office Add-in part.
    pub const ADD_IN_RELATIONSHIP: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextension";
    /// Content type of a persisted task-pane part.
    pub const TASK_PANES_CONTENT_TYPE: &str = "application/vnd.ms-office.webextensiontaskpanes+xml";
    /// Content type of an Office Add-in part.
    pub const ADD_IN_CONTENT_TYPE: &str = "application/vnd.ms-office.webextension+xml";
}

use raw::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};

const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

const STANDARD_XML_BYTES: usize = 4 * 1024 * 1024;
const STANDARD_TOTAL_XML_BYTES: usize = 64 * 1024 * 1024;
const STANDARD_DEPTH: usize = 128;
const STANDARD_NODES: usize = 65_536;
const STANDARD_ITEMS: usize = 4096;
const STANDARD_STRING_BYTES: usize = 8 * 1024 * 1024;
const STANDARD_TOTAL_STRING_BYTES: usize = 128 * 1024 * 1024;
const STANDARD_IMAGE_BYTES: usize = 64 * 1024 * 1024;
const STANDARD_TOTAL_IMAGE_BYTES: usize = 256 * 1024 * 1024;
const STANDARD_PACKAGE_PARTS: usize = 65_536;
const STANDARD_PACKAGE_RELATIONSHIPS: usize = 262_144;
const STANDARD_PART_ALLOCATIONS: usize = 8_192;
const STANDARD_PART_DELETIONS: usize = 8_192;

pub use model::*;
pub use package::*;

#[cfg(test)]
mod tests {
    use super::codec::*;
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::constants::relationship_type as rt;
    use litchi_opc::part::XmlPart;

    const LOCAL_OMEX_EXTENSION: &[u8] =
        include_bytes!("../../../../test-data/ooxml/web_extensions/omex_webextension.xml");
    const LOCAL_REGISTRY_EXTENSION: &[u8] =
        include_bytes!("../../../../test-data/ooxml/web_extensions/registry_webextension.xml");
    const LOCAL_VISIBLE_TASK_PANES: &[u8] =
        include_bytes!("../../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");
    const LOCAL_HIDDEN_TASK_PANES: &[u8] =
        include_bytes!("../../../../test-data/ooxml/web_extensions/hidden_taskpanes.xml");
    const LOCAL_SNAPSHOT_EFFECTS_EXTENSION: &[u8] = include_bytes!(
        "../../../../test-data/ooxml/web_extensions/snapshot_effects_webextension.xml"
    );
    const LOCAL_EXTENSION_LISTS_EXTENSION: &[u8] = include_bytes!(
        "../../../../test-data/ooxml/web_extensions/extension_lists_webextension.xml"
    );
    const LOCAL_EXTENSION_LISTS_TASK_PANES: &[u8] =
        include_bytes!("../../../../test-data/ooxml/web_extensions/extension_lists_taskpanes.xml");

    #[test]
    fn loads_local_omex_and_registry_fixtures_inertly() {
        let omex = local_fixture_package(LOCAL_VISIBLE_TASK_PANES, LOCAL_OMEX_EXTENSION);
        let panes = load(&omex).unwrap().unwrap();
        assert_eq!(panes.panes.len(), 1);
        assert_eq!(panes.panes[0].add_in.reference.store, Store::Omex);
        assert!(panes.panes[0].visible);

        let registry = local_fixture_package(LOCAL_HIDDEN_TASK_PANES, LOCAL_REGISTRY_EXTENSION);
        let panes = load(&registry).unwrap().unwrap();
        assert_eq!(panes.panes[0].add_in.reference.store, Store::Registry);
        assert!(!panes.panes[0].visible);
    }

    #[test]
    fn file_references_require_an_atomic_nonempty_location() {
        let error = Reference::new("Example3", "15.0", Store::FileSystem).unwrap_err();
        assert!(error.to_string().contains("Reference::file"));
        assert!(Reference::file("Example3", "15.0", "").is_err());

        let reference = Reference::file("Example3", "15.0", r"C:\Example").unwrap();
        assert_eq!(reference.store(), Store::FileSystem);
        assert_eq!(reference.location_name(), Some(r"C:\Example"));

        let extension = AddIn::new("plain-instance-id", reference)
            .unwrap()
            .bind(Binding::new("Matrix1", "matrix", "plain-app-ref").unwrap())
            .unwrap();
        let xml = write_add_in(&extension, Conformance::Transitional).unwrap();
        assert!(
            std::str::from_utf8(&xml)
                .unwrap()
                .contains(r#"store="C:\Example" storeType="FileSystem""#)
        );
        assert_eq!(parse_add_in(&xml).unwrap(), extension);
    }

    #[test]
    fn office_safe_parser_rejects_a_storeless_file_reference() {
        let xml = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="plain-instance-id">
                 <we:reference id="Example3" version="15.0" storeType="FileSystem"/>
                 <we:properties/><we:bindings/>
               </we:webextension>"#
        );
        let error = parse_add_in(xml.as_bytes()).unwrap_err();
        assert!(error.to_string().contains("non-empty location"));
    }

    #[test]
    fn strict_writer_is_deterministic_and_round_trips() {
        let extension = sample_extension();
        let first = write_add_in(&extension, Conformance::Strict).unwrap();
        let second = write_add_in(&extension, Conformance::Strict).unwrap();
        assert_eq!(first, second);
        assert!(
            std::str::from_utf8(&first)
                .unwrap()
                .contains(STRICT_RELATIONSHIPS_NAMESPACE)
        );
        assert_eq!(parse_add_in(&first).unwrap(), extension);
    }

    #[test]
    fn snapshot_compression_and_effect_trees_round_trip() {
        let extension = parse_add_in(LOCAL_SNAPSHOT_EFFECTS_EXTENSION).unwrap();
        let snapshot = extension.snapshot.as_ref().unwrap();
        assert_eq!(
            snapshot.compression_state,
            Some(Compression::HighQualityPrint)
        );
        assert_eq!(
            snapshot
                .effects
                .iter()
                .map(Effect::kind)
                .collect::<Vec<_>>(),
            vec![
                EffectKind::AlphaModulateFixed,
                EffectKind::Duotone,
                EffectKind::Blur,
            ]
        );
        assert!(snapshot.effects[1].xml().contains("srgbClr"));

        let written = write_add_in(&extension, Conformance::Strict).unwrap();
        let reparsed = parse_add_in(&written).unwrap();
        assert_eq!(reparsed, extension);
        let written = std::str::from_utf8(&written).unwrap();
        assert!(written.contains("cstate=\"hqprint\""));
        assert!(written.contains(STRICT_RELATIONSHIPS_NAMESPACE));
    }

    #[test]
    fn preserves_all_extension_list_sites_with_inherited_namespaces_and_mixed_content() {
        let extension = parse_add_in(LOCAL_EXTENSION_LISTS_EXTENSION).unwrap();
        let reference_extension = extension.reference.extension_list.as_ref().unwrap();
        assert_eq!(reference_extension.kind(), ExtKind::AddIn);
        assert!(reference_extension.xml().contains("xmlns:vendor="));
        assert!(reference_extension.xml().contains("xmlns:r="));
        assert!(reference_extension.xml().contains("reference text"));
        assert!(reference_extension.xml().contains("<![CDATA[<opaque>]]>"));
        assert!(reference_extension.xml().contains("<!--kept-->"));
        assert!(extension.alternate_references[0].extension_list.is_some());
        assert!(extension.bindings[0].extension_list.is_some());
        assert_eq!(
            extension
                .snapshot
                .as_ref()
                .unwrap()
                .extension_list
                .as_ref()
                .unwrap()
                .kind(),
            ExtKind::DrawingMl
        );
        assert!(extension.extension_list.is_some());

        let written = write_add_in(&extension, Conformance::Strict).unwrap();
        assert_eq!(parse_add_in(&written).unwrap(), extension);

        let panes = parse_panes(LOCAL_EXTENSION_LISTS_TASK_PANES).unwrap();
        let pane_extension = panes[0].extension_list.as_ref().unwrap();
        assert_eq!(pane_extension.kind(), ExtKind::TaskPane);
        assert!(pane_extension.xml().contains("xmlns:vendor="));
        assert!(pane_extension.xml().contains("<![CDATA[<pane-data>]]>"));
        assert!(pane_extension.xml().contains("<!--pane comment-->"));
    }

    #[test]
    fn package_crud_round_trips_every_inert_extension_list() {
        let package = local_fixture_package(
            LOCAL_EXTENSION_LISTS_TASK_PANES,
            LOCAL_EXTENSION_LISTS_EXTENSION,
        );
        let loaded = load(&package).unwrap().unwrap();
        assert!(loaded.panes[0].extension_list.is_some());
        assert!(loaded.panes[0].add_in.extension_list.is_some());

        let mut stored = OpcPackage::new();
        put(&mut stored, loaded.clone(), Conformance::Strict).unwrap();
        assert_eq!(load(&stored).unwrap(), Some(loaded));
    }

    #[test]
    fn authored_extension_lists_validate_namespace_placement_and_security() {
        let web = ExtList::from_xml(
            br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="urn:test"><v:data xmlns:v="urn:test">text<![CDATA[data]]></v:data></a:ext></we:extLst>"#,
        )
        .unwrap();
        assert_eq!(web.kind(), ExtKind::AddIn);
        assert_eq!(ExtList::from_xml(web.as_xml()).unwrap(), web);

        let mut extension = sample_extension();
        extension.extension_list = Some(web.clone());
        assert!(write_add_in(&extension, Conformance::Transitional).is_ok());

        let mut panes = sample_task_panes();
        panes.panes[0].extension_list = Some(web);
        assert!(write_panes(&panes, Conformance::Transitional).is_err());
        assert!(
            ExtList::from_xml(
                br#"<!DOCTYPE extLst [<!ENTITY x "boom">]><we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">&x;</we:extLst>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(br#"<v:extLst xmlns:v="urn:not-an-office-namespace"/>"#).is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" unexpected="1"/>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext/></we:extLst>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:v="urn:test"><v:data/></we:extLst>"#
            )
            .is_err()
        );
        assert!(
            ExtList::from_xml(
                br#"<we:extLst xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">text</we:extLst>"#
            )
            .is_err()
        );
    }

    #[test]
    fn accepts_every_ct_blip_effect_kind_and_rejects_invalid_markup() {
        let names = [
            "alphaBiLevel",
            "alphaCeiling",
            "alphaFloor",
            "alphaInv",
            "alphaMod",
            "alphaModFix",
            "alphaRepl",
            "biLevel",
            "blur",
            "clrChange",
            "clrRepl",
            "duotone",
            "fillOverlay",
            "grayscl",
            "hsl",
            "lum",
            "tint",
        ];
        for name in names {
            let xml = format!(r#"<a:{name} xmlns:a="{DRAWINGML_NAMESPACE}"/>"#);
            let effect = Effect::from_xml(xml.as_bytes()).unwrap();
            assert_eq!(effect.kind().local_name(), name);
            assert_eq!(Effect::from_xml(effect.xml().as_bytes()).unwrap(), effect);
        }

        assert!(
            Effect::from_xml(
                br#"<a:reflection xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .is_err()
        );
        assert!(
            Effect::from_xml(
                br#"<!DOCTYPE x><a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .is_err()
        );
        assert!(
            Effect::from_xml(
                br#"<a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">text</a:blur>"#
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_invalid_snapshot_compression_and_effect_order() {
        let invalid_compression = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:reference id="a" version="1"/><we:properties/><we:bindings/><we:snapshot cstate="lossless"/></we:webextension>"#
        );
        assert!(parse_add_in(invalid_compression.as_bytes()).is_err());

        let misplaced_extension_list = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}" id="x"><we:reference id="a" version="1"/><we:properties/><we:bindings/><we:snapshot><a:extLst/><a:blur/></we:snapshot></we:webextension>"#
        );
        assert!(parse_add_in(misplaced_extension_list.as_bytes()).is_err());
    }

    #[test]
    fn accepts_mce_alternate_content_and_strict_relationship_attributes() {
        let xml = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:r="{STRICT_RELATIONSHIPS_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" id="x"><we:reference id="a" version="1"/><mc:AlternateContent><mc:Choice Requires="we"><we:alternateReferences/></mc:Choice><mc:Fallback/></mc:AlternateContent><we:properties/><we:bindings/><we:snapshot r:embed="rId1"/></we:webextension>"#
        );
        let extension = parse_add_in(xml.as_bytes()).unwrap();
        assert_eq!(
            extension
                .snapshot
                .unwrap()
                .embedded_relationship_id
                .as_deref(),
            Some("rId1")
        );
    }

    #[test]
    fn rejects_dtd_bad_order_bad_store_and_nonfinite_width() {
        assert!(parse_add_in(br#"<!DOCTYPE x><x/>"#).is_err());
        let bad_order = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:properties/><we:reference id="a" version="1"/><we:bindings/></we:webextension>"#
        );
        assert!(parse_add_in(bad_order.as_bytes()).is_err());
        let bad_store = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" id="x"><we:reference id="a" version="1" storeType="Network"/><we:properties/><we:bindings/></we:webextension>"#
        );
        assert!(parse_add_in(bad_store.as_bytes()).is_err());
        let bad_width = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="NaN" row="0"><wetp:webextensionref r:id="rId1"/></wetp:taskpane></wetp:taskpanes>"#
        );
        assert!(parse_panes(bad_width.as_bytes()).is_err());
        let obsolete_float = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="1" width="320" row="0"><wetp:webextensionref r:id="rId1"/><wetp:float/></wetp:taskpane></wetp:taskpanes>"#
        );
        assert!(parse_panes(obsolete_float.as_bytes()).is_err());
    }

    #[test]
    fn enforces_input_and_list_caps() {
        assert!(parse_add_in(&vec![b' '; MAX_WEB_EXTENSION_XML_BYTES + 1]).is_err());
        let mut model = Panes::default();
        model
            .panes
            .resize_with(MAX_WEB_EXTENSION_ITEMS + 1, || Pane {
                dock_state: Dock::Right,
                visible: false,
                width: 320.0,
                row: 0,
                locked: false,
                relationship_id: "rId1".into(),
                add_in: sample_extension(),
                snapshot_resources: vec![],
                extension_list: None,
            });
        assert!(write_panes(&model, Conformance::Transitional).is_err());
        let mut extension = sample_extension();
        extension.id = "x".repeat(MAX_WEB_EXTENSION_XML_BYTES);
        assert!(write_add_in(&extension, Conformance::Transitional).is_err());

        let mut excessive_nodes = format!(
            r#"<we:extLst xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}" xmlns:v="urn:test"><a:ext uri="urn:test"><v:data>"#
        );
        excessive_nodes.push_str(&"<v:n/>".repeat(MAX_WEB_EXTENSION_XML_NODES));
        excessive_nodes.push_str("</v:data></a:ext></we:extLst>");
        assert!(ExtList::from_xml(excessive_nodes.as_bytes()).is_err());
    }

    #[test]
    fn rejects_external_wrong_content_type_and_dangling_package_graphs() {
        let external = synthetic_package(true, ADD_IN_CONTENT_TYPE, "rId1");
        assert!(load(&external).is_err());

        let wrong_type = synthetic_package(false, "application/xml", "rId1");
        assert!(matches!(load(&wrong_type), Err(Error::ContentType { .. })));

        let dangling = synthetic_package(false, ADD_IN_CONTENT_TYPE, "missing");
        assert!(load(&dangling).is_err());
    }

    #[test]
    fn package_crud_round_trips_embedded_and_linked_snapshots() {
        let mut package = OpcPackage::new();
        let authored = sample_task_panes();
        put(&mut package, authored.clone(), Conformance::Transitional).unwrap();
        assert_eq!(load(&package).unwrap(), Some(authored.clone()));

        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let extension_name = package
            .get_part(&task_panes_name)
            .unwrap()
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap();
        let extension = package.get_part(&extension_name).unwrap();
        assert!(!extension.rels().get("rIdSnapshot").unwrap().is_external());
        assert!(extension.rels().get("rIdLinked").unwrap().is_external());

        let mut replacement = authored;
        replacement.panes[0].add_in.snapshot = None;
        replacement.panes[0].snapshot_resources.clear();
        replacement.panes[0].visible = false;
        put(&mut package, replacement.clone(), Conformance::Strict).unwrap();
        assert_eq!(load(&package).unwrap(), Some(replacement));
        assert!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .is_err()
        );

        assert!(remove(&mut package).unwrap());
        assert!(load(&package).unwrap().is_none());
        assert!(!remove(&mut package).unwrap());
        assert_eq!(package.part_count(), 0);
    }

    #[test]
    fn byte_identical_put_is_a_signature_preserving_no_op() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(package.is_signed());

        let loaded = load(&package).unwrap().unwrap();
        let patch = plan_put(&package, loaded.clone(), Conformance::Transitional).unwrap();
        assert!(patch.is_empty());
        assert!(patch.inverse().is_empty());
        assert!(!patch.apply(&mut package).unwrap());
        assert_eq!(
            format!("{patch:?}"),
            "Patch { empty: true, affected_parts: 0, .. }"
        );
        put(&mut package, loaded, Conformance::Transitional).unwrap();

        assert!(package.is_signed());
    }

    #[test]
    fn put_patch_is_exact_reversible_and_arc_shared() {
        let mut package = OpcPackage::new();
        let original = sample_task_panes();
        put(&mut package, original.clone(), Conformance::Transitional).unwrap();
        let image_name = PackURI::new("/media/web-extension-snapshot.png").unwrap();
        let original_image = package.get_part(&image_name).unwrap().blob_arc();
        let original_root = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .map(RelationshipState::capture)
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);

        let replacement_image = Arc::new(vec![1, 2, 3, 4, 5, 6]);
        let mut replacement = load(&package).unwrap().unwrap();
        assert!(
            replacement
                .edit(0usize, |pane| {
                    pane.set_visible(false);
                    pane.set_image(
                        image_name.as_str(),
                        "image/png",
                        Arc::clone(&replacement_image),
                    )?;
                    Ok(())
                })
                .unwrap()
        );

        let patch = plan_put(&package, replacement.clone(), Conformance::Transitional).unwrap();
        assert!(!patch.is_empty());
        let image_change = patch
            .parts
            .iter()
            .find(|part| part.name == image_name)
            .unwrap();
        assert!(Arc::ptr_eq(
            &image_change.after.as_ref().unwrap().data,
            &replacement_image,
        ));
        let inverse = patch.inverse();
        let inverse_image = inverse
            .parts
            .iter()
            .find(|part| part.name == image_name)
            .unwrap();
        assert!(Arc::ptr_eq(
            &image_change.before.as_ref().unwrap().data,
            &inverse_image.after.as_ref().unwrap().data,
        ));
        assert!(!format!("{patch:?}").contains("rId"));
        assert!(!format!("{patch:?}").contains("/webextensions"));

        assert!(patch.apply(&mut package).unwrap());
        assert!(!package.is_signed());
        let applied = load(&package).unwrap().unwrap();
        assert!(!applied.get(0usize).unwrap().visible());
        assert_eq!(
            applied.get(0usize).unwrap().image().unwrap().bytes(),
            replacement_image.as_slice(),
        );
        assert_eq!(
            applied.get(0usize).unwrap().link().unwrap().external(),
            Some("https://example.invalid/inert-snapshot.png"),
        );
        assert!(Arc::ptr_eq(
            &package.get_part(&image_name).unwrap().blob_arc(),
            &replacement_image,
        ));

        assert!(inverse.apply(&mut package).unwrap());
        let restored = load(&package).unwrap().unwrap();
        assert!(restored.get(0usize).unwrap().visible());
        assert_eq!(
            restored.get(0usize).unwrap().add_in(),
            original.get(0usize).unwrap().add_in(),
        );
        assert_eq!(
            restored.get(0usize).unwrap().image().unwrap().bytes(),
            original.get(0usize).unwrap().image().unwrap().bytes(),
        );
        assert!(Arc::ptr_eq(
            &package.get_part(&image_name).unwrap().blob_arc(),
            &original_image,
        ));
        let restored_root = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap();
        assert!(original_root.matches(restored_root));
    }

    #[test]
    fn remove_patch_restores_exact_shared_parts() {
        let mut package = OpcPackage::new();
        let original = sample_task_panes();
        put(&mut package, original.clone(), Conformance::Transitional).unwrap();
        let image_name = PackURI::new("/media/web-extension-snapshot.png").unwrap();
        let original_image = package.get_part(&image_name).unwrap().blob_arc();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);

        let patch = plan_remove(&package).unwrap();
        let inverse = patch.inverse();
        let restored_image = inverse
            .parts
            .iter()
            .find(|part| part.name == image_name)
            .and_then(|part| part.after.as_ref())
            .unwrap();
        assert!(Arc::ptr_eq(&restored_image.data, &original_image));

        assert!(patch.apply(&mut package).unwrap());
        assert!(!package.is_signed());
        assert!(load(&package).unwrap().is_none());
        assert!(package.get_part(&image_name).is_err());

        assert!(inverse.apply(&mut package).unwrap());
        assert_eq!(load(&package).unwrap(), Some(original));
        assert!(Arc::ptr_eq(
            &package.get_part(&image_name).unwrap().blob_arc(),
            &original_image,
        ));
    }

    #[test]
    fn patch_rejects_changed_root_and_part_before_mutation() {
        let mut root_changed = OpcPackage::new();
        put(
            &mut root_changed,
            sample_task_panes(),
            Conformance::Transitional,
        )
        .unwrap();
        root_changed.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let mut replacement = load(&root_changed).unwrap().unwrap();
        replacement.panes[0].visible = false;
        let patch = plan_put(&root_changed, replacement, Conformance::Transitional).unwrap();
        let root = root_changed
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .map(RelationshipState::capture)
            .unwrap();
        root_changed.rels_mut().remove(&root.id);
        root_changed.rels_mut().add_relationship(
            root.relationship_type,
            format!("{}#changed", root.target),
            root.id,
            root.external,
        );
        let task_name = PackURI::new("/webextensions/taskpanes.xml").unwrap();
        let task_before = root_changed.get_part(&task_name).unwrap().blob_arc();
        assert!(patch.apply(&mut root_changed).is_err());
        assert!(root_changed.is_signed());
        assert!(Arc::ptr_eq(
            &root_changed.get_part(&task_name).unwrap().blob_arc(),
            &task_before,
        ));

        let mut part_changed = OpcPackage::new();
        put(
            &mut part_changed,
            sample_task_panes(),
            Conformance::Transitional,
        )
        .unwrap();
        part_changed.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let mut replacement = load(&part_changed).unwrap().unwrap();
        replacement.panes[0].visible = false;
        let patch = plan_put(&part_changed, replacement, Conformance::Transitional).unwrap();
        let image_name = PackURI::new("/media/web-extension-snapshot.png").unwrap();
        part_changed
            .get_part_mut(&image_name)
            .unwrap()
            .set_blob(vec![9, 8, 7]);
        let task_before = part_changed.get_part(&task_name).unwrap().blob_arc();
        assert!(patch.apply(&mut part_changed).is_err());
        assert!(part_changed.is_signed());
        assert!(Arc::ptr_eq(
            &part_changed.get_part(&task_name).unwrap().blob_arc(),
            &task_before,
        ));
        assert_eq!(
            part_changed.get_part(&image_name).unwrap().blob(),
            &[9, 8, 7]
        );
    }

    #[test]
    fn patch_rechecks_shared_graph_protection_before_mutation() {
        let mut stale = OpcPackage::new();
        put(&mut stale, sample_task_panes(), Conformance::Transitional).unwrap();
        stale.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let patch = plan_remove(&stale).unwrap();
        let task_name = add_shared_task_pane_ingress(&mut stale);
        let before_parts = stale.part_count();

        assert!(patch.apply(&mut stale).is_err());
        assert!(stale.is_signed());
        assert_eq!(stale.part_count(), before_parts);
        assert!(stale.get_part(&task_name).is_ok());
        assert!(load(&stale).unwrap().is_some());

        let mut shared_before_planning = OpcPackage::new();
        put(
            &mut shared_before_planning,
            sample_task_panes(),
            Conformance::Transitional,
        )
        .unwrap();
        let task_name = add_shared_task_pane_ingress(&mut shared_before_planning);
        assert!(remove(&mut shared_before_planning).unwrap());
        assert!(load(&shared_before_planning).unwrap().is_none());
        assert!(shared_before_planning.get_part(&task_name).is_ok());
    }

    #[test]
    fn patch_rejects_new_inbound_relationship_to_absent_destination() {
        let mut package = OpcPackage::new();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let patch = plan_put(&package, sample_task_panes(), Conformance::Transitional).unwrap();
        package.rels_mut().add_relationship(
            "urn:litchi:test:new-shared-target".into(),
            "media/web-extension-snapshot.png".into(),
            "rIdNewSharedTarget".into(),
            false,
        );

        assert!(patch.apply(&mut package).is_err());
        assert!(package.is_signed());
        assert_eq!(package.part_count(), 0);
        assert!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .is_err()
        );
        assert!(
            package
                .rels()
                .get("rIdNewSharedTarget")
                .is_some_and(|relationship| {
                    relationship.reltype() == "urn:litchi:test:new-shared-target"
                })
        );
    }

    #[test]
    fn changed_shared_task_pane_part_is_rejected_without_mutation() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let target = task_panes_name.as_str().trim_start_matches('/').to_owned();
        package.rels_mut().add_relationship(
            "urn:litchi:test:shared-task-panes".into(),
            target,
            "rIdSharedTaskPanes".into(),
            false,
        );
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let before = package.get_part(&task_panes_name).unwrap().blob().to_vec();
        let mut changed = load(&package).unwrap().unwrap();
        changed.panes[0].visible = false;

        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&task_panes_name).unwrap().blob(), before);
        assert!(package.is_signed());
        assert!(load(&package).unwrap().unwrap().panes[0].visible);
    }

    #[test]
    fn shared_task_pane_ingress_protects_descendant_add_in() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_panes_name = add_shared_task_pane_ingress(&mut package);
        let extension_name = package
            .get_part(&task_panes_name)
            .unwrap()
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap();
        let before = package.get_part(&extension_name).unwrap().blob().to_vec();
        let mut changed = load(&package).unwrap().unwrap();
        assert!(
            changed
                .edit(0usize, |pane| {
                    pane.add_in_mut().set_frozen(false);
                    Ok(())
                })
                .unwrap()
        );

        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&extension_name).unwrap().blob(), before);
        assert!(
            load(&package)
                .unwrap()
                .unwrap()
                .get(0usize)
                .unwrap()
                .add_in()
                .is_frozen()
        );
    }

    #[test]
    fn shared_task_pane_ingress_protects_descendant_image() {
        let mut package = OpcPackage::new();
        put(&mut package, sample_task_panes(), Conformance::Transitional).unwrap();
        add_shared_task_pane_ingress(&mut package);
        let image_name = PackURI::new("/media/web-extension-snapshot.png").unwrap();
        let before = package.get_part(&image_name).unwrap().blob().to_vec();
        let mut changed = load(&package).unwrap().unwrap();
        assert!(
            changed
                .edit(0usize, |pane| {
                    pane.set_image(
                        "/media/web-extension-snapshot.png",
                        "image/png",
                        Arc::new(vec![1, 2, 3, 4]),
                    )?;
                    Ok(())
                })
                .unwrap()
        );

        assert!(put(&mut package, changed, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&image_name).unwrap().blob(), before);
    }

    #[test]
    fn internal_web_relationship_targets_reject_queries_and_fragments() {
        let mut root = OpcPackage::new();
        put(&mut root, sample_task_panes(), Conformance::Transitional).unwrap();
        let root_id = root
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .r_id()
            .to_owned();
        root.rels_mut().remove(&root_id);
        root.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP.into(),
            "webextensions/taskpanes.xml?version=1".into(),
            root_id,
            false,
        );
        assert!(matches!(load(&root), Err(Error::Relationship(_))));

        let mut add_in = OpcPackage::new();
        put(&mut add_in, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_name = add_in
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let task_part = add_in.get_part_mut(&task_name).unwrap();
        task_part.rels_mut().remove("rId1");
        task_part.rels_mut().add_relationship(
            ADD_IN_RELATIONSHIP.into(),
            "webextension1.xml#instance".into(),
            "rId1".into(),
            false,
        );
        assert!(matches!(load(&add_in), Err(Error::Relationship(_))));

        let mut image = OpcPackage::new();
        put(&mut image, sample_task_panes(), Conformance::Transitional).unwrap();
        let task_name = image
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        let add_in_name = image
            .get_part(&task_name)
            .unwrap()
            .rels()
            .get("rId1")
            .unwrap()
            .target_partname()
            .unwrap();
        let add_in_part = image.get_part_mut(&add_in_name).unwrap();
        add_in_part.rels_mut().remove("rIdSnapshot");
        add_in_part.rels_mut().add_relationship(
            IMAGE_RELATIONSHIP_TYPE.into(),
            "../media/web-extension-snapshot.png?size=large".into(),
            "rIdSnapshot".into(),
            false,
        );
        assert!(matches!(load(&image), Err(Error::Relationship(_))));
    }

    #[test]
    fn case_equivalent_existing_parts_are_rejected_before_put_mutates() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/Data.bin").unwrap(),
            "application/octet-stream".into(),
            vec![1],
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/CUSTOM/data.bin").unwrap(),
            "application/octet-stream".into(),
            vec![2],
        )));
        let before = package.part_count();
        assert!(put(&mut package, sample_task_panes(), Conformance::Transitional,).is_err());
        assert_eq!(package.part_count(), before);
        assert_eq!(package.rels().len(), 0);
    }

    #[test]
    fn package_store_rejects_resource_mismatches_without_mutation() {
        let mut package = OpcPackage::new();
        let mut malformed = sample_task_panes();
        malformed.panes[0].snapshot_resources.pop();
        assert!(put(&mut package, malformed, Conformance::Transitional,).is_err());
        assert_eq!(package.part_count(), 0);
        assert_eq!(package.rels().iter().count(), 0);

        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/media/web-extension-snapshot.png").unwrap(),
            "image/png".into(),
            vec![9, 9, 9],
        )));
        assert!(put(&mut package, sample_task_panes(), Conformance::Transitional,).is_err());
        assert_eq!(package.part_count(), 1);
        assert_eq!(
            package
                .get_part(&PackURI::new("/media/web-extension-snapshot.png").unwrap())
                .unwrap()
                .blob(),
            &[9, 9, 9]
        );
    }

    #[test]
    fn semantic_facade_authors_selects_and_removes_without_raw_ids() {
        let reference = Reference::new("wa1", "1.0.0.0", Store::Omex).unwrap();
        let binding = Binding::new("binding-1", BindingKind::Matrix, "app-ref").unwrap();
        let add_in = AddIn::new("add-in-1", reference)
            .unwrap()
            .bind(binding)
            .unwrap();
        let bytes = Arc::new(vec![0x89, b'P', b'N', b'G']);
        let effect = Effect::from_xml(
            br#"<a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rad="1000"/>"#,
        )
        .unwrap();
        let pane = Pane::new(add_in)
            .show(false)
            .dock(Dock::Left)
            .unwrap()
            .width(420.0)
            .unwrap()
            .embed("/media/add-in-preview.png", "image/png", Arc::clone(&bytes))
            .unwrap()
            .linked("https://example.invalid/inert.png#preview")
            .unwrap()
            .compress(Compression::Print)
            .effect(effect)
            .unwrap();

        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        assert_eq!(panes.get("add-in-1").unwrap().dock_kind(), &Dock::Left);
        assert!(!panes.get(0usize).unwrap().visible());
        assert!(panes.get(1usize).is_none());
        let image = panes.get("add-in-1").unwrap().image().unwrap();
        assert!(Arc::ptr_eq(&bytes, &image.shared()));
        assert_eq!(image.content_type(), "image/png");
        assert_eq!(
            panes.get("add-in-1").unwrap().link().unwrap().external(),
            Some("https://example.invalid/inert.png#preview")
        );
        assert!(panes.remove(99usize).is_none());
        assert_eq!(panes.remove("add-in-1").unwrap().add_in().id(), "add-in-1");
        assert!(panes.is_empty());
    }

    #[test]
    fn panes_push_rekeys_colliding_hidden_relationship_ids() {
        let first_add_in = AddIn::new(
            "add-in-1",
            Reference::new("ref-1", "1", Store::Omex).unwrap(),
        )
        .unwrap();
        let second_add_in = AddIn::new(
            "add-in-2",
            Reference::new("ref-2", "1", Store::Registry).unwrap(),
        )
        .unwrap();
        let mut first_source = Panes::new();
        first_source.push(Pane::new(first_add_in)).unwrap();
        let mut second_source = Panes::new();
        second_source.push(Pane::new(second_add_in)).unwrap();
        let first = first_source.remove(0usize).unwrap();
        let second = second_source.remove(0usize).unwrap();
        let mut panes = Panes::new();
        panes.push(first).unwrap();
        panes.push(second).unwrap();

        assert_eq!(panes.len(), 2);
        assert_ne!(
            panes.panes[0].relationship_id,
            panes.panes[1].relationship_id
        );
        assert!(!panes.panes[1].relationship_id.is_empty());
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
        assert_eq!(load(&package).unwrap().unwrap().len(), 2);
    }

    #[test]
    fn panes_push_canonicalizes_equivalent_resources_within_one_pane() {
        let bytes = Arc::new(vec![1, 2, 3]);
        let pane = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", Arc::clone(&bytes))
        .unwrap()
        .linked_image("/MEDIA/preview.png", "image/png", Arc::clone(&bytes))
        .unwrap();

        let mut panes = Panes::new();
        panes.push(pane).unwrap();

        let pane = panes.get("add-in-1").unwrap();
        let embedded = pane.image().unwrap();
        let linked = pane.link().unwrap().internal().unwrap();
        assert_eq!(embedded.name(), linked.name());
        assert_eq!(embedded.name().as_str(), "/media/Preview.png");
        assert!(Arc::ptr_eq(&embedded.shared(), &linked.shared()));
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
    }

    #[test]
    fn panes_push_rejects_conflicting_resources_within_one_pane_atomically() {
        let pane = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", vec![1, 2, 3])
        .unwrap()
        .linked_image("/MEDIA/preview.png", "image/png", vec![4, 5, 6])
        .unwrap();

        let mut panes = Panes::new();
        assert!(panes.push(pane).is_err());
        assert!(panes.is_empty());
    }

    #[test]
    fn panes_edit_is_checked_canonical_and_transactional() {
        let first = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", vec![1, 2, 3])
        .unwrap();
        let second = Pane::new(
            AddIn::new(
                "add-in-2",
                Reference::new("ref-2", "1", Store::Registry).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/other.png", "image/png", vec![4, 5, 6])
        .unwrap();
        let mut panes = Panes::new();
        panes.push(first).unwrap().push(second).unwrap();
        let before = panes.get("add-in-2").unwrap().clone();

        assert!(
            panes
                .edit("add-in-2", |pane| {
                    pane.set_visible(false);
                    pane.set_image("/MEDIA/preview.png", "image/png", vec![9])?;
                    Ok(())
                })
                .is_err()
        );
        assert_eq!(panes.get("add-in-2"), Some(&before));

        let shared = panes.get("add-in-1").unwrap().image().unwrap().shared();
        assert!(
            panes
                .edit("add-in-2", |pane| {
                    pane.set_visible(false);
                    pane.set_image("/MEDIA/preview.png", "image/png", Arc::clone(&shared))?;
                    Ok(())
                })
                .unwrap()
        );
        let edited = panes.get("add-in-2").unwrap();
        assert!(!edited.visible());
        assert_eq!(
            edited.image().unwrap().name().as_str(),
            "/media/Preview.png"
        );
        assert!(Arc::ptr_eq(&shared, &edited.image().unwrap().shared()));

        let mut invoked = false;
        assert!(
            !panes
                .edit(99usize, |_| {
                    invoked = true;
                    Ok(())
                })
                .unwrap()
        );
        assert!(!invoked);
    }

    #[test]
    fn conflicting_case_equivalent_snapshot_resources_are_rejected_atomically() {
        let first = Pane::new(
            AddIn::new(
                "add-in-1",
                Reference::new("ref-1", "1", Store::Omex).unwrap(),
            )
            .unwrap(),
        )
        .embed("/media/Preview.png", "image/png", vec![1, 2, 3])
        .unwrap();
        let second = Pane::new(
            AddIn::new(
                "add-in-2",
                Reference::new("ref-2", "1", Store::Registry).unwrap(),
            )
            .unwrap(),
        )
        .embed("/MEDIA/preview.png", "image/png", vec![4, 5, 6])
        .unwrap();
        let mut first_source = Panes::new();
        first_source.push(first).unwrap();
        let mut second_source = Panes::new();
        second_source.push(second).unwrap();
        let first = first_source.remove(0usize).unwrap();
        let second = second_source.remove(0usize).unwrap();

        let mut panes = Panes::new();
        panes.push(first).unwrap();
        assert!(panes.push(second).is_err());
        assert_eq!(panes.len(), 1);
        assert_eq!(
            panes.get("add-in-1").unwrap().image().unwrap().bytes(),
            &[1, 2, 3]
        );
    }

    #[test]
    fn absent_snapshot_metadata_no_ops_do_not_create_a_snapshot() {
        let add_in = AddIn::new(
            "add-in-1",
            Reference::new("ref-1", "1", Store::Omex).unwrap(),
        )
        .unwrap();
        let mut pane = Pane::new(add_in);
        assert!(pane.add_in().snapshot().is_none());

        assert!(!pane.clear_compression());
        assert!(pane.add_in().snapshot().is_none());

        let effect = Effect::from_xml(
            br#"<a:blur xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        )
        .unwrap();
        assert!(pane.replace_effect(0, effect).unwrap().is_none());
        assert!(pane.add_in().snapshot().is_none());
    }

    #[test]
    fn semantic_snapshot_authoring_rejects_bad_mime_and_uri_reference() {
        let reference = Reference::new("wa1", "1", Store::Registry).unwrap();
        let add_in = AddIn::new("add-in-1", reference).unwrap();
        assert!(
            Pane::new(add_in.clone())
                .embed("/media/a.bin", "image", vec![1, 2, 3])
                .is_err()
        );
        assert!(
            Pane::new(add_in.clone())
                .embed("/media/a.bin", "image/png; charset=binary", vec![1])
                .is_err()
        );
        assert!(Pane::new(add_in.clone()).linked("bad target").is_err());
        assert!(
            Pane::new(add_in)
                .linked("https://example.invalid/%GG")
                .is_err()
        );
    }

    #[test]
    fn internal_linked_image_is_typed_and_round_trips() {
        let reference = Reference::new("wa1", "1", Store::Omex).unwrap();
        let add_in = AddIn::new("add-in-1", reference).unwrap();
        let pane = Pane::new(add_in)
            .linked_image(
                "/media/linked-preview.png",
                "image/png",
                Arc::new(vec![1, 2, 3, 4]),
            )
            .unwrap();
        assert_eq!(
            pane.link().unwrap().internal().unwrap().bytes(),
            &[1, 2, 3, 4]
        );
        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
        let mut loaded = load(&package).unwrap().unwrap();
        assert_eq!(
            loaded
                .get(0usize)
                .unwrap()
                .link()
                .unwrap()
                .internal()
                .unwrap()
                .bytes(),
            &[1, 2, 3, 4]
        );
        assert!(
            loaded
                .edit(0usize, |pane| {
                    pane.set_external_link("https://example.invalid/inert.png")?;
                    Ok(())
                })
                .unwrap()
        );
        assert_eq!(
            loaded.get(0usize).unwrap().link().unwrap().external(),
            Some("https://example.invalid/inert.png")
        );
    }

    #[test]
    fn checked_update_crud_covers_collections_metadata_and_all_ext_sites() {
        let add_ext = ExtList::from_xml(
            format!(r#"<we:extLst xmlns:we="{WEB_EXTENSION_NAMESPACE}"/>"#).as_bytes(),
        )
        .unwrap();
        let pane_ext = ExtList::from_xml(
            format!(r#"<wetp:extLst xmlns:wetp="{TASK_PANES_NAMESPACE}"/>"#).as_bytes(),
        )
        .unwrap();
        let drawing_ext =
            ExtList::from_xml(format!(r#"<a:extLst xmlns:a="{DRAWINGML_NAMESPACE}"/>"#).as_bytes())
                .unwrap();

        let reference = Reference::new("primary", "1", Store::Omex).unwrap();
        let mut add_in = AddIn::new("add-in-1", reference).unwrap();
        add_in
            .set_reference(Reference::new("primary-2", "2", Store::Registry).unwrap())
            .unwrap();
        add_in
            .push_reference(Reference::file("alternate", "1", r"C:\AddIns").unwrap())
            .unwrap();
        add_in
            .upsert_reference(Reference::new("alternate", "2", Store::Registry).unwrap())
            .unwrap();
        assert_eq!(
            add_in.alternate_reference("alternate").unwrap().version(),
            "2"
        );
        assert!(add_in.remove_reference(9usize).is_none());

        add_in
            .push_property(Property::new("mode", "old").unwrap())
            .unwrap();
        add_in
            .upsert_property(Property::new("mode", "new").unwrap())
            .unwrap();
        assert_eq!(add_in.property("mode").unwrap().value(), "new");
        assert!(add_in.remove_property(4usize).is_none());

        add_in
            .push_binding(Binding::new("binding", BindingKind::Matrix, "app-ref").unwrap())
            .unwrap();
        add_in
            .upsert_binding(Binding::new("binding", BindingKind::Table, "app-ref-2").unwrap())
            .unwrap();
        assert_eq!(
            add_in.binding("binding").unwrap().kind(),
            &BindingKind::Table
        );
        assert!(add_in.remove_binding(7usize).is_none());

        add_in.reference_mut().set_ext(add_ext.clone()).unwrap();
        assert!(add_in.reference_mut().clear_ext().is_some());
        add_in.reference_mut().set_ext(add_ext.clone()).unwrap();
        add_in
            .binding_mut("binding")
            .unwrap()
            .set_ext(add_ext.clone())
            .unwrap();
        add_in.set_ext(add_ext.clone()).unwrap();
        assert!(add_in.set_ext(pane_ext.clone()).is_err());

        let mut pane = Pane::new(add_in);
        pane.set_visible(false).set_row(3).set_locked(true);
        pane.set_width(480.0)
            .unwrap()
            .set_dock(Dock::Bottom)
            .unwrap();
        pane.snapshot_mut().set_ext(drawing_ext).unwrap();
        pane.set_ext(pane_ext).unwrap();
        assert!(!pane.visible());
        assert_eq!(pane.row(), 3);
        assert!(pane.locked());
        assert_eq!(pane.pane_width(), 480.0);
        assert_eq!(pane.dock_kind(), &Dock::Bottom);
        assert!(pane.add_in().reference().ext().is_some());
        assert!(pane.add_in().binding("binding").unwrap().ext().is_some());
        assert!(pane.add_in().ext().is_some());
        assert!(pane.add_in().snapshot().unwrap().ext().is_some());
        assert!(pane.ext().is_some());

        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        let mut package = OpcPackage::new();
        put(&mut package, panes, Conformance::Transitional).unwrap();
        assert!(load(&package).unwrap().is_some());
    }

    #[test]
    fn package_graph_limits_bound_index_allocation_relationships_and_deletion() {
        let mut authored = OpcPackage::new();
        let no_allocations = Limits {
            part_allocations: 0,
            ..Limits::standard()
        };
        assert!(
            put_with(
                &mut authored,
                sample_task_panes(),
                Conformance::Transitional,
                &no_allocations,
            )
            .is_err()
        );
        assert_eq!(authored.part_count(), 0);

        put(
            &mut authored,
            sample_task_panes(),
            Conformance::Transitional,
        )
        .unwrap();
        let no_parts = Limits {
            package_parts: 0,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&authored, &no_parts),
            Err(Error::Limit { .. })
        ));
        let no_relationships = Limits {
            package_relationships: 0,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&authored, &no_relationships),
            Err(Error::Limit { .. })
        ));
        let no_deletions = Limits {
            part_deletions: 0,
            ..Limits::standard()
        };
        assert!(remove_with(&mut authored, &no_deletions).is_err());
        assert!(load(&authored).unwrap().is_some());
    }

    #[test]
    fn absent_feature_skips_the_bounded_full_graph_index() {
        let sentinel_name = PackURI::new("/custom/sentinel.bin").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            sentinel_name.clone(),
            "application/octet-stream".into(),
            vec![1, 2, 3],
        )));
        let limits = Limits {
            package_parts: 0,
            ..Limits::standard()
        };

        assert!(load_with(&package, &limits).unwrap().is_none());
        assert!(!remove_with(&mut package, &limits).unwrap());
        assert_eq!(package.get_part(&sentinel_name).unwrap().blob(), &[1, 2, 3]);
    }

    #[test]
    fn part_name_allocation_probes_are_operation_wide() {
        let occupied_name = PackURI::new("/webextensions/taskpanes.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            occupied_name.clone(),
            "application/octet-stream".into(),
            vec![7],
        )));
        let mut panes = Panes::new();
        panes
            .push(Pane::new(
                AddIn::new(
                    "add-in-1",
                    Reference::new("ref-1", "1", Store::Omex).unwrap(),
                )
                .unwrap(),
            ))
            .unwrap();
        let limits = Limits {
            part_allocations: 2,
            ..Limits::standard()
        };

        assert!(matches!(
            put_with(&mut package, panes, Conformance::Transitional, &limits,),
            Err(Error::Limit { .. })
        ));
        assert_eq!(package.part_count(), 1);
        assert_eq!(package.get_part(&occupied_name).unwrap().blob(), &[7]);
    }

    #[test]
    fn aggregate_xml_and_retained_string_budgets_bound_put_and_load() {
        let panes = sample_task_panes();
        let task_xml = write_panes(&panes, Conformance::Transitional).unwrap();
        let add_in_xml = write_add_in(
            panes.get(0usize).unwrap().add_in(),
            Conformance::Transitional,
        )
        .unwrap();
        let combined_xml = task_xml.len().checked_add(add_in_xml.len()).unwrap();

        let xml_tight = Limits {
            total_xml_bytes: combined_xml - 1,
            ..Limits::standard()
        };
        let mut rejected = OpcPackage::new();
        assert!(matches!(
            put_with(
                &mut rejected,
                panes.clone(),
                Conformance::Transitional,
                &xml_tight,
            ),
            Err(Error::Limit { .. })
        ));
        assert_eq!(rejected.part_count(), 0);

        let strings_tight = Limits {
            total_string_bytes: combined_xml - 1,
            ..Limits::standard()
        };
        assert!(matches!(
            put_with(
                &mut rejected,
                panes.clone(),
                Conformance::Transitional,
                &strings_tight,
            ),
            Err(Error::Limit { .. })
        ));
        assert_eq!(rejected.part_count(), 0);

        let mut stored = OpcPackage::new();
        put(&mut stored, panes, Conformance::Transitional).unwrap();
        assert!(matches!(
            load_with(&stored, &xml_tight),
            Err(Error::Limit { .. })
        ));
        assert!(matches!(
            load_with(&stored, &strings_tight),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn inherited_namespace_expansion_is_charged_before_fragment_retention() {
        let mut declarations = String::new();
        for index in 0..32 {
            declarations.push_str(&format!(
                " xmlns:v{index}=\"urn:litchi:{}\"",
                "n".repeat(128)
            ));
        }
        let mut bindings = String::new();
        for index in 0..32 {
            bindings.push_str(&format!(
                r#"<we:binding id="b{index}" type="table" appref="a{index}"><we:extLst><a:ext uri="urn:e"><v0:data/></a:ext></we:extLst></we:binding>"#
            ));
        }
        let xml = format!(
            r#"<we:webextension xmlns:we="{WEB_EXTENSION_NAMESPACE}" xmlns:a="{DRAWINGML_NAMESPACE}"{declarations} id="x"><we:reference id="r" version="1"/><we:properties/><we:bindings>{bindings}</we:bindings></we:webextension>"#
        );
        let limits = Limits {
            string_bytes: 32 * 1024,
            ..Limits::standard()
        };

        assert!(matches!(
            parse_add_in_with(xml.as_bytes(), &limits),
            Err(Error::Limit { .. })
        ));
        assert_eq!(parse_add_in(xml.as_bytes()).unwrap().bindings().len(), 32);
    }

    #[test]
    fn package_and_authored_relationship_metadata_share_the_string_budget() {
        let mut stored = local_fixture_package(LOCAL_VISIBLE_TASK_PANES, LOCAL_OMEX_EXTENSION);
        stored.rels_mut().add_relationship(
            "urn:litchi:test:opaque".into(),
            format!("https://example.invalid/{}", "x".repeat(16 * 1024)),
            "rIdOpaque".into(),
            true,
        );
        let tight = Limits {
            total_string_bytes: 32 * 1024,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&stored, &tight),
            Err(Error::Limit { .. })
        ));

        let add_in = AddIn::new(
            "add-in-1",
            Reference::new("ref-1", "1", Store::Omex).unwrap(),
        )
        .unwrap();
        let pane = Pane::new(add_in)
            .linked(format!("https://example.invalid/{}", "y".repeat(16 * 1024)))
            .unwrap();
        let mut panes = Panes::new();
        panes.push(pane).unwrap();
        let mut package = OpcPackage::new();
        assert!(matches!(
            put_with(&mut package, panes, Conformance::Transitional, &tight,),
            Err(Error::Limit { .. })
        ));
        assert_eq!(package.part_count(), 0);
    }

    #[test]
    fn explicit_limits_bound_both_put_and_load() {
        let reference = Reference::new("wa1", "1", Store::Omex).unwrap();
        let add_in = AddIn::new("add-in-1", reference).unwrap();
        let mut panes = Panes::new();
        panes.push(Pane::new(add_in)).unwrap();
        let tight = Limits {
            xml_bytes: 128,
            ..Limits::standard()
        };

        let mut rejected = OpcPackage::new();
        assert!(
            put_with(
                &mut rejected,
                panes.clone(),
                Conformance::Transitional,
                &tight,
            )
            .is_err()
        );
        assert_eq!(rejected.part_count(), 0);
        assert_eq!(rejected.rels().len(), 0);

        let mut stored = OpcPackage::new();
        put(&mut stored, panes, Conformance::Transitional).unwrap();
        assert!(load_with(&stored, &tight).is_err());
    }

    fn synthetic_package(
        external_extension: bool,
        extension_content_type: &str,
        pane_relationship_id: &str,
    ) -> OpcPackage {
        let task_panes_xml = format!(
            r#"<wetp:taskpanes xmlns:wetp="{TASK_PANES_NAMESPACE}" xmlns:r="{TRANSITIONAL_RELATIONSHIPS_NAMESPACE}"><wetp:taskpane dockstate="right" visibility="0" width="320" row="0"><wetp:webextensionref r:id="{pane_relationship_id}"/></wetp:taskpane></wetp:taskpanes>"#
        );
        let extension_xml = write_add_in(&sample_extension(), Conformance::Transitional).unwrap();
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP.into(),
            "word/webextensions/taskpanes.xml".into(),
            "rIdTaskPanes".into(),
            false,
        );
        let mut task_panes_part = XmlPart::new(
            PackURI::new("/word/webextensions/taskpanes.xml").unwrap(),
            TASK_PANES_CONTENT_TYPE.into(),
            task_panes_xml.into_bytes(),
        );
        task_panes_part.rels_mut().add_relationship(
            ADD_IN_RELATIONSHIP.into(),
            if external_extension {
                "https://example.invalid/add-in".into()
            } else {
                "webextension1.xml".into()
            },
            "rId1".into(),
            external_extension,
        );
        package.add_part(Box::new(task_panes_part));
        if !external_extension {
            package.add_part(Box::new(XmlPart::new(
                PackURI::new("/word/webextensions/webextension1.xml").unwrap(),
                extension_content_type.into(),
                extension_xml,
            )));
        }
        package
    }

    fn add_shared_task_pane_ingress(package: &mut OpcPackage) -> PackURI {
        let task_panes_name = package
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
            .unwrap()
            .target_partname()
            .unwrap();
        package.rels_mut().add_relationship(
            "urn:litchi:test:shared-task-panes".into(),
            task_panes_name.as_str().trim_start_matches('/').to_owned(),
            "rIdSharedTaskPanes".into(),
            false,
        );
        task_panes_name
    }

    fn local_fixture_package(task_panes_xml: &[u8], extension_xml: &[u8]) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            TASK_PANES_RELATIONSHIP.into(),
            "webextensions/taskpanes.xml".into(),
            "rIdTaskPanes".into(),
            false,
        );
        let mut task_panes_part = XmlPart::new(
            PackURI::new("/webextensions/taskpanes.xml").unwrap(),
            TASK_PANES_CONTENT_TYPE.into(),
            task_panes_xml.to_vec(),
        );
        task_panes_part.rels_mut().add_relationship(
            ADD_IN_RELATIONSHIP.into(),
            "webextension1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(task_panes_part));
        package.add_part(Box::new(XmlPart::new(
            PackURI::new("/webextensions/webextension1.xml").unwrap(),
            ADD_IN_CONTENT_TYPE.into(),
            extension_xml.to_vec(),
        )));
        package
    }

    fn sample_extension() -> AddIn {
        AddIn {
            id: "{00000000-0000-0000-0000-000000000001}".into(),
            frozen: true,
            reference: Reference {
                id: "wa1".into(),
                version: "1.0.0.0".into(),
                location: Some("en-us".into()),
                store: Store::Omex,
                extension_list: None,
            },
            alternate_references: vec![],
            properties: vec![Property {
                name: "Office.AutoShowTaskpaneWithDocument".into(),
                value: "false".into(),
            }],
            bindings: vec![Binding {
                id: "binding-1".into(),
                kind: BindingKind::Matrix,
                app_ref: "app-ref".into(),
                extension_list: None,
            }],
            snapshot: Some(Snapshot::default()),
            extension_list: None,
        }
    }

    fn sample_task_panes() -> Panes {
        let mut extension = sample_extension();
        extension.snapshot = Some(Snapshot {
            embedded_relationship_id: Some("rIdSnapshot".into()),
            linked_relationship_id: Some("rIdLinked".into()),
            compression_state: Some(Compression::HighQualityPrint),
            effects: vec![
                Effect::from_xml(
                    br#"<a:alphaModFix xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" amt="50000"/>"#,
                )
                .unwrap(),
            ],
            extension_list: None,
        });
        Panes {
            panes: vec![Pane {
                dock_state: Dock::Right,
                visible: true,
                width: 320.0,
                row: 0,
                locked: false,
                relationship_id: "rId1".into(),
                add_in: extension,
                snapshot_resources: vec![
                    SnapshotResource {
                        relationship_id: "rIdSnapshot".into(),
                        target: SnapshotTarget::Internal {
                            part_name: PackURI::new("/media/web-extension-snapshot.png").unwrap(),
                            content_type: "image/png".into(),
                            data: Arc::new(vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]),
                        },
                    },
                    SnapshotResource {
                        relationship_id: "rIdLinked".into(),
                        target: SnapshotTarget::External {
                            target: "https://example.invalid/inert-snapshot.png".into(),
                        },
                    },
                ],
                extension_list: None,
            }],
        }
    }
}
