//! Borrowed, semantic `PresentationML` shape scenes.
//!
//! [`Scene`] indexes the full owner XML once, preserving inherited namespace
//! aliases and markup-compatibility branch selection. Shape XML is exposed by
//! checked spans into that processed owner; no element subtree is copied.
//!
//! ```
//! use litchi_pptx::shape::{Scene, Shape};
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//!
//! let xml = br#"<p:spTree
//!     xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
//!     <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr></p:sp>
//! </p:spTree>"#;
//! let scene = Scene::read(xml)?;
//! let Some(Shape::Auto(title)) = scene.get("Title")? else {
//!     return Err("missing title shape".into());
//! };
//! assert_eq!(title.common().name(), Some("Title"));
//! # Ok(())
//! # }
//! ```

pub mod classification;
pub mod designer;
pub mod diagram;
pub mod text;
pub mod theme;
pub mod zoom;

mod model;
mod reader;

pub use model::{
    Auto, Bounds, Chart, Common, Connector, Content, Diagram, Frame, Group, Ole, Picture,
    Placeholder, Shape, Shapes, Span, Table, Unknown,
};
pub use reader::{Key, Limits, LookupError, Scene};

/// Read a borrowed-by-default scene with conservative finite limits.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
#[inline]
pub fn read(xml: &[u8]) -> crate::Result<Scene<'_>> {
    Scene::read(xml)
}

/// Read a scene using explicit finite limits.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
#[inline]
pub fn read_with(xml: &[u8], limits: Limits) -> crate::Result<Scene<'_>> {
    Scene::read_with(xml, limits)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";

    #[test]
    fn indexes_transitional_shapes_without_copying_fragments() {
        let xml = format!(
            r#"<q:sld xmlns:q="{PML}" xmlns:d="{DML}">
                <q:cSld><q:spTree>
                    <q:nvGrpSpPr/><q:grpSpPr/>
                    <q:sp><q:nvSpPr><q:cNvPr id="2" name="Title &amp; Body"/><q:nvPr><q:ph type="title" idx="4"/></q:nvPr></q:nvSpPr><q:spPr><d:xfrm><d:off x="-10" y="20"/><d:ext cx="30" cy="40"/></d:xfrm></q:spPr><q:txBody><d:p><d:r><d:t>Hello</d:t></d:r></d:p><d:p><d:r><d:t>world</d:t></d:r></d:p></q:txBody></q:sp>
                    <q:pic><q:nvPicPr><q:cNvPr id="3" name="Photo"/></q:nvPicPr></q:pic>
                </q:spTree></q:cSld>
            </q:sld>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("scene");
        assert!(!scene.is_rewritten());
        assert_eq!(scene.len(), 2);

        let title = scene.get("Title & Body").expect("lookup").expect("title");
        assert!(matches!(title, Shape::Auto(_)));
        assert_eq!(title.id(), Some(2));
        assert_eq!(title.text(), Some("Hello\nworld"));
        assert_eq!(
            title.placeholder().expect("placeholder").kind(),
            Some("title")
        );
        assert_eq!(title.placeholder().expect("placeholder").index(), 4);
        assert_eq!(scene.placeholders().count(), 1);
        let bounds = title.bounds().expect("bounds");
        assert_eq!((bounds.x(), bounds.y()), (-10, 20));
        assert_eq!((bounds.width(), bounds.height()), (30, 40));

        let raw = title.xml().expect("borrowed shape XML");
        let owner = scene.xml();
        let span = title.span().expect("span");
        let start = usize::try_from(span.start()).expect("u32 fits usize");
        assert_eq!(raw.as_ptr(), owner[start..].as_ptr());
        assert!(raw.starts_with(b"<q:sp>"));
    }

    #[test]
    fn accepts_strict_namespaces_and_classifies_frames() {
        let chart = "http://purl.oclc.org/ooxml/drawingml/chart";
        let xml = format!(
            r#"<s:spTree xmlns:s="{STRICT_PML}" xmlns:d="{STRICT_DML}" xmlns:c="{chart}">
                <s:nvGrpSpPr/><s:grpSpPr/>
                <s:graphicFrame><s:nvGraphicFramePr><s:cNvPr id="7" name="Revenue"/></s:nvGraphicFramePr><d:graphic><d:graphicData><c:chart/></d:graphicData></d:graphic></s:graphicFrame>
                <s:graphicFrame><s:nvGraphicFramePr><s:cNvPr id="8" name="Generic"/></s:nvGraphicFramePr></s:graphicFrame>
            </s:spTree>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("strict scene");
        assert!(matches!(scene.at(0).expect("chart"), Shape::Chart(_)));
        assert!(matches!(scene.at(1).expect("frame"), Shape::Frame(_)));
    }

    #[test]
    fn groups_are_hierarchical_while_scene_order_is_preorder() {
        let xml = format!(
            r#"<p:spTree xmlns:p="{PML}">
                <p:nvGrpSpPr/><p:grpSpPr/>
                <p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="Outer"/></p:nvGrpSpPr><p:grpSpPr/>
                    <p:sp><p:nvSpPr><p:cNvPr id="2" name="First"/></p:nvSpPr></p:sp>
                    <p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Inner"/></p:nvGrpSpPr><p:grpSpPr/>
                        <p:pic><p:nvPicPr><p:cNvPr id="4" name="Nested"/></p:nvPicPr></p:pic>
                    </p:grpSp>
                </p:grpSp>
                <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="Line"/></p:nvCxnSpPr></p:cxnSp>
            </p:spTree>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("nested scene");
        let names: Vec<_> = scene.iter().filter_map(Shape::name).collect();
        assert_eq!(names, ["Outer", "First", "Inner", "Nested", "Line"]);
        let roots: Vec<_> = scene.roots().filter_map(Shape::name).collect();
        assert_eq!(roots, ["Outer", "Line"]);

        let Shape::Group(outer) = scene.at(0).expect("outer") else {
            panic!("expected group");
        };
        let children: Vec<_> = outer.shapes().filter_map(Shape::name).collect();
        assert_eq!(children, ["First", "Inner"]);
        let Shape::Group(inner) = scene.at(2).expect("inner") else {
            panic!("expected nested group");
        };
        assert_eq!(
            inner.shapes().filter_map(Shape::name).collect::<Vec<_>>(),
            ["Nested"]
        );
    }

    #[test]
    fn mce_selects_one_branch_and_excludes_inactive_fallback_shapes() {
        let p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        let mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        let xml = format!(
            r#"<p:spTree xmlns:p="{PML}" xmlns:p14="{p14}" xmlns:mc="{mc}">
                <p:nvGrpSpPr/><p:grpSpPr/>
                <mc:AlternateContent>
                    <mc:Choice Requires="p14"><p14:contentPart/></mc:Choice>
                    <mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="9" name="Duplicate"/></p:nvPicPr></p:pic></mc:Fallback>
                </mc:AlternateContent>
                <p:sp><p:nvSpPr><p:cNvPr id="10" name="Duplicate"/></p:nvSpPr></p:sp>
            </p:spTree>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("MCE scene");
        assert!(scene.is_rewritten());
        assert_eq!(scene.len(), 2);
        assert!(matches!(scene.at(0).expect("content"), Shape::Content(_)));
        assert!(matches!(
            scene.get("Duplicate").expect("unambiguous active name"),
            Some(Shape::Auto(_))
        ));
    }

    #[test]
    fn indexes_standard_presentationml_content_parts_as_content_shapes() {
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let xml = format!(
            r#"<p:spTree xmlns:p="{PML}" xmlns:r="{rel}">
                <p:nvGrpSpPr/><p:grpSpPr/>
                <p:contentPart r:id="rIdOpaque"/>
            </p:spTree>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("content-part scene");
        assert_eq!(scene.len(), 1);
        assert!(matches!(
            scene.at(0).expect("content part"),
            Shape::Content(_)
        ));
        assert_eq!(
            scene.at(0).expect("content part").xml().unwrap(),
            b"<p:contentPart r:id=\"rIdOpaque\"/>"
        );
    }

    #[test]
    fn nested_ole_fallback_picture_is_not_a_scene_shape() {
        let xml = format!(
            r#"<p:spTree xmlns:p="{PML}" xmlns:a="{DML}">
                <p:nvGrpSpPr/><p:grpSpPr/>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="11" name="Object"/></p:nvGraphicFramePr><a:graphic><a:graphicData><p:oleObj><p:pic><p:nvPicPr><p:cNvPr id="12" name="Fallback preview"/></p:nvPicPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>
            </p:spTree>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("OLE scene");
        assert_eq!(scene.len(), 1);
        assert!(matches!(scene.at(0).expect("OLE"), Shape::Ole(_)));
        assert!(scene.get("Fallback preview").expect("lookup").is_none());
    }

    #[test]
    fn duplicate_active_names_and_checked_positions_are_errors() {
        let xml = format!(
            r#"<p:spTree xmlns:p="{PML}"><p:nvGrpSpPr/><p:grpSpPr/>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="Same"/></p:nvSpPr></p:sp>
                <p:pic><p:nvPicPr><p:cNvPr id="3" name="Same"/></p:nvPicPr></p:pic>
            </p:spTree>"#
        );
        let scene = Scene::read(xml.as_bytes()).expect("scene");
        assert!(scene.get("Same").is_err());
        assert!(scene.at(2).is_err());
        assert!(matches!(
            scene.get(2_usize),
            Err(LookupError::IndexOutOfBounds { index: 2, len: 2 })
        ));
        assert!(matches!(
            scene.shape("Missing"),
            Err(LookupError::NameNotFound { .. })
        ));
    }
}
