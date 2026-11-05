//! PPTX writes the OfficeCLI example corpus needed (record 0269): text-box
//! formatting, fill and preset geometry on `add`, shape fill and geometry on
//! `set`, OfficeCLI's shape aliases (`w`, `h`, `fontSize`, `preset`), and
//! slide transitions on `set`. Each runs through both native owner
//! routes, reads back before and after a byte round trip, keeps an exact
//! repeat as a no-op, and is refused on signed and encrypted sources.
#![cfg(feature = "pptx")]

use litchi_ops::{Code, File, NewNode, OpenOptions, PropValue, SourceMode, props};

const SHAPE: &str = "/slide[1]/shape[1]";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder
        .presentation()
        .unwrap()
        .add_slide()
        .unwrap()
        .add_text_box("alpha", 0, 0, 1_828_800, 914_400);
    builder.build().unwrap().to_bytes().unwrap()
}

fn open(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

/// The first slide's XML, read from the saved package.
fn slide_xml(file: &File) -> String {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut entry = archive.by_name("ppt/slides/slide1.xml").unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut entry, &mut xml).unwrap();
    xml
}

#[test]
fn text_box_add_takes_formatting_fill_geometry_and_officecli_aliases() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let node = NewNode::new("shape")
            .with_props(props! {
                "text" => "Hello",
                "x" => "1cm", "y" => "1cm", "w" => "10cm", "h" => "2cm",
                "fontSize" => 24, "bold" => true, "color" => "FF0000",
                "font" => "Georgia", "align" => "center",
                "fill" => "1E3A5F", "preset" => "ellipse",
            })
            .unwrap();
        let added = file.add("/slide[1]", node).unwrap();
        let path = added.path.to_string();
        assert_eq!(path, "/slide[1]/shape[2]");
        for reread in [
            added.file.clone(),
            open(&added.file.to_bytes().unwrap(), eager),
        ] {
            assert_eq!(prop(&reread, &path, "size").as_deref(), Some("24pt"));
            assert_eq!(prop(&reread, &path, "bold").as_deref(), Some("true"));
            assert_eq!(prop(&reread, &path, "color").as_deref(), Some("#FF0000"));
            assert_eq!(prop(&reread, &path, "width").as_deref(), Some("10cm"));
            let xml = slide_xml(&reread);
            assert!(xml.contains("<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>"));
            assert!(
                xml.contains("<a:solidFill><a:srgbClr val=\"1E3A5F\"/></a:solidFill></p:spPr>")
            );
        }
        // The input is unchanged; a refused value adds nothing.
        assert_eq!(file.to_bytes().unwrap(), bytes);
        let bad = NewNode::new("shape")
            .with_props(props! { "w" => "1cm", "h" => "1cm", "geometry" => "blob" })
            .unwrap();
        let error = file.add("/slide[1]", bad).unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("geometry"));
    }
}

#[test]
fn shape_fill_and_geometry_set_read_back_and_repeat_as_no_ops() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let filled = file
            .set(
                SHAPE,
                props! { "fill" => "none", "geometry" => "roundRect" },
            )
            .unwrap();
        let xml = slide_xml(&filled);
        assert!(xml.contains("<a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom>"));
        assert!(xml.contains("<a:noFill/></p:spPr>"));
        let reopened = open(&filled.to_bytes().unwrap(), eager);
        assert_eq!(slide_xml(&reopened), xml);
        // The same assignments again change nothing.
        let commit = filled
            .edit(|edit| edit.set(SHAPE, props! { "fill" => "none", "shape" => "roundRect" }))
            .unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(
            commit.file().to_bytes().unwrap(),
            filled.to_bytes().unwrap()
        );
        // Null removes the local fill choice; geometry cannot be cleared.
        let cleared = filled
            .set(SHAPE, props! { "fill" => PropValue::Null })
            .unwrap();
        assert!(!slide_xml(&cleared).contains("<a:noFill/></p:spPr>"));
        assert_eq!(
            filled
                .set(SHAPE, props! { "geometry" => PropValue::Null })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        // A scheme colour is not writable natively; the input is unchanged.
        assert_eq!(
            file.set(SHAPE, props! { "fill" => "accent1" })
                .unwrap_err()
                .code(),
            Code::UnsupportedCapability
        );
        assert_eq!(file.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn slide_transitions_set_and_clear() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let pushed = file
            .set("/slide[1]", props! { "transition" => "push-right-slow" })
            .unwrap();
        let xml = slide_xml(&pushed);
        assert!(xml.contains("<p:transition spd=\"slow\"><p:push dir=\"r\"/></p:transition>"));
        let reopened = open(&pushed.to_bytes().unwrap(), eager);
        assert_eq!(slide_xml(&reopened), xml);
        let again = pushed
            .edit(|edit| edit.set("/slide[1]", props! { "transition" => "push-right-slow" }))
            .unwrap();
        assert!(!again.diagnostics().changed());

        let morphed = pushed
            .set("/slide[1]", props! { "transition" => "morph-byword" })
            .unwrap();
        let xml = slide_xml(&morphed);
        assert!(xml.contains("option=\"byWord\""));
        assert!(!xml.contains("<p:push"));

        for clear in [PropValue::from("none"), PropValue::Null] {
            let clear: PropValue<'static> = clear;
            let cleared = morphed
                .set("/slide[1]", props! { "transition" => clear.clone() })
                .unwrap();
            // The leaf clears the active branch of a markup-compatibility
            // owner and keeps its fallback for older readers (record 0269).
            let xml = slide_xml(&cleared);
            assert!(!xml.contains("morph"), "{xml}");
            assert!(
                xml.contains("<mc:Choice Requires=\"p159\"></mc:Choice>"),
                "{xml}"
            );
            // A plain transition clears completely.
            let plain = pushed
                .set("/slide[1]", props! { "transition" => clear.clone() })
                .unwrap();
            assert!(!slide_xml(&plain).contains("<p:transition"));
        }
        let error = file
            .set("/slide[1]", props! { "transition" => "sparkle" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("transition"));
        assert_eq!(file.to_bytes().unwrap(), bytes);

        // A transition is set after the slide is added, not on `add`.
        let node = NewNode::new("slide")
            .with_props(props! { "layout" => "Blank", "transition" => "fade" })
            .unwrap();
        assert_eq!(
            file.add("/", node).unwrap_err().code(),
            Code::UnsupportedProperty
        );
    }
}

#[test]
fn signed_sources_refuse_the_new_writes() {
    let bytes =
        std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx").unwrap();
    let signed = open(&bytes, false);
    assert!(signed.is_signed());
    for (path, props) in [
        (SHAPE, props! { "fill" => "FF0000" }),
        (SHAPE, props! { "geometry" => "ellipse" }),
        ("/slide[1]", props! { "transition" => "fade" }),
    ] {
        assert_eq!(
            signed.set(path, props).unwrap_err().code(),
            Code::DocumentProtected
        );
    }
    let node = NewNode::new("shape")
        .with_props(props! { "w" => "1cm", "h" => "1cm", "size" => 12 })
        .unwrap();
    assert_eq!(
        signed.add("/slide[1]", node).unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(signed.to_bytes().unwrap(), bytes);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_sources_refuse_the_new_writes_until_reencrypted() {
    use litchi_ops::Password;
    use litchi_pptx::encryption::{self, Mode};
    let encrypted = encryption::encrypt(
        fixture(),
        "secret",
        Mode::Standard,
        &encryption::Limits::default(),
    )
    .unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(encrypted).unwrap() else {
        panic!("encrypted fixture")
    };
    let source = locked.unlock(Password::new("secret".to_owned())).unwrap();
    for (path, props) in [
        (SHAPE, props! { "fill" => "FF0000" }),
        ("/slide[1]", props! { "transition" => "fade" }),
    ] {
        assert_eq!(
            source.set(path, props).unwrap_err().code(),
            Code::DocumentProtected
        );
    }
}
