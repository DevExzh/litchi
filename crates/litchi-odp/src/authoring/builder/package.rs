//! ODP package assembly for the authoring snapshot.

use super::{Builder, validation};
use crate::core::{PackageWriter, Structure};
use litchi_core::Result;

pub(super) fn build(builder: &Builder) -> Result<Vec<u8>> {
    validation::validate(builder.snapshot())?;

    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.presentation")?;

    let content_xml = builder.generate_content_xml()?;
    writer.add_file("content.xml", content_xml.as_bytes())?;

    let mut styles_xml = Structure::default_styles_xml();
    for layout in &builder.page_layouts.layouts {
        styles_xml = crate::model::page_layout::set_xml(&styles_xml, layout)?;
    }
    writer.add_file("styles.xml", styles_xml.as_bytes())?;

    let meta_xml = builder.generate_meta_xml();
    writer.add_file("meta.xml", meta_xml.as_bytes())?;

    for (path, media) in &builder.media_files {
        writer.add_file_with_media_type(path, &media.bytes, &media.media_type)?;
    }

    writer.finish_to_bytes()
}
