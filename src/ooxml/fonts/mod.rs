#[cfg(feature = "fonts")]
pub mod obfuscation;

#[cfg(feature = "fonts")]
pub use obfuscation::*;

#[cfg(feature = "fonts")]
use crate::common::id::{format_guid_braced, generate_guid_bytes};
#[cfg(feature = "fonts")]
use crate::fonts::{AllsortsSubsetter, CollectGlyphs, FontData, FontSubsetter};
#[cfg(feature = "fonts")]
use crate::ooxml::error::{OoxmlError, Result};
#[cfg(feature = "fonts")]
use crate::ooxml::opc::constants::relationship_type as rt;
#[cfg(feature = "fonts")]
use crate::ooxml::opc::part::BlobPart;
#[cfg(feature = "fonts")]
use crate::ooxml::opc::{OpcPackage, PackURI};
#[cfg(feature = "fonts")]
use allsorts::{
    binary::read::ReadScope,
    tables::{
        FontTableProvider, OpenTypeFont,
        cmap::{Cmap, CmapSubtable},
    },
};
#[cfg(feature = "fonts")]
use roaring::RoaringBitmap;
#[cfg(feature = "fonts")]
use std::collections::HashMap;

/// Trait for document types that support font embedding.
#[cfg(feature = "fonts")]
pub trait EmbedFonts: CollectGlyphs {
    /// Embed fonts into the given OPC package based on used glyphs and save options.
    fn embed_fonts(&mut self) -> Result<()>;
}

/// Map Unicode code points to glyph IDs using the font's cmap table
#[cfg(feature = "fonts")]
fn map_codepoints_to_glyph_ids(font_data: &FontData, codepoints: &[u32]) -> Result<Vec<u16>> {
    let scope = ReadScope::new(&font_data.data);
    let font_file = scope
        .read::<OpenTypeFont>()
        .map_err(|e| OoxmlError::Other(format!("Failed to parse font: {}", e)))?;

    let provider = font_file
        .table_provider(font_data.index as usize)
        .map_err(|e| OoxmlError::Other(format!("Failed to get table provider: {}", e)))?;

    // Get cmap table - table_data returns Result<Option<Rc<dyn AsRef<[u8]>>>>
    let cmap_data = provider
        .table_data(allsorts::tag::CMAP)
        .map_err(|e| OoxmlError::Other(format!("Failed to get cmap table: {}", e)))?
        .ok_or_else(|| OoxmlError::Other("Font has no cmap table".to_string()))?;

    let cmap_scope = ReadScope::new(cmap_data.as_ref());
    let cmap = cmap_scope
        .read::<Cmap>()
        .map_err(|e| OoxmlError::Other(format!("Failed to read cmap: {}", e)))?;

    // Find the best cmap subtable (prefer Unicode BMP or full Unicode)
    // encoding_records() already returns an iterator, iterate directly
    let mut subtable_opt = None;
    for record in cmap.encoding_records() {
        if let Ok(st) = cmap_scope
            .offset(record.offset as usize)
            .read::<CmapSubtable>()
        {
            subtable_opt = Some(st);
            break;
        }
    }

    let subtable = subtable_opt
        .ok_or_else(|| OoxmlError::Other("No usable cmap subtable found".to_string()))?;

    // Map each code point to glyph ID
    let mut glyph_ids = Vec::with_capacity(codepoints.len() + 1);

    // Always include glyph 0 (.notdef)
    glyph_ids.push(0);

    for &cp in codepoints {
        if let Ok(Some(glyph_id)) = subtable.map_glyph(cp)
            && glyph_id != 0
            && !glyph_ids.contains(&glyph_id)
        {
            glyph_ids.push(glyph_id);
        }
    }

    // Sort glyph IDs for better subsetting results
    glyph_ids.sort_unstable();
    glyph_ids.dedup();

    Ok(glyph_ids)
}

/// Information about an embedded font
#[cfg(feature = "fonts")]
pub struct EmbeddedFontInfo {
    pub relationship_id: String,
    pub font_key: String,
    pub properties: Option<crate::fonts::FontProperties>,
}

#[cfg(feature = "fonts")]
pub fn embed_fonts_in_package(
    used_glyphs: HashMap<String, RoaringBitmap>,
    package: &mut OpcPackage,
    font_dir: &str,
    rel_source_uri: &PackURI,
) -> Result<HashMap<String, EmbeddedFontInfo>> {
    let mut embedded_fonts = HashMap::new();
    let options = package.save_options().clone();

    if !options.embed_fonts {
        return Ok(embedded_fonts);
    }

    if used_glyphs.is_empty() {
        return Ok(embedded_fonts);
    }

    let loader = crate::fonts::loader::FontLoader::new();
    let subsetter = AllsortsSubsetter::new();

    for (font_name, glyphs) in &used_glyphs {
        // 1. Find and load the font
        let font_data = match loader.load_system_font(font_name) {
            Ok(data) => data,
            Err(_) => {
                // Font not found on system, skip it
                continue;
            },
        };

        // 2. Subset if requested
        let final_data = if options.subset_fonts {
            // RoaringBitmap stores u32 code points directly
            let codepoints: Vec<u32> = glyphs.iter().collect();

            match map_codepoints_to_glyph_ids(&font_data, &codepoints) {
                Ok(glyph_ids) => {
                    // Subset the font using the mapped glyph IDs
                    match subsetter.subset(&font_data, &glyph_ids) {
                        Ok(subsetted) => subsetted,
                        Err(_) => {
                            // If subsetting fails, fall back to full font
                            font_data.data.clone()
                        },
                    }
                },
                Err(_) => {
                    // If cmap mapping fails, fall back to full font
                    font_data.data.clone()
                },
            }
        } else {
            // Subsetting disabled, use full font
            font_data.data.clone()
        };

        // 3. OOXML Font Obfuscation
        // Generate a GUID for the font (used as fontKey in XML)
        let guid_bytes = generate_guid_bytes();
        let mut obfuscated_data = final_data;
        obfuscate_font_data_bytes(&mut obfuscated_data, &guid_bytes);
        let guid = format_guid_braced(&guid_bytes);

        // 4. Add to package
        let font_filename = format!("{}.odttf", font_name.replace(' ', "_"));
        let font_partname = format!("{}/{}", font_dir.trim_end_matches('/'), font_filename);
        let font_uri = PackURI::new(&font_partname)
            .map_err(|e| OoxmlError::Other(format!("Invalid font URI: {}", e)))?;

        // Content type for obfuscated OpenType font
        let content_type = "application/vnd.openxmlformats-officedocument.obfuscatedFont";

        let font_part = BlobPart::new(font_uri.clone(), content_type.to_string(), obfuscated_data);
        package.add_part(Box::new(font_part));

        // 5. Build relationship with RELATIVE path (Office expects relative paths)
        if let Ok(source_part) = package.get_part_mut(rel_source_uri) {
            // Use relative path: "fonts/font_name.odttf" instead of absolute "/word/fonts/..."
            let relative_target = format!("fonts/{}", font_filename);
            let rid = source_part.relate_to(&relative_target, rt::FONT);
            embedded_fonts.insert(
                font_name.clone(),
                EmbeddedFontInfo {
                    relationship_id: rid,
                    font_key: guid,
                    properties: font_data.properties.clone(),
                },
            );
        }
    }

    Ok(embedded_fonts)
}
