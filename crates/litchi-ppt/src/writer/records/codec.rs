use super::super::env_data::{
    SHEET_PROPERTIES_CHILD_TYPE, SheetPropertiesAtom, SlideViewInfoAtom, SrKinsokuAtom,
    TxCFStyleAtom, TxPFStyleAtom, TxSIStyleAtom, VBAInfoAtom,
};
use super::super::spec::{
    BinaryTagData, MAIN_MASTER_PLACEHOLDERS, MAIN_MASTER_SLIDE_ATOM_RESERVED, SlideLayoutType,
    Tag10, color_schemes,
};
use super::super::tx_style::{
    build_tx_master_style_body, build_tx_master_style_center_body,
    build_tx_master_style_center_title, build_tx_master_style_half_body,
    build_tx_master_style_notes, build_tx_master_style_other, build_tx_master_style_quarter_body,
    build_tx_master_style_title, tx_style_instance,
};
use super::model::{Error, RecordBuilder, record_type};

use crate::view_info::SlideViewInfo;
use litchi_core::unit::emu_u32_to_ppt_master_u32;

use super::validation::validate_complete_record;

/// Create a document container record
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_document_container(slides: &[Vec<u8>]) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

    // Add slide list
    for slide in slides {
        builder.write_child(slide);
    }

    builder.build()
}

/// Create a `MainMaster` container aligned with POI's empty.ppt structure.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_main_master_container(ppdrawing: &[u8]) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 0, record_type::MAIN_MASTER);

    // 1) SlideAtom (MS-PPT 2.4.7) - master slide atom
    let mut slide_atom = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
    let mut atom_data = Vec::with_capacity(24);
    // SSlideLayoutAtom: geometry + placeholder types
    atom_data.extend_from_slice(&(SlideLayoutType::TitleBody as u32).to_le_bytes());
    atom_data.extend_from_slice(&MAIN_MASTER_PLACEHOLDERS);
    // masterID=0 (masters don't reference another master), notesID=0
    atom_data.extend_from_slice(&0u32.to_le_bytes());
    atom_data.extend_from_slice(&0u32.to_le_bytes());
    // flags = 0 for MainMaster (doesn't follow anything)
    atom_data.extend_from_slice(&0u16.to_le_bytes());
    // reserved field from POI
    atom_data.extend_from_slice(&MAIN_MASTER_SLIDE_ATOM_RESERVED.to_le_bytes());
    slide_atom.write_data(&atom_data);
    builder.write_child(&slide_atom.build()?);

    // 2) 12 ColorSchemeAtom records (inst=6) - MS-PPT 2.4.17
    for scheme in &color_schemes::ALL {
        let mut cs = RecordBuilder::new(0x00, 6, record_type::COLOR_SCHEME_ATOM);
        cs.write_data(&scheme.to_bytes());
        builder.write_child(&cs.build()?);
    }

    // 3) TxMasterStyleAtom for Title (instance=0) - MS-PPT 2.9.45
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::TITLE,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_title());
        builder.write_child(&tx.build()?);
    }

    // 4) TxMasterStyleAtom for Body (instance=1)
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::BODY,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_body());
        builder.write_child(&tx.build()?);
    }

    // 5) TxMasterStyleAtom for Notes (instance=2)
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::NOTES,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_notes());
        builder.write_child(&tx.build()?);
    }

    // 6) TxMasterStyleAtom for CENTER_BODY (instance=5)
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::CENTER_BODY,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_center_body());
        builder.write_child(&tx.build()?);
    }

    // 7) TxMasterStyleAtom for CENTER_TITLE (instance=6)
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::CENTER_TITLE,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_center_title());
        builder.write_child(&tx.build()?);
    }

    // 8) TxMasterStyleAtom for HALF_BODY (instance=7)
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::HALF_BODY,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_half_body());
        builder.write_child(&tx.build()?);
    }

    // 9) TxMasterStyleAtom for QUARTER_BODY (instance=8)
    {
        let mut tx = RecordBuilder::new(
            0x00,
            tx_style_instance::QUARTER_BODY,
            record_type::TX_MASTER_STYLE_ATOM,
        );
        tx.write_data(&build_tx_master_style_quarter_body());
        builder.write_child(&tx.build()?);
    }

    // 10) PPDrawing for the master (Escher DgContainer)
    builder.write_child(ppdrawing);

    // 11) Tail ColorSchemeAtom (instance=1) - same as scheme 0
    {
        let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
        color.write_data(&color_schemes::DEFAULT_LIGHT.to_bytes());
        builder.write_child(&color.build()?);
    }

    // 12) ProgTags with PPT10 binary tag
    {
        let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
        let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
        let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
        cstr.write_data(&Tag10::to_bytes());
        prog_bin.write_child(&cstr.build()?);
        let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
        bin.write_data(&BinaryTagData::MAIN_MASTER.to_bytes());
        prog_bin.write_child(&bin.build()?);
        prog_tags.write_child(&prog_bin.build()?);
        builder.write_child(&prog_tags.build()?);
    }

    builder.build()
}

/// Create a `SlideListWithText` (instance=MASTER) containing `SlidePersistAtom` entries for masters.
/// Each entry is (`persist_id_ref`, `slide_identifier`).
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_slide_list_with_text_master(entries: &[(u32, u32)]) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 1, record_type::SLIDE_LIST_WITH_TEXT);

    for &(persist_id_ref, slide_identifier) in entries {
        let mut spa = RecordBuilder::new(0x00, 0, record_type::SLIDE_PERSIST_ATOM);
        let mut data = Vec::with_capacity(24);
        data.extend_from_slice(&persist_id_ref.to_le_bytes()); // refID
        data.extend_from_slice(&0u32.to_le_bytes()); // flags: masters typically 0
        data.extend_from_slice(&0u32.to_le_bytes()); // numPlaceholderTexts
        data.extend_from_slice(&slide_identifier.to_le_bytes()); // slideIdentifier (e.g., 0x8000_0000)
        // reserved 4 bytes
        data.extend_from_slice(&0u32.to_le_bytes());
        spa.write_data(&data);
        builder.write_child(&spa.build()?);
    }

    builder.build()
}

/// Create a `DocumentAtom` record with minimal fields.
///
/// The input slide dimensions are specified in EMUs (as used by the writer
/// API) and converted into PPT "master units" (576 units per inch) to
/// match the values written by `LibreOffice` and `PowerPoint`.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_document_atom(
    slide_size_width_master: u32,
    slide_size_height_master: u32,
    slide_count: u32,
    notes_count: u32,
    master_count: u32,
) -> Result<Vec<u8>, Error> {
    create_document_atom_with_font_embedding(
        slide_size_width_master,
        slide_size_height_master,
        slide_count,
        notes_count,
        master_count,
        false,
    )
}

/// Create a `DocumentAtom` with an explicit `fSaveWithFonts` value.
///
/// The compatibility helper [`create_document_atom`] retains the historical
/// `false` value. Fresh writers use this variant after validating the complete
/// font catalog, so the flag cannot disagree with emitted embedded faces.
pub(crate) fn create_document_atom_with_font_embedding(
    slide_size_width_master: u32,
    slide_size_height_master: u32,
    _slide_count: u32,
    _notes_count: u32,
    _master_count: u32,
    save_with_fonts: bool,
) -> Result<Vec<u8>, Error> {
    // Match PowerPoint/POI files: recVer = 1, recInstance = 0, 40-byte payload
    let mut builder = RecordBuilder::new(0x01, 0, record_type::DOCUMENT_ATOM);
    let mut data = Vec::with_capacity(40);

    let slide_width_mu = emu_u32_to_ppt_master_u32(slide_size_width_master);
    let slide_height_mu = emu_u32_to_ppt_master_u32(slide_size_height_master);

    // slideSize (width, height)
    data.extend_from_slice(&slide_width_mu.to_le_bytes());
    data.extend_from_slice(&slide_height_mu.to_le_bytes());
    // notesSize - POI uses (height, width) for portrait orientation
    data.extend_from_slice(&slide_height_mu.to_le_bytes()); // notesSizeX = height
    data.extend_from_slice(&slide_width_mu.to_le_bytes()); // notesSizeY = width
    // serverZoom ratio 5:10 per POI empty.ppt
    data.extend_from_slice(&5u32.to_le_bytes()); // serverZoomFrom
    data.extend_from_slice(&10u32.to_le_bytes()); // serverZoomTo
    // master persists
    data.extend_from_slice(&0u32.to_le_bytes()); // notesMasterPersist
    data.extend_from_slice(&0u32.to_le_bytes()); // handoutMasterPersist
    // first slide number and slide size type (use ON_SCREEN for default)
    data.extend_from_slice(&1u16.to_le_bytes()); // firstSlideNum
    data.extend_from_slice(&0u16.to_le_bytes()); // slideSizeType = ON_SCREEN
    // flags bytes
    data.push(u8::from(save_with_fonts)); // saveWithFonts
    data.push(0u8); // omitTitlePlace
    data.push(0u8); // rightToLeft
    data.push(1u8); // showComments (visible comments, matches LibreOffice)
    // reserved padding to reach 48-byte payload (POI/LibreOffice-compatible)
    builder.write_data(&data);
    builder.build()
}

/// Create a slide container record
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_slide_container(_slide_id: u32, text: &str) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 0, record_type::SLIDE);

    // Add slide atom
    let mut slide_atom = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
    let mut atom_data = Vec::with_capacity(24);
    // Embedded SSlideLayoutAtom (12 bytes): geometry + 8 bytes placeholder IDs
    // SL_Blank = 0x000D per MS-PPT section 2.13.27 SSlideLayoutType
    let geometry_blank: u32 = 0x000D;
    atom_data.extend_from_slice(&geometry_blank.to_le_bytes());
    atom_data.extend_from_slice(&[0u8; 8]);
    // masterID (USES_MASTER_SLIDE_ID = 0x80000000), notesID=0
    atom_data.extend_from_slice(&0x8000_0000u32.to_le_bytes());
    atom_data.extend_from_slice(&0u32.to_le_bytes());
    // flags (follow master objects/scheme/background) = 0x0007, reserved=0x0000
    atom_data.extend_from_slice(&7u16.to_le_bytes());
    atom_data.extend_from_slice(&0u16.to_le_bytes());
    slide_atom.write_data(&atom_data);
    builder.write_child(&slide_atom.build()?);

    // Add text if provided
    if !text.is_empty() {
        let text_atom = create_text_atom(text)?;
        builder.write_child(&text_atom);
    }

    builder.build()
}

/// Wrap an Escher `DggContainer` blob into a `PPDrawingGroup` PPT record.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn wrap_dgg_into_ppdrawing_group(dgg_blob: &[u8]) -> Result<Vec<u8>, Error> {
    // Align with POI: version 0x0F (container) but payload is raw Escher DGG data
    let mut builder = RecordBuilder::new(0x0F, 0, record_type::PP_DRAWING_GROUP);
    builder.write_data(dgg_blob);
    builder.build()
}

/// Wrap an Escher `DgContainer` blob (plus any following Escher children) into a `PPDrawing` PPT record.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn wrap_dg_into_ppdrawing(dg_blob_and_children: &[u8]) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 0, record_type::PP_DRAWING);
    builder.write_data(dg_blob_and_children);
    builder.build()
}

/// Create a minimal Environment container with an empty `FontCollection` child.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_environment_minimal() -> Result<Vec<u8>, Error> {
    let mut fc = RecordBuilder::new(0x0F, 0, record_type::FONT_COLLECTION);
    let mut fea = RecordBuilder::new(0x00, 0, record_type::FONT_ENTITY_ATOM);
    let mut fe_data = vec![0u8; 68];
    for (i, ch) in "Arial\0".encode_utf16().enumerate() {
        if i * 2 + 1 < 64 {
            let bytes = ch.to_le_bytes();
            fe_data[i * 2] = bytes[0];
            fe_data[i * 2 + 1] = bytes[1];
        }
    }
    fe_data[66] = 4;
    fea.write_data(&fe_data);
    fc.write_child(&fea.build()?);
    create_environment_with_font_collection(&fc.build()?)
}

/// Create the canonical Environment around one checked base font collection.
///
/// `font_collection` must be one complete `FontCollectionContainer` record.
/// Its bytes are inserted unchanged; all unrelated Environment children retain
/// their historical order and values.
pub(crate) fn create_environment_with_font_collection(
    font_collection: &[u8],
) -> Result<Vec<u8>, Error> {
    validate_complete_record(font_collection, 0x0f, record_type::FONT_COLLECTION)?;
    // Environment container (1010)
    let mut env = RecordBuilder::new(0x0F, 0, record_type::ENVIRONMENT);

    // Children in POI order for empty.ppt:
    // 1) SrKinsoku (4040) container with SrKinsokuAtom
    let mut kinsoku = RecordBuilder::new(0x0F, 2, record_type::SR_KINSOKU);
    let mut kinsoku_atom = RecordBuilder::new(0x00, 3, record_type::SR_KINSOKU_ATOM);
    kinsoku_atom.write_data(&SrKinsokuAtom::DEFAULT.to_bytes());
    kinsoku.write_child(&kinsoku_atom.build()?);
    env.write_child(&kinsoku.build()?);

    // 2) Caller-provided deterministic FontCollection container.
    env.write_child(font_collection);

    // 3) TxCFStyleAtom (character formatting defaults)
    let mut txcf = RecordBuilder::new(0x00, 0, record_type::TX_CF_STYLE_ATOM);
    txcf.write_data(&TxCFStyleAtom::DEFAULT.to_bytes());
    env.write_child(&txcf.build()?);

    // 4) TxPFStyleAtom (paragraph formatting defaults)
    let mut txpf = RecordBuilder::new(0x00, 0, record_type::TX_PF_STYLE_ATOM);
    txpf.write_data(&TxPFStyleAtom::DEFAULT.to_bytes());
    env.write_child(&txpf.build()?);

    // 5) TxSIStyleAtom (special info formatting)
    let mut txsi = RecordBuilder::new(0x00, 0, record_type::TX_SI_STYLE_ATOM);
    txsi.write_data(&TxSIStyleAtom::DEFAULT.to_bytes());
    env.write_child(&txsi.build()?);

    // 6) TxMasterStyleAtom for OTHER (instance=4)
    let mut tx = RecordBuilder::new(
        0x00,
        tx_style_instance::OTHER,
        record_type::TX_MASTER_STYLE_ATOM,
    );
    tx.write_data(&build_tx_master_style_other());
    env.write_child(&tx.build()?);

    env.build()
}

/// Create a `SlideListWithText` (instance=SLIDES) containing `SlidePersistAtom` entries.
/// Each entry is (`persist_id_ref`, `slide_identifier`).
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_slide_list_with_text_slides(entries: &[(u32, u32)]) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 0, record_type::SLIDE_LIST_WITH_TEXT);

    for &(persist_id_ref, slide_identifier) in entries {
        let mut spa = RecordBuilder::new(0x00, 0, record_type::SLIDE_PERSIST_ATOM);
        let mut data = Vec::with_capacity(24);
        data.extend_from_slice(&persist_id_ref.to_le_bytes()); // refID
        data.extend_from_slice(&4u32.to_le_bytes()); // flags: HAS_SHAPES_OTHER_THAN_PLACEHOLDERS
        data.extend_from_slice(&0u32.to_le_bytes()); // numPlaceholderTexts
        data.extend_from_slice(&slide_identifier.to_le_bytes()); // slideIdentifier
        // reserved 4 bytes
        data.extend_from_slice(&0u32.to_le_bytes());
        spa.write_data(&data);
        builder.write_child(&spa.build()?);
    }

    builder.build()
}

/// Create a `SlideListWithText` (instance=NOTES) containing `SlidePersistAtom` entries for notes.
/// Each entry is (`persist_id_ref`, `notes_identifier`).
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_slide_list_with_text_notes(entries: &[(u32, u32)]) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x0F, 2, record_type::SLIDE_LIST_WITH_TEXT); // instance=2 for NOTES

    for &(persist_id_ref, notes_identifier) in entries {
        let mut spa = RecordBuilder::new(0x00, 0, record_type::SLIDE_PERSIST_ATOM);
        let mut data = Vec::with_capacity(24);
        data.extend_from_slice(&persist_id_ref.to_le_bytes()); // refID
        data.extend_from_slice(&4u32.to_le_bytes()); // flags: HAS_SHAPES_OTHER_THAN_PLACEHOLDERS
        data.extend_from_slice(&0u32.to_le_bytes()); // numPlaceholderTexts
        data.extend_from_slice(&notes_identifier.to_le_bytes()); // slideIdentifier (for notes, often slide# + offset)
        // reserved 4 bytes
        data.extend_from_slice(&0u32.to_le_bytes());
        spa.write_data(&data);
        builder.write_child(&spa.build()?);
    }

    builder.build()
}

/// Create a `DocInfo` List container with optional typed slide and notes editing views.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_docinfo_list_container(
    slide_view_info: Option<&SlideViewInfo>,
    notes_view_info: Option<&SlideViewInfo>,
) -> Result<Vec<u8>, Error> {
    create_docinfo_list_container_with_binary_tags(
        slide_view_info,
        notes_view_info,
        VBAInfoAtom::DEFAULT,
        std::iter::empty(),
    )
}

/// Create a `DocInfo` List and append validated `ProgBinaryTag` records to its
/// single `DocProgTagsContainer`.
pub(crate) fn create_docinfo_list_container_with_binary_tags<'a>(
    slide_view_info: Option<&SlideViewInfo>,
    notes_view_info: Option<&SlideViewInfo>,
    vba_info: VBAInfoAtom,
    additional_binary_tags: impl IntoIterator<Item = &'a [u8]>,
) -> Result<Vec<u8>, Error> {
    create_docinfo_list_container_with_extensions(
        slide_view_info,
        notes_view_info,
        vba_info,
        std::iter::empty::<&[u8]>(),
        None,
        additional_binary_tags,
    )
}

/// Create one `DocInfo` list whose single `DocProgTagsContainer` owns the
/// canonical `___PPT10` payload and every additional versioned binary tag.
///
/// `PowerPoint` 10 font records are appended to the existing document extension
/// payload rather than authored as a duplicate tag or duplicate `ProgTags`
/// owner. Only the two font record kinds permitted by MS-PPT are accepted.
pub(crate) fn create_docinfo_list_container_with_extensions<'a, 'b>(
    slide_view_info: Option<&SlideViewInfo>,
    notes_view_info: Option<&SlideViewInfo>,
    vba_info: VBAInfoAtom,
    ppt10_font_records: impl IntoIterator<Item = &'a [u8]>,
    modify_password: Option<&[u8]>,
    additional_binary_tags: impl IntoIterator<Item = &'b [u8]>,
) -> Result<Vec<u8>, Error> {
    let mut list = RecordBuilder::new(0x0F, 0, record_type::DOC_INFO_LIST);

    // SheetProperties (1044) container with timestamp atom
    let mut sheet = RecordBuilder::new(0x0F, 1, record_type::SHEET_PROPERTIES);
    let mut sheet_child = RecordBuilder::new(0x00, 0, SHEET_PROPERTIES_CHILD_TYPE);
    sheet_child.write_data(&SheetPropertiesAtom::DEFAULT.to_bytes());
    sheet.write_child(&sheet_child.build()?);
    list.write_child(&sheet.build()?);

    if let Some(view) = slide_view_info {
        let bytes = view
            .to_bytes()
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::InvalidData, error))?;
        list.write_child(&bytes);
    } else {
        // Preserve the historical canonical default slide view when no override is set.
        let mut svi = RecordBuilder::new(0x0F, 0, record_type::SLIDE_VIEW_INFO);
        let mut svia = RecordBuilder::new(0x00, 0, record_type::SLIDE_VIEW_INFO_ATOM);
        svia.write_data(&SlideViewInfoAtom::DEFAULT.to_bytes());
        svi.write_child(&svia.build()?);
        list.write_child(&svi.build()?);
    }
    if let Some(view) = notes_view_info {
        let bytes = view
            .to_bytes()
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::InvalidData, error))?;
        list.write_child(&bytes);
    }

    // VBAInfo (1023) container with VBAInfoAtom
    let mut vba = RecordBuilder::new(0x0F, 1, record_type::VBA_INFO);
    let mut vba_atom = RecordBuilder::new(0x02, 0, record_type::VBA_INFO_ATOM);
    vba_atom.write_data(&vba_info.to_bytes());
    vba.write_child(&vba_atom.build()?);
    list.write_child(&vba.build()?);

    // ProgTags with PPT10 binary tag
    let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
    let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
    let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
    cstr.write_data(&Tag10::to_bytes());
    prog_bin.write_child(&cstr.build()?);
    let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
    let mut collection = None;
    let mut flags = None;
    for record in ppt10_font_records {
        let kind = complete_record_type(record)?;
        match kind {
            record_type::FONT_COLLECTION_10 if collection.is_none() && flags.is_none() => {
                validate_complete_record(record, 0x0f, kind)?;
                collection = Some(record);
            },
            record_type::FONT_EMBED_FLAGS_10_ATOM if flags.is_none() => {
                validate_complete_record(record, 0x00, kind)?;
                flags = Some(record);
            },
            record_type::FONT_COLLECTION_10 => {
                return Err(Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "PowerPoint 10 font collection is duplicated or follows its flags",
                ));
            },
            record_type::FONT_EMBED_FLAGS_10_ATOM => {
                return Err(Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "PowerPoint 10 font embedding flags are duplicated",
                ));
            },
            _ => {
                return Err(Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "unsupported record in PowerPoint 10 font extension",
                ));
            },
        }
    }
    if let Some(record) = collection {
        bin.write_child(record);
    }
    // The canonical optional GridSpacing10Atom precedes embedding flags.
    bin.write_data(&BinaryTagData::DOCINFO.to_bytes());
    if let Some(record) = flags {
        bin.write_child(record);
    }
    if let Some(record) = modify_password {
        validate_complete_record(record, 0x00, record_type::CSTRING)?;
        let version_and_instance = u16::from_le_bytes([record[0], record[1]]);
        if version_and_instance >> 4 != 3 {
            return Err(Error::new(
                std::io::ErrorKind::InvalidInput,
                "PowerPoint modify-password atom has an invalid record instance",
            ));
        }
        bin.write_child(record);
    }
    prog_bin.write_child(&bin.build()?);
    prog_tags.write_child(&prog_bin.build()?);
    for tag in additional_binary_tags {
        validate_complete_record(tag, 0x0f, record_type::PROG_BINARY_TAG)?;
        prog_tags.write_child(tag);
    }
    list.write_child(&prog_tags.build()?);

    list.build()
}

fn complete_record_type(record: &[u8]) -> Result<u16, Error> {
    if record.len() < 8 {
        return Err(Error::new(
            std::io::ErrorKind::InvalidInput,
            "PowerPoint extension record is truncated",
        ));
    }
    let payload_len = usize::try_from(u32::from_le_bytes([
        record[4], record[5], record[6], record[7],
    ]))
    .map_err(|_err| Error::new(std::io::ErrorKind::InvalidInput, "record length overflows"))?;
    if payload_len.checked_add(8) != Some(record.len()) {
        return Err(Error::new(
            std::io::ErrorKind::InvalidInput,
            "PowerPoint extension record length is inconsistent",
        ));
    }
    Ok(u16::from_le_bytes([record[2], record[3]]))
}

/// Create the historical minimal `DocInfo` List with a default slide view.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_docinfo_list_container_minimal() -> Result<Vec<u8>, Error> {
    create_docinfo_list_container(None, None)
}

/// Create an `EndDocument` record.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_end_document() -> Result<Vec<u8>, Error> {
    let builder = RecordBuilder::new(0x00, 0, record_type::END_DOCUMENT);
    builder.build()
}

/// Create a text chars atom (for Unicode text)
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn create_text_atom(text: &str) -> Result<Vec<u8>, Error> {
    let mut builder = RecordBuilder::new(0x00, 0, record_type::TEXT_CHARS_ATOM);

    // Convert text to UTF-16LE
    let utf16: Vec<u16> = text.encode_utf16().collect();
    let mut text_data = Vec::new();
    for ch in utf16 {
        text_data.extend_from_slice(&ch.to_le_bytes());
    }

    builder.write_data(&text_data);
    builder.build()
}
