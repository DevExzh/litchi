#![allow(
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "decoding steps deliberately rebind a working value as it is refined through the parse pipeline"
)]
use super::{
    ControlWord, Cow, Destination, MAX_PICTURE_DATA_BYTES, ParsedBodyStoryEvent, Parser, RtfError,
    RtfResult, Token, control_symbol_text,
};

impl<'a> Parser<'a> {
    /// Parse picture/image content.
    pub(super) fn parse_body_picture_compatibility(
        &mut self,
        kind: crate::PictureCompatibilityKind,
        starred: bool,
    ) -> RtfResult<()> {
        if self.current_state()?.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility wrapper must occur in document body text".to_string(),
            ));
        }
        if self.picture_compatibility_records.len() >= crate::MAX_PICTURE_COMPATIBILITY_RECORDS {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility record count exceeds the safety limit".to_string(),
            ));
        }
        self.pos += if starred { 2 } else { 1 };
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF picture-compatibility wrapper contains text outside pict".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos..self.pos + 2),
            Some([Token::OpenBrace, Token::Control(ControlWord::Picture)])
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility wrapper must contain exactly one pict destination"
                    .to_string(),
            ));
        }
        self.pos += 1;
        let picture_index = self.pictures.len();
        self.parse_picture()?;
        if self.pictures.len() != picture_index + 1 {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility payload cannot be empty".to_string(),
            ));
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility pict destination is not closed".to_string(),
            ));
        }
        self.pos += 1;
        while let Some(Token::Text(text)) = self.tokens.get(self.pos) {
            if !text.chars().all(char::is_whitespace) {
                return Err(RtfError::MalformedDocument(
                    "RTF picture-compatibility wrapper contains trailing text".to_string(),
                ));
            }
            self.pos += 1;
        }
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility wrapper must contain exactly one pict destination"
                    .to_string(),
            ));
        }
        self.pos += 1;
        if self
            .picture_compatibility_records
            .last()
            .is_some_and(|record| {
                record.position > self.body_text_len
                    || (record.position == self.body_text_len && record.kind == kind)
            })
        {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility wrappers are duplicated or out of body order"
                    .to_string(),
            ));
        }
        let index = self.picture_compatibility_records.len();
        self.picture_compatibility_records
            .push(crate::PictureCompatibilityRecord {
                position: self.body_text_len,
                kind,
                picture_index,
            });
        self.body_story_events.push(ParsedBodyStoryEvent::Resolved(
            crate::BodyStoryEvent::PictureCompatibility(index),
        ));
        Ok(())
    }

    /// Parse picture/image content.
    ///
    /// Pictures in RTF have the format:
    /// {\pict\emfblip\picw<width>\pich<height>...<hex data>}
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn parse_picture(&mut self) -> RtfResult<()> {
        let shape_properties = self.scan_picture_shape_properties()?;
        self.pos += 1; // Skip \pict

        let mut image_type = super::super::super::picture::ImageType::Unknown;
        let mut width = None;
        let mut height = None;
        let mut goal_width = None;
        let mut goal_height = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut scaled = false;
        let mut crop = crate::PictureCrop::default();
        let mut bitmap = crate::PictureBitmapMetadata::default();
        let mut blip_tag = None;
        let mut blip_upi = None;
        let mut blip_uid = None;
        let mut identity_stage = 0u8;
        let mut data_started = false;
        let mut data = Vec::new();
        let mut high_nibble = None;

        // Parse picture properties and data
        while let Some(token) = self.tokens.get(self.pos) {
            match token {
                Token::CloseBrace => {
                    break;
                },
                Token::Control(control) => {
                    self.pos += 1;
                    match control {
                        ControlWord::Emfblip => {
                            image_type = super::super::super::picture::ImageType::Emf;
                        },
                        ControlWord::Pngblip => {
                            image_type = super::super::super::picture::ImageType::Png;
                        },
                        ControlWord::Jpegblip => {
                            image_type = super::super::super::picture::ImageType::Jpeg;
                        },
                        ControlWord::Macpict => {
                            image_type = super::super::super::picture::ImageType::Pict;
                        },
                        ControlWord::Wmetafile(_) | ControlWord::Pmmetafile(_) => {
                            image_type = super::super::super::picture::ImageType::Wmf;
                        },
                        ControlWord::Dibitmap(_) => {
                            image_type = super::super::super::picture::ImageType::Dib;
                        },
                        ControlWord::Wbitmap(_) => {
                            image_type = super::super::super::picture::ImageType::Dib;
                            bitmap.windows_bitmap = true;
                        },
                        ControlWord::PictureWidth(w) => width = Some(*w),
                        ControlWord::PictureHeight(h) => height = Some(*h),
                        ControlWord::PictureGoalWidth(w) => goal_width = Some(*w),
                        ControlWord::PictureGoalHeight(h) => goal_height = Some(*h),
                        ControlWord::PictureScaleX(s) => scale_x = Some(*s),
                        ControlWord::PictureScaleY(s) => scale_y = Some(*s),
                        ControlWord::PictureScaled(parameter) => {
                            if data_started || scaled || parameter.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF picscaled is duplicated, parameterized, or late"
                                        .to_string(),
                                ));
                            }
                            scaled = true;
                        },
                        ControlWord::PictureBitmap(parameter) => {
                            if data_started || bitmap.bitmap_source || parameter.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF picbmp is duplicated, parameterized, or late".to_string(),
                                ));
                            }
                            bitmap.bitmap_source = true;
                        },
                        ControlWord::PictureBitsPerPixel(value) => {
                            if data_started || bitmap.bits_per_pixel.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF picbpp is duplicated or late".to_string(),
                                ));
                            }
                            bitmap.bits_per_pixel =
                                Some(Self::positive_picture_u16(*value, "picbpp")?);
                        },
                        ControlWord::WindowsBitmapBitsPerPixel(value) => {
                            if data_started || bitmap.windows_bits_per_pixel.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF wbmbitspixel is duplicated or late".to_string(),
                                ));
                            }
                            bitmap.windows_bits_per_pixel =
                                Some(Self::positive_picture_u16(*value, "wbmbitspixel")?);
                        },
                        ControlWord::WindowsBitmapPlanes(value) => {
                            if data_started || bitmap.planes.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF wbmplanes is duplicated or late".to_string(),
                                ));
                            }
                            bitmap.planes = Some(Self::positive_picture_u16(*value, "wbmplanes")?);
                        },
                        ControlWord::WindowsBitmapWidthBytes(value) => {
                            if data_started || bitmap.width_bytes.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF wbmwidthbytes is duplicated or late".to_string(),
                                ));
                            }
                            let value = value.ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF wbmwidthbytes requires a parameter".to_string(),
                                )
                            })?;
                            bitmap.width_bytes = Some(
                                u32::try_from(value)
                                    .ok()
                                    .filter(|value| *value != 0)
                                    .ok_or_else(|| {
                                        RtfError::MalformedDocument(
                                            "RTF wbmwidthbytes must be positive".to_string(),
                                        )
                                    })?,
                            );
                        },
                        ControlWord::PictureCropLeft(value) => Self::set_picture_crop(
                            &mut crop.left,
                            *value,
                            data_started,
                            "piccropl",
                        )?,
                        ControlWord::PictureCropRight(value) => Self::set_picture_crop(
                            &mut crop.right,
                            *value,
                            data_started,
                            "piccropr",
                        )?,
                        ControlWord::PictureCropTop(value) => {
                            Self::set_picture_crop(
                                &mut crop.top,
                                *value,
                                data_started,
                                "piccropt",
                            )?;
                        },
                        ControlWord::PictureCropBottom(value) => Self::set_picture_crop(
                            &mut crop.bottom,
                            *value,
                            data_started,
                            "piccropb",
                        )?,
                        ControlWord::BlipTag(value) => {
                            if data_started || blip_tag.is_some() || identity_stage > 1 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF bliptag is duplicated, late, or out of order".to_string(),
                                ));
                            }
                            blip_tag = Some(*value);
                            identity_stage = 1;
                        },
                        ControlWord::BlipUnitsPerInch(value) => {
                            if data_started || blip_upi.is_some() || identity_stage > 2 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF blipupi is duplicated, late, or out of order".to_string(),
                                ));
                            }
                            let value = u16::try_from(*value).map_err(|_err| {
                                RtfError::MalformedDocument(
                                    "RTF blipupi is outside 1..=65535".to_string(),
                                )
                            })?;
                            if value == 0 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF blipupi must be positive".to_string(),
                                ));
                            }
                            blip_upi = Some(value);
                            identity_stage = 2;
                        },
                        ControlWord::BlipUid => {
                            return Err(RtfError::MalformedDocument(
                                "RTF blipuid destination must be starred and grouped".to_string(),
                            ));
                        },
                        _ => {},
                    }
                },
                Token::Text(text) => {
                    data_started |= text.bytes().any(|byte| !byte.is_ascii_whitespace());
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF picture contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            data.push((high << 4) | nibble);
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    if data.len() > MAX_PICTURE_DATA_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture data exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Token::Binary(bytes) => {
                    data_started = true;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture binary payload splits a hexadecimal byte".to_string(),
                        ));
                    }
                    data.extend_from_slice(bytes);
                    if data.len() > MAX_PICTURE_DATA_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture data exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Token::OpenBrace => {
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::BlipUid),
                        ])
                    ) {
                        if data_started || blip_uid.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF blipuid is duplicated or occurs after picture data"
                                    .to_string(),
                            ));
                        }
                        blip_uid = Some(self.parse_picture_uid()?);
                        identity_stage = 3;
                    } else if matches!(
                        self.tokens.get(self.pos..self.pos + 2),
                        Some([Token::OpenBrace, Token::Control(ControlWord::BlipUid)])
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF blipuid destination must be starred".to_string(),
                        ));
                    } else {
                        self.skip_group()?;
                    }
                },
            }
        }

        if high_nibble.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF picture contains an odd number of hexadecimal digits".to_string(),
            ));
        }

        if !data.is_empty() {
            // If type not specified, try to detect from data
            if image_type == super::super::super::picture::ImageType::Unknown {
                image_type = super::super::super::picture::detect_image_type(&data);
            }

            // Allocate in arena and create picture
            let data_alloc = self.arena.alloc_slice_copy(&data);
            let mut picture =
                super::super::super::picture::Picture::new(image_type, Cow::Borrowed(data_alloc));
            picture.width = width;
            picture.height = height;
            picture.goal_width = goal_width;
            picture.goal_height = goal_height;
            picture.scale_x = scale_x;
            picture.scale_y = scale_y;
            picture.scaled = scaled;
            picture.crop = crop;
            picture.bitmap = bitmap;
            if blip_tag.is_some() || blip_upi.is_some() || blip_uid.is_some() {
                let identity = super::super::super::picture::PictureIdentity {
                    tag: blip_tag,
                    units_per_inch: blip_upi,
                    uid: blip_uid.map(|uid| Cow::Borrowed(self.arena.alloc_slice_copy(&uid))),
                };
                identity.validate()?;
                picture.identity = Some(identity);
            }

            picture.shape_properties = shape_properties;
            picture.validate()?;
            self.pictures.push(picture);
        }

        Ok(())
    }

    pub(super) fn positive_picture_u16(value: Option<i32>, name: &str) -> RtfResult<u16> {
        value
            .and_then(|value| u16::try_from(value).ok())
            .filter(|value| *value != 0)
            .ok_or_else(|| {
                RtfError::MalformedDocument(format!(
                    "RTF {name} must have a positive 16-bit parameter"
                ))
            })
    }

    pub(super) fn set_picture_crop(
        slot: &mut Option<i32>,
        value: Option<i32>,
        data_started: bool,
        name: &str,
    ) -> RtfResult<()> {
        if data_started || slot.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {name} is duplicated or occurs after picture data"
            )));
        }
        *slot = Some(value.ok_or_else(|| {
            RtfError::MalformedDocument(format!("RTF {name} requires a parameter"))
        })?);
        Ok(())
    }

    pub(super) fn scan_picture_shape_properties(
        &mut self,
    ) -> RtfResult<Option<crate::PictureShapeProperties<'a>>> {
        let original_pos = self.pos;
        let result = self.scan_picture_shape_properties_inner();
        self.pos = original_pos;
        result
    }

    pub(super) fn scan_picture_shape_properties_inner(
        &mut self,
    ) -> RtfResult<Option<crate::PictureShapeProperties<'a>>> {
        let mut index = self.pos + 1;
        let mut depth = 0usize;
        let mut properties = None;
        let mut saw_picture_content = false;
        while index < self.tokens.len() {
            match self.tokens.get(index) {
                Some(Token::OpenBrace) if depth == 0 => {
                    let starred = matches!(
                        self.tokens.get(index + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    );
                    let control_index = index + if starred { 2 } else { 1 };
                    if matches!(
                        self.tokens.get(control_index),
                        Some(Token::Control(ControlWord::PictureProperties(_)))
                    ) {
                        if !starred {
                            return Err(RtfError::MalformedDocument(
                                "RTF picprop destination must be starred".to_string(),
                            ));
                        }
                        if properties.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF picture contains multiple picprop destinations".to_string(),
                            ));
                        }
                        if saw_picture_content {
                            return Err(RtfError::MalformedDocument(
                                "RTF picprop must precede picture payload data".to_string(),
                            ));
                        }
                        self.pos = index;
                        properties = Some(self.parse_picture_shape_properties()?);
                        index = self.pos;
                        continue;
                    }
                    depth = 1;
                    index += 1;
                },
                Some(Token::OpenBrace) => {
                    depth += 1;
                    index += 1;
                },
                Some(Token::CloseBrace) if depth == 0 => break,
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    index += 1;
                },
                Some(Token::Control(ControlWord::PictureProperties(_))) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF picprop destination must be a direct child group of pict".to_string(),
                    ));
                },
                Some(Token::Text(value))
                    if value.bytes().all(|byte| byte.is_ascii_whitespace()) =>
                {
                    index += 1;
                },
                Some(Token::Text(_) | Token::Binary(_)) => {
                    if depth == 0 {
                        saw_picture_content = true;
                    }
                    index += 1;
                },
                Some(_) => {
                    index += 1;
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        Ok(properties)
    }

    pub(super) fn parse_picture_shape_properties(
        &mut self,
    ) -> RtfResult<crate::PictureShapeProperties<'a>> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || !matches!(
                self.tokens.get(self.pos + 1),
                Some(Token::Control(ControlWord::IgnorableDestination))
            )
        {
            return Err(RtfError::MalformedDocument(
                "invalid RTF picprop destination".to_string(),
            ));
        }
        match self.tokens.get(self.pos + 2) {
            Some(Token::Control(ControlWord::PictureProperties(None))) => self.pos += 3,
            Some(Token::Control(ControlWord::PictureProperties(Some(_)))) => {
                return Err(RtfError::MalformedDocument(
                    "RTF picprop destination must not have a parameter".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF picprop destination".to_string(),
                ));
            },
        }

        let mut result = crate::PictureShapeProperties::default();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::Text(value))
                    if value.bytes().all(|byte| byte.is_ascii_whitespace()) =>
                {
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::PictureShapeId(Some(value)))) => {
                    if result.shape_id.replace(*value).is_some() || !result.properties.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF shplid must occur at most once before picture properties"
                                .to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::PictureShapeId(None))) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF shplid requires a parameter".to_string(),
                    ));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ShapeProperty))
                    ) =>
                {
                    if result.properties.len() >= crate::MAX_PICTURE_SHAPE_PROPERTIES {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture shape-property count exceeds the safety limit".to_string(),
                        ));
                    }
                    result.properties.push(self.parse_picture_shape_property()?);
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    result.validate()?;
                    return Ok(result);
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content or ordering in RTF picprop destination".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    pub(super) fn parse_picture_shape_property(&mut self) -> RtfResult<crate::ShapeProperty<'a>> {
        self.pos += 2; // opening brace and \sp
        self.skip_picture_property_whitespace();
        let name = self.parse_picture_property_text(ControlWord::ShapePropertyName, "sn")?;
        self.skip_picture_property_whitespace();
        let (value, binary_value) = self.parse_picture_property_value()?;
        self.skip_picture_property_whitespace();
        let theme_value = if matches!(
            self.tokens.get(self.pos..self.pos + 3),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::ShapeThemeValue(_)),
            ])
        ) {
            let value = self.parse_shape_theme_value()?;
            self.skip_picture_property_whitespace();
            Some(value)
        } else {
            None
        };
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF picture sp must contain one sn, one sv, and at most one trailing hsv"
                    .to_string(),
            ));
        }
        self.pos += 1;
        Ok(crate::ShapeProperty {
            name: Cow::Borrowed(self.arena.alloc_str(&name)),
            value: Cow::Borrowed(self.arena.alloc_str(&value)),
            binary_value: binary_value.map(Cow::Owned),
            theme_value,
            hyperlink: None,
        })
    }

    pub(super) fn parse_picture_property_value(&mut self) -> RtfResult<(String, Option<Vec<u8>>)> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || !matches!(
                self.tokens.get(self.pos + 1),
                Some(Token::Control(ControlWord::ShapePropertyValue))
            )
        {
            return Err(RtfError::MalformedDocument(
                "RTF picture sp requires an unstarred sv destination".to_string(),
            ));
        }
        self.pos += 2;
        let mut text = String::new();
        let mut binary_value = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if binary_value.is_some() && !text.trim().is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture property cannot mix scalar and binary values".to_string(),
                        ));
                    }
                    return Ok((text.trim().to_string(), binary_value));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) && matches!(
                        self.tokens.get(self.pos + 2),
                        Some(Token::Control(ControlWord::ShapeBinaryValue(_)))
                    ) =>
                {
                    if binary_value.is_some() || !text.trim().is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF svb must be the only picture property value".to_string(),
                        ));
                    }
                    binary_value = Some(self.parse_shape_binary_value()?);
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "nested groups are not allowed in RTF picture sv".to_string(),
                    ));
                },
                Some(Token::Control(ControlWord::Unicode(code))) if binary_value.is_none() => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                },
                Some(Token::Control(control))
                    if binary_value.is_none() && control_symbol_text(control).is_some() =>
                {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "active controls are not allowed in RTF picture sv".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    /// Parse the `\hl` shape-hyperlink group inside a shape property.
    ///
    /// Expects `self.pos` at the group's opening brace and consumes tokens
    /// through its closing brace. The `\hlloc`, `\hlsrc`, and `\hlfr` string
    /// groups may appear in any order (RTF "Hyperlink Property for Shapes").
    pub(super) fn parse_shape_hyperlink_destination(
        &mut self,
    ) -> RtfResult<crate::ShapeHyperlink<'a>> {
        self.pos += 2; // opening brace and hl control
        let mut hyperlink = crate::ShapeHyperlink::default();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    let slot = match self.tokens.get(self.pos + 1) {
                        Some(Token::Control(ControlWord::ShapeHyperlinkLocation)) => {
                            &mut hyperlink.location
                        },
                        Some(Token::Control(ControlWord::ShapeHyperlinkSource)) => {
                            &mut hyperlink.source
                        },
                        Some(Token::Control(ControlWord::ShapeHyperlinkFriendlyName)) => {
                            &mut hyperlink.friendly_name
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF hl may contain only hlloc, hlsrc, and hlfr destinations"
                                    .to_string(),
                            ));
                        },
                    };
                    if slot.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF shape-hyperlink string destination".to_string(),
                        ));
                    }
                    *slot = Some(self.parse_shape_hyperlink_string()?);
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {
                    self.pos += 1;
                },
                _ => {
                    return Err(RtfError::MalformedDocument(
                        "RTF hl contains ungrouped, binary, or active data".to_string(),
                    ));
                },
            }
        }
        hyperlink.validate()?;
        Ok(hyperlink)
    }

    /// Parse one `\hlloc`/`\hlsrc`/`\hlfr` string destination group;
    /// `self.pos` is at its opening brace.
    pub(super) fn parse_shape_hyperlink_string(&mut self) -> RtfResult<Cow<'a, str>> {
        self.pos += 2; // opening brace and destination control
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0).cast_unsigned() as usize;
        let mut fallback_skip = 0usize;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                None => return Err(RtfError::UnexpectedEof),
                _ => {
                    if !self.consume_destination_text_token(
                        &mut value,
                        &mut unicode_skip,
                        &mut fallback_skip,
                        "shape-hyperlink string",
                    )? {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-hyperlink string destination contains grouped, binary, or active data"
                                .to_string(),
                        ));
                    }
                    if value.len() > crate::shape::MAX_SHAPE_HYPERLINK_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape-hyperlink string exceeds the safety limit".to_string(),
                        ));
                    }
                },
            }
        }
        Ok(Cow::Owned(value.trim_end_matches(['\r', '\n']).to_string()))
    }

    pub(super) fn skip_picture_property_whitespace(&mut self) {
        while matches!(
            self.tokens.get(self.pos),
            Some(Token::Text(value)) if value.bytes().all(|byte| byte.is_ascii_whitespace())
        ) {
            self.pos += 1;
        }
    }

    pub(super) fn parse_picture_property_text(
        &mut self,
        expected: ControlWord<'_>,
        name: &str,
    ) -> RtfResult<String> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || self.tokens.get(self.pos + 1) != Some(&Token::Control(expected))
        {
            return Err(RtfError::MalformedDocument(format!(
                "RTF picture sp requires an unstarred {name} destination"
            )));
        }
        self.pos += 2;
        let mut text = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(text.trim().to_string());
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(format!(
                        "nested groups are not allowed in RTF picture {name}"
                    )));
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(format!(
                        "active controls are not allowed in RTF picture {name}"
                    )));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    pub(super) fn parse_picture_uid(&mut self) -> RtfResult<Vec<u8>> {
        self.pos += 3; // opening brace, ignorable marker, and blipuid
        let mut bytes = Vec::with_capacity(16);
        let mut high_nibble = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF blipuid contains an odd number of hexadecimal digits".to_string(),
                        ));
                    }
                    if !bytes.is_empty() && bytes.len() != 16 {
                        return Err(RtfError::MalformedDocument(
                            "RTF blipuid must contain exactly 16 bytes or be empty".to_string(),
                        ));
                    }
                    return Ok(bytes);
                },
                Some(Token::Text(text)) => {
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF blipuid contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            bytes.push((high << 4) | nibble);
                            if bytes.len() > 16 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF blipuid exceeds 16 bytes".to_string(),
                                ));
                            }
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Control(_) | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF blipuid contains active, nested, or binary content".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }
}
