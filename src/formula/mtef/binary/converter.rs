//! MTEF to AST conversion logic
//!
//! This module implements the conversion from parsed MTEF objects to formula AST nodes.
//! Based on rtf2latex2e Eqn_TranslateObjects and related conversion functions.
//!
//! The conversion process involves:
//! - Character translation using typeface lookup tables
//! - Template parsing and AST node construction
//! - Embellishment application
//! - Mode switching (math/text) based on typeface attributes

use super::charset::*;
use super::objects::*;
use crate::formula::ast::{Fence, LargeOperator, LineStyle, MathNode, MatrixFence};
use crate::formula::mtef::MtefError;
use crate::formula::mtef::constants::*;
use crate::formula::mtef::templates::{TemplateArgs, TemplateParser};
use std::borrow::Cow;

/// Type alias for subscript/superscript parsing result (base, subscript, superscript)
type SubSupResult<'a> =
    Result<(Vec<MathNode<'a>>, Vec<MathNode<'a>>, Vec<MathNode<'a>>), MtefError>;

/// Type alias for large operator parsing result (lower_limit, upper_limit, integrand)
type LargeOpResult<'a> = Result<
    (
        Option<Vec<MathNode<'a>>>,
        Option<Vec<MathNode<'a>>>,
        Vec<MathNode<'a>>,
    ),
    MtefError,
>;

/// Implementation of AST conversion methods for MtefBinaryParser
impl<'arena> super::parser::MtefBinaryParser<'arena> {
    pub fn convert_objects_to_ast(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<Vec<MathNode<'arena>>, MtefError> {
        let mut nodes = Vec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            match obj.tag {
                MtefRecordType::Char => {
                    if let Some(char_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefChar>() {
                        // Special handling based on rtf2latex2e Eqn_TranslateObjects logic
                        match char_obj.typeface {
                            130 => {
                                // Function typeface - auto-recognize functions
                                let (node, skip_count) = self.convert_function_to_node(current)?;
                                nodes.push(node);
                                // Skip the consumed characters
                                for _ in 0..skip_count {
                                    current = current.and_then(|c| c.next.as_deref());
                                }
                                continue;
                            },
                            129 if self.mode != crate::formula::mtef::constants::EQN_MODE_TEXT => {
                                // Text in math mode
                                let (node, skip_count) = self.convert_text_run_to_node(current)?;
                                nodes.push(node);
                                // Skip the consumed characters
                                for _ in 0..skip_count {
                                    current = current.and_then(|c| c.next.as_deref());
                                }
                                continue;
                            },
                            _ => {
                                // Regular character
                                nodes.push(self.convert_char_to_node(char_obj)?);
                            },
                        }
                    }
                },
                MtefRecordType::Tmpl => {
                    if let Some(tmpl_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefTemplate>() {
                        nodes.push(self.convert_template_to_node(tmpl_obj)?);
                    }
                },
                MtefRecordType::Line => {
                    if let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                        && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
                    {
                        nodes.extend(line_nodes);
                    }
                },
                MtefRecordType::Pile => {
                    if let Some(pile_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefPile>() {
                        nodes.push(self.convert_pile_to_node(pile_obj)?);
                    }
                },
                MtefRecordType::Matrix => {
                    if let Some(matrix_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefMatrix>() {
                        nodes.push(self.convert_matrix_to_node(matrix_obj)?);
                    }
                },
                MtefRecordType::Font => {
                    // Font objects affect character rendering but don't generate output
                    // In a full implementation, this would update the current font context
                },
                MtefRecordType::Size
                | MtefRecordType::Full
                | MtefRecordType::Sub
                | MtefRecordType::Sub2
                | MtefRecordType::Sym
                | MtefRecordType::SubSym => {
                    // Size objects affect character size but don't generate output
                    // In a full implementation, this would update the current size context
                },
                _ => {
                    // Skip other record types for now
                },
            }
            current = obj.next.as_deref();
        }

        Ok(nodes)
    }

    fn convert_char_to_node(&self, char_obj: &MtefChar) -> Result<MathNode<'arena>, MtefError> {
        let text = self.convert_char_to_text(char_obj).map_err(|e| {
            MtefError::ParseError(format!(
                "Failed to convert character (typeface={}, char={}): {}",
                char_obj.typeface, char_obj.character, e
            ))
        })?;
        Ok(MathNode::Text(text))
    }

    /// Convert a function sequence to a MathNode (handles typeface 130 functions)
    fn convert_function_to_node(
        &self,
        start_obj: Option<&MtefObjectList>,
    ) -> Result<(MathNode<'arena>, usize), MtefError> {
        use crate::formula::mtef::binary::charset::lookup_function;

        let mut function_name = String::new();
        let mut current = start_obj;
        let mut skip_count = 0;

        // Gather function name from consecutive characters with typeface 130
        while let Some(obj) = current {
            if let MtefRecordType::Char = obj.tag
                && let Some(char_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefChar>()
                && char_obj.typeface == 130
                && (char_obj.character as u8).is_ascii_alphabetic()
                && let Some(ch) = char::from_u32(char_obj.character as u32)
            {
                function_name.push(ch);
                skip_count += 1;
                current = obj.next.as_deref();
                continue;
            }
            break;
        }

        if function_name.is_empty() {
            return Err(MtefError::ParseError("Empty function name".to_string()));
        }

        // Look up the function in the table
        let latex_text = if let Some(func) = lookup_function(&function_name) {
            Cow::Borrowed(func.trim_end()) // Remove trailing space
        } else {
            // Fallback: wrap in \mathrm{}
            Cow::Owned(format!("\\mathrm{{{}}}", function_name))
        };

        Ok((MathNode::Text(latex_text), skip_count))
    }

    /// Convert a text run to a MathNode (handles typeface 129 text in math)
    fn convert_text_run_to_node(
        &self,
        start_obj: Option<&MtefObjectList>,
    ) -> Result<(MathNode<'arena>, usize), MtefError> {
        let mut text_run = String::new();
        let mut current = start_obj;
        let mut skip_count = 0;

        // Gather text from consecutive characters with typeface 129, also skip SIZE objects
        while let Some(obj) = current {
            match obj.tag {
                MtefRecordType::Char => {
                    if let Some(char_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefChar>()
                        && char_obj.typeface == 129
                        && let Some(ch) = char::from_u32(char_obj.character as u32)
                    {
                        text_run.push(ch);
                        skip_count += 1;
                        current = obj.next.as_deref();
                        continue;
                    }
                    break;
                },
                MtefRecordType::Size
                | MtefRecordType::Full
                | MtefRecordType::Sub
                | MtefRecordType::Sub2
                | MtefRecordType::Sym
                | MtefRecordType::SubSym => {
                    // Skip size objects
                    skip_count += 1;
                    current = obj.next.as_deref();
                    continue;
                },
                _ => break,
            }
        }

        // Wrap text in \text{} for LaTeX
        let latex_text = format!("\\text{{{}}}", text_run);
        Ok((MathNode::Text(Cow::Owned(latex_text)), skip_count))
    }

    fn convert_char_to_text(&self, char_obj: &MtefChar) -> Result<Cow<'arena, str>, MtefError> {
        // Implement proper character translation based on rtf2latex2e Eqn_GetTexChar logic
        let typeface = char_obj.typeface as usize;
        let character = char_obj.character;

        let mut _math_attr = 0; // Default math attribute (MA_NONE)
        let mut current_mode = self.mode; // Current mode for this character

        // Get base character representation
        let mut base_text = if (129..129 + NUM_TYPEFACE_SLOTS).contains(&typeface) {
            let charset_index = typeface - 129;
            let charset_atts = get_charset_attributes(charset_index);

            _math_attr = charset_atts.math_attr;

            // Handle mode switching based on _math_attr (following rtf2latex2e logic)
            let _mode_changed = match _math_attr {
                MA_FORCE_TEXT => {
                    let old_mode = current_mode;
                    current_mode = EQN_MODE_TEXT;
                    Some(old_mode)
                },
                MA_FORCE_MATH => {
                    let old_mode = current_mode;
                    // For forced math mode, use inline if equation is inline, otherwise display
                    current_mode = if self.inline != 0 {
                        EQN_MODE_INLINE
                    } else {
                        EQN_MODE_DISPLAY
                    };
                    Some(old_mode)
                },
                MA_TEXT | MA_MATH => {
                    // For special case: mode depends on variation (like spaces)
                    if typeface == 152 && _math_attr == MA_TEXT {
                        let old_mode = current_mode;
                        current_mode = EQN_MODE_TEXT;
                        Some(old_mode)
                    } else if typeface == 152 && _math_attr == MA_MATH {
                        let old_mode = current_mode;
                        current_mode = if self.inline != 0 {
                            EQN_MODE_INLINE
                        } else {
                            EQN_MODE_DISPLAY
                        };
                        Some(old_mode)
                    } else {
                        None
                    }
                },
                _ => None,
            };

            // Try character lookup first using PHF map
            let lookup_result = if charset_atts.do_lookup {
                // Special handling for typefaces with mode-dependent lookups
                let lookup_math_attr = if typeface == 152 {
                    // Space characters have different meanings in math vs text
                    _math_attr
                } else {
                    _math_attr
                };

                lookup_character(typeface, character, lookup_math_attr)
            } else {
                None
            };

            if let Some(latex_char) = lookup_result {
                latex_char.to_string()
            } else if charset_atts.use_codepoint {
                self.convert_codepoint(character, typeface)?.to_string()
            } else {
                format!("\\char{}", character)
            }
        } else {
            // Fallback for unknown typefaces
            format!("\\char{}", character)
        };

        // Apply embellishments if present (following rtf2latex2e logic)
        if let Some(embellishments) = &char_obj.embellishment_list {
            self.apply_embellishments(&mut base_text, embellishments, current_mode)?;
        }

        Ok(Cow::Owned(base_text))
    }

    fn apply_embellishments(
        &self,
        base_text: &mut String,
        embellishments: &MtefEmbell,
        mode: i32,
    ) -> Result<(), MtefError> {
        // Apply embellishments to the base character, following rtf2latex2e Eqn_GetTexChar logic
        let mut current = Some(embellishments);

        while let Some(embell) = current {
            if embell.embell > 0
                && usize::from(embell.embell)
                    < crate::formula::mtef::binary::charset::EMBELLISHMENT_TEMPLATES.len()
            {
                let template = get_embellishment_template(embell.embell);
                if !template.is_empty() {
                    // Split template on comma to get math and text versions
                    // Use appropriate version based on current mode
                    let template_part = if let Some(comma_pos) = template.find(',') {
                        if mode != EQN_MODE_TEXT {
                            &template[..comma_pos] // Math version
                        } else {
                            &template[comma_pos + 1..] // Text version
                        }
                    } else {
                        template // Whole template if no comma
                    };

                    // Replace %1 with the base character
                    let new_text = template_part.replace("%1", base_text);
                    *base_text = new_text;
                }
            }
            current = embell.next.as_deref();
        }

        Ok(())
    }

    fn convert_codepoint(
        &self,
        character: u16,
        typeface: usize,
    ) -> Result<Cow<'arena, str>, MtefError> {
        // Handle special characters and formatting based on rtf2latex2e logic
        if (32..=127).contains(&character) {
            let ch = character as u8 as char;

            // Special handling for ampersand
            if character == 38 {
                // '&'
                return Ok(Cow::Borrowed("\\&"));
            }

            // Special handling for certain typefaces (like bold)
            if typeface == 135 {
                // Bold typeface - matches rtf2latex2e logic
                return Ok(Cow::Owned(format!("\\mathbf{{{}}}", ch)));
            }

            // Regular character
            return Ok(Cow::Owned(ch.to_string()));
        }

        // For non-ASCII characters, try to convert as Unicode
        if let Some(c) = char::from_u32(character as u32) {
            Ok(Cow::Owned(c.to_string()))
        } else {
            // Fallback for unmappable characters
            Ok(Cow::Owned(format!("\\char{}", character)))
        }
    }

    fn convert_template_to_node(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Handle templates based on selector type
        // Some templates have specific AST representations, others use generic template parsing
        match tmpl_obj.selector {
            0..=9 => {
                // Fences (parentheses, brackets, braces, etc.)
                self.convert_fence_template(tmpl_obj)
            },
            10 => {
                // Root
                self.convert_legacy_template(tmpl_obj)
            },
            11 => {
                // Fraction
                self.convert_legacy_template(tmpl_obj)
            },
            12..=13 => {
                // Underline/overline
                self.convert_decoration_template(tmpl_obj)
            },
            14 => {
                // Arrows
                self.convert_arrow_template(tmpl_obj)
            },
            15 | 21 => {
                // Integrals
                self.convert_legacy_template(tmpl_obj)
            },
            16..=20 => {
                // Large operators (sum, product, etc.)
                self.convert_large_op_template(tmpl_obj)
            },
            22 => {
                // Sum (alternate form)
                self.convert_large_op_template(tmpl_obj)
            },
            23 => {
                // Limit
                self.convert_limit_template(tmpl_obj)
            },
            24..=25 => {
                // Horizontal braces
                self.convert_brace_template(tmpl_obj)
            },
            27..=29 => {
                // Scripts (subscript, superscript, sub+sup)
                self.convert_legacy_template(tmpl_obj)
            },
            _ => {
                // Try template table lookup for unknown templates
                let variation = tmpl_obj.variation;
                if let Some(template_def) =
                    TemplateParser::find_template(tmpl_obj.selector, variation)
                {
                    // Parse subobjects into arguments
                    let args = if let Some(obj_list) = &tmpl_obj.subobject_list {
                        self.parse_template_arguments(obj_list)?
                    } else {
                        smallvec::SmallVec::new()
                    };

                    // Apply the template
                    Ok(TemplateParser::parse_template_arguments(
                        template_def.template,
                        &args,
                    ))
                } else {
                    // Fallback for completely unknown templates
                    Ok(MathNode::Text(Cow::Owned(format!(
                        "\\unknown_template_{}_{{{}}}",
                        tmpl_obj.selector, tmpl_obj.variation
                    ))))
                }
            },
        }
    }

    fn convert_legacy_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Template handling based on MTEF selector values from rtf2latex2e
        match tmpl_obj.selector {
            14 => {
                // Fraction (ffract)
                // Fraction template - should have numerator and denominator subobjects
                if let Some(obj_list) = &tmpl_obj.subobject_list {
                    let (numerator, denominator) = self.parse_fraction_subobjects(obj_list)?;
                    Ok(TemplateParser::parse_fraction(numerator, denominator))
                } else {
                    Ok(MathNode::Text(Cow::Borrowed("\\frac{}{}")))
                }
            },
            13 => {
                // Root (sqroot/nthroot)
                // Root template - may have index and base
                if let Some(obj_list) = &tmpl_obj.subobject_list {
                    let (base, index) = self.parse_root_subobjects(obj_list)?;
                    Ok(TemplateParser::parse_root(
                        base,
                        if index.is_empty() { None } else { Some(index) },
                    ))
                } else {
                    Ok(MathNode::Text(Cow::Borrowed("\\sqrt{}")))
                }
            },
            15 => {
                // Scripts (super, sub, subsup based on variation)
                match tmpl_obj.variation {
                    0 => {
                        // Superscript
                        if let Some(obj_list) = &tmpl_obj.subobject_list {
                            let (base, superscript) =
                                self.parse_superscript_subobjects(obj_list)?;
                            Ok(TemplateParser::parse_superscript(base, superscript))
                        } else {
                            Ok(MathNode::Text(Cow::Borrowed("^{}")))
                        }
                    },
                    1 => {
                        // Subscript
                        if let Some(obj_list) = &tmpl_obj.subobject_list {
                            let (base, subscript) = self.parse_subscript_subobjects(obj_list)?;
                            Ok(TemplateParser::parse_subscript(base, subscript))
                        } else {
                            Ok(MathNode::Text(Cow::Borrowed("_{}")))
                        }
                    },
                    2 => {
                        // Sub+Sup
                        if let Some(obj_list) = &tmpl_obj.subobject_list {
                            let (base, subscript, superscript) =
                                self.parse_subsup_subobjects(obj_list)?;
                            Ok(TemplateParser::parse_subsup(base, subscript, superscript))
                        } else {
                            Ok(MathNode::Text(Cow::Borrowed("_{}^{}")))
                        }
                    },
                    _ => Ok(MathNode::Text(Cow::Borrowed("_{}^{}"))), // fallback
                }
            },
            21 => {
                // Integrals
                // For now, just create a simple integral node
                // This should be expanded to handle limits properly
                if let Some(obj_list) = &tmpl_obj.subobject_list {
                    let integrand = self.parse_single_subobject(obj_list)?;
                    Ok(MathNode::LargeOp {
                        operator: crate::formula::ast::LargeOperator::Integral,
                        lower_limit: None,
                        upper_limit: None,
                        integrand: Some(integrand),
                        hide_lower: true,
                        hide_upper: true,
                    })
                } else {
                    Ok(MathNode::Text(Cow::Borrowed("\\int ")))
                }
            },
            _ => {
                // Unknown template - return as placeholder
                Ok(MathNode::Text(Cow::Owned(format!(
                    "\\unknown_template_{}_{{{}}}",
                    tmpl_obj.selector, tmpl_obj.variation
                ))))
            },
        }
    }

    fn parse_template_arguments(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<TemplateArgs<'arena>, MtefError> {
        // Parse template arguments from subobjects
        // This follows the rtf2latex2e pattern where arguments are separated by LINE objects
        let mut args = TemplateArgs::new();
        let mut current_arg = smallvec::SmallVec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            match obj.tag {
                MtefRecordType::Line => {
                    if let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                        && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
                    {
                        current_arg.extend(line_nodes);
                    }
                },
                MtefRecordType::Pile => {
                    // Piles can separate arguments
                    if !current_arg.is_empty() {
                        args.push(current_arg);
                        current_arg = smallvec::SmallVec::new();
                    }
                    if let Some(pile_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefPile>() {
                        let pile_node = self.convert_pile_to_node(pile_obj)?;
                        current_arg.push(pile_node);
                    }
                },
                _ => {
                    // Other objects go into current argument
                    let nodes = self.convert_single_object_to_ast(obj)?;
                    current_arg.extend(nodes);
                },
            }
            current = obj.next.as_deref();
        }

        // Add the last argument if not empty
        if !current_arg.is_empty() {
            args.push(current_arg);
        }

        Ok(args)
    }

    fn parse_fraction_subobjects(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<(Vec<MathNode<'arena>>, Vec<MathNode<'arena>>), MtefError> {
        // Parse LINE objects as numerator and denominator
        let mut numerator = Vec::new();
        let mut denominator = Vec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if numerator.is_empty() {
                    numerator = line_nodes;
                } else {
                    denominator = line_nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((numerator, denominator))
    }

    fn parse_root_subobjects(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<(Vec<MathNode<'arena>>, Vec<MathNode<'arena>>), MtefError> {
        // Parse LINE objects as index and base
        let mut index = Vec::new();
        let mut base = Vec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if index.is_empty() {
                    index = line_nodes;
                } else {
                    base = line_nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((base, index))
    }

    fn parse_subscript_subobjects(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<(Vec<MathNode<'arena>>, Vec<MathNode<'arena>>), MtefError> {
        // Parse LINE objects as base and subscript
        let mut base = Vec::new();
        let mut subscript = Vec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if base.is_empty() {
                    base = line_nodes;
                } else {
                    subscript = line_nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((base, subscript))
    }

    fn parse_superscript_subobjects(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<(Vec<MathNode<'arena>>, Vec<MathNode<'arena>>), MtefError> {
        // Parse LINE objects as base and superscript
        let mut base = Vec::new();
        let mut superscript = Vec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if base.is_empty() {
                    base = line_nodes;
                } else {
                    superscript = line_nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((base, superscript))
    }

    fn parse_subsup_subobjects(&self, obj_list: &MtefObjectList) -> SubSupResult<'arena> {
        // Parse LINE objects as base, subscript, and superscript
        let mut base = Vec::new();
        let mut subscript = Vec::new();
        let mut superscript = Vec::new();
        let mut current = Some(obj_list);

        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if base.is_empty() {
                    base = line_nodes;
                } else if subscript.is_empty() {
                    subscript = line_nodes;
                } else {
                    superscript = line_nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((base, subscript, superscript))
    }

    fn parse_single_subobject(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<Vec<MathNode<'arena>>, MtefError> {
        // Parse a single subobject (typically for templates with one content area)
        let mut current = Some(obj_list);
        let mut result = Vec::new();

        while let Some(obj) = current {
            match obj.tag {
                MtefRecordType::Line => {
                    if let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                        && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
                    {
                        result.extend(line_nodes);
                    }
                },
                _ => {
                    // Convert other object types directly
                    let nodes = self.convert_single_object_to_ast(obj)?;
                    result.extend(nodes);
                },
            }
            current = obj.next.as_deref();
        }

        Ok(result)
    }

    fn convert_single_object_to_ast(
        &self,
        obj: &MtefObjectList,
    ) -> Result<Vec<MathNode<'arena>>, MtefError> {
        // Convert a single object to AST nodes
        let mut nodes = Vec::new();

        match obj.tag {
            MtefRecordType::Char => {
                if let Some(char_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefChar>() {
                    nodes.push(self.convert_char_to_node(char_obj)?);
                }
            },
            MtefRecordType::Tmpl => {
                if let Some(tmpl_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefTemplate>() {
                    nodes.push(self.convert_template_to_node(tmpl_obj)?);
                }
            },
            MtefRecordType::Pile => {
                if let Some(pile_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefPile>() {
                    nodes.push(self.convert_pile_to_node(pile_obj)?);
                }
            },
            MtefRecordType::Matrix => {
                if let Some(matrix_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefMatrix>() {
                    nodes.push(self.convert_matrix_to_node(matrix_obj)?);
                }
            },
            MtefRecordType::Font => {
                // Font objects affect character rendering but don't generate output
            },
            MtefRecordType::Size
            | MtefRecordType::Full
            | MtefRecordType::Sub
            | MtefRecordType::Sub2
            | MtefRecordType::Sym
            | MtefRecordType::SubSym => {
                // Size objects affect character size but don't generate output
            },
            _ => {
                // Skip other record types for now
            },
        }

        Ok(nodes)
    }

    fn convert_line_to_nodes(
        &self,
        line_obj: &MtefLine,
    ) -> Result<Option<Vec<MathNode<'arena>>>, MtefError> {
        if let Some(obj_list) = &line_obj.object_list {
            Ok(Some(self.convert_objects_to_ast(obj_list)?))
        } else {
            Ok(None)
        }
    }

    fn convert_pile_to_node(&self, pile_obj: &MtefPile) -> Result<MathNode<'arena>, MtefError> {
        // Convert pile to appropriate AST node
        // Piles are vertical stacks of elements, often used for fractions, limits, etc.
        if let Some(line_list) = &pile_obj.line_list {
            let mut rows = Vec::new();
            let mut current: Option<&MtefObjectList> = Some(line_list);

            while let Some(obj) = current {
                if obj.tag == MtefRecordType::Line
                    && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                {
                    if let Some(line_nodes) = self.convert_line_to_nodes(line_obj)? {
                        // Each line becomes a row in the pile
                        rows.push(vec![line_nodes]);
                    } else {
                        // Empty line - add empty row
                        rows.push(vec![Vec::new()]);
                    }
                }
                current = obj.next.as_deref();
            }

            if rows.len() == 1 {
                // Single row - just return the content
                Ok(MathNode::Row(
                    rows.into_iter().flatten().flatten().collect(),
                ))
            } else if rows.len() == 2 {
                // Two rows - could be a fraction or other binary operation
                // For now, represent as a simple vertical stack
                Ok(MathNode::Matrix {
                    rows,
                    fence_type: MatrixFence::None,
                    properties: None,
                })
            } else if !rows.is_empty() {
                // Multiple rows - create a matrix structure
                Ok(MathNode::Matrix {
                    rows,
                    fence_type: MatrixFence::None,
                    properties: None,
                })
            } else {
                Ok(MathNode::Text(Cow::Borrowed("\\pile")))
            }
        } else {
            Ok(MathNode::Text(Cow::Borrowed("\\pile")))
        }
    }

    fn convert_matrix_to_node(
        &self,
        matrix_obj: &MtefMatrix,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert matrix to proper matrix AST node
        // MTEF matrices store elements in row-major order
        if let Some(element_list) = &matrix_obj.element_list {
            let mut rows = Vec::new();
            let mut current: Option<&MtefObjectList> = Some(element_list);
            let mut cell_index = 0;
            let total_cells = (matrix_obj.rows as usize) * (matrix_obj.cols as usize);

            // Initialize rows
            for _ in 0..(matrix_obj.rows as usize) {
                let mut row = Vec::new();
                for _ in 0..(matrix_obj.cols as usize) {
                    row.push(Vec::new()); // Initialize empty cells
                }
                rows.push(row);
            }

            // Fill matrix cells
            while let Some(obj) = current {
                if obj.tag == MtefRecordType::Line
                    && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                    && let Some(line_nodes) = self.convert_line_to_nodes(line_obj)?
                {
                    // Calculate row and column from cell index
                    let row_idx = cell_index / (matrix_obj.cols as usize);
                    let col_idx = cell_index % (matrix_obj.cols as usize);

                    if row_idx < rows.len() && col_idx < rows[row_idx].len() {
                        rows[row_idx][col_idx] = line_nodes;
                    }
                    cell_index += 1;
                }
                current = obj.next.as_deref();

                // Safety check to prevent infinite loops
                if cell_index >= total_cells {
                    break;
                }
            }

            // Determine fence type based on matrix properties
            // This is a simplified approach - in a full implementation,
            // this might be determined by context or additional MTEF data
            let fence_type = match (matrix_obj.rows, matrix_obj.cols) {
                (1, _) => MatrixFence::None, // Row vector
                (_, 1) => MatrixFence::None, // Column vector
                _ => MatrixFence::Paren,     // General matrix with parentheses
            };

            Ok(MathNode::Matrix {
                rows,
                fence_type,
                properties: None,
            })
        } else {
            // Empty matrix
            Ok(MathNode::Matrix {
                rows: Vec::new(),
                fence_type: MatrixFence::None,
                properties: None,
            })
        }
    }

    fn convert_fence_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert fence templates (parentheses, brackets, braces, etc.) to Fence AST nodes
        let fence_type = match tmpl_obj.selector {
            0 => match tmpl_obj.variation {
                1 | 2 => Fence::Angle, // left/right only or both
                3 => Fence::Angle,
                _ => Fence::Angle,
            },
            1 => Fence::Paren,
            2 => Fence::Brace,
            3 => Fence::Bracket,
            4 => Fence::Pipe,
            5 => Fence::DoublePipe,
            6 => Fence::Floor,
            7 => Fence::Ceiling,
            8 => Fence::SquareBracket,
            9 => match tmpl_obj.variation {
                0 => Fence::SquareBracket,
                16 => Fence::Paren,
                17 => Fence::Paren,
                18 => Fence::Bracket,
                19 => Fence::Bracket,
                32 => Fence::Paren,
                33 => Fence::Paren,
                34 => Fence::SquareBracket,
                35 => Fence::SquareBracket,
                48 => Fence::Paren,
                49 => Fence::Paren,
                50 => Fence::Bracket,
                51 => Fence::Bracket,
                _ => Fence::Paren,
            },
            _ => Fence::Paren,
        };

        // Parse the content inside the fence
        let content = if let Some(obj_list) = &tmpl_obj.subobject_list {
            self.parse_single_subobject(obj_list)?
        } else {
            Vec::new()
        };

        Ok(TemplateParser::parse_fence(fence_type, content))
    }

    fn convert_decoration_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert underline/overline templates
        let content = if let Some(obj_list) = &tmpl_obj.subobject_list {
            self.parse_single_subobject(obj_list)?
        } else {
            Vec::new()
        };

        match tmpl_obj.selector {
            12 => {
                // Underline
                let underline_style = if tmpl_obj.variation == 1 {
                    LineStyle::Double
                } else {
                    LineStyle::Single
                };
                Ok(MathNode::Run {
                    content,
                    literal: None,
                    style: None,
                    font: None,
                    color: None,
                    underline: Some(underline_style),
                    overline: None,
                    strike_through: None,
                    double_strike_through: None,
                })
            },
            13 => {
                // Overline
                let overline_style = if tmpl_obj.variation == 1 {
                    LineStyle::Double
                } else {
                    LineStyle::Single
                };
                Ok(MathNode::Run {
                    content,
                    literal: None,
                    style: None,
                    font: None,
                    color: None,
                    underline: None,
                    overline: Some(overline_style),
                    strike_through: None,
                    double_strike_through: None,
                })
            },
            _ => Ok(MathNode::Text(Cow::Borrowed("\\decoration"))),
        }
    }

    fn convert_arrow_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert arrow templates to appropriate AST nodes
        // For now, fall back to template parsing
        let variation = tmpl_obj.variation;
        if let Some(template_def) = TemplateParser::find_template(tmpl_obj.selector, variation) {
            let args = if let Some(obj_list) = &tmpl_obj.subobject_list {
                self.parse_template_arguments(obj_list)?
            } else {
                smallvec::SmallVec::new()
            };
            Ok(TemplateParser::parse_template_arguments(
                template_def.template,
                &args,
            ))
        } else {
            Ok(MathNode::Text(Cow::Borrowed("\\arrow")))
        }
    }

    fn convert_large_op_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert large operator templates (sum, product, etc.)
        let operator = match tmpl_obj.selector {
            16 | 22 => LargeOperator::Sum,
            17 => LargeOperator::Product,
            18 => LargeOperator::Coproduct,
            19 => LargeOperator::Union,
            20 => LargeOperator::Intersection,
            _ => LargeOperator::Sum,
        };

        // Parse limits from subobjects
        let (lower_limit, upper_limit, integrand) = if let Some(obj_list) = &tmpl_obj.subobject_list
        {
            self.parse_large_op_subobjects(obj_list)?
        } else {
            (None, None, Vec::new())
        };

        Ok(TemplateParser::parse_large_op(
            operator,
            lower_limit.unwrap_or_default(),
            upper_limit.unwrap_or_default(),
            integrand,
        ))
    }

    fn convert_limit_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert limit templates
        // Parse the limit expression and the approaching value
        let (function, approaching) = if let Some(obj_list) = &tmpl_obj.subobject_list {
            self.parse_limit_subobjects(obj_list)?
        } else {
            (Vec::new(), Vec::new())
        };

        // For now, create a simple limit node - combine function and approaching value
        let mut content = function;
        if !approaching.is_empty() {
            content.push(MathNode::Text(Cow::Borrowed(" \\to ")));
            content.extend(approaching);
        }

        Ok(MathNode::Limit {
            content: Box::new(content),
            limit_type: crate::formula::ast::LimitType::Upper, // Default to upper for general limits
        })
    }

    fn convert_brace_template(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<MathNode<'arena>, MtefError> {
        // Convert horizontal brace templates
        let _is_upper = tmpl_obj.variation == 1;

        let (_content, _brace_text) = if let Some(obj_list) = &tmpl_obj.subobject_list {
            self.parse_brace_subobjects(obj_list)?
        } else {
            (Vec::new(), Vec::new())
        };

        // For now, fall back to template parsing
        let variation = tmpl_obj.variation;
        if let Some(template_def) = TemplateParser::find_template(tmpl_obj.selector, variation) {
            let args = if let Some(obj_list) = &tmpl_obj.subobject_list {
                self.parse_template_arguments(obj_list)?
            } else {
                smallvec::SmallVec::new()
            };
            Ok(TemplateParser::parse_template_arguments(
                template_def.template,
                &args,
            ))
        } else {
            Ok(MathNode::Text(Cow::Borrowed("\\brace")))
        }
    }

    fn parse_large_op_subobjects(&self, obj_list: &MtefObjectList) -> LargeOpResult<'arena> {
        // Parse subobjects for large operators: lower_limit, upper_limit, integrand
        let mut lower_limit = None;
        let mut upper_limit = None;
        let mut integrand = Vec::new();

        // Large operators typically have integrand first, then limits
        // This is a simplified parsing - real implementation would be more complex
        let mut current = Some(obj_list);
        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if integrand.is_empty() {
                    integrand = nodes;
                } else if lower_limit.is_none() {
                    lower_limit = Some(nodes);
                } else if upper_limit.is_none() {
                    upper_limit = Some(nodes);
                }
            }
            current = obj.next.as_deref();
        }

        Ok((lower_limit, upper_limit, integrand))
    }

    fn parse_limit_subobjects(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<(Vec<MathNode<'arena>>, Vec<MathNode<'arena>>), MtefError> {
        // Parse subobjects for limits: function and approaching value
        let mut function = Vec::new();
        let mut approaching = Vec::new();

        let mut current = Some(obj_list);
        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if function.is_empty() {
                    function = nodes;
                } else {
                    approaching = nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((function, approaching))
    }

    fn parse_brace_subobjects(
        &self,
        obj_list: &MtefObjectList,
    ) -> Result<(Vec<MathNode<'arena>>, Vec<MathNode<'arena>>), MtefError> {
        // Parse subobjects for braces: content and brace symbol
        let mut content = Vec::new();
        let mut brace_text = Vec::new();

        let mut current = Some(obj_list);
        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
                && let Some(nodes) = self.convert_line_to_nodes(line_obj)?
            {
                if content.is_empty() {
                    content = nodes;
                } else {
                    brace_text = nodes;
                }
            }
            current = obj.next.as_deref();
        }

        Ok((content, brace_text))
    }
}
