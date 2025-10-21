use crate::formula::ast::{MathNode, StrikeStyle};
use crate::formula::omml::attributes::*;
use crate::formula::omml::elements::*;
use crate::formula::omml::error::OmmlError;
use crate::formula::omml::handlers::*;
use crate::formula::omml::lookup::*;
use crate::formula::omml::properties::*;
use crate::formula::omml::utils::{validate_element_nesting, validate_omml_structure, *};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::borrow::Cow;

/// OMML parser that converts OMML XML to our formula AST
pub struct OmmlParser<'arena> {
    arena: &'arena bumpalo::Bump,
}

impl<'arena> OmmlParser<'arena> {
    /// Create a new OMML parser with the given arena
    pub fn new(arena: &'arena bumpalo::Bump) -> Self {
        Self { arena }
    }

    /// Parse OMML from a string
    ///
    /// # Example
    /// ```ignore
    /// let formula = Formula::new();
    /// let parser = OmmlParser::new(formula.arena());
    /// let nodes = parser.parse("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
    /// ```
    pub fn parse(&self, xml: &str) -> Result<Vec<MathNode<'arena>>, OmmlError> {
        // Validate input
        if xml.trim().is_empty() {
            return Err(OmmlError::InvalidStructure("Empty XML input".to_string()));
        }

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::with_capacity(1024); // Pre-allocate buffer for performance

        // Use high-performance element stack with capacity hint and context pooling
        let mut stack = ElementStack::with_capacity(64);
        let mut context_pool = ContextPool::new(32);
        let mut result = Vec::new();
        let mut depth = 0;
        const MAX_DEPTH: usize = 1000; // Prevent stack overflow attacks

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(OmmlError::InvalidStructure(format!(
                            "Maximum XML depth {} exceeded",
                            MAX_DEPTH
                        )));
                    }
                    self.handle_start_element(e, &mut stack, &mut context_pool)?;
                },
                Ok(Event::End(ref e)) => {
                    let name = e.local_name();
                    self.handle_end_element(
                        name.as_ref(),
                        &mut stack,
                        &mut result,
                        &mut context_pool,
                    )?;
                    depth = depth.saturating_sub(1);
                },
                Ok(Event::Text(ref e)) => {
                    self.handle_text_element(e, &mut stack)?;
                },
                Ok(Event::CData(ref e)) => {
                    self.handle_cdata_element(e, &mut stack)?;
                },
                Ok(Event::Empty(ref e)) => {
                    // Handle self-closing tags
                    self.handle_empty_element(e, &mut stack, &mut result, &mut context_pool)?;
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    let position = reader.buffer_position();
                    return Err(OmmlError::XmlError(format!(
                        "XML parsing error at position {}: {}",
                        position, e
                    )));
                },
                _ => {}, // Skip other events (comments, processing instructions, etc.)
            }
            buf.clear();
        }

        // Validate that we have a properly closed document
        if depth != 0 {
            return Err(OmmlError::InvalidStructure(format!(
                "Unclosed elements detected, final depth: {}",
                depth
            )));
        }

        // Validate result structure
        if result.is_empty() && !xml.contains("<m:oMath") {
            return Err(OmmlError::InvalidStructure(
                "No mathematical content found in OMML".to_string(),
            ));
        }

        // Validate the parsed structure
        validate_omml_structure(&result)?;

        // Return any remaining contexts in the stack to the pool
        while let Some(context) = stack.pop() {
            context_pool.put(context);
        }

        Ok(result)
    }

    fn handle_start_element(
        &self,
        elem: &BytesStart,
        stack: &mut ElementStack<'arena>,
        context_pool: &mut ContextPool<'arena>,
    ) -> Result<(), OmmlError> {
        let name = elem.local_name();
        let name_str =
            std::str::from_utf8(name.as_ref()).map_err(|e| OmmlError::ParseError(e.to_string()))?;
        let element_type = get_element_type(name_str);

        // Handle context-dependent element types
        let element_type = match (name_str, stack.last().map(|ctx| ctx.element_type)) {
            ("e" | "m:e", Some(ElementType::Nary)) => ElementType::Integrand,
            ("e" | "m:e", Some(ElementType::Radical)) => ElementType::Base,
            (
                "e" | "m:e",
                Some(ElementType::Superscript)
                | Some(ElementType::Subscript)
                | Some(ElementType::SubSup),
            ) => ElementType::Base,
            ("e" | "m:e", Some(ElementType::Fraction)) => ElementType::Denominator, // Actually, fraction has num/den, but e might be used differently
            ("e" | "m:e", Some(ElementType::MatrixRow)) => ElementType::MatrixCell, // Matrix cells within rows
            _ => element_type,
        };

        // Validate element nesting
        let parent_type = stack.last().map(|ctx| ctx.element_type);
        validate_element_nesting(&element_type, parent_type.as_ref())?;

        // Create new context for this element using the context pool
        let mut context = context_pool.get(element_type);

        // Parse attributes using SIMD-accelerated parsing with caching
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Store raw attributes for property element handlers
        // SAFETY: The attributes are valid for the duration of the XML parsing
        context.attributes = unsafe {
            std::mem::transmute::<
                Vec<quick_xml::events::attributes::Attribute<'_>>,
                Vec<quick_xml::events::attributes::Attribute<'_>>,
            >(attrs.clone())
        };

        // Use batch attribute parsing with PHF lookups and caching for better performance
        let mut cache = AttributeCache::new(&attrs);

        // Element-specific attribute parsing using handlers
        match element_type {
            ElementType::Delimiter => {
                DelimiterHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Nary => {
                NaryHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Accent => {
                AccentHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Matrix => {
                MatrixHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Fraction => {
                FractionHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::GroupChar => {
                GroupCharHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::EqArr => {
                EqArrHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Spacing => {
                SpacingHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::MatrixCell => {
                // Matrix cells don't need special start handling
            },
            ElementType::EqArrPr => {
                // Equation array properties - no special start handling needed
            },
            ElementType::Limit => {
                // Limit elements - no special start handling needed
            },
            ElementType::PreScript => {
                // Pre-script elements - no special start handling needed
            },
            ElementType::PostScript => {
                // Post-script elements - no special start handling needed
            },
            ElementType::Properties => {
                // Parse properties based on the element name
                context.properties = match name_str {
                    "dPr" | "m:dPr" => parse_delimiter_properties(&attrs),
                    "fPr" | "m:fPr" => parse_fraction_properties(&attrs),
                    "naryPr" | "m:naryPr" => parse_nary_properties(&attrs),
                    "accPr" | "m:accPr" => parse_accent_properties(&attrs),
                    "radPr" | "m:radPr" => parse_radical_properties(&attrs),
                    "sSupPr" | "m:sSupPr" => parse_general_properties(&attrs),
                    "sSubPr" | "m:sSubPr" => parse_general_properties(&attrs),
                    "funcPr" | "m:funcPr" => parse_general_properties(&attrs),
                    "limPr" | "m:limPr" => parse_limit_properties(&attrs),
                    "barPr" | "m:barPr" => parse_bar_properties(&attrs),
                    "boxPr" | "m:boxPr" => parse_box_properties(&attrs),
                    "borderBoxPr" | "m:borderBoxPr" => parse_border_box_properties(&attrs),
                    "phantomPr" | "m:phantomPr" => parse_phantom_properties(&attrs),
                    "spacingPr" | "m:spacingPr" => parse_spacing_properties(&attrs),
                    _ => parse_general_properties(&attrs),
                };
            },
            ElementType::AccentProperties => {
                context.properties = parse_accent_properties(&attrs);
            },
            _ => {
                // For elements that don't need special handling, properties are already parsed
                context.properties = parse_attributes_batch_with_cache(&mut cache);
            },
        }

        stack.push(context);
        Ok(())
    }

    fn handle_end_element(
        &self,
        name: &[u8],
        stack: &mut ElementStack<'arena>,
        result: &mut Vec<MathNode<'arena>>,
        _context_pool: &mut ContextPool<'arena>,
    ) -> Result<(), OmmlError> {
        if stack.is_empty() {
            return Ok(());
        }

        let name_str =
            std::str::from_utf8(name).map_err(|e| OmmlError::ParseError(e.to_string()))?;
        let element_type = get_element_type(name_str);
        let mut context = stack.pop().unwrap();

        // Get parent context for passing results up
        let parent_context = stack.last_mut();

        // Use element-specific handlers
        match element_type {
            ElementType::Math => {
                // Root element - add all children to result
                result.extend(context.children);
            },
            ElementType::Run => {
                // Check if run has any properties - if so, create a Run node
                let has_properties = context.properties.run_literal.is_some()
                    || context.properties.math_variant.is_some()
                    || context.properties.run_normal_text.is_some()
                    || context.properties.color.is_some()
                    || context.properties.underline.is_some()
                    || context.properties.overline.is_some()
                    || context.properties.strike_through.is_some()
                    || context.properties.double_strike_through.is_some();

                if has_properties {
                    // Create a Run node with properties
                    let run_node = MathNode::Run {
                        content: context.children.clone(),
                        literal: context.properties.run_literal,
                        style: context
                            .properties
                            .math_variant
                            .as_ref()
                            .and_then(|s| parse_style_value(s)),
                        font: context
                            .properties
                            .run_normal_text
                            .as_ref()
                            .map(|s| std::borrow::Cow::Borrowed(self.arena.alloc_str(s))),
                        color: context
                            .properties
                            .color
                            .as_ref()
                            .map(|s| std::borrow::Cow::Borrowed(self.arena.alloc_str(s))),
                        underline: context
                            .properties
                            .underline
                            .as_ref()
                            .and_then(|s| parse_line_style(Some(s))),
                        overline: context
                            .properties
                            .overline
                            .as_ref()
                            .and_then(|s| parse_line_style(Some(s))),
                        strike_through: context
                            .properties
                            .strike_through
                            .and_then(|b| if b { Some(StrikeStyle::Single) } else { None }),
                        double_strike_through: context.properties.double_strike_through,
                    };

                    if let Some(parent) = parent_context {
                        parent.children.push(run_node);
                    }
                } else {
                    // No properties - pass children up directly
                    if let Some(parent) = parent_context {
                        extend_vec_efficient(&mut parent.children, context.children);
                    }
                }
            },
            ElementType::Text => {
                // Create text node and pass up
                if !context.text.is_empty() {
                    let text = intern_string(self.arena, context.text.as_str());
                    let node = MathNode::Text(Cow::Borrowed(text));
                    if let Some(parent) = parent_context {
                        parent.children.push(node);
                    }
                    // Text node recorded
                }
            },
            ElementType::Delimiter => {
                DelimiterHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Nary => {
                NaryHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Function => {
                FunctionHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::FunctionName => {
                FunctionNameHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Accent => {
                AccentHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Bar => {
                BarHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Box => {
                BoxHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Phantom => {
                PhantomHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Matrix => {
                MatrixHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::MatrixRow => {
                MatrixRowHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Fraction => {
                FractionHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Radical => {
                RadicalHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Superscript => {
                SuperscriptHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Subscript => {
                SubscriptHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::SubSup => {
                SubSupHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::GroupChar => {
                GroupCharHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::BorderBox => {
                BorderBoxHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::EqArr => {
                EqArrHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Character => {
                CharHandler::handle_end(name, &mut context, parent_context, self.arena);
            },
            ElementType::Spacing => {
                SpacingHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Numerator => {
                NumeratorHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Denominator => {
                DenominatorHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Degree => {
                DegreeHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Base => {
                BaseHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::SuperscriptElement => {
                SuperscriptElementHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::SubscriptElement => {
                SubscriptElementHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::LowerLimit => {
                LowerLimitHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::UpperLimit => {
                UpperLimitHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::LimLow => {
                LimLowHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::LimUpp => {
                LimUppHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Integrand => {
                IntegrandHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Limit => {
                LimitHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::PreScript => {
                PreScriptHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::PostScript => {
                PostScriptHandler::handle_end(&mut context, parent_context, self.arena);
            },
            // Handle run properties (rPr) and control properties (ctrlPr) specifically
            ElementType::Properties if name_str == "rPr" || name_str == "m:rPr" => {
                RunPropsHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Properties if name_str == "ctrlPr" || name_str == "m:ctrlPr" => {
                CtrlPropsHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Properties if name_str == "groupChrPr" || name_str == "m:groupChrPr" => {
                GroupChrPrHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Position => {
                PosHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::VerticalAlignment => {
                VertJcHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Lit => {
                LitHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Scr => {
                ScrHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Sty => {
                StyHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Nor => {
                NorHandler::handle_end(&mut context, parent_context, self.arena);
            },
            // Handle property elements - store properties in parent and pass children up
            ElementType::Properties => {
                if let Some(parent) = parent_context {
                    // Store the parsed properties in the parent context
                    parent.properties = context.properties.clone();
                    extend_vec_efficient(&mut parent.children, context.children);
                }
            },
            ElementType::AccentProperties => {
                if let Some(parent) = parent_context {
                    // Store the parsed properties in the parent context
                    parent.properties = context.properties.clone();
                    extend_vec_efficient(&mut parent.children, context.children);
                }
            },
            // Handle structural elements that just pass children up
            ElementType::MatrixCell => {
                if let Some(parent) = parent_context {
                    extend_vec_efficient(&mut parent.children, context.children);
                }
            },
            _ => {
                // For unknown or unhandled elements, pass children up
                if let Some(parent) = parent_context {
                    extend_vec_efficient(&mut parent.children, context.children);
                }
            },
        }

        Ok(())
    }

    fn handle_text_element(
        &self,
        event: &[u8],
        stack: &mut ElementStack<'arena>,
    ) -> Result<(), OmmlError> {
        if let Some(context) = stack.last_mut() {
            // For OMML, text content is typically plain and doesn't need unescaping
            let text_str =
                std::str::from_utf8(event).map_err(|e| OmmlError::ParseError(e.to_string()))?;

            // Process text efficiently
            let processed_text = process_text_zero_copy(text_str);
            context.text.push_str(processed_text.as_ref());
        }

        Ok(())
    }

    fn handle_cdata_element(
        &self,
        event: &[u8],
        stack: &mut ElementStack<'arena>,
    ) -> Result<(), OmmlError> {
        if let Some(context) = stack.last_mut() {
            // CDATA content is already unescaped by quick-xml
            let text_str =
                std::str::from_utf8(event).map_err(|e| OmmlError::ParseError(e.to_string()))?;

            // Process text efficiently
            let processed_text = process_text_zero_copy(text_str);
            context.text.push_str(processed_text.as_ref());
        }

        Ok(())
    }

    fn handle_empty_element(
        &self,
        elem: &BytesStart,
        stack: &mut ElementStack<'arena>,
        _result: &mut Vec<MathNode<'arena>>,
        context_pool: &mut ContextPool<'arena>,
    ) -> Result<(), OmmlError> {
        let name = elem.local_name();
        let name_str =
            std::str::from_utf8(name.as_ref()).map_err(|e| OmmlError::ParseError(e.to_string()))?;
        let element_type = get_element_type(name_str);

        // For self-closing elements, we need to handle both start and end logic
        let mut context = context_pool.get(element_type);

        // Parse attributes
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Store raw attributes for property element handlers
        // SAFETY: The attributes are valid for the duration of the XML parsing
        context.attributes = unsafe {
            std::mem::transmute::<
                Vec<quick_xml::events::attributes::Attribute<'_>>,
                Vec<quick_xml::events::attributes::Attribute<'_>>,
            >(attrs.clone())
        };

        context.properties = parse_attributes_batch(&attrs);

        // Handle element-specific start logic
        match element_type {
            ElementType::Delimiter => {
                DelimiterHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Nary => {
                NaryHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Accent => {
                AccentHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Matrix => {
                MatrixHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Fraction => {
                FractionHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::GroupChar => {
                GroupCharHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::EqArr => {
                EqArrHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::Spacing => {
                SpacingHandler::handle_start(elem, &mut context, self.arena);
            },
            ElementType::MatrixCell => {
                // Matrix cells don't need special start handling
            },
            ElementType::EqArrPr => {
                // Equation array properties - no special start handling needed
            },
            ElementType::Limit => {
                // Limit elements - no special start handling needed
            },
            ElementType::PreScript => {
                // Pre-script elements - no special start handling needed
            },
            ElementType::PostScript => {
                // Post-script elements - no special start handling needed
            },
            _ => {
                // For other elements, properties are already parsed
            },
        }

        // Handle element-specific end logic (since it's self-closing)
        let parent_context = stack.last_mut();
        match element_type {
            ElementType::Math => {
                // Root element - should not be self-closing in valid OMML
                return Err(OmmlError::InvalidStructure(
                    "Math element cannot be self-closing".to_string(),
                ));
            },
            ElementType::Run => {
                // Pass empty content up to parent
                if let Some(_parent) = parent_context {
                    // Empty run contributes nothing
                }
            },
            ElementType::Text => {
                // Empty text node
                if let Some(parent) = parent_context {
                    let text = intern_string(self.arena, "");
                    let node = MathNode::Text(std::borrow::Cow::Borrowed(text));
                    parent.children.push(node);
                }
            },
            ElementType::Delimiter => {
                DelimiterHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Nary => {
                NaryHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Function => {
                FunctionHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Accent => {
                AccentHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Bar => {
                BarHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Box => {
                BoxHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Phantom => {
                PhantomHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Matrix => {
                MatrixHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::MatrixRow => {
                MatrixRowHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Fraction => {
                FractionHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Radical => {
                RadicalHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Superscript => {
                SuperscriptHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Subscript => {
                SubscriptHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::SubSup => {
                SubSupHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::GroupChar => {
                GroupCharHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::BorderBox => {
                BorderBoxHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::EqArr => {
                EqArrHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Spacing => {
                SpacingHandler::handle_end(&mut context, parent_context, self.arena);
            },
            ElementType::Character => {
                CharHandler::handle_end(name.as_ref(), &mut context, parent_context, self.arena);
            },
            _ => {
                // For unknown or unhandled self-closing elements, do nothing
            },
        }

        // Return context to pool for reuse
        context_pool.put(context);

        Ok(())
    }
}
