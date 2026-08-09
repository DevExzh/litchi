//! Structural and inert-authoring validation for BIFF8 array formulas.

use crate::{Error, Result};

use super::{Limits, Owner};

pub(crate) const RECORD_TYPE: u16 = 0x0221;
pub(crate) const FIXED_BYTES: usize = 14;
pub(crate) const MAX_RECORD_BYTES: usize = 8_224;
pub(crate) const MAX_RGCE_BYTES: usize = 1_800;
const MAX_SEMANTIC_STACK: usize = 40;

pub(crate) fn validate(owner: &Owner, limits: Limits, authored: bool) -> Result<()> {
    let count = checked_cell_count(owner)?;
    if count > limits.max_cells() {
        return Err(invalid(
            "array-formula rectangle exceeds the configured cell limit",
        ));
    }
    if owner.tokens().is_empty() {
        return Err(invalid("ArrayParsedFormula cce must be greater than zero"));
    }
    if owner.tokens().len() > limits.max_token_bytes() || owner.tokens().len() > MAX_RGCE_BYTES {
        return Err(invalid("ArrayParsedFormula rgce exceeds 1,800 bytes"));
    }
    if owner.extra().len() > limits.max_extra_bytes() {
        return Err(invalid(
            "ArrayParsedFormula rgcb exceeds the configured limit",
        ));
    }
    let total = FIXED_BYTES
        .checked_add(owner.tokens().len())
        .and_then(|value| value.checked_add(owner.extra().len()))
        .ok_or_else(|| invalid("Array payload length overflows"))?;
    if total > limits.max_record_bytes() || total > MAX_RECORD_BYTES {
        return Err(invalid("Array payload exceeds the BIFF8 record limit"));
    }

    let arrays = validate_tokens(owner.tokens(), authored, limits)?;
    if arrays == 0 {
        if !owner.extra().is_empty() {
            return Err(invalid("ArrayParsedFormula has unowned RgbExtra bytes"));
        }
    } else {
        validate_array_extra(owner.extra(), arrays)?;
    }
    Ok(())
}

fn checked_cell_count(owner: &Owner) -> Result<usize> {
    let first = owner.range().first();
    let last = owner.range().last();
    let rows = usize::from(last.row() - first.row()) + 1;
    let cols = usize::from(last.col() - first.col()) + 1;
    rows.checked_mul(cols)
        .ok_or_else(|| invalid("array-formula rectangle cardinality overflows"))
}

fn validate_tokens(tokens: &[u8], authored: bool, limits: Limits) -> Result<usize> {
    let mut position = 0usize;
    let mut count = 0usize;
    let mut arrays = 0usize;
    let mut actual_bytes = 0usize;
    let mut stack = SemanticStack::new();
    let mut controls = ControlStack::new();
    let mut memories = MemoryStack::new();

    while position < tokens.len() {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("array-formula token count overflows"))?;
        if count > limits.max_tokens() {
            return Err(invalid(
                "array-formula token count exceeds the configured limit",
            ));
        }
        let opcode = tokens[position];
        if opcode & 0x80 != 0 {
            return Err(invalid("array formula contains an invalid Ptg opcode"));
        }
        if opcode == 0x01 || opcode == 0x02 {
            return Err(invalid(
                "ArrayParsedFormula contains a forbidden PtgExp or PtgTbl",
            ));
        }
        if opcode == 0x18 {
            return Err(Error::UnsupportedFeature(
                "extended natural-language formula tokens cannot be validated losslessly"
                    .to_string(),
            ));
        }

        let shape = if opcode < 0x20 {
            base_token(
                tokens,
                position,
                opcode,
                authored,
                limits,
                &mut stack,
                &mut controls,
            )?
        } else {
            classified_token(
                tokens,
                position,
                opcode,
                authored,
                &mut arrays,
                &mut stack,
                &mut controls,
                &mut memories,
            )?
        };
        position = position
            .checked_add(shape.wire_bytes)
            .ok_or_else(|| invalid("array-formula token length overflows"))?;
        if position > tokens.len() {
            return Err(invalid("array formula contains a truncated Ptg"));
        }
        actual_bytes = actual_bytes
            .checked_add(shape.actual_bytes)
            .ok_or_else(|| invalid("actual array-formula size overflows"))?;
        if actual_bytes > MAX_RGCE_BYTES {
            return Err(invalid(
                "ArrayParsedFormula actual expression size exceeds 1,800 bytes",
            ));
        }
        memories.advance(position, matches!(opcode, 0x0f..=0x11), &mut stack)?;
        stack.check_limits(limits)?;
    }

    controls.finish()?;
    memories.finish()?;
    stack.finish(limits)?;
    Ok(arrays)
}

fn base_token(
    tokens: &[u8],
    position: usize,
    opcode: u8,
    authored: bool,
    limits: Limits,
    stack: &mut SemanticStack,
    controls: &mut ControlStack,
) -> Result<TokenShape> {
    match opcode {
        0x03..=0x0e => {
            stack.binary(true, true)?;
            Ok(TokenShape::fixed(1))
        },
        0x0f..=0x11 => {
            stack.binary(false, false)?;
            Ok(TokenShape::fixed(1))
        },
        0x12..=0x14 => {
            stack.unary(true, true)?;
            Ok(TokenShape::fixed(1))
        },
        0x15 => {
            stack.parenthesize()?;
            Ok(TokenShape::fixed(1))
        },
        0x16 => {
            stack.operand(true)?;
            Ok(TokenShape::fixed(1))
        },
        0x17 => {
            let count = usize::from(*tokens.get(position + 1).ok_or_else(truncated)?);
            let flags = *tokens.get(position + 2).ok_or_else(truncated)?;
            if flags & !0x01 != 0 {
                return Err(invalid("PtgStr has reserved option bits"));
            }
            if count > limits.max_string_utf16_units() {
                return Err(invalid("PtgStr exceeds the configured UTF-16 unit limit"));
            }
            let width = if flags & 1 == 0 { 1 } else { 2 };
            let wire_bytes = 3usize
                .checked_add(count.checked_mul(width).ok_or_else(truncated)?)
                .ok_or_else(truncated)?;
            let actual_bytes = 3usize
                .checked_add(count.checked_mul(2).ok_or_else(truncated)?)
                .ok_or_else(truncated)?;
            stack.operand(true)?;
            Ok(TokenShape {
                wire_bytes,
                actual_bytes,
            })
        },
        0x19 => {
            let options = *tokens.get(position + 1).ok_or_else(truncated)?;
            let wire_bytes =
                validate_attribute(tokens, position, options, authored, stack, controls)?;
            Ok(TokenShape::fixed(wire_bytes))
        },
        0x1c => {
            let error = *tokens.get(position + 1).ok_or_else(truncated)?;
            if !matches!(error, 0x00 | 0x07 | 0x0f | 0x17 | 0x1d | 0x24 | 0x2a) {
                return Err(invalid("PtgErr contains an invalid BIFF error code"));
            }
            stack.operand(true)?;
            Ok(TokenShape::fixed(2))
        },
        0x1d => {
            if !matches!(tokens.get(position + 1), Some(0 | 1)) {
                return Err(invalid("PtgBool contains an invalid Boolean"));
            }
            stack.operand(true)?;
            Ok(TokenShape::fixed(2))
        },
        0x1e => {
            stack.operand(true)?;
            Ok(TokenShape::fixed(3))
        },
        0x1f => {
            validate_xnum(
                tokens
                    .get(position + 1..position + 9)
                    .ok_or_else(truncated)?,
                "PtgNum",
            )?;
            stack.operand(true)?;
            Ok(TokenShape::fixed(9))
        },
        _ => Err(invalid(
            "array formula contains an unsupported base Ptg opcode",
        )),
    }
}

fn classified_token(
    tokens: &[u8],
    position: usize,
    opcode: u8,
    authored: bool,
    arrays: &mut usize,
    stack: &mut SemanticStack,
    controls: &mut ControlStack,
    memories: &mut MemoryStack,
) -> Result<TokenShape> {
    let base = (opcode & 0x1f) | 0x20;
    let value_type = classified_value_type(opcode)?;
    match base {
        0x20 => {
            if opcode != 0x40 && opcode != 0x60 {
                return Err(invalid("PtgArray operand type must be value or array"));
            }
            *arrays = arrays
                .checked_add(1)
                .ok_or_else(|| invalid("PtgArray count overflows"))?;
            stack.operand(true)?;
            Ok(TokenShape {
                wire_bytes: 8,
                actual_bytes: 15,
            })
        },
        0x21 => {
            let index = read_u16(tokens, position + 1)?;
            let arity = crate::formula::fixed_function_arity(index).ok_or_else(|| {
                invalid("PtgFunc has an unknown function index or variable arity")
            })?;
            if authored {
                validate_safe_function(index)?;
            }
            stack.function(arity, value_type)?;
            Ok(TokenShape::fixed(3))
        },
        0x22 => {
            let arity = usize::from(*tokens.get(position + 1).ok_or_else(truncated)?);
            let raw = read_u16(tokens, position + 2)?;
            if authored {
                if raw == 0x00ff || raw & 0x8000 != 0 {
                    return Err(unsafe_formula(
                        "external, UDF, and command-equivalent functions are forbidden",
                    ));
                }
                validate_safe_function(raw)?;
            }
            controls.maybe_close(raw, arity, position + 4)?;
            stack.function(arity, value_type)?;
            Ok(TokenShape::fixed(4))
        },
        0x23 | 0x39 if authored => Err(unsafe_formula(
            "defined-name and external-name operands are forbidden for authoring",
        )),
        0x3a..=0x3d if authored => Err(unsafe_formula(
            "unresolved 3-D or external references are forbidden for authoring",
        )),
        0x23 => {
            stack.operand(value_type)?;
            Ok(TokenShape::fixed(5))
        },
        0x24 => {
            validate_column(tokens, position + 3)?;
            stack.operand(value_type)?;
            Ok(TokenShape {
                wire_bytes: 5,
                actual_bytes: 7,
            })
        },
        0x25 => {
            validate_column(tokens, position + 5)?;
            validate_column(tokens, position + 7)?;
            stack.operand(value_type)?;
            Ok(TokenShape {
                wire_bytes: 9,
                actual_bytes: 13,
            })
        },
        0x26 => Err(Error::UnsupportedFeature(
            "PtgMemArea ancillary data cannot yet be proven lossless".to_string(),
        )),
        0x27 => {
            let error = *tokens.get(position + 1).ok_or_else(truncated)?;
            if !matches!(error, 0x00 | 0x07 | 0x0f | 0x17 | 0x1d | 0x24 | 0x2a) {
                return Err(invalid("PtgMemErr contains an invalid BIFF error code"));
            }
            memories.push(
                position + 7,
                read_u16(tokens, position + 5)?,
                stack.len(),
                false,
                value_type,
            )?;
            Ok(TokenShape::fixed(7))
        },
        0x29 => {
            memories.push(
                position + 3,
                read_u16(tokens, position + 1)?,
                stack.len(),
                true,
                value_type,
            )?;
            Ok(TokenShape::fixed(3))
        },
        0x28 => Err(Error::UnsupportedFeature(
            "PtgMemNoMem is outside the current formula codec boundary".to_string(),
        )),
        0x2a => {
            stack.operand(true)?;
            Ok(TokenShape {
                wire_bytes: 5,
                actual_bytes: 7,
            })
        },
        0x2b => {
            stack.operand(true)?;
            Ok(TokenShape {
                wire_bytes: 9,
                actual_bytes: 13,
            })
        },
        0x2c | 0x2d => Err(invalid("ArrayParsedFormula contains PtgRefN or PtgAreaN")),
        0x39 => {
            if read_u16(tokens, position + 5)? != 0 {
                return Err(invalid("PtgNameX reserved field must be zero"));
            }
            stack.operand(value_type)?;
            Ok(TokenShape::fixed(7))
        },
        0x3a => {
            validate_column(tokens, position + 5)?;
            stack.operand(value_type)?;
            Ok(TokenShape {
                wire_bytes: 7,
                actual_bytes: 9,
            })
        },
        0x3b => {
            validate_column(tokens, position + 7)?;
            validate_column(tokens, position + 9)?;
            stack.operand(value_type)?;
            Ok(TokenShape {
                wire_bytes: 11,
                actual_bytes: 15,
            })
        },
        0x3c => {
            stack.operand(true)?;
            Ok(TokenShape {
                wire_bytes: 7,
                actual_bytes: 9,
            })
        },
        0x3d => {
            stack.operand(true)?;
            Ok(TokenShape {
                wire_bytes: 11,
                actual_bytes: 16,
            })
        },
        _ => Err(invalid("array formula contains an unsupported Ptg opcode")),
    }
}

fn validate_attribute(
    tokens: &[u8],
    position: usize,
    options: u8,
    authored: bool,
    stack: &mut SemanticStack,
    controls: &mut ControlStack,
) -> Result<usize> {
    let data = read_u16(tokens, position + 2)?;
    match options {
        0x01 => {
            validate_prefix_attribute(position, stack)?;
            if authored {
                return Err(unsafe_formula(
                    "PtgAttrSemi is not available for inert authoring",
                ));
            }
            Ok(4)
        },
        0x02 => {
            if stack.len() == 0 {
                return Err(invalid("PtgAttrIf must follow its condition expression"));
            }
            controls.push_if(position + 4, data)?;
            Ok(4)
        },
        0x04 => {
            if stack.len() == 0 {
                return Err(invalid("PtgAttrChoose must follow its selector expression"));
            }
            let branch_count = usize::from(data);
            if !(1..=29).contains(&branch_count) {
                return Err(invalid("PtgAttrChoose cOffset must be between 1 and 29"));
            }
            let wire_bytes = 4usize
                .checked_add((branch_count + 1).checked_mul(2).ok_or_else(truncated)?)
                .ok_or_else(truncated)?;
            let end = position.checked_add(wire_bytes).ok_or_else(truncated)?;
            if end > tokens.len() {
                return Err(truncated());
            }
            let mut offsets = [0u16; 30];
            for (index, offset) in offsets[..=branch_count].iter_mut().enumerate() {
                *offset = read_u16(tokens, position + 4 + index * 2)?;
            }
            if usize::from(offsets[0]) != wire_bytes - 4 {
                return Err(invalid(
                    "PtgAttrChoose first offset does not match its token size",
                ));
            }
            controls.push_choose(end, branch_count, offsets)?;
            Ok(wire_bytes)
        },
        0x08 => {
            controls.push_goto(position + 4, data)?;
            Ok(4)
        },
        0x10 => {
            stack.function(1, true)?;
            Ok(4)
        },
        0x20 | 0x21 => {
            validate_prefix_attribute(position, stack)?;
            if authored {
                return Err(unsafe_formula(
                    "macro-sheet assignment tokens are forbidden",
                ));
            }
            Ok(4)
        },
        0x40 => {
            validate_space(tokens, position, data, stack, false)?;
            Ok(4)
        },
        0x41 => {
            validate_prefix_attribute(position, stack)?;
            validate_space(tokens, position, data, stack, true)?;
            if authored {
                return Err(unsafe_formula(
                    "PtgAttrSpaceSemi is not available for inert authoring",
                ));
            }
            Ok(4)
        },
        _ => Err(invalid("PtgAttr contains an invalid attribute form")),
    }
}

fn validate_prefix_attribute(position: usize, stack: &SemanticStack) -> Result<()> {
    if position != 0 || stack.len() != 0 {
        return Err(invalid(
            "PtgAttr prefix form must be the first token in the expression",
        ));
    }
    Ok(())
}

fn validate_space(
    tokens: &[u8],
    position: usize,
    data: u16,
    stack: &SemanticStack,
    semi: bool,
) -> Result<()> {
    let subtype = crate::utils::truncate_u16_to_u8(data);
    let next = tokens.get(position + 4).copied();
    match subtype {
        0 | 1 | 6 => {
            if !next.is_some_and(expression_start) {
                return Err(invalid(
                    "PtgAttrSpace must precede an expression or base expression",
                ));
            }
        },
        2..=5 if !semi => {
            if stack.len() == 0 || next != Some(0x15) {
                return Err(invalid(
                    "parenthesis PtgAttrSpace must follow an expression and precede PtgParen",
                ));
            }
        },
        _ => return Err(invalid("PtgAttrSpace contains an invalid space subtype")),
    }
    Ok(())
}

fn expression_start(opcode: u8) -> bool {
    if matches!(opcode, 0x16 | 0x17 | 0x1c..=0x1f) {
        return true;
    }
    if opcode == 0x19 {
        return true;
    }
    if opcode < 0x20 || opcode & 0x80 != 0 {
        return false;
    }
    matches!(
        (opcode & 0x1f) | 0x20,
        0x20 | 0x23..=0x2d | 0x39..=0x3d
    )
}

const MAX_CONTROL_DEPTH: usize = 8;
const MAX_BRANCHES: usize = 29;

#[derive(Clone, Copy)]
struct Goto {
    end: usize,
    offset: u16,
}

impl Goto {
    const EMPTY: Self = Self { end: 0, offset: 0 };
}

#[derive(Clone, Copy)]
struct ControlFrame {
    kind: u8,
    attr_end: usize,
    first_offset: u16,
    branch_count: usize,
    offsets: [u16; MAX_BRANCHES + 1],
    gotos: [Goto; MAX_BRANCHES],
    goto_count: usize,
}

impl ControlFrame {
    const EMPTY: Self = Self {
        kind: 0,
        attr_end: 0,
        first_offset: 0,
        branch_count: 0,
        offsets: [0; MAX_BRANCHES + 1],
        gotos: [Goto::EMPTY; MAX_BRANCHES],
        goto_count: 0,
    };
}

struct ControlStack {
    frames: [ControlFrame; MAX_CONTROL_DEPTH],
    len: usize,
}

impl ControlStack {
    const IF: u8 = 1;
    const CHOOSE: u8 = 2;

    const fn new() -> Self {
        Self {
            frames: [ControlFrame::EMPTY; MAX_CONTROL_DEPTH],
            len: 0,
        }
    }

    fn push_if(&mut self, attr_end: usize, first_offset: u16) -> Result<()> {
        self.push(ControlFrame {
            kind: Self::IF,
            attr_end,
            first_offset,
            ..ControlFrame::EMPTY
        })
    }

    fn push_choose(
        &mut self,
        attr_end: usize,
        branch_count: usize,
        offsets: [u16; MAX_BRANCHES + 1],
    ) -> Result<()> {
        self.push(ControlFrame {
            kind: Self::CHOOSE,
            attr_end,
            branch_count,
            offsets,
            ..ControlFrame::EMPTY
        })
    }

    fn push(&mut self, frame: ControlFrame) -> Result<()> {
        let slot = self
            .frames
            .get_mut(self.len)
            .ok_or_else(|| invalid("PtgAttr control nesting exceeds eight"))?;
        *slot = frame;
        self.len += 1;
        Ok(())
    }

    fn push_goto(&mut self, end: usize, offset: u16) -> Result<()> {
        let frame =
            self.frames
                .get_mut(self.len.checked_sub(1).ok_or_else(|| {
                    invalid("PtgAttrGoto appears without PtgAttrIf or PtgAttrChoose")
                })?)
                .ok_or_else(|| invalid("PtgAttr control stack is invalid"))?;
        let maximum = if frame.kind == Self::IF {
            2
        } else {
            frame.branch_count
        };
        if frame.goto_count >= maximum {
            return Err(invalid("PtgAttr control expression has too many branches"));
        }
        let span = end
            .checked_sub(frame.attr_end)
            .ok_or_else(|| invalid("PtgAttr branch offset underflows"))?;
        if frame.goto_count == 0
            && frame.kind == Self::IF
            && usize::from(frame.first_offset) != span
        {
            return Err(invalid(
                "PtgAttrIf branch offset does not match the token span",
            ));
        }
        if frame.kind == Self::CHOOSE && usize::from(frame.offsets[frame.goto_count + 1]) != span {
            return Err(invalid(
                "PtgAttrChoose branch offset does not match the token span",
            ));
        }
        frame.gotos[frame.goto_count] = Goto { end, offset };
        frame.goto_count += 1;
        Ok(())
    }

    fn maybe_close(&mut self, raw_function: u16, arity: usize, end: usize) -> Result<()> {
        if self.len == 0 {
            return Ok(());
        }
        let frame = self.frames[self.len - 1];
        let index = raw_function & 0x7fff;
        let matching = (frame.kind == Self::IF && index == 1 && frame.goto_count >= 1)
            || (frame.kind == Self::CHOOSE
                && index == 100
                && frame.goto_count == frame.branch_count);
        if !matching {
            return Ok(());
        }
        let expected_arity = frame.goto_count + 1;
        let offsets_match = frame.gotos[..frame.goto_count].iter().all(|branch| {
            end.checked_sub(branch.end)
                .and_then(|span| span.checked_sub(1))
                == Some(usize::from(branch.offset))
        });
        if frame.kind == Self::IF && frame.goto_count == 1 && !offsets_match {
            return Ok(());
        }
        if raw_function & 0x8000 != 0 || arity != expected_arity || !offsets_match {
            return Err(invalid(
                "optimized PtgAttr control expression has invalid function or offsets",
            ));
        }
        self.len -= 1;
        Ok(())
    }

    fn finish(&self) -> Result<()> {
        if self.len != 0 {
            return Err(invalid("PtgAttr control expression is not terminated"));
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct MemoryFrame {
    expression_start: usize,
    expression_bytes: usize,
    stack_len: usize,
    allows_nested: bool,
    value_type: bool,
}

impl MemoryFrame {
    const EMPTY: Self = Self {
        expression_start: 0,
        expression_bytes: 0,
        stack_len: 0,
        allows_nested: false,
        value_type: false,
    };
}

struct MemoryStack {
    frames: [MemoryFrame; MAX_CONTROL_DEPTH],
    len: usize,
}

impl MemoryStack {
    const fn new() -> Self {
        Self {
            frames: [MemoryFrame::EMPTY; MAX_CONTROL_DEPTH],
            len: 0,
        }
    }

    fn push(
        &mut self,
        expression_start: usize,
        expression_bytes: u16,
        stack_len: usize,
        allows_nested: bool,
        value_type: bool,
    ) -> Result<()> {
        if expression_bytes == 0 {
            return Err(invalid("memory Ptg cce must be greater than zero"));
        }
        if self.len != 0 && !self.frames[self.len - 1].allows_nested {
            return Err(invalid("nested memory Ptg is forbidden in this expression"));
        }
        let slot = self
            .frames
            .get_mut(self.len)
            .ok_or_else(|| invalid("memory Ptg nesting exceeds eight"))?;
        *slot = MemoryFrame {
            expression_start,
            expression_bytes: usize::from(expression_bytes),
            stack_len,
            allows_nested,
            value_type,
        };
        self.len += 1;
        Ok(())
    }

    fn advance(
        &mut self,
        end: usize,
        reference_operator: bool,
        stack: &mut SemanticStack,
    ) -> Result<()> {
        if self.len == 0 {
            return Ok(());
        }
        let frame = self.frames[self.len - 1];
        let span = end
            .checked_sub(frame.expression_start)
            .ok_or_else(|| invalid("memory Ptg expression span underflows"))?;
        if span > frame.expression_bytes {
            return Err(invalid("memory Ptg cce is smaller than its expression"));
        }
        if span == frame.expression_bytes {
            if !reference_operator || stack.len() != frame.stack_len + 1 {
                return Err(invalid(
                    "memory Ptg must own exactly one binary reference expression",
                ));
            }
            stack.set_top_value_type(frame.value_type)?;
            self.len -= 1;
        }
        Ok(())
    }

    fn finish(&self) -> Result<()> {
        if self.len != 0 {
            return Err(invalid("memory Ptg expression is incomplete"));
        }
        Ok(())
    }
}

fn validate_array_extra(extra: &[u8], array_count: usize) -> Result<()> {
    let mut position = 0usize;
    for _ in 0..array_count {
        let header = extra
            .get(position..position + 3)
            .ok_or_else(|| invalid("PtgExtraArray header is truncated"))?;
        let columns = usize::from(header[0]) + 1;
        let rows = usize::from(u16::from_le_bytes([header[1], header[2]])) + 1;
        let values = rows
            .checked_mul(columns)
            .ok_or_else(|| invalid("PtgExtraArray cardinality overflows"))?;
        position += 3;
        for _ in 0..values {
            position = validate_ser_ar(extra, position)?;
        }
    }
    if position != extra.len() {
        return Err(invalid("RgbExtra has trailing or unowned bytes"));
    }
    Ok(())
}

fn validate_ser_ar(extra: &[u8], position: usize) -> Result<usize> {
    let kind = *extra
        .get(position)
        .ok_or_else(|| invalid("SerAr value is truncated"))?;
    match kind {
        0x00 => fixed_ser_ar(extra, position, false, None),
        0x01 => {
            let end = fixed_ser_ar(extra, position, false, None)?;
            validate_xnum(&extra[position + 1..end], "SerNum")?;
            Ok(end)
        },
        0x02 => {
            let header = extra
                .get(position + 1..position + 4)
                .ok_or_else(|| invalid("SerStr header is truncated"))?;
            let count = usize::from(u16::from_le_bytes([header[0], header[1]]));
            if count > 255 || header[2] & !1 != 0 {
                return Err(invalid("SerStr has an invalid length or option flags"));
            }
            let width = if header[2] & 1 == 0 { 1 } else { 2 };
            let bytes = count
                .checked_mul(width)
                .ok_or_else(|| invalid("SerStr length overflows"))?;
            let end = position
                .checked_add(4)
                .and_then(|value| value.checked_add(bytes))
                .ok_or_else(|| invalid("SerStr length overflows"))?;
            if end > extra.len() {
                return Err(invalid("SerStr value is truncated"));
            }
            if width == 2 {
                let valid = extra[position + 4..end]
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
                if char::decode_utf16(valid).any(|value| value.is_err()) {
                    return Err(invalid("SerStr contains invalid UTF-16"));
                }
            }
            Ok(end)
        },
        0x04 => fixed_ser_ar(extra, position, true, Some(0..=1)),
        0x10 => fixed_ser_ar(extra, position, true, None).and_then(|end| {
            if matches!(
                extra[position + 1],
                0x00 | 0x07 | 0x0f | 0x17 | 0x1d | 0x24 | 0x2a
            ) {
                Ok(end)
            } else {
                Err(invalid("SerErr contains an invalid BIFF error code"))
            }
        }),
        _ => Err(invalid("SerAr contains an unknown value kind")),
    }
}

fn fixed_ser_ar(
    extra: &[u8],
    position: usize,
    reserved2_zero: bool,
    value_range: Option<std::ops::RangeInclusive<u8>>,
) -> Result<usize> {
    let end = position
        .checked_add(9)
        .ok_or_else(|| invalid("SerAr length overflows"))?;
    let value = extra
        .get(position..end)
        .ok_or_else(|| invalid("SerAr value is truncated"))?;
    if reserved2_zero && value[2] != 0 {
        return Err(invalid("SerAr reserved byte must be zero"));
    }
    if value_range.is_some_and(|range| !range.contains(&value[1])) {
        return Err(invalid("SerAr contains an invalid scalar value"));
    }
    Ok(end)
}

fn validate_xnum(bytes: &[u8], owner: &str) -> Result<()> {
    let raw: [u8; 8] = bytes
        .try_into()
        .map_err(|_error| invalid(format!("{owner} Xnum is truncated")))?;
    let bits = u64::from_le_bytes(raw);
    let exponent = (bits >> 52) & 0x7ff;
    let fraction = bits & 0x000f_ffff_ffff_ffff;
    if exponent == 0x7ff || (exponent == 0 && fraction != 0) || bits == 0x8000_0000_0000_0000 {
        return Err(invalid(format!(
            "{owner} Xnum cannot be non-finite, subnormal, or negative zero"
        )));
    }
    Ok(())
}

fn validate_safe_function(index: u16) -> Result<()> {
    if matches!(
        index,
        0 | 1 | 4 | 5 | 6 | 7 | 24 | 27 | 31 | 32 | 102 | 115 | 116 | 336
    ) {
        Ok(())
    } else {
        Err(unsafe_formula(
            "function is outside the inert array-formula authoring allow-list",
        ))
    }
}

fn classified_value_type(opcode: u8) -> Result<bool> {
    match (opcode >> 5) & 0x03 {
        1 => Ok(false),
        2 | 3 => Ok(true),
        _ => Err(invalid("classified Ptg has an invalid operand type")),
    }
}

fn validate_column(bytes: &[u8], offset: usize) -> Result<()> {
    if read_u16(bytes, offset)? & 0x3f00 != 0 {
        return Err(invalid("Ptg cell reference contains reserved column bits"));
    }
    Ok(())
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let pair = bytes.get(offset..offset + 2).ok_or_else(truncated)?;
    Ok(u16::from_le_bytes([pair[0], pair[1]]))
}

#[derive(Debug, Clone, Copy)]
struct TokenShape {
    wire_bytes: usize,
    actual_bytes: usize,
}

impl TokenShape {
    const fn fixed(bytes: usize) -> Self {
        Self {
            wire_bytes: bytes,
            actual_bytes: bytes,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct Node {
    value_type: bool,
    nesting: usize,
    operands: usize,
    depth: usize,
}

impl Node {
    const EMPTY: Self = Self {
        value_type: false,
        nesting: 0,
        operands: 0,
        depth: 0,
    };
}

struct SemanticStack {
    nodes: [Node; MAX_SEMANTIC_STACK],
    len: usize,
}

impl SemanticStack {
    const fn new() -> Self {
        Self {
            nodes: [Node::EMPTY; MAX_SEMANTIC_STACK],
            len: 0,
        }
    }

    const fn len(&self) -> usize {
        self.len
    }

    fn set_top_value_type(&mut self, value_type: bool) -> Result<()> {
        let index = self
            .len
            .checked_sub(1)
            .ok_or_else(|| invalid("memory Ptg completed without a result node"))?;
        self.nodes[index].value_type = value_type;
        Ok(())
    }

    fn operand(&mut self, value_type: bool) -> Result<()> {
        self.push(Node {
            value_type,
            nesting: 0,
            operands: 1,
            depth: 0,
        })
    }

    fn unary(&mut self, requires_value: bool, value_type: bool) -> Result<()> {
        let child = self.pop()?;
        if requires_value && !child.value_type {
            return Err(invalid("value operator received a reference expression"));
        }
        self.push(Node {
            value_type,
            nesting: child.nesting,
            operands: child.operands,
            depth: child
                .depth
                .checked_add(1)
                .ok_or_else(|| invalid("array-formula expression depth overflows"))?,
        })
    }

    fn parenthesize(&mut self) -> Result<()> {
        let child = self.pop()?;
        self.push(Node {
            depth: child
                .depth
                .checked_add(1)
                .ok_or_else(|| invalid("array-formula expression depth overflows"))?,
            ..child
        })
    }

    fn binary(&mut self, requires_value: bool, value_type: bool) -> Result<()> {
        let right = self.pop()?;
        let left = self.pop()?;
        if requires_value != left.value_type || requires_value != right.value_type {
            return Err(invalid(
                "binary formula operator received incompatible value/reference operands",
            ));
        }
        let right_pressure = right
            .operands
            .checked_add(1)
            .ok_or_else(|| invalid("array-formula operand count overflows"))?;
        self.push(Node {
            value_type,
            nesting: left.nesting.max(right.nesting),
            operands: left.operands.max(right_pressure),
            depth: left
                .depth
                .max(right.depth)
                .checked_add(1)
                .ok_or_else(|| invalid("array-formula expression depth overflows"))?,
        })
    }

    fn function(&mut self, arity: usize, value_type: bool) -> Result<()> {
        let start = self
            .len
            .checked_sub(arity)
            .ok_or_else(|| invalid("formula function has insufficient RPN arguments"))?;
        let mut nesting = 0usize;
        let mut operands = 0usize;
        let mut depth = 0usize;
        for (index, child) in self.nodes[start..self.len].iter().enumerate() {
            nesting = nesting.max(child.nesting);
            depth = depth.max(child.depth);
            operands = operands.max(
                child
                    .operands
                    .checked_add(index)
                    .ok_or_else(|| invalid("array-formula operand count overflows"))?,
            );
        }
        self.len = start;
        self.push(Node {
            value_type,
            nesting: nesting
                .checked_add(1)
                .ok_or_else(|| invalid("array-formula function nesting overflows"))?,
            operands,
            depth: depth
                .checked_add(1)
                .ok_or_else(|| invalid("array-formula expression depth overflows"))?,
        })
    }

    fn push(&mut self, node: Node) -> Result<()> {
        let slot = self
            .nodes
            .get_mut(self.len)
            .ok_or_else(|| invalid("array-formula semantic stack exceeds 40 operands"))?;
        *slot = node;
        self.len += 1;
        Ok(())
    }

    fn pop(&mut self) -> Result<Node> {
        self.len = self
            .len
            .checked_sub(1)
            .ok_or_else(|| invalid("array-formula RPN stack underflows"))?;
        Ok(self.nodes[self.len])
    }

    fn check_limits(&self, limits: Limits) -> Result<()> {
        for node in &self.nodes[..self.len] {
            if node.nesting > limits.max_nesting_depth() {
                return Err(invalid(
                    "array-formula function nesting exceeds the configured or normative limit",
                ));
            }
            if node.operands > limits.max_operands() {
                return Err(invalid(
                    "array-formula operand count exceeds the configured or normative limit",
                ));
            }
            if node.depth > limits.max_operator_depth() {
                return Err(invalid(
                    "array-formula expression depth exceeds the configured limit",
                ));
            }
        }
        Ok(())
    }

    fn finish(&self, limits: Limits) -> Result<()> {
        if self.len != 1 {
            return Err(invalid(
                "ArrayParsedFormula RPN expression must have exactly one root",
            ));
        }
        if !self.nodes[0].value_type {
            return Err(invalid(
                "ArrayParsedFormula root expression must be a value type",
            ));
        }
        self.check_limits(limits)
    }
}

fn truncated() -> Error {
    invalid("array formula contains a truncated Ptg")
}

fn unsafe_formula(message: impl Into<String>) -> Error {
    Error::UnsupportedFeature(message.into())
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: RECORD_TYPE,
        message: message.into(),
    }
}
