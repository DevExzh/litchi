//! Shared catalog for reference geometry and workbook metadata functions.

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, UnsupportedKind,
    WorkingValue,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Function {
    Areas,
    Column,
    Columns,
    IsRef,
    Row,
    Rows,
    Sheet,
    Sheets,
}

impl Function {
    pub(super) fn from_name(name: &str) -> Option<Self> {
        let first = name.as_bytes().first()?.to_ascii_uppercase();
        // Separate common existing names such as SUM before comparing the
        // metadata spellings without allocating a normalized name.
        match (first, name.len()) {
            (b'A', 5) if name.eq_ignore_ascii_case("AREAS") => Some(Self::Areas),
            (b'C', 6) if name.eq_ignore_ascii_case("COLUMN") => Some(Self::Column),
            (b'C', 7) if name.eq_ignore_ascii_case("COLUMNS") => Some(Self::Columns),
            (b'I', 5) if name.eq_ignore_ascii_case("ISREF") => Some(Self::IsRef),
            (b'R', 3) if name.eq_ignore_ascii_case("ROW") => Some(Self::Row),
            (b'R', 4) if name.eq_ignore_ascii_case("ROWS") => Some(Self::Rows),
            (b'S', 5) if name.eq_ignore_ascii_case("SHEET") => Some(Self::Sheet),
            (b'S', 6) if name.eq_ignore_ascii_case("SHEETS") => Some(Self::Sheets),
            _ => None,
        }
    }

    pub(super) const fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Column | Self::Row | Self::Sheet | Self::Sheets => count <= 1,
            Self::Areas | Self::Columns | Self::IsRef | Self::Rows => count == 1,
        }
    }
}

pub(super) fn is_reference_metadata_function(name: &str) -> bool {
    Function::from_name(name).is_some()
}

/// Apply the context-free scalar subset. Reference and array expressions are
/// refused by the scalar scheduler before this bridge; the value evaluator
/// owns the corresponding descriptor and workbook operations.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let function = Function::from_name(name)
        .ok_or(EvaluationFailure::Unsupported(UnsupportedKind::Function))?;
    if !function.valid_arity(node.child_count()) {
        return evaluator.finish_invalid_arity(node);
    }
    evaluator.charge_work(1)?;
    if node.child_count() == 0 {
        return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
    }
    let argument = evaluator.pop_value()?;
    let result = if function == Function::IsRef {
        WorkingValue::Logical(false)
    } else {
        match argument {
            WorkingValue::Error(error) => WorkingValue::Error(error),
            WorkingValue::Text(_) | WorkingValue::Number(_) | WorkingValue::Logical(_)
                if function == Function::Sheet =>
            {
                return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
            },
            _ => WorkingValue::Error(ScalarError::Value),
        }
    };
    evaluator.push_value(result)
}
