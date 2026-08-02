//! Worksheet range bindings for inert Office web extensions.
//!
//! Package relationships and XML vocabulary details live in [`crate::raw::web`].
//! This module keeps the ordinary API semantic: application references are the
//! primary key, numeric positions are checked, and collection edits commit only
//! after every invariant has been revalidated.

use std::collections::HashSet;

use litchi_ooxml_common::web::Panes;

use crate::error::{Result, allocation, invalid};

pub(crate) const MAX_BINDINGS: usize = 65_536;
pub(crate) const MAX_STRING_BYTES: usize = 32_767;
pub(crate) const MAX_TOTAL_STRING_BYTES: usize = 16 * 1024 * 1024;

/// One worksheet range connected to an MS-OWEXML package binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    app_ref: String,
    formula: String,
}

impl Binding {
    /// Create a checked worksheet binding.
    pub fn new(app_ref: impl Into<String>, formula: impl Into<String>) -> Result<Self> {
        let value = Self {
            app_ref: app_ref.into(),
            formula: formula.into(),
        };
        value.validate()?;
        Ok(value)
    }

    /// Package-level application reference used as this collection's key.
    #[must_use]
    pub fn app_ref(&self) -> &str {
        &self.app_ref
    }

    /// Sheet-qualified A1 range formula persisted in `xm:f`.
    #[must_use]
    pub fn formula(&self) -> &str {
        &self.formula
    }

    /// Replace the application reference after validating its local domain.
    ///
    /// Use [`Bindings::edit`] when the binding belongs to a collection so
    /// uniqueness is checked transactionally as well.
    pub fn set_app_ref(&mut self, app_ref: impl Into<String>) -> Result<&mut Self> {
        let app_ref = app_ref.into();
        bounded_nonempty(&app_ref, "web-extension appRef")?;
        self.app_ref = app_ref;
        Ok(self)
    }

    /// Replace the sheet-qualified A1 range formula.
    pub fn set_formula(&mut self, formula: impl Into<String>) -> Result<&mut Self> {
        let formula = formula.into();
        validate_formula(&formula)?;
        self.formula = formula;
        Ok(self)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        bounded_nonempty(&self.app_ref, "web-extension appRef")?;
        validate_formula(&self.formula)
    }

    fn string_bytes(&self) -> Result<usize> {
        self.app_ref
            .len()
            .checked_add(self.formula.len())
            .ok_or_else(|| invalid("worksheet web-binding string size overflow"))
    }
}

/// Checked lookup key. Application references are the semantic primary key;
/// positions remain available for ordered workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Selector<'a> {
    Id(&'a str),
    Index(usize),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Id(value)
    }
}

impl<'a> From<&'a String> for Selector<'a> {
    fn from(value: &'a String) -> Self {
        Self::Id(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

/// A bounded, application-reference-unique worksheet binding collection.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Bindings {
    values: Vec<Binding>,
    string_bytes: usize,
}

impl Bindings {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            values: Vec::new(),
            string_bytes: 0,
        }
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Binding> {
        self.values.iter()
    }

    /// Look up a binding by application reference or checked position.
    #[must_use]
    pub fn get<'a, 'key>(&'a self, selector: impl Into<Selector<'key>>) -> Option<&'a Binding> {
        match selector.into() {
            Selector::Id(app_ref) => self.values.iter().find(|value| value.app_ref == app_ref),
            Selector::Index(index) => self.values.get(index),
        }
    }

    /// Add a new binding, rejecting a duplicate application reference.
    pub fn add(&mut self, binding: Binding) -> Result<&mut Self> {
        binding.validate()?;
        if self.values.len() >= MAX_BINDINGS {
            return Err(invalid(format!(
                "worksheet web bindings exceed the {MAX_BINDINGS} item limit"
            )));
        }
        if self
            .values
            .iter()
            .any(|value| value.app_ref == binding.app_ref)
        {
            return Err(invalid(format!(
                "duplicate worksheet web-extension appRef '{}'",
                binding.app_ref
            )));
        }
        let string_bytes = checked_total(self.string_bytes, binding.string_bytes()?)?;
        self.values
            .try_reserve(1)
            .map_err(|source| allocation("worksheet web bindings", source))?;
        self.values.push(binding);
        self.string_bytes = string_bytes;
        Ok(self)
    }

    /// Insert by application reference or replace the existing value.
    ///
    /// The replaced value is returned without copying. `None` means a new
    /// binding was inserted.
    pub fn put(&mut self, binding: Binding) -> Result<Option<Binding>> {
        binding.validate()?;
        let Some(index) = self
            .values
            .iter()
            .position(|value| value.app_ref == binding.app_ref)
        else {
            self.add(binding)?;
            return Ok(None);
        };
        let old_bytes = self
            .values
            .get(index)
            .ok_or_else(|| invalid("worksheet web-binding index disappeared"))?
            .string_bytes()?;
        let base = self
            .string_bytes
            .checked_sub(old_bytes)
            .ok_or_else(|| invalid("worksheet web-binding size accounting underflow"))?;
        let string_bytes = checked_total(base, binding.string_bytes()?)?;
        let slot = self
            .values
            .get_mut(index)
            .ok_or_else(|| invalid("worksheet web-binding index disappeared"))?;
        let old = std::mem::replace(slot, binding);
        self.string_bytes = string_bytes;
        Ok(Some(old))
    }

    /// Edit one binding transactionally while preserving collection invariants.
    ///
    /// Returns `false` when the selector is absent. A rejected edit leaves the
    /// original binding and collection accounting unchanged.
    pub fn edit<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
        edit: impl FnOnce(&mut Binding) -> Result<()>,
    ) -> Result<bool> {
        let index = match selector.into() {
            Selector::Id(app_ref) => self
                .values
                .iter()
                .position(|value| value.app_ref == app_ref),
            Selector::Index(index) => (index < self.values.len()).then_some(index),
        };
        let Some(index) = index else {
            return Ok(false);
        };
        let original = self
            .values
            .get(index)
            .ok_or_else(|| invalid("worksheet web-binding index disappeared"))?;
        let old_bytes = original.string_bytes()?;
        let mut candidate = original.clone();
        edit(&mut candidate)?;
        candidate.validate()?;
        if self
            .values
            .iter()
            .enumerate()
            .any(|(other, value)| other != index && value.app_ref == candidate.app_ref)
        {
            return Err(invalid(format!(
                "duplicate worksheet web-extension appRef '{}'",
                candidate.app_ref
            )));
        }
        let base = self
            .string_bytes
            .checked_sub(old_bytes)
            .ok_or_else(|| invalid("worksheet web-binding size accounting underflow"))?;
        let string_bytes = checked_total(base, candidate.string_bytes()?)?;
        let slot = self
            .values
            .get_mut(index)
            .ok_or_else(|| invalid("worksheet web-binding index disappeared"))?;
        *slot = candidate;
        self.string_bytes = string_bytes;
        Ok(true)
    }

    /// Remove by application reference or checked position.
    pub fn remove<'key>(&mut self, selector: impl Into<Selector<'key>>) -> Option<Binding> {
        let index = match selector.into() {
            Selector::Id(app_ref) => self
                .values
                .iter()
                .position(|value| value.app_ref == app_ref)?,
            Selector::Index(index) if index < self.values.len() => index,
            Selector::Index(_) => return None,
        };
        let value = self.values.remove(index);
        self.string_bytes = self
            .string_bytes
            .saturating_sub(value.app_ref.len().saturating_add(value.formula.len()));
        Some(value)
    }

    /// Remove every binding and report whether the collection changed.
    pub fn clear(&mut self) -> bool {
        let changed = !self.values.is_empty();
        self.values.clear();
        self.string_bytes = 0;
        changed
    }

    #[must_use]
    pub fn into_vec(self) -> Vec<Binding> {
        self.values
    }

    pub(crate) fn validate_all(&self) -> Result<()> {
        if validate_values(&self.values)? != self.string_bytes {
            return Err(invalid("worksheet web-binding size accounting mismatch"));
        }
        Ok(())
    }
}

impl TryFrom<Vec<Binding>> for Bindings {
    type Error = crate::Error;

    fn try_from(values: Vec<Binding>) -> Result<Self> {
        let string_bytes = validate_values(&values)?;
        Ok(Self {
            values,
            string_bytes,
        })
    }
}

impl<'a> IntoIterator for &'a Bindings {
    type Item = &'a Binding;
    type IntoIter = std::slice::Iter<'a, Binding>;

    fn into_iter(self) -> Self::IntoIter {
        self.values.iter()
    }
}

impl IntoIterator for Bindings {
    type Item = Binding;
    type IntoIter = std::vec::IntoIter<Binding>;

    fn into_iter(self) -> Self::IntoIter {
        self.values.into_iter()
    }
}

/// Borrowed package application references used to validate worksheet links.
///
/// The set borrows the package model rather than copying every identifier.
#[derive(Debug, Clone)]
pub struct Refs<'a> {
    values: HashSet<&'a str>,
}

impl<'a> Refs<'a> {
    /// Build a bounded, unique reference set.
    pub fn new(values: impl IntoIterator<Item = &'a str>) -> Result<Self> {
        let mut refs = HashSet::new();
        let mut count = 0usize;
        for value in values {
            count = count
                .checked_add(1)
                .ok_or_else(|| invalid("MS-OWEXML binding count overflow"))?;
            if count > MAX_BINDINGS {
                return Err(invalid(format!(
                    "MS-OWEXML bindings exceed the {MAX_BINDINGS} item limit"
                )));
            }
            bounded_nonempty(value, "MS-OWEXML binding appRef")?;
            refs.try_reserve(1)
                .map_err(|source| allocation("MS-OWEXML binding reference set", source))?;
            if !refs.insert(value) {
                return Err(invalid(format!(
                    "duplicate MS-OWEXML binding appRef '{value}'"
                )));
            }
        }
        Ok(Self { values: refs })
    }

    /// Borrow every package binding reference from a task-pane graph.
    pub fn from_panes(panes: &'a Panes) -> Result<Self> {
        Self::new(
            panes
                .iter()
                .flat_map(|pane| pane.add_in().bindings().iter())
                .map(litchi_ooxml_common::web::Binding::app_ref),
        )
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    #[must_use]
    pub fn contains(&self, app_ref: &str) -> bool {
        self.values.contains(app_ref)
    }

    /// Require every worksheet binding to resolve to one package binding.
    pub fn check(&self, bindings: &Bindings) -> Result<()> {
        bindings.validate_all()?;
        for binding in bindings {
            if !self.contains(binding.app_ref()) {
                return Err(invalid(format!(
                    "worksheet web-extension appRef '{}' has no MS-OWEXML binding",
                    binding.app_ref()
                )));
            }
        }
        Ok(())
    }
}

fn checked_total(current: usize, additional: usize) -> Result<usize> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| invalid("worksheet web-binding string size overflow"))?;
    if total > MAX_TOTAL_STRING_BYTES {
        return Err(invalid(format!(
            "worksheet web-binding strings exceed the {MAX_TOTAL_STRING_BYTES} byte limit"
        )));
    }
    Ok(total)
}

fn validate_values(values: &[Binding]) -> Result<usize> {
    if values.len() > MAX_BINDINGS {
        return Err(invalid(format!(
            "worksheet web bindings exceed the {MAX_BINDINGS} item limit"
        )));
    }
    let mut seen = HashSet::new();
    seen.try_reserve(values.len())
        .map_err(|source| allocation("worksheet web-binding validation state", source))?;
    let mut string_bytes = 0usize;
    for value in values {
        value.validate()?;
        if !seen.insert(value.app_ref.as_str()) {
            return Err(invalid(format!(
                "duplicate worksheet web-extension appRef '{}'",
                value.app_ref
            )));
        }
        string_bytes = checked_total(string_bytes, value.string_bytes()?)?;
    }
    Ok(string_bytes)
}

fn validate_formula(value: &str) -> Result<()> {
    bounded_nonempty(value, "web-extension range formula")?;
    if value.trim() != value {
        return Err(invalid(
            "web-extension formula cannot contain surrounding whitespace",
        ));
    }
    let bang = find_unquoted_bang(value)?;
    let (sheet, range) = value.split_at(bang);
    validate_sheet_name(sheet)?;
    validate_a1_range(&range[1..])
}

fn find_unquoted_bang(value: &str) -> Result<usize> {
    let mut quoted = false;
    let bytes = value.as_bytes();
    let mut index = 0usize;
    let mut found = None;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' => {
                if quoted && bytes.get(index + 1) == Some(&b'\'') {
                    index += 1;
                } else {
                    quoted = !quoted;
                }
            },
            b'!' if !quoted && found.is_some() => {
                return Err(invalid(
                    "web-extension formula must contain one sheet qualifier",
                ));
            },
            b'!' if !quoted => found = Some(index),
            _ => {},
        }
        index += 1;
    }
    if quoted {
        return Err(invalid("unterminated quoted worksheet name"));
    }
    found.ok_or_else(|| invalid("web-extension formula requires a sheet qualifier"))
}

fn validate_sheet_name(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid("empty worksheet name in web-extension formula"));
    }
    if value.starts_with('\'') {
        if !value.ends_with('\'') || value.len() < 3 {
            return Err(invalid("invalid quoted worksheet name"));
        }
    } else if value
        .bytes()
        .any(|byte| !byte.is_ascii_alphanumeric() && byte != b'_' && byte != b'.')
    {
        return Err(invalid(
            "worksheet name must be quoted in web-extension formula",
        ));
    }
    Ok(())
}

fn validate_a1_range(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    validate_a1_cell(parts.next().unwrap_or_default())?;
    if let Some(last) = parts.next() {
        validate_a1_cell(last)?;
    }
    if parts.next().is_some() {
        return Err(invalid(
            "web-extension formula contains more than one range operator",
        ));
    }
    Ok(())
}

fn validate_a1_cell(value: &str) -> Result<()> {
    let value = value.strip_prefix('$').unwrap_or(value);
    let column_end = value
        .bytes()
        .position(|byte| !byte.is_ascii_alphabetic())
        .unwrap_or(value.len());
    if column_end == 0 || column_end > 3 {
        return Err(invalid("invalid column in web-extension range"));
    }
    let column = value[..column_end].bytes().fold(0u32, |number, byte| {
        number * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1)
    });
    let row = value[column_end..]
        .strip_prefix('$')
        .unwrap_or(&value[column_end..]);
    if column > 16_384
        || row.is_empty()
        || !row.bytes().all(|byte| byte.is_ascii_digit())
        || row
            .parse::<u32>()
            .ok()
            .filter(|row| (1..=1_048_576).contains(row))
            .is_none()
    {
        return Err(invalid("invalid cell in web-extension range"));
    }
    Ok(())
}

fn bounded_nonempty(value: &str, field: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!(
            "{field} must contain 1..={MAX_STRING_BYTES} bytes"
        )));
    }
    if value.chars().any(char::is_control) {
        return Err(invalid(format!("{field} contains a control character")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn binding(app_ref: &str, formula: &str) -> Binding {
        Binding::new(app_ref, formula).unwrap()
    }

    #[test]
    fn supports_semantic_and_checked_positional_crud() {
        let mut bindings = Bindings::new();
        bindings.add(binding("sales", "Sheet1!A1")).unwrap();
        bindings.add(binding("point", "Sheet1!B2")).unwrap();

        assert_eq!(
            bindings.get("sales").map(Binding::formula),
            Some("Sheet1!A1")
        );
        assert_eq!(bindings.get(1).map(Binding::app_ref), Some("point"));
        assert!(bindings.get(usize::MAX).is_none());

        let old = bindings
            .put(binding("sales", "Sheet1!C3"))
            .unwrap()
            .unwrap();
        assert_eq!(old.formula(), "Sheet1!A1");
        assert_eq!(
            bindings.get("sales").map(Binding::formula),
            Some("Sheet1!C3")
        );

        assert_eq!(
            bindings.remove("point").map(|value| value.app_ref),
            Some("point".into())
        );
        assert!(bindings.remove(99).is_none());
    }

    #[test]
    fn failed_edits_leave_the_collection_unchanged() {
        let mut bindings = Bindings::try_from(vec![
            binding("sales", "Sheet1!A1"),
            binding("point", "Sheet1!B2"),
        ])
        .unwrap();
        let before = bindings.clone();

        assert!(
            bindings
                .edit("sales", |value| {
                    value.set_app_ref("point")?;
                    Ok(())
                })
                .is_err()
        );
        assert_eq!(bindings, before);
        assert!(!bindings.edit(99, |_| Ok(())).unwrap());
    }

    #[test]
    fn validates_ranges_and_package_references() {
        for invalid_formula in [
            "A1",
            "!A1",
            "Sheet1!A0",
            "Sheet1!XFE1",
            "Sheet1:Sheet2!A1",
            "Sheet1!A1:B2:C3",
            "A!B!C1",
        ] {
            assert!(
                Binding::new("binding", invalid_formula).is_err(),
                "{invalid_formula}"
            );
        }

        let bindings = Bindings::try_from(vec![binding("sales", "Sheet1!A1")]).unwrap();
        let refs = Refs::new(["sales", "other"]).unwrap();
        refs.check(&bindings).unwrap();
        assert!(Refs::new(["sales", "sales"]).is_err());
        assert!(Refs::new(["other"]).unwrap().check(&bindings).is_err());
    }
}
