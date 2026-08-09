use super::super::Result;
use super::super::codec::{invalid, limit, require_nonempty};
use super::super::validation::{
    validate_binding, validate_extension_list, validate_model, validate_store_reference,
};
use super::{
    Binding, ExtKind, ExtList, MAX_WEB_EXTENSION_ITEMS, Property, Reference, Selector, Snapshot,
};
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddIn {
    pub(in crate::web) id: String,
    pub(in crate::web) frozen: bool,
    pub(in crate::web) reference: Reference,
    pub(in crate::web) alternate_references: Vec<Reference>,
    pub(in crate::web) properties: Vec<Property>,
    pub(in crate::web) bindings: Vec<Binding>,
    pub(in crate::web) snapshot: Option<Snapshot>,
    pub(in crate::web) extension_list: Option<ExtList>,
}

impl AddIn {
    pub fn new(id: impl Into<String>, reference: Reference) -> Result<Self> {
        let value = Self {
            id: id.into(),
            frozen: false,
            reference,
            alternate_references: Vec::new(),
            properties: Vec::new(),
            bindings: Vec::new(),
            snapshot: None,
            extension_list: None,
        };
        validate_model(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn frozen(mut self, frozen: bool) -> Self {
        self.frozen = frozen;
        self
    }

    pub fn set_frozen(&mut self, frozen: bool) -> &mut Self {
        self.frozen = frozen;
        self
    }

    pub fn bind(mut self, binding: Binding) -> Result<Self> {
        self.push_binding(binding)?;
        Ok(self)
    }

    pub fn push_binding(&mut self, binding: Binding) -> Result<&mut Self> {
        validate_binding(&binding)?;
        if self.bindings.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "web extension bindings",
                MAX_WEB_EXTENSION_ITEMS,
                self.bindings.len().saturating_add(1),
            );
        }
        if self.bindings.iter().any(|value| value.id == binding.id) {
            return invalid(format!("duplicate binding id '{}'", binding.id));
        }
        if self
            .bindings
            .iter()
            .any(|value| value.app_ref == binding.app_ref)
        {
            return invalid(format!("duplicate binding appRef '{}'", binding.app_ref));
        }
        self.bindings.push(binding);
        Ok(self)
    }

    pub fn upsert_binding(&mut self, binding: Binding) -> Result<&mut Self> {
        validate_binding(&binding)?;
        if let Some(index) = self
            .bindings
            .iter()
            .position(|value| value.id == binding.id)
        {
            if self
                .bindings
                .iter()
                .enumerate()
                .any(|(other, value)| other != index && value.app_ref == binding.app_ref)
            {
                return invalid(format!("duplicate binding appRef '{}'", binding.app_ref));
            }
            self.bindings[index] = binding;
            return Ok(self);
        }
        self.push_binding(binding)
    }

    #[must_use]
    pub fn binding<'key>(&self, selector: impl Into<Selector<'key>>) -> Option<&Binding> {
        match selector.into() {
            Selector::Id(id) => self.bindings.iter().find(|value| value.id == id),
            Selector::Index(index) => self.bindings.get(index),
        }
    }

    #[must_use]
    pub fn binding_mut<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<&mut Binding> {
        match selector.into() {
            Selector::Id(id) => self.bindings.iter_mut().find(|value| value.id == id),
            Selector::Index(index) => self.bindings.get_mut(index),
        }
    }

    pub fn remove_binding<'key>(&mut self, selector: impl Into<Selector<'key>>) -> Option<Binding> {
        let index = match selector.into() {
            Selector::Id(id) => self.bindings.iter().position(|value| value.id == id)?,
            Selector::Index(index) if index < self.bindings.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.bindings.remove(index))
    }

    pub fn prop(mut self, property: Property) -> Result<Self> {
        self.push_property(property)?;
        Ok(self)
    }

    pub fn push_property(&mut self, property: Property) -> Result<&mut Self> {
        if self.properties.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "web extension properties",
                MAX_WEB_EXTENSION_ITEMS,
                self.properties.len().saturating_add(1),
            );
        }
        if self
            .properties
            .iter()
            .any(|value| value.name == property.name)
        {
            return invalid(format!("duplicate property name '{}'", property.name));
        }
        self.properties.push(property);
        Ok(self)
    }

    pub fn upsert_property(&mut self, property: Property) -> Result<&mut Self> {
        require_nonempty("property name", &property.name)?;
        if let Some(index) = self
            .properties
            .iter()
            .position(|value| value.name == property.name)
        {
            self.properties[index] = property;
            return Ok(self);
        }
        self.push_property(property)
    }

    #[must_use]
    pub fn property<'key>(&self, selector: impl Into<Selector<'key>>) -> Option<&Property> {
        match selector.into() {
            Selector::Id(name) => self.properties.iter().find(|value| value.name == name),
            Selector::Index(index) => self.properties.get(index),
        }
    }

    pub fn remove_property<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<Property> {
        let index = match selector.into() {
            Selector::Id(name) => self
                .properties
                .iter()
                .position(|value| value.name == name)?,
            Selector::Index(index) if index < self.properties.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.properties.remove(index))
    }

    pub fn push_reference(&mut self, reference: Reference) -> Result<&mut Self> {
        validate_store_reference(&reference)?;
        if self.alternate_references.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "alternate references",
                MAX_WEB_EXTENSION_ITEMS,
                self.alternate_references.len().saturating_add(1),
            );
        }
        if reference.id == self.reference.id
            || self
                .alternate_references
                .iter()
                .any(|value| value.id == reference.id)
        {
            return invalid(format!("duplicate reference id '{}'", reference.id));
        }
        self.alternate_references.push(reference);
        Ok(self)
    }

    pub fn upsert_reference(&mut self, reference: Reference) -> Result<&mut Self> {
        validate_store_reference(&reference)?;
        if reference.id == self.reference.id {
            return invalid(format!(
                "alternate reference id '{}' duplicates the primary reference",
                reference.id
            ));
        }
        if let Some(index) = self
            .alternate_references
            .iter()
            .position(|value| value.id == reference.id)
        {
            self.alternate_references[index] = reference;
            return Ok(self);
        }
        self.push_reference(reference)
    }

    #[must_use]
    pub fn alternate_reference<'key>(
        &self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<&Reference> {
        match selector.into() {
            Selector::Id(id) => self
                .alternate_references
                .iter()
                .find(|value| value.id == id),
            Selector::Index(index) => self.alternate_references.get(index),
        }
    }

    #[must_use]
    pub fn alternate_reference_mut<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<&mut Reference> {
        match selector.into() {
            Selector::Id(id) => self
                .alternate_references
                .iter_mut()
                .find(|value| value.id == id),
            Selector::Index(index) => self.alternate_references.get_mut(index),
        }
    }

    pub fn remove_reference<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
    ) -> Option<Reference> {
        let index = match selector.into() {
            Selector::Id(id) => self
                .alternate_references
                .iter()
                .position(|value| value.id == id)?,
            Selector::Index(index) if index < self.alternate_references.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.alternate_references.remove(index))
    }

    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    #[must_use]
    pub const fn is_frozen(&self) -> bool {
        self.frozen
    }

    #[must_use]
    pub const fn reference(&self) -> &Reference {
        &self.reference
    }

    pub fn set_reference(&mut self, reference: Reference) -> Result<&mut Self> {
        validate_store_reference(&reference)?;
        if self
            .alternate_references
            .iter()
            .any(|alternate| alternate.id == reference.id)
        {
            return invalid(format!(
                "primary reference id '{}' duplicates an alternate reference",
                reference.id
            ));
        }
        self.reference = reference;
        Ok(self)
    }

    pub const fn reference_mut(&mut self) -> &mut Reference {
        &mut self.reference
    }

    #[must_use]
    pub fn alternate_references(&self) -> &[Reference] {
        &self.alternate_references
    }

    #[must_use]
    pub fn properties(&self) -> &[Property] {
        &self.properties
    }

    #[must_use]
    pub fn bindings(&self) -> &[Binding] {
        &self.bindings
    }

    #[must_use]
    pub const fn snapshot(&self) -> Option<&Snapshot> {
        self.snapshot.as_ref()
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::AddIn])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}
