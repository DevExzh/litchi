use super::Package;

use litchi_iwa_protos::package_metadata_codec::{
    ComponentDescriptor, ComponentSelector, DataReferenceOwnerDescriptor,
    ExternalReferenceDescriptor, ObjectUuidDescriptor, PackageMetadataVisitor, RewriteError,
};

pub(super) const ENTRY_NAME: &str = "Index/Metadata.iwa";
pub(super) const MESSAGE_TYPE: u32 = 11_006;

#[derive(Debug, Clone, Copy)]
pub(super) struct MessageRoute {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
}

pub(super) fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

pub(super) fn unique_message_route(source: &Package) -> Option<MessageRoute> {
    let mut route = None;
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != MESSAGE_TYPE {
                    continue;
                }
                if route.is_some() || component.name() != ENTRY_NAME {
                    return None;
                }
                route = Some(MessageRoute {
                    component_index,
                    object_index,
                    message_index,
                });
            }
        }
    }
    route
}

pub(super) struct ComponentSelectorVisitor<'source> {
    target_locators: &'source [&'source str],
    identifiers: Vec<Option<u64>>,
    duplicate: bool,
}

impl<'source> ComponentSelectorVisitor<'source> {
    pub(super) fn new(target_locators: &'source [&'source str]) -> Result<Self, usize> {
        let mut identifiers = Vec::new();
        identifiers
            .try_reserve_exact(target_locators.len())
            .map_err(|_| target_locators.len())?;
        identifiers.resize(target_locators.len(), None);
        Ok(Self {
            target_locators,
            identifiers,
            duplicate: false,
        })
    }

    pub(super) fn into_selectors(self) -> Result<Vec<ComponentSelector<'source>>, usize> {
        if self.duplicate || self.identifiers.iter().any(Option::is_none) {
            return Err(0);
        }
        let mut selectors = Vec::new();
        selectors
            .try_reserve_exact(self.identifiers.len())
            .map_err(|_| self.identifiers.len())?;
        for (index, identifier) in self.identifiers.into_iter().enumerate() {
            selectors.push(ComponentSelector::new(
                identifier.ok_or(0usize)?,
                self.target_locators[index],
            ));
        }
        Ok(selectors)
    }
}

impl PackageMetadataVisitor for ComponentSelectorVisitor<'_> {
    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if !component.is_current() {
            return Ok(());
        }
        for (index, locator) in self.target_locators.iter().enumerate() {
            if component.effective_locator() == *locator {
                if self.identifiers[index].is_some() {
                    self.duplicate = true;
                } else {
                    self.identifiers[index] = Some(component.identifier());
                }
                break;
            }
        }
        Ok(())
    }
}

pub(super) struct ObjectOwnershipVisitor {
    object_identifier: u64,
    owned: bool,
}

impl ObjectOwnershipVisitor {
    pub(super) const fn new(object_identifier: u64) -> Self {
        Self {
            object_identifier,
            owned: false,
        }
    }

    pub(super) const fn is_owned(&self) -> bool {
        self.owned
    }
}

impl PackageMetadataVisitor for ObjectOwnershipVisitor {
    fn visit_object_uuid(&mut self, binding: ObjectUuidDescriptor<'_>) -> Result<(), RewriteError> {
        self.owned |= binding.object_identifier() == self.object_identifier;
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        self.owned |= reference.object_identifier() == Some(self.object_identifier);
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        self.owned |= owner.object_identifier() == self.object_identifier;
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), RewriteError> {
        self.owned |= identifier == self.object_identifier;
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), RewriteError> {
        self.owned |= object_identifier == self.object_identifier;
        Ok(())
    }
}
