//! Presentation-structure OPC facade.

pub use super::codec::{
    add_custom_show, add_custom_show_slide, add_section, add_section_slide, find_custom_show,
    find_section, load, remove_custom_show, remove_custom_show_slide, remove_section,
    remove_section_slide, reorder_custom_show_slides, reorder_custom_shows, reorder_section_slides,
    reorder_sections, replace_custom_show, replace_section, store,
    synchronize_after_slide_mutation, update_custom_show, update_section,
};
