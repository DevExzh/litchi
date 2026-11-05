"""Classify every `pub fn` in crates/litchi-iwa/src into an exit family.

Input: pubfns.tsv (path:line<TAB>name) from a `pub fn` scan.
Output: TSV with format, family, destination, parity, evidence.
"""
import re
import sys

ROWS = [l.rstrip('\n').split('\t') for l in open(sys.argv[1]) if l.strip()]

# (regex on name, family) checked in order; the format prefix is stripped first.
FAMILIES = [
    (r'conditional_highlight', 'table.conditional_highlight'),
    (r'table_cell_comment|^(set|clear)_cell_comment|cell_comment_repl', 'table.cell_comment'),
    (r'drawable_comment', 'drawable.comment'),
    (r'text_box_comment|_comment_repl|text_box_comments', 'text.comment'),
    (r'(table|slide_table).*(merge|unmerge)|^(merge|unmerge)_', 'table.merge'),
    (r'formula', 'table.formula'),
    (r'sort', 'table.sort'),
    (r'hidden_axes', 'table.hidden_axes'),
    (r'(insert|remove)_.*table_(row|column)|insert_table_(row|column)|remove_table_(row|column)', 'table.topology'),
    (r'table_cell_.*(_format|data_format)$|table_cell_(numeral|checkbox|star|slider|stepper|pop_up|custom|currency|percentage|scientific|fraction|date_time|duration|text)_format', 'table.cell_format'),
    (r'table_cell_(text|paragraph)', 'table.cell_rich_text'),
    (r'table_cell_(fill|border|layout)', 'table.cell_appearance'),
    (r'pivot', 'table.pivot'),
    (r'(add|duplicate|remove|resize)_(empty_)?(slide_)?table$|set_slide_table_geometry|^slide_tables?$|^tables$|add_empty_table|duplicate_table|remove_table$|resize_table', 'table.lifecycle'),
    (r'set_(slide_)?table_cells?$|clear_slide_table_cell$|get_cell|iter_cells|cell_count|^merges$|get_comment|iter_comments|comment_count', 'table.cells'),
    (r'chart', 'chart'),
    (r'date_time_field', 'text.date_time_field'),
    (r'number_attachment', 'text.number_attachment'),
    (r'bookmark', 'text.bookmark'),
    (r'named_paragraph_style|following_style', 'text.paragraph_style'),
    (r'hyperlink', 'text.hyperlink'),
    (r'highlight', 'text.highlight'),
    (r'drop_cap', 'text.drop_cap'),
    (r'text_box_text_language|language_boundary', 'text.language'),
    (r'text_box_paragraph_list', 'text.paragraph_list'),
    (r'text_box_paragraph', 'text.paragraph'),
    (r'text_box_text', 'text.character'),
    (r'text_box_(geometry|properties|text_layout|columns)', 'drawable.text_box_properties'),
    (r'(add|duplicate|remove)_.*text_box$|^add_text_box$|^duplicate_text_box$|^remove_text_box$', 'drawable.text_box_lifecycle'),
    (r'(_shape|_line|rectangle)', 'drawable.shape'),
    (r'image', 'drawable.image'),
    (r'movie', 'drawable.movie'),
    (r'audio', 'drawable.audio'),
    (r'media', 'media.asset'),
    (r'drawable_order|move_.*drawable$', 'drawable.order'),
    (r'drawables$', 'drawable.list'),
    (r'build|_in$|_out$|_action$|move_along_path|timing_curve|^set_(start|delay|duration|effect|move_path|move_alignment|action_acceleration)$|^with_|^(start|duration|delay|effect)$|^(straight|cubic|linear|from_path|semantic|recalculate_natural_size|as_u64)$', 'keynote.build'),
    (r'slide_layout|^add_slide$|^insert_slide$|^duplicate_slide$', 'slide.lifecycle'),
    (r'section', 'section.lifecycle'),
    (r'text_storage|drawable_text|body_storage|body_text|clear_body|set_body_text|replace_body_text', 'text.storage'),
    (r'^(add_empty_sheet|remove_sheet|duplicate_sheet|sheets)$', 'sheet.lifecycle'),
    (r'^(open|from_bytes|from_package|from_shared_bytes|from_read_at|package|into_package|to_bytes|save|write_to|validate|snapshot|create|create_with_text|builder|new|edit|apply|commit|reset|len|is_empty|slides)$', 'io'),
]

PARITY = {
    # (format, family): (status, evidence)
    ('numbers', 'table.merge'): ('Y', 'litchi_numbers::Package::edit_table_merges / MergeReader'),
    ('pages', 'table.merge'): ('Y', 'litchi_pages::Package::edit_body_table_merges'),
    ('keynote', 'table.merge'): ('Y', 'litchi_keynote::Package::edit_slide_table_merges'),
    ('numbers', 'table.sort'): ('Y', 'litchi_numbers::Package::edit_table_sort_order / apply_table_sort_order (host delegates)'),
    ('pages', 'table.sort'): ('Y', 'litchi_pages::Package::edit_body_table_sort_order (host delegates)'),
    ('keynote', 'table.sort'): ('Y', 'litchi_keynote::Package::edit_slide_table_sort_order / edit_slide_table_physical_sort'),
    ('pages', 'table.hidden_axes'): ('Y', 'litchi_pages::Package::edit_body_table_hidden_axes'),
    ('numbers', 'table.hidden_axes'): ('N', '-'),
    ('keynote', 'table.hidden_axes'): ('N', '-'),
    ('numbers', 'table.topology'): ('Y', 'litchi_numbers::Package::edit_table_topology (insert/remove row/column)'),
    ('pages', 'table.topology'): ('N', '-'),
    ('keynote', 'table.topology'): ('P', 'host delegates to a crate-internal edit_slide_table_topology; no public family'),
    ('numbers', 'table.cell_format'): ('P', 'litchi_numbers::Package::edit_table_cell_{number,currency,percentage,scientific,fraction,date_time,duration,text,custom,control,pop_up_menu}_format; generic DataFormat/numeral-system routes host-only'),
    ('keynote', 'table.cell_format'): ('P', 'litchi_keynote::Package::edit_slide_table_cell_number_format only'),
    ('pages', 'table.cell_format'): ('N', '-'),
    ('numbers', 'table.cell_rich_text'): ('P', 'litchi_numbers::Package::edit_table_cell_text_style (character properties); paragraph/list properties host-only'),
    ('pages', 'table.cell_rich_text'): ('N', '-'),
    ('keynote', 'table.cell_rich_text'): ('N', '-'),
    ('numbers', 'table.cell_appearance'): ('P', 'litchi_numbers::Package::edit_table_appearance (table-level); per-cell fill/border/layout host-only'),
    ('pages', 'table.cell_appearance'): ('P', 'litchi_pages::Package::edit_body_table_appearance (table-level only)'),
    ('keynote', 'table.cell_appearance'): ('P', 'litchi_keynote::Package::edit_slide_table_appearance (table-level only)'),
    ('numbers', 'table.conditional_highlight'): ('B', 'branch refactor/numbers-conditional-highlight-admission-review-20261024 (edit_table_cell_conditional_highlighting)'),
    ('pages', 'table.conditional_highlight'): ('N', '-'),
    ('keynote', 'table.conditional_highlight'): ('N', '-'),
    ('numbers', 'table.cell_comment'): ('P', 'litchi_numbers::Package::edit_table_cell_comment[_replies]; storage creation and identity fallbacks host-only'),
    ('pages', 'table.cell_comment'): ('B', 'branch feat/pages-table-comment-ownership-20261024 (body_table_comments)'),
    ('keynote', 'table.cell_comment'): ('N', '-'),
    ('numbers', 'table.formula'): ('P', 'litchi_numbers::Package::edit_table_cells set_formula/set_formula_cached; AST construction host-only'),
    ('pages', 'table.formula'): ('N', '-'),
    ('keynote', 'table.formula'): ('N', '-'),
    ('numbers', 'table.lifecycle'): ('P', 'litchi_numbers::Package::edit_table_lifecycle add/remove; duplicate and resize host-only (branch refactor/iwa-table-lifecycle-native-tests-20261023)'),
    ('pages', 'table.lifecycle'): ('P', 'litchi_pages::Package::edit_body_table_creation / edit_body_table_deletion; duplicate host-only'),
    ('keynote', 'table.lifecycle'): ('P', 'litchi_keynote::Package::edit_slide_drawables add_table/remove; duplicate/resize/geometry on branch feat/keynote-table-lifecycle-focused-20261024'),
    ('numbers', 'table.cells'): ('Y', 'litchi_numbers::Package::edit_table_cells'),
    ('pages', 'table.cells'): ('Y', 'litchi_pages::Package::edit_body_table_values'),
    ('keynote', 'table.cells'): ('Y', 'litchi_keynote::Package::edit_slide_drawables set_table_cell / slide_table_cells'),
    ('numbers', 'table.pivot'): ('N', '-'),
    ('numbers', 'chart.lifecycle'): ('P', 'litchi_numbers::Package::edit_sheet_chart_data (data only); create/duplicate/remove/kind/direction/geometry host-only'),
    ('pages', 'chart.lifecycle'): ('P', 'litchi_pages::Package::edit_body_chart_data (data only); create/duplicate/remove/kind/direction/geometry host-only'),
    ('keynote', 'chart.lifecycle'): ('B', 'litchi_keynote::Package::edit_slide_chart_data; create/duplicate/remove on integration line refactor/iwa-exit-current-20261023 + repair/keynote-chart-deploy-20261024'),
    ('keynote', 'chart.text'): ('Y', 'litchi_keynote::Package::edit_slide_chart_{title,caption,axis_title}'),
    ('keynote', 'chart.legend_axis'): ('Y', 'litchi_keynote::Package::edit_slide_chart_{legend,value_axis_settings}'),
    ('shared', 'chart.style'): ('P', 'neutral values partly in litchi_iwa_common::chart; host codecs over generated messages'),
    ('pages', 'text.date_time_field'): ('N', '-'),
    ('numbers', 'text.date_time_field'): ('N', '-'),
    ('keynote', 'text.date_time_field'): ('N', '-'),
    ('pages', 'text.number_attachment'): ('B', 'branch refactor/iwa-attachment-reader-admission-20261023 (body_number_attachments)'),
    ('pages', 'text.bookmark'): ('N', '-'),
    ('pages', 'section.lifecycle'): ('Y', 'litchi_pages::Package::edit_section_lifecycle (add/remove)'),
    ('pages', 'text.storage'): ('P', 'litchi_pages::Package::edit_body_text / edit_section_text / edit_header_footer_text; drawable storages host-only'),
    ('keynote', 'text.storage'): ('P', 'litchi_keynote::Package::edit_slide_text / edit_slide_body / edit_slide_title / edit_slide_notes'),
    ('keynote', 'slide.lifecycle'): ('P', 'litchi_keynote::Package::edit_slide_creation (add/duplicate) / edit_slide_deletion; set_slide_layout on integration line (slide_layout_reassignment)'),
    ('numbers', 'sheet.lifecycle'): ('P', 'litchi_numbers::Package::edit_sheet_lifecycle add/remove; duplicate on integration line (sheet_duplication)'),
    ('keynote', 'keynote.build'): ('P', 'litchi_keynote::Package::edit_slide_build_order (move only); build add/set/remove and effect values host-only'),
    ('pages', 'drawable.order'): ('Y', 'litchi_pages::Package::edit_body_drawable_order'),
    ('keynote', 'drawable.order'): ('P', 'litchi_keynote::Package::edit_slide_drawables move_to'),
    ('numbers', 'drawable.order'): ('N', '-'),
    ('keynote', 'drawable.comment'): ('Y', 'litchi_keynote::Package::edit_slide_drawable_comment'),
    ('pages', 'drawable.comment'): ('B', 'integration line refactor/iwa-exit-current-20261023 (pages package/drawable_comments.rs)'),
    ('numbers', 'drawable.comment'): ('B', 'branch feat/numbers-drawable-comments-focused-rereview-20261024'),
    ('keynote', 'drawable.text_box_lifecycle'): ('P', 'litchi_keynote::Package::edit_slide_drawables add_text_box/remove; duplicate host-only (branch refactor/keynote-shape-text-repair-20261024-r5)'),
    ('keynote', 'drawable.text_box_properties'): ('P', 'litchi_keynote::Package::edit_slide_drawables set_geometry'),
    ('keynote', 'drawable.shape'): ('P', 'litchi_keynote::Package::edit_slide_drawables add_shape/set_geometry/set_text; style properties host-only (branch refactor/keynote-shape-text-repair-20261024-r5)'),
    ('pages', 'drawable.shape'): ('B', 'branch repair/pages-body-shape-a408-20261024 (body_shape_lifecycle, body_shape_properties)'),
    ('keynote', 'drawable.image'): ('P', 'litchi_keynote::Package::edit_slide_drawables add_image / edit_slide_image_adjustments; lifecycle on integration line (slide_image_lifecycle)'),
    ('numbers', 'drawable.image'): ('P', 'litchi_numbers::Package::edit_sheet_image_adjustments only'),
    ('pages', 'drawable.image'): ('P', 'litchi_pages::Package::edit_body_image_adjustments only; lifecycle on branch refactor/pages-media-lifecycle-admission-20261024'),
    ('keynote', 'drawable.movie'): ('Y', 'litchi_keynote::Package::edit_slide_movie_{creation,geometry,playback_settings,caption,title} / edit_slide_media_lifecycle'),
    ('keynote', 'drawable.audio'): ('Y', 'litchi_keynote::Package::edit_slide_audio_{creation,position} / edit_slide_media_lifecycle'),
    ('pages', 'drawable.movie'): ('B', 'branch refactor/pages-media-lifecycle-admission-20261024 (body_media)'),
    ('pages', 'drawable.audio'): ('B', 'branch refactor/pages-media-lifecycle-admission-20261024 (body_media)'),
    ('keynote', 'media.asset'): ('P', 'litchi_keynote::Package::edit_slide_media_data (slide media); generic asset replace host-only'),
}


def fmt_of(path):
    head = path.split('/')[0].split('.')[0]
    return head if head in ('numbers', 'pages', 'keynote') else 'shared'


def family_of(name, path):
    if '/creation' in path:
        return 'builder'
    stripped = name
    for p in ('slide_', 'sheet_', 'body_'):
        pass
    for rx, fam in FAMILIES:
        if re.search(rx, stripped):
            if fam == 'chart':
                if re.search(r'^(add|duplicate|remove)_\w*chart$|chart_(kind|data|direction|geometry)$|charts$', name):
                    return 'chart.lifecycle'
                if re.search(r'chart_(title|caption|axis_title)$', name):
                    return 'chart.text'
                if re.search(r'chart_legend_visible$|value_axis_(bounds|scale|steps)$', name):
                    return 'chart.legend_axis'
                return 'chart.style'
            return fam
    top = path.split('/')[0].split('.')[0]
    if top in ('text', 'charts', 'shapes', 'media', 'comments'):
        return {'text': 'text.shared', 'charts': 'chart', 'shapes': 'drawable.shape', 'media': 'media.asset', 'comments': 'drawable.comment'}[top]
    return 'other'


DEST = {
    'numbers': 'litchi-numbers', 'pages': 'litchi-pages', 'keynote': 'litchi-keynote',
}
SHARED_DEST = {
    'chart': 'litchi-iwa-common::chart (values) + litchi-iwa-protos chart codecs (private)',
    'text': 'litchi-iwa-text (values) + litchi-iwa-text-wire (storage transforms)',
    'drawable': 'litchi-iwa-common::shape (values) + litchi-iwa-protos shape codecs',
    'media': 'litchi-iwa-archive (Data/ members, unstable) + leaf media families',
    'table': 'litchi-numbers-wire (TST/BNC writers) + litchi-iwa-common::table (values)',
    'io': 'litchi-iwa-archive::unstable (physical) / leaf Package',
    'builder': 'leaf Builder',
}


def dest_of(fmt, fam):
    if fmt in DEST:
        return DEST[fmt]
    return SHARED_DEST.get(fam.split('.')[0], 'litchi-iwa-archive / litchi-iwa-common (see record 0363 section 3)')


print('\t'.join(['path', 'name', 'format', 'family', 'destination', 'leaf_parity', 'evidence']))
for loc, name in ROWS:
    path = loc.split(':')[0]
    fmt = fmt_of(path)
    fam = family_of(name, path)
    status, ev = PARITY.get((fmt, fam), ('Y' if fam == 'io' else 'N', 'leaf Package / Builder' if fam == 'io' else '-'))
    if fam == 'builder':
        status, ev = ('P', {'numbers': 'litchi_numbers::create::Builder (identity, sheet_name, table_name); table_dimensions/language/locale host-only',
                            'pages': 'litchi_pages::builder::Builder (identity, body_text); language/locale/body_table host-only',
                            'keynote': 'litchi_keynote::Builder (identity, title, subtitle, notes, slide_size); language/locale host-only'}.get(fmt, '-'))
    print('\t'.join([loc, name, fmt, fam, dest_of(fmt, fam), status, ev]))
