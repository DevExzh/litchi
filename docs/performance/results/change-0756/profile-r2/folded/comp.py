import sys, re
from collections import Counter
exec(open(__import__('os').path.join(__import__('os').path.dirname(__import__('os').path.abspath(__file__)), 'an.py')).read().split("if __name__")[0])
CASES['pptx_stream'] = ('pptx', r'^write_pptx_stream', r'.')
CASES['xlsx_stream'] = ('xlsx', r'^write_streaming_xlsx', r'.')
RULES = [
 ('memset', r'^(__memset|memset)'),
 ('memcpy/memmove', r'^(__memmove|__memcpy|memcpy|memmove|copy_nonoverlapping|copy_from_slice)'),
 ('memcmp', r'^(__memcmp|__bcmp|bcmp|memcmp)'),
 ('alloc/free', r'^(__libc_malloc|__libc_free|__libc_realloc|__libc_calloc|_int_malloc|_int_free|_int_realloc|malloc$|free$|realloc$|tcache_|sysmalloc|__brk|__glibc_morecore|unlink_chunk|malloc_consolidate|__rust_alloc|__rust_dealloc|__rust_realloc|__rustc::__rust_|__rdl_|finish_grow|_int_free_chunk|_int_free_merge|_int_free_create|__libc_malloc2|alloc_impl|deallocate<|allocate<)'),
 ('crc32', r'^(crc32_chunk|update_fast_16|update_slow|reduce128|crc32fast|crc32|fold_by_4|zlib_rs::crc32)|crc32'),
 ('sha256', r'sha2::|sha256|compress256|Sha256|_mm_sha256'),
 ('deflate', r'zlib_rs|^deflate|longest_match|insert_string|quick_insert|send_bits|send_code|emit_lit|emit_dist|compress_block|build_tree|pqdownheap|gen_bitlen|scan_tree|send_tree|gen_codes|fill_window|slide_hash|^tally_|hash_calc|update_hash|^is_match<|compare256|zng_tr|flush_block|init_block|construct_huffman|flate2|compress_once|compress_uninit|OwnedCompressor|ReusableDeflateState|fizzle_matches|insert_match|flush_pending|bi_flush|bi_windup|push_lit|push_dist|^compress$|sym_buf|begin_member|^initialize$'),
 ('budget accounting', r'litchi_core::budget|budget::Node|BudgetedOutput|BudgetedSink'),
 ('XML audit', r'xml_minifier::audit|verify_with_policy|verify_source|inspect_attributes|audit_published_xml|publication_accepts_preserved_xml|check_attribute_layout|check_for_duplicates'),
 ('namespace binding', r'BindingTracker|resolve_element|^resolve_event|resolve_prefix'),
 ('XML parse (quick-xml)', r'quick_xml|read_event_impl|read_until_close|^read_event$|read_with<|^feed$|^emit_start|^emit_end|^emit_text|^read_text|read_bang|decoded_and_normalized|memchr|find_avx2|find_raw|search_chunk|search_slice|into_owned|^local_name$'),
 ('std SipHash', r'c_rounds|d_rounds|Sip13|SipHasher|^hash_one|^make_hash|^write_str<|hashbrown|core::hash|reserve_rehash|^hash<'),
 ('char iteration / UTF-16 count', r'next_code_point|utf16_code_unit_len|encode_utf16|^count<|utf8_first_byte|len_utf8|encode_utf8'),
 ('UTF-8 validation', r'from_utf8|run_utf8_validation'),
 ('text escaping', r'escape|escaped|is_plain_text_character|is_xml_character|append_character'),
 ('number/text formatting', r'pad_integral|format_inner|fmt::|writer_payload_text|semantic_docx_text|float_to|int_to|itoa|ryu|push_str'),
 ('docx document scan (model build)', r'^process_event<|^scan_document_with_context|^from_shared_xml'),
 ('docx text extraction', r'for_each_word_text_chunk|extract_word_text|word_special_character|append_xml_text_chunks|append_utf8_str_chunks|push_scanned|is_fragment_word_name|decompose'),
 ('vec growth / extend', r'^reserve<|needs_to_grow|^extend_from_slice|append_elements|spec_extend|do_reserve_and_handle|grow_amortized|grow_one|try_reserve'),
 ('OPC/ZIP part-name policy', r'packuri|PackURI|part_name_policy|zip_name_policy|validate_percent|validate_part_name|preservation_member_name'),
]
RC = [(n, re.compile(r)) for n, r in RULES]
def classify(fr):
    k = 0
    while k < len(fr) and fr[-1-k] == 'K': k += 1
    user = fr[:len(fr)-k]
    for f in reversed(user):
        s = nm(f)
        for n, rx in RC:
            if rx.search(s):
                return ('page faults/kernel <- ' + n) if k else n, s
    return ('page faults/kernel <- other' if k else 'other'), (short(user[-1]) if user else '?')
def run(case, top=16, other_top=8):
    tot, runner, rows = load(case)
    T = sum(p for p, n, f in rows)
    c = Counter(); oth = Counter(); kern = 0
    for p, n, fr in rows:
        cat, s = classify(fr)
        if cat.startswith('page faults'):
            kern += p
        c[cat] += p
        if cat == 'other': oth[re.sub(r'<.*', '<..>', s)] += p
    print(f'== {case}: timed share of process cycles {100*T/tot:.1f}%; kernel-leaf {100*kern/T:.1f}%')
    for k, v in c.most_common(top): print(f'  {100*v/T:6.2f}%  {k}')
    print('  other top:', ', '.join(f'{k[:40]} {100*v/T:.1f}%' for k, v in oth.most_common(other_top)))
for case in sys.argv[1:]:
    run(case)
