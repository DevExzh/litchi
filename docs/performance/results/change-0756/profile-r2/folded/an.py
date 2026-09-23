#!/usr/bin/env python3
"""an.py CASE [mode]: timed-region attribution over folded stacks (period-weighted).
modes: phases, leaf, incl, comp, all"""
import re, sys
from collections import Counter, defaultdict
import gzip, os
D = os.path.dirname(os.path.abspath(__file__)) + '/'
GLUE = re.compile(r'^(branch<|map_err<|and_then<|map<|into<|from$|from<|unwrap_or_else<|ok_or_else<|call_once|\{closure|try_fold|try_for_each|for_each<|fold<|catch_unwind|do_call|__rust_begin)')
def nm(f): return f.split('|R:')[0]
def short(f): return re.sub(r'<.*', '<..>', nm(f))
# case -> (file, runner regex, allowed first-child regex)
CASES = {
 'full_text': ('docx_semantic_full_text', r'^run_semantic_docx$', r'^(extract_word_text|text|read_event_impl|for_each_word_text_chunk|litchi_ooxml_common::binding_tracker|push_scanned|append_|is_fragment_word_name|pop$|result$)'),
 'noop_save': ('docx_semantic_noop_edit_save', r'^run_semantic_docx$', r'^(edit_document|publish_document_edit|write_plain|replace_paragraph_text|to_stream|reserve_operation)'),
 'one_save': ('docx_semantic_one_edit_save', r'^run_semantic_docx$', r'^(edit_document|publish_document_edit|write_plain|replace_paragraph_text|to_stream|reserve_operation|format_inner|semantic_docx_text)'),
 'doc_fresh': ('doc_fresh_write_to', r'^run_fresh_writer', r'^write_fresh_doc'),
 'ppt_fresh': ('ppt_fresh_write_to', r'^run_fresh_writer', r'^write_fresh_ppt'),
 'xls_fresh': ('xls_fresh_write_to', r'^run_fresh_writer', r'^write_fresh_xls'),
}
def load(case):
    fn, rrx, crx = CASES[case]
    rrx = re.compile(rrx); crx = re.compile(crx)
    tot = 0; out = []; runner = 0
    for line in (open(D + fn + '.folded') if os.path.exists(D + fn + '.folded') else gzip.open(D + fn + '.folded.gz', 'rt')):
        p, n, comm, st = line.rstrip('\n').split('\t'); p = int(p)
        if not comm.startswith('litchi-perf'): continue
        tot += p
        fr = st.split(' ;; ')
        idx = None
        for i, f in enumerate(fr):
            if rrx.search(nm(f)): idx = i; break
        if idx is None: continue
        runner += p
        j = idx + 1
        while j < len(fr) and GLUE.search(nm(fr[j])): j += 1
        if j >= len(fr) or not crx.search(nm(fr[j])): continue
        out.append((p, int(n), fr[j:]))
    return tot, runner, out
def phases(rows, depth, top=30, skip_first=0):
    c = Counter(); T = sum(p for p, n, f in rows)
    for p, n, fr in rows:
        ch = []
        for f in fr[skip_first:]:
            if GLUE.search(nm(f)): continue
            ch.append(short(f)[:55] + ('*' if '|R:' in f else ''))
            if len(ch) >= depth: break
        c[' > '.join(ch)] += p
    for k, v in c.most_common(top): print(f'{100*v/T:6.2f}%  {k}')
def leaf(rows, top=40, ctx=2):
    c = Counter(); T = sum(p for p, n, f in rows)
    for p, n, fr in rows:
        # leaf-most user frame(s)
        k = 0
        while k < len(fr) and fr[-1-k] == 'K': k += 1
        user = [short(x)[:50] for x in fr[:len(fr)-k]]
        key = ('[K] ' if k else '') + ' < '.join(reversed(user[-ctx:]))
        c[key] += p
    for key, v in c.most_common(top): print(f'{100*v/T:6.2f}%  {key}')
def incl(rows, top=60):
    c = Counter(); T = sum(p for p, n, f in rows)
    for p, n, fr in rows:
        seen = set()
        for f in fr:
            s = short(f)
            if s in seen or GLUE.search(nm(f)): continue
            seen.add(s); c[s] += p
    for key, v in c.most_common(top): print(f'{100*v/T:6.2f}%  {key[:100]}')
if __name__ == '__main__':
    case = sys.argv[1]; mode = sys.argv[2] if len(sys.argv) > 2 else 'all'
    tot, runner, rows = load(case)
    T = sum(p for p, n, f in rows); ns = sum(n for p, n, f in rows)
    kern = sum(p for p, n, f in rows if f[-1] == 'K')
    print(f'== {case}: timed={100*T/tot:.1f}% of process cycles, {100*T/runner:.1f}% of runner; samples={ns}; kernel-leaf share of timed={100*kern/T:.1f}%')
    if mode in ('phases', 'all'):
        d = int(sys.argv[3]) if len(sys.argv) > 3 else 2
        print('-- phases'); phases(rows, d)
    if mode in ('leaf', 'all'):
        print('-- leaf'); leaf(rows)
    if mode in ('incl', 'all'):
        print('-- inclusive'); incl(rows)
