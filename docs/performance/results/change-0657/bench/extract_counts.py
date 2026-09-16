#!/usr/bin/env python3
import os,re,subprocess
SCR="/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0657/bench"
CG=os.path.join(SCR,"cg2")
SELS=["xlsx_producer_medium_source_planning","xlsx_producer_medium_source_one_edit_save"]
LEGS=["before","after"]; M=40
MOD="litchi_xlsx::cell_values::validation::"
# The raw worksheet parser re-enters the module through the observer callback,
# and that call already sits inside worksheet_xml_and_parse_source's inclusive
# cost, so its context must not be added to the module total.
CALLBACK="litchi_xlsx::raw::worksheet::codec::parse_source_with_observer"
row_re=re.compile(r'^\s*([\d,]+)\s*\([^)]*\)\s+(\S+?):(.+?)\s+\[')
def totals(tag):
    with open(os.path.join(CG,tag+".out"),'rb') as f:
        d=f.read().decode('utf-8','replace')
    return int(re.findall(r'^totals:\s*(\d+)',d,re.M)[-1])
def ann(tag):
    out=subprocess.run(["callgrind_annotate","--inclusive=yes","--threshold=100",
                        os.path.join(CG,tag+".out")],capture_output=True,text=True).stdout
    rows=[]
    for line in out.splitlines():
        m=row_re.match(line)
        if m: rows.append((int(m.group(1).replace(",","")),m.group(3).strip()))
    return rows
def split_ctx(n):
    i=n.rfind("'")
    return (n[:i],n[i+1:]) if i>=0 else (n,None)
def toplevel(rows):
    """Module contexts entered from outside, excluding the nested observer callback."""
    r=[]
    for ir,name in rows:
        fn,cal=split_ctx(name)
        if fn.startswith(MOD) and cal is not None and not cal.startswith(MOD) and cal!=CALLBACK:
            r.append((ir,fn,cal))
    return r
def callback_ctx(rows):
    return [(ir,split_ctx(n)[0]) for ir,n in rows
            if split_ctx(n)[0].startswith(MOD) and split_ctx(n)[1]==CALLBACK]
def named(rows,fn):
    return sum(ir for ir,n in rows if split_ctx(n)[0]==MOD+fn)
o=[]
A=o.append
A("=== change 0657 - deterministic instruction counts (valgrind callgrind, Ir) ===")
A("")
A("Method: whole-process Ir at --samples 2 and --samples 42 with --warmup 0;")
A("        Ir/op = (T42 - T2) / 40. Process start, corpus construction and all")
A("        one-time setup cancel exactly in the difference. One 'op' is one")
A("        harness sample iteration, which also carries the harness's own")
A("        per-sample verification outside its timer, so Ir/op is an upper")
A("        bound on the timed operation. taskset -c 12; callgrind with")
A("        --separate-callers=1 --cache-sim=no --branch-sim=no.")
A("")
cache={}
for sel in SELS:
    A("="*76); A(f"SELECTOR: {sel}"); A("="*76)
    A(f"{'leg':<7} {'Ir @ N=2':>18} {'Ir @ N=42':>18} {'Ir / op':>16}")
    per={}
    for leg in LEGS:
        t2,t42=totals(f"{leg}__{sel}__N2"),totals(f"{leg}__{sel}__N42")
        per[leg]=(t42-t2)/M
        A(f"{leg:<7} {t2:>18,} {t42:>18,} {per[leg]:>16,.0f}")
    A(f"  --> after/before Ir per operation = {per['after']/per['before']:.5f}"
      f"   ({(per['after']-per['before'])/per['before']*100:+.3f}%)")
    A("")
    A("  Inclusive Ir of litchi_xlsx::cell_values::validation (the changed module).")
    A("  Sum of every context entered from outside the module, EXCLUDING the")
    A("  Validator::observe callback invoked by raw::worksheet::codec (that call")
    A("  is already inside worksheet_xml_and_parse_source's inclusive cost).")
    A(f"  {'leg':<7} {'incl @ N=2':>16} {'incl @ N=42':>16} {'incl / op':>14}")
    mper={}
    for leg in LEGS:
        r2=ann(f"{leg}__{sel}__N2"); r42=ann(f"{leg}__{sel}__N42")
        cache[(sel,leg)]=r42
        s2=sum(x[0] for x in toplevel(r2)); s42=sum(x[0] for x in toplevel(r42))
        mper[leg]=(s42-s2)/M
        A(f"  {leg:<7} {s2:>16,} {s42:>16,} {mper[leg]:>14,.0f}")
    A(f"  --> after/before validation inclusive Ir per operation = {mper['after']/mper['before']:.5f}"
      f"   ({(mper['after']-mper['before'])/mper['before']*100:+.2f}%)")
    A("")
    A("  Module entry contexts at N=42 (these are the summed rows above):")
    for leg in LEGS:
        A(f"    [{leg}]")
        for ir,fn,cal in sorted(toplevel(cache[(sel,leg)]),reverse=True):
            A(f"      {ir:>14,}  {fn.replace(MOD,'validation::')}  <=  {cal}")
    A("")
    A("  Nested observer callback (already counted inside the entries above):")
    for leg in LEGS:
        for ir,fn in callback_ctx(cache[(sel,leg)]):
            A(f"    [{leg}] {ir:>14,}  {fn.replace(MOD,'validation::')}  <=  {CALLBACK}")
    A("")
    A("  Per-function inclusive Ir at N=42, summed over caller contexts.")
    A("  NOTE: inlining moved work between these names, so only the module")
    A("  total above is comparable leg to leg; these are a breakdown, not a claim.")
    A(f"  {'function':<40} {'before':>16} {'after':>16} {'ratio':>9}")
    for fn in ["worksheet_xml_and_parse_source","validate_xml","Validator::observe",
               "validate_element","validate_attributes","bind_dialect","text_allowed",
               "validate_close","allowed_unqualified_attribute"]:
        b=named(cache[(sel,"before")],fn); a=named(cache[(sel,"after")],fn)
        if b==0 and a==0: continue
        A(f"  {fn:<40} {b:>16,} {a:>16,} {(f'{a/b:.4f}' if b else 'n/a'):>9}")
    A("")
t="\n".join(o); print(t)
open(os.path.join(SCR,"counts2.txt"),"w").write(t+"\n")
