import json, re, collections, sys
W='/home/zhuhe/code/litchi-worktrees/0770-quick-xml-fail-fast-attribute-checks/'
sites = json.load(open('survey-sites.json'))
attr_sites = [s for s in sites if 'BytesStart::attributes' in s[4]]
texts = {}
def text(f):
    if f not in texts: texts[f] = open(W+f).read()
    return texts[f]
def offset(f, l, c):
    t = text(f); lines = t.split('\n')
    return sum(len(x)+1 for x in lines[:l-1]) + (c-1)
def skip_ws(t, i):
    while i < len(t) and t[i] in ' \t\r\n': i += 1
    return i
def match_brace(t, i):
    # t[i] == '{' ; return index of matching '}' (skip strings, chars, comments roughly)
    depth = 0; j = i
    while j < len(t):
        ch = t[j]
        if t.startswith('//', j):
            j = t.find('\n', j); continue
        if t.startswith('/*', j):
            j = t.find('*/', j) + 2; continue
        if ch == '"':
            # raw strings not handled; ordinary string
            j += 1
            while t[j] != '"':
                if t[j] == '\\': j += 1
                j += 1
            j += 1; continue
        if ch == "'" and j+2 < len(t) and (t[j+2] == "'" or (t[j+1]=='\\' and t[j+3]=="'")):
            j = t.find("'", j+2 if t[j+1] != '\\' else j+3) + 1; continue
        if ch == '{': depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0: return j
        j += 1
    return None
def statement_from(t, i):
    # from i to the next ';' or '{' at depth 0 of () [] (include the block for match/let-else)
    depth = 0; j = i
    while j < len(t):
        ch = t[j]
        if ch in '([': depth += 1
        elif ch in ')]': depth -= 1
        elif ch == ';' and depth == 0: return t[i:j+1]
        elif ch == '{' and depth == 0:
            k = match_brace(t, j)
            # include block and continue to ';' if let-else / match expression statement
            rest = t[k+1:k+3]
            seg = t[i:k+1]
            if re.match(r'\s*;', t[k+1:k+5]) or ' else' in t[i:j] :
                m = re.match(r'\s*;', t[k+1:])
                return t[i:k+1+(m.end() if m else 0)]
            return seg
        j += 1
    return t[i:]
out = []
for pkg, f, l, c, m in attr_sites:
    t = text(f)
    o = offset(f, l, c)
    lines = t.split('\n'); line = lines[l-1]
    head = line[:c-1]
    fm = re.search(r'\bfor\s+(\w+|\([^)]*\))\s+in\s+[\w\.\(\)&*]*$', head.strip())
    after = t[o:]
    mchain = re.match(r'attributes\(\)\s*(\.with_checks\((true|false)\))?', after)
    wc = mchain.group(2) or 'default'
    rec = {'pkg': pkg, 'file': f, 'line': l, 'col': c, 'checks': wc, 'src': line.strip()}
    if fm and fm.group(1).isidentifier():
        var = fm.group(1)
        brace = t.find('{', o)
        end = match_brace(t, brace)
        body = t[brace+1:end]
        um = re.search(r'\b' + var + r'\b', body)
        if not um:
            rec['kind'] = 'for-unused'
        else:
            # find statement start: go back to previous ';' or '{' or '}' in body
            s = max(body.rfind(';', 0, um.start()), body.rfind('{', 0, um.start()), body.rfind('}', 0, um.start())) + 1
            st = statement_from(body, s).strip()
            rec['kind'] = 'for'
            rec['var'] = var
            rec['stmt'] = re.sub(r'\s+', ' ', st)[:400]
    else:
        rec['kind'] = 'other'
        rec['ctx'] = '\n'.join(lines[l-3:l+4])
    out.append(rec)
json.dump(out, open('classified-raw.json', 'w'), indent=1)
print(len(out))
