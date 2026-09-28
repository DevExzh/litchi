"""Replay exact pre-dispatch assembly windows without inferring execution cost."""
import json
import re
import sys
import custody as c

def analyze():
    records=[]
    for leg in ('before','after'):
        receipt=c.read(c.P/f'codegen-{leg}/receipt.json')
        assert receipt['schema']=='litchi.performance.0813.codegen.v1' and receipt['leg']==leg
        assert [r['variant'] for r in receipt['rows']]==['native','profile']
        build=c.read(c.P/f'build-{leg}/build.json')
        for row in receipt['rows']:
            assert row['binary']==build['binaries'][row['variant']] and row['source']==build['source']
            for artifact in (row['nm']['symbols'],row['objdump']['assembly']):
                assert c.artifact(artifact['path'])==artifact
            assert row['nm']['exit_code']==row['objdump']['exit_code']==0
            assert row['nm']['matched_count']==1
            symbol=open(row['nm']['symbols']['path']).read().split()
            assert len(symbol)==4 and int(symbol[0],16)==row['address'] and int(symbol[1],16)==row['size']
            instructions=[]
            for line in open(row['objdump']['assembly']['path']):
                match=re.fullmatch(r'\s*([0-9a-f]+):\s+(.+)\n?',line)
                if match:
                    offset=int(match[1],16)-row['address']
                    assert 0<=offset<row['size']
                    instructions.append({'offset':offset,'instruction':match[2]})
            calls=[i for i,v in enumerate(instructions) if 'call' in v['instruction'] and
                   '<quick_xml::reader::Reader<R>::read_event_impl>' in v['instruction']]
            assert len(calls)==1,calls
            start=calls[0]
            window=[];terminated=False
            for item in instructions[start:start+81]:
                window.append(item)
                if re.match(r'jmp\s+\*',item['instruction']):
                    terminated=True;break
            records.append({'leg':leg,'variant':row['variant'],'binary':row['binary'],
                'symbol_bytes':row['size'],'instructions':len(instructions),
                'reader_to_indirect_dispatch':window,'indirect_dispatch_found':terminated,
                'vector_moves_in_window':sum(bool(re.match(r'(?:v?movups|v?movaps|v?movdqu|v?movdqa)\s',v['instruction'])) for v in window),
                'receipt':c.artifact(c.P/f'codegen-{leg}/receipt.json')})
    return {'schema':'litchi.performance.0813.codegen-analysis.v1','rows':records,
        'scope':'Static exact-binary instruction windows. Counts are not runtime frequency, causal savings, or adoption evidence.'}

result=analyze();encoded=json.dumps(result,indent=2,sort_keys=True)+'\n';out=c.P/'codegen-analysis.json'
if '--check' in sys.argv:assert out.read_text()==encoded
else:
    assert sys.argv[1:]==['--write'] and not out.exists();out.write_text(encoded)
print('0813 code-generation window replay PASS')
