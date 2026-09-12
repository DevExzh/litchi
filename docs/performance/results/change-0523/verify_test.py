"""Reject in-memory semantic corruptions without changing retained evidence."""
import copy
import json
from analyze import read, validate_row
from run import HERE, FOLDER


def main():
    cfb=read(FOLDER/'native-r1-cfb.json')['results'][0]
    xls=next(r for r in read(FOLDER/'native-r1-xls.json')['results'] if r['case']=='xls_source_backed_open')
    allocation=read(FOLDER/'alloc-r1-cfb.json')['results'][0]
    for row, count, enabled in [(cfb,1000,False),(xls,1000,False),(allocation,30,True)]:
        validate_row(row,count,enabled)
    short=copy.deepcopy(cfb);short['elapsed_ns']['samples'].pop()
    alignment=copy.deepcopy(cfb);alignment['operation_metrics']['sample_indices'][0]=alignment['operation_metrics']['sample_indices'][1]
    counter=copy.deepcopy(xls);counter['source']['read_calls'][0]+=1
    balance=copy.deepcopy(allocation);balance['operation_metrics']['allocation']['allocated_bytes']['values'][0]+=1
    checks=[]
    for name,row,count,enabled in [('short_native_vector',short,1000,False),
        ('operation_sample_alignment',alignment,1000,False),('source_read_contract',counter,1000,False),
        ('allocation_live_byte_balance',balance,30,True)]:
        try:
            validate_row(row,count,enabled)
        except (AssertionError,ValueError) as error:
            checks.append(dict(case=name,rejected=True,reason=str(error) or 'allocation/source invariant assertion'))
        else:
            raise AssertionError('corruption accepted: '+name)
    result=dict(status='pass',valid_controls_pass=True,checks=checks,
        scope='Four in-memory raw semantic mutations; no retained artifact modified')
    (HERE/'verifier-tests.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps(result))


if __name__=='__main__':
    main()
