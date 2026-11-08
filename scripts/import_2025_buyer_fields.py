"""Map organization and individual names to their separate invoice fields."""
import hashlib,json,re,unicodedata

def is_organization(name):
    folded=unicodedata.normalize('NFKD',str(name or '')).encode('ascii','ignore').decode().upper().strip()
    return bool(re.match(r'^(CONG TY\b|CTY\b|CHI NHANH\b|DOANH NGHIEP\b|HOP TAC XA\b|HO KINH DOANH\b|TRUONG\b|BENH VIEN\b|UY BAN\b)',folded))

def apply_buyer_fields(payload,out,root):
    policy=out/'buyer-field-confirmation.json'
    assert policy.exists()
    stats={}
    # Adjustment identity follows the original invoice's separate H/L fields.
    targets=[o for o in payload['outputs'] if o['kind'] not in ['report','adjustment']]
    targets += [s for o in payload['outputs'] if o['kind']=='report' for s in o['sheets'] if s['name']=='Tra lai cho so HD']
    for table in targets:
        h=table['headers'];ci=2;pi=h.index('Người mua hàng');changed=0
        assert h[ci] in ['Tên đơn vị mua hàng','Tên khách hàng']
        for row in table['rows']:
            if row[ci] and not is_organization(row[ci]):
                row[pi]=row[pi] or row[ci]
                row[ci]=None;changed+=1
        stats[table.get('name') or table['kind']]=changed
    path=str(policy.relative_to(root));payload['manifest']=[m for m in payload['manifest'] if m['file']!=path]
    payload['manifest'].append({'file':path,'sha256':hashlib.sha256(policy.read_bytes()).hexdigest()})
    return stats
