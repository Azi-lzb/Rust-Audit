import csv,json,sys,hashlib,collections
from pathlib import Path
import openpyxl
root=Path(sys.argv[1]); template_source=Path(sys.argv[2]); external_file=Path(sys.argv[3]); config_rust=Path(sys.argv[4]); config_python=Path(sys.argv[5])
report=json.loads((root/'rust-with-external'/'acceptance-report.json').read_text(encoding='utf-8'))
inner=json.loads(report['steps'][0]['stdout']); pre=inner['preflight']
def clean_path(s): return str(s).replace('\\\\?\\','')
def base(s): return clean_path(s).replace('/','\\').split('\\')[-1]
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest().upper()
def read_rows(p):
    wb=openpyxl.load_workbook(p,read_only=True,data_only=True); ws=wb.worksheets[0]; out=[]
    for row in ws.iter_rows(min_row=2,values_only=True):
        if any(v is not None for v in row): out.append(tuple('' if v is None else (v.isoformat() if hasattr(v,'isoformat') else str(v).strip() if isinstance(v,str) else v) for v in row))
    wb.close(); return out
source_rows=[]
with (root/'source-manifest.csv').open(encoding='utf-8-sig',newline='') as f:
    for row in list(csv.reader(f))[1:]: source_rows.append({'business':row[0],'file':row[1],'path':row[3],'sha256_before':row[5]})
by_file={x['file']:x for x in source_rows}; categories=sorted({x['business'] for x in source_rows})
python_data=json.loads((root/'python-results.json').read_text(encoding='utf-8'))
python_rows=[]; python_batch={}
for b in python_data:
    f=b['result_fields']; sp=f.get('summary_path'); rows=read_rows(Path(sp)) if sp else []
    python_rows.extend(rows)
    python_batch[b['name']]={'status':b['status'],'elapsed_ms':b['elapsed_ms'],'files':len(f.get('files') or []),'issues':len(rows),'issues_reported':len(f.get('current_issues') or []),'output':sp or f.get('log_path')}
rust_wb=Path(clean_path(pre['resultPath'])); rust_rows=read_rows(rust_wb)
rust_result_by_file={base(x.get('sourceFile','')):x for x in pre.get('fileResults',[])}
def type_is_cf(x): return len(x)>3 and x[3]=='\u6761\u4ef6\u683c\u5f0f\u89e6\u53d1'
def type_is_formula(x): return len(x)>3 and x[3]=='\u9519\u8bef'
def rows_per_file(rows):
    d=collections.defaultdict(list)
    for r in rows:
        if r: d[str(r[0])].append(r)
    return d
py_file_rows=rows_per_file(python_rows); rust_file_rows=rows_per_file(rust_rows)
py_cat=collections.defaultdict(list); rust_cat=collections.defaultdict(list)
for row in python_rows:
    cat=by_file.get(str(row[0]),{}).get('business','UNKNOWN'); py_cat[cat].append(row)
for row in rust_rows:
    cat=by_file.get(str(row[0]),{}).get('business','UNKNOWN'); rust_cat[cat].append(row)
per_file=[]
for source in source_rows:
    name=source['file']; pf=py_file_rows.get(name,[]); rf=rust_file_rows.get(name,[]); status=rust_result_by_file.get(name,{})
    per_file.append({'business':source['business'],'source':name,'rust_status':status.get('status','MISSING'),'python_issues':len(pf),'rust_issues':len(rf),'python_formula':sum(type_is_formula(x) for x in pf),'rust_formula':sum(type_is_formula(x) for x in rf),'python_conditional':sum(type_is_cf(x) for x in pf),'rust_conditional':sum(type_is_cf(x) for x in rf),'counts_equal':len(pf)==len(rf),'reason':status.get('reason','')})
per_category=[]
for cat in categories:
    files=[x for x in per_file if x['business']==cat]; pr=py_cat.get(cat,[]); rr=rust_cat.get(cat,[])
    status_counts=dict(collections.Counter(x['rust_status'] for x in files))
    pc=collections.Counter(pr); rc=collections.Counter(rr); keys_py=collections.Counter(tuple(x[:4]) for x in pr); keys_rust=collections.Counter(tuple(x[:4]) for x in rr)
    per_category.append({'business':cat,'files':len(files),'rust_statuses':status_counts,'python_elapsed_ms':python_batch.get(cat,{}).get('elapsed_ms'),'python_formula':sum(type_is_formula(x) for x in pr),'rust_formula':sum(type_is_formula(x) for x in rr),'python_conditional':sum(type_is_cf(x) for x in pr),'rust_conditional':sum(type_is_cf(x) for x in rr),'python_issues':len(pr),'rust_issues':len(rr),'exact_rows':sum((pc&rc).values()),'matching_keys_file_sheet_cell_type':sum((keys_py&keys_rust).values()),'file_issue_counts_equal':sum(x['counts_equal'] for x in files),'files_count':len(files)})
# Categorize failure points by exact runtime reason evidence.
failed=[x for x in per_file if x['rust_status']!='COMPLETED']
external_link_failures=[x for x in failed if 'unverified workbook parts' in x['reason']]
formula_verify_failures=[x for x in failed if '公式位置或公式文本与来源不一致' in x['reason']]
other_failures=[x for x in failed if x not in external_link_failures and x not in formula_verify_failures]
source_hashes_ok=all(sha(x['path'])==x['sha256_before'] for x in source_rows)
template_copies=[]
for copy in sorted((root/'business_templates').glob('*.xlsx')):
    original=template_source/copy.name
    template_copies.append({'template':copy.name,'original_exists':original.is_file(),'sha256_equal':original.is_file() and sha(original)==sha(copy)})
old_config=Path(r'C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\config-real\1.逐筆統計系统_配置.xlsx')
# The path above is ASCII source with a Unicode filename; search by the known project config names if encoding changed.
if not old_config.exists(): old_config=next(Path(r'C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\config-real').glob('1.*'))
original_config=Path(r'C:\Users\AZI\Desktop\VibeCode\Rust-Audit\audit-core\config-real')
if not config_rust.exists(): raise FileNotFoundError(config_rust)
result={'source_count':len(source_rows),'category_count':len(categories),'templates_matched':len(json.loads((root/'template-mapping.json').read_text(encoding='utf-8'))),'python_batches':len(python_data),'python_completed_batches':sum(x['status']=='completed' for x in python_data),'python_files':sum(x['files'] for x in python_batch.values()),'python_elapsed_ms':sum(x['elapsed_ms'] for x in python_data),'python_issues':len(python_rows),'python_formula_issues':sum(type_is_formula(x) for x in python_rows),'python_conditional_issues':sum(type_is_cf(x) for x in python_rows),'rust_status':report.get('status'),'rust_elapsed_ms':report.get('elapsed_ms'),'rust_files_total':pre.get('filesTotal'),'rust_structure_passed':pre.get('structurePassed'),'rust_reviewed_files':pre.get('reviewedFiles'),'rust_completed_files':sum(v.get('status')=='COMPLETED' for v in rust_result_by_file.values()),'rust_manual_recalc':pre.get('waitingManualRecalc'),'rust_failed_files':pre.get('failedFiles'),'rust_formula_issues':sum(type_is_formula(x) for x in rust_rows),'rust_conditional_issues':sum(type_is_cf(x) for x in rust_rows),'rust_issues':len(rust_rows),'exact_row_matches':sum((collections.Counter(python_rows)&collections.Counter(rust_rows)).values()),'matching_issue_keys':sum((collections.Counter(tuple(x[:4]) for x in python_rows)&collections.Counter(tuple(x[:4]) for x in rust_rows)).values()),'count_equal_files':sum(x['counts_equal'] for x in per_file),'files_count':len(per_file),'completed_files_count_equal':sum(x['counts_equal'] and x['rust_status']=='COMPLETED' for x in per_file),'completed_files_count_total':sum(x['rust_status']=='COMPLETED' for x in per_file),'python_only_rows':sum((collections.Counter(python_rows)-collections.Counter(rust_rows)).values()),'rust_only_rows':sum((collections.Counter(rust_rows)-collections.Counter(python_rows)).values()),'external_link_failure_files':[x['source'] for x in external_link_failures],'formula_verification_failure_files':[x['source'] for x in formula_verify_failures],'other_failure_files':[x['source'] for x in other_failures],'source_sha256_unchanged':source_hashes_ok,'template_copies':template_copies,'external_sha256_after':sha(external_file),'external_sha256_expected':'5DF655ABC0074D9A70896794F478E82A33E0992AAA70C6ADD96FAA3232908F6E','rust_acceptance_input_unchanged':report['steps'][0].get('input_files_unchanged'),'python_engine_mix':{'excel':sum('Microsoft Excel' in (x.get('summary') or '') for x in python_data),'wps':sum('WPS' in (x.get('summary') or '') for x in python_data)},'rust_recalculation_engines':pre.get('recalculationEngines'),'per_category':per_category,'per_file':per_file,'python_batches_detail':python_batch,'rust_result_workbook':str(rust_wb)}
(root/'final-comparison.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
with (root/'per-file-parity-with-external.csv').open('w',encoding='utf-8-sig',newline='') as f:
    w=csv.DictWriter(f,fieldnames=list(per_file[0].keys()));w.writeheader();w.writerows(per_file)
with (root/'business-parity-with-external.csv').open('w',encoding='utf-8-sig',newline='') as f:
    w=csv.DictWriter(f,fieldnames=list(per_category[0].keys()));w.writeheader();w.writerows(per_category)
print('RUST',result['rust_status'],result['rust_elapsed_ms'],'ms; files',result['rust_completed_files'],'/',result['rust_files_total'],'failed',result['rust_failed_files'],'formula',result['rust_formula_issues'],'conditional',result['rust_conditional_issues'])
print('PYTHON',result['python_completed_batches'],'batches;',result['python_files'],'/',result['source_count'],'files;',result['python_elapsed_ms'],'ms; formula',result['python_formula_issues'],'conditional',result['python_conditional_issues'])
print('PARITY exact',result['exact_row_matches'],'keys',result['matching_issue_keys'],'count-equal',result['count_equal_files'],'/',result['files_count'],'completed-equal',result['completed_files_count_equal'],'/',result['completed_files_count_total'])
print('PYTHON_ONLY',result['python_only_rows'],'RUST_ONLY',result['rust_only_rows'],'source_hash_ok',source_hashes_ok,'template_copies_ok',all(x['sha256_equal'] for x in template_copies),'external_hash',result['external_sha256_after'])
print('FAILURES external_links',len(external_link_failures),'formula_verify',len(formula_verify_failures),'other',len(other_failures))
print('ENGINES PY',result['python_engine_mix'],'RUST',pre.get('recalculationEngines'))