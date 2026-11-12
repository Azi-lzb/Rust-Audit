"""Local real GUI/API acceptance: drives production workers, never fabricated log strings."""
import argparse, ctypes, hashlib, json, os, pathlib, re, socket, subprocess, time, urllib.request, zipfile

def save(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')

def api(port, method, args=()):
    request=urllib.request.Request(f'http://127.0.0.1:{port}/api/{method}',data=json.dumps(list(args)).encode(),headers={'Content-Type':'application/json'})
    with urllib.request.urlopen(request,timeout=15) as response: return json.load(response)['result']

def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def generate_fixtures(root):
    # Developer test dependency only; the production program never imports this.
    import openpyxl
    from openpyxl.workbook.defined_name import DefinedName
    root.mkdir(parents=True,exist_ok=False)
    good='=IF(TRUE,"通过","错|数值核对|合成指标|测试|1|1|0")'
    issue='=IF(TRUE,"错|数值核对|合成指标|测试异常|2|1|1","通过")'
    fallback='=IF(INFO("system")<>"","错|数值核对|合成指标|测试回退|2|1|1","通过")'
    def book(path,formula,sheet,label,rows):
        workbook=openpyxl.Workbook();ws=workbook.active;ws.title=sheet
        ws['A1']=formula;ws['B1']=label
        for row in range(2,rows+2): ws.cell(row,3,row-1)
        for name,cell in [('校验区域','$A$1'),('条件格式区域','$D$1'),('表结构区域','$B$1')]:
            workbook.defined_names.add(DefinedName(name,attr_text=f"'{sheet}'!{cell}"))
        workbook.save(path)
        # Match the UTF-8 literal encoding of production Excel fixtures. The existing
        # baseline OOXML importer mishandles numeric character entities in formulas;
        # this logging acceptance must not change that separate Excel implementation.
        with zipfile.ZipFile(path) as archive:
            parts=[(info,archive.read(info.filename)) for info in archive.infolist()]
        with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as archive:
            for info,data in parts:
                if info.filename.endswith('.xml'):
                    data=re.sub(r'&#(\d+);',lambda match:chr(int(match[1])) if int(match[1])>127 else match[0],data.decode('utf-8')).encode('utf-8')
                archive.writestr(info,data)
    cases=[]
    for name in ['passed','issues','skipped','missing_sheet','missing_field','calculation_failed','engine_fallback','partial_failure','same_name_failure','output_failure']:
        directory=root/name;input_dir=directory/'input';templates=directory/'templates'
        input_dir.mkdir(parents=True);templates.mkdir()
        formula=good if name=='passed' else fallback if name in ('calculation_failed','engine_fallback') else issue
        template=templates/'template.xlsx';book(template,formula,'检查表','目标字段',0)
        count=1 if name in ('missing_sheet','missing_field') else 2
        for index in range(1,count+1):
            sheet,label=('其他表','目标字段') if name=='missing_sheet' else ('检查表','不匹配字段') if name=='skipped' and index==2 else ('检查表','目标字段')
            if name=='missing_field': label=None
            source=input_dir/f'机构{index:02}_2026-07-31_核查.xlsx'
            if name=='same_name_failure':
                folder=input_dir/f'机构{index:02}';folder.mkdir();source=folder/'同名_2026-07-31_核查.xlsx'
            book(source,good,sheet,label,5000 if index==1 else 0)
        cases.append({'name':name,'input':str(input_dir),'templates':str(templates),'template':str(template),'files':count})
    save(root/'cases.json',cases);save(root/'generator.json',{'tool':'openpyxl','version':openpyxl.__version__,'purpose':'local synthetic acceptance inputs only'})

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--exe',type=pathlib.Path,required=True)
    parser.add_argument('--fixtures',type=pathlib.Path,required=True);parser.add_argument('--output',type=pathlib.Path,required=True)
    parser.add_argument('--cases',nargs='*',help='Run only the named acceptance cases')
    arguments=parser.parse_args(); arguments.fixtures=arguments.fixtures.resolve()
    if not arguments.fixtures.exists(): generate_fixtures(arguments.fixtures)
    root=arguments.output.resolve();root.mkdir(parents=True,exist_ok=False)
    state_dir=root/'state';state_dir.mkdir()
    save(state_dir/'settings.json',{'calc_engine':'auto','edit_engine':'auto','preferences':{'boCalcEngine':'betteroffice','conditionalFormatEvaluator':'PYTHON','s1ConditionalFormatSource':'audit_copy'}})
    with socket.socket() as sock: sock.bind(('127.0.0.1',0));port=sock.getsockname()[1]
    env=os.environ.copy();env['AUDIT_STATE_DIR']=str(state_dir)
    stdout=(root/'server.stdout.log').open('w',encoding='utf-8');stderr=(root/'server.stderr.log').open('w',encoding='utf-8')
    server=subprocess.Popen([str(arguments.exe.resolve()),'serve','--port',str(port)],env=env,stdout=stdout,stderr=stderr,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    save(root/'server.json',{'pid':server.pid,'port':port,'exe':str(arguments.exe.resolve()),'exeSha256':sha(arguments.exe.resolve())})
    handle=None;results=[];busy=False
    try:
        for _ in range(200):
            if server.poll() is not None: raise RuntimeError('test server exited before ready')
            try: api(port,'get_state');break
            except Exception: time.sleep(.1)
        else: raise RuntimeError('test server not ready')
        cases=json.loads((arguments.fixtures/'cases.json').read_text(encoding='utf-8'))
        if arguments.cases: cases=[case for case in cases if case['name'] in arguments.cases]
        for case in cases:
            name=case['name'];directory=root/name;directory.mkdir()
            before={str(path):sha(path) for path in pathlib.Path(case['input']).rglob('*.xlsx')}
            before[str(pathlib.Path(case['template']))]=sha(pathlib.Path(case['template']))
            settings={'boCalcEngine':'betteroffice','boCalcFallback':'none' if name=='calculation_failed' else 'auto'}
            # These existing BO preference keys are read by the worker, but the current GUI update API rejects them (501).
            # Configure the owned test state directly between idle runs; do not change the product API in a logging task.
            config={'calc_engine':'auto','edit_engine':'auto','preferences':{'conditionalFormatEvaluator':'PYTHON','s1ConditionalFormatSource':'audit_copy',**settings}}
            save(state_dir/'settings.json',config);save(directory/'test-settings.json',config)
            api(port,'update',[{'template':case['template'],'historyConfig':''}])
            initial=api(port,'get_state');start_seq=initial.get('logVersion',0)
            values={'input':case['input'],'templateDir':case['templates'],'template':case['template'],'output':str(directory/'output'),'mappingConfig':'','external':''}
            started=time.monotonic();api(port,'start_dag_flow',['dag:汇总核查表校验',values,[]]);busy=True;injected=False;work=[];last=None
            while True:
                snapshot=api(port,'get_state');current=snapshot.get('currentWork')
                if current!=last: work.append({'elapsed':time.monotonic()-started,'currentWork':current});last=current
                if snapshot.get('busy') is not True: busy=False;break
                if not injected and name in ('partial_failure','same_name_failure') and current and '复制' in str(current.get('stage','')):
                    target=sorted(pathlib.Path(case['input']).rglob('*.xlsx'))[-1].resolve()
                    assert target.is_relative_to(arguments.fixtures.resolve())
                    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
                    kernel.CreateFileW.argtypes=[ctypes.c_wchar_p,ctypes.c_uint32,ctypes.c_uint32,ctypes.c_void_p,ctypes.c_uint32,ctypes.c_uint32,ctypes.c_void_p]
                    kernel.CreateFileW.restype=ctypes.c_void_p
                    handle=kernel.CreateFileW(str(target),0x80000000,0,None,3,0x80,None)
                    if handle==ctypes.c_void_p(-1).value: raise ctypes.WinError(ctypes.get_last_error())
                    injected=True
                if not injected and name=='output_failure':
                    folders=list((directory/'output').glob('S1-F1_审核结果_*'))
                    if folders:
                        (folders[0]/'汇总核查表校验.xlsx').mkdir();injected=True
                if time.monotonic()-started>900: raise TimeoutError(f'{name}: worker still running; server preserved')
                time.sleep(.01)
            if handle:
                ctypes.windll.kernel32.CloseHandle.argtypes=[ctypes.c_void_p];ctypes.windll.kernel32.CloseHandle(handle);handle=None
            save(directory/'final-state.json',snapshot);save(directory/'progress.json',work)
            user_logs=[entry for entry in snapshot.get('log',[]) if entry.get('seq',0)>start_seq]
            save(directory/'user-log.json',user_logs)
            (directory/'user-log.txt').write_text('\n'.join(f"[{entry.get('time','')}] [{entry.get('level','info')}] {entry.get('text','')}" for entry in user_logs),encoding='utf-8')
            after={path:sha(pathlib.Path(path)) for path in before};save(directory/'input-hashes.json',{'before':before,'after':after,'unchanged':before==after})
            # The same async GUI writer is allowed a bounded drain before copying the evidence.
            log=pathlib.Path(snapshot.get('diagnosticLogPath',''));deadline=time.monotonic()+3
            while time.monotonic()<deadline:
                if log.is_file() and '最终摘要' in log.read_text(encoding='utf-8'): break
                time.sleep(.05)
            journal=log.read_text(encoding='utf-8') if log.is_file() else ''
            (directory/'developer.jsonl').write_text(journal,encoding='utf-8')
            events=[json.loads(line) for line in journal.splitlines() if line]
            result={'case':name,'runId':snapshot.get('runId'),'status':snapshot.get('s1F1Preflight',{}).get('status'),'statistics':snapshot.get('runStatistics'),'inputUnchanged':before==after,'userLogCount':len(user_logs),'injected':injected,'elapsedSeconds':time.monotonic()-started,'diagnosticIncomplete':snapshot.get('diagnosticIncomplete',False),'fallbackObserved':any(event.get('stage')=='betteroffice_refused' for event in events)}
            errors=[];stats=result['statistics'] or {};preflight=snapshot.get('s1F1Preflight',{})
            if not result['runId'] or not result['status']: errors.append('missing real run state')
            if not result['inputUnchanged']: errors.append('inputs changed')
            if result['diagnosticIncomplete']: errors.append('diagnostic journal incomplete')
            if not events: errors.append('diagnostic journal missing')
            if events and any(event.get('runId')!=result['runId'] for event in events): errors.append('diagnostic run identity mismatch')
            if not any('最终摘要' in event.get('message','') for event in events): errors.append('final summary not persisted')
            if any(entry.get('detail') for entry in user_logs): errors.append('technical detail leaked into GUI buffer')
            if name in ('passed','issues','engine_fallback'):
                if len(user_logs)!=2: errors.append('normal task should have start and final summary only')
                if any(re.search(r'OOXML|openpyxl|Umya|BetterOffice|betteroffice|COM|worker|IPC|缓存',entry.get('text','')) for entry in user_logs): errors.append('normal technical engine information leaked into GUI log')
            if name=='passed' and (preflight.get('totalIssues')!=0 or preflight.get('reviewedFiles')!=2): errors.append('expected two completed files and zero audit issues')
            if name=='issues' and (preflight.get('totalIssues')!=2 or preflight.get('reviewedFiles')!=2): errors.append('expected two actual formula issues')
            if name=='skipped' and (preflight.get('skipped')!=1 or preflight.get('reviewedFiles')!=1): errors.append('expected one skipped and one completed')
            if name=='missing_sheet' and (preflight.get('skipped')!=1 or result['status']!='no_structure_passed'): errors.append('expected missing-sheet input skipped')
            if name=='missing_field' and (preflight.get('skipped')!=1 or result['status']!='no_structure_passed'): errors.append('expected missing required structure field skipped')
            if name=='calculation_failed' and (not result['fallbackObserved'] or not (preflight.get('waitingManualRecalc',0)>0 or preflight.get('failedFiles',0)>0)): errors.append('actual calculation refusal/failure not observed; earlier copy failures do not qualify')
            if name=='engine_fallback' and (not result['fallbackObserved'] or result['status']!='completed'): errors.append('successful internal engine fallback not observed')
            if name in ('partial_failure','same_name_failure') and (not injected or preflight.get('failedFiles')!=1 or preflight.get('reviewedFiles')!=1): errors.append('one failed and one completed file not observed')
            if name in ('partial_failure','same_name_failure') and not any('文件无法读取' in entry.get('text','') for entry in user_logs if entry.get('level')=='error'): errors.append('file read failure was not accurately described')
            if name=='same_name_failure' and not any('机构02' in entry.get('text','') and '同名' in entry.get('text','') for entry in user_logs if entry.get('level')=='error'): errors.append('failure did not distinguish the same-named file by directory')
            if name=='output_failure' and (not injected or preflight.get('resultPath') is not None or not stats.get('resultError')): errors.append('actual output publication failure not observed; earlier copy failures do not qualify')
            if not stats: errors.append('structured run statistics missing')
            if name in ('calculation_failed','output_failure') and '审核失败' not in snapshot.get('status',''): errors.append('main GUI status must not call an unpublished failed task partially completed')
            result['validationErrors']=errors
            results.append(result);save(root/'results.json',results);print(json.dumps(result,ensure_ascii=True),flush=True)
    finally:
        if handle:
            ctypes.windll.kernel32.CloseHandle.argtypes=[ctypes.c_void_p];ctypes.windll.kernel32.CloseHandle(handle)
        if not busy:
            server.terminate();server.wait(timeout=15)
        stdout.close();stderr.close()
    if any(result['validationErrors'] for result in results): raise SystemExit(1)
if __name__=='__main__': main()
