"""Read-only real-input S1 GUI/API capture; output must be a fresh owned directory."""
import argparse, json, os, pathlib, socket, subprocess, time
from gui_logging_acceptance_20261112 import api, save, sha


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--exe',type=pathlib.Path,required=True)
    parser.add_argument('--manifest',type=pathlib.Path,required=True);parser.add_argument('--output',type=pathlib.Path,required=True)
    args=parser.parse_args();root=args.output.resolve();root.mkdir(parents=True,exist_ok=False)
    inputs=json.loads(args.manifest.read_text(encoding='utf-8-sig'))['s1']
    source=pathlib.Path(inputs['sourceDirectory']);templates=pathlib.Path(inputs['templateDirectory'])
    files=list(source.rglob('*.xlsx'))+list(templates.rglob('*.xlsx'))
    files += [pathlib.Path(inputs[key]) for key in ('mappingConfig','externalData') if inputs.get(key)]
    before={str(path.resolve()):sha(path) for path in files};save(root/'input-before.json',before)
    state=root/'state';state.mkdir()
    save(state/'settings.json',{'calc_engine':'auto','edit_engine':'auto','preferences':{'boCalcEngine':'betteroffice','conditionalFormatEvaluator':'PYTHON','s1ConditionalFormatSource':'audit_copy'}})
    with socket.socket() as sock: sock.bind(('127.0.0.1',0));port=sock.getsockname()[1]
    env=os.environ.copy();env['AUDIT_STATE_DIR']=str(state)
    stdout=(root/'server.stdout.log').open('w');stderr=(root/'server.stderr.log').open('w')
    server=subprocess.Popen([str(args.exe.resolve()),'serve','--port',str(port)],env=env,stdout=stdout,stderr=stderr,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    save(root/'server.json',{'pid':server.pid,'port':port,'exe':str(args.exe.resolve()),'exeSha256':sha(args.exe.resolve())})
    busy=False
    try:
        for _ in range(200):
            if server.poll() is not None: raise RuntimeError('owned server exited before ready')
            try: api(port,'get_state');break
            except Exception: time.sleep(.1)
        api(port,'update',[{'historyConfig':inputs.get('mappingConfig','')}])
        start_seq=api(port,'get_state').get('logVersion',0)
        values={'input':str(source),'templateDir':str(templates),'output':str(root/'output'),
                'mappingConfig':inputs.get('mappingConfig',''),'external':inputs.get('externalData','')}
        started=time.monotonic();api(port,'start_dag_flow',['dag:汇总核查表校验',values,[]]);busy=True
        changes=[];last=None;latencies=[]
        while True:
            polled=time.monotonic();snapshot=api(port,'get_state');latencies.append((time.monotonic()-polled)*1000)
            current=snapshot.get('currentWork')
            if current!=last:
                changes.append({'elapsedSeconds':time.monotonic()-started,'currentWork':current});last=current
                save(root/'progress.json',changes)
            if snapshot.get('busy') is not True: busy=False;break
            if time.monotonic()-started>3600: raise TimeoutError('live batch preserved; do not restart')
            time.sleep(.2)
        elapsed=time.monotonic()-started;save(root/'final-state.json',snapshot)
        logs=[row for row in snapshot.get('log',[]) if row.get('seq',0)>start_seq]
        save(root/'user-log.json',logs)
        (root/'user-log.txt').write_text('\n'.join(f"[{row.get('time','')}] [{row.get('level','')}] {row['text']}" for row in logs),encoding='utf-8')
        after={path:sha(pathlib.Path(path)) for path in before};save(root/'input-hashes.json',{'before':before,'after':after,'unchanged':before==after})
        journal=pathlib.Path(snapshot['diagnosticLogPath']);deadline=time.monotonic()+3
        while time.monotonic()<deadline:
            if journal.is_file() and '最终摘要' in journal.read_text(encoding='utf-8'): break
            time.sleep(.05)
        for part in journal.parent.glob(journal.stem+'*.jsonl'): (root/part.name).write_bytes(part.read_bytes())
        latencies.sort();summary={'runId':snapshot.get('runId'),'statistics':snapshot.get('runStatistics'),
            'elapsedSeconds':elapsed,'inputUnchanged':before==after,'userLogCount':len(logs),'diagnosticIncomplete':snapshot.get('diagnosticIncomplete'),
            'polls':len(latencies),'pollP95Ms':latencies[int((len(latencies)-1)*.95)],'pollMaxMs':max(latencies),
            'status':snapshot.get('status'),'preflightStatus':snapshot.get('s1F1Preflight',{}).get('status')}
        save(root/'capture-summary.json',summary);print(json.dumps(summary,ensure_ascii=True),flush=True)
        if before!=after or snapshot.get('diagnosticIncomplete') or not snapshot.get('runStatistics'): raise SystemExit(1)
    finally:
        if not busy: server.terminate();server.wait(timeout=15)
        stdout.close();stderr.close()

if __name__=='__main__': main()
