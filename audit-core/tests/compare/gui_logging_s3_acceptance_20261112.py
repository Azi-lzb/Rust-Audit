"""Real S3-F03/F05 GUI API acceptance with disposable business workbooks."""
import argparse, json, os, pathlib, socket, subprocess, time
import openpyxl
from gui_logging_acceptance_20261112 import api, save, sha


def book(path, sheets):
    workbook=openpyxl.Workbook();workbook.remove(workbook.active)
    for name,rows in sheets.items():
        sheet=workbook.create_sheet(name)
        for row in rows: sheet.append(row)
    workbook.save(path)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--exe',type=pathlib.Path,required=True)
    parser.add_argument('--output',type=pathlib.Path,required=True);args=parser.parse_args()
    root=args.output.resolve();root.mkdir(parents=True,exist_ok=False)
    state=root/'state';state.mkdir();target=root/'说明文件.xlsx';feedback=root/'机构反馈.xlsx'
    headers=['机构代码','机构名称','指标代码','指标名称','数据值','增减额(亿元)','变动原因']
    book(target,{'说明文件':[headers,['0001','测试机构','I001','测试指标',100,20,'']]})
    book(feedback,{'说明文件':[headers,['0001','测试机构','I001','测试指标',100,20,'核实后调整']]})
    common=root/'通用配置.xlsx';forms=root/'拆分配置.xlsx';comparison=root/'比较.xlsx'
    (root/'forms-output').mkdir()
    book(common,{'环比警戒':[['下限','上限','备注文字','填充颜色','是否整行填充','金额变动阈值','是否说明'],['10','30','','6']]})
    book(forms,{'报表清单':[['报表代码','报表名称','频度','批次','是否启用写入'],['测试A1','测试表','月','1','是']],
                '转表设置':[['参数','值'],['隐藏空行','是'],['空表删除','是']],
                '测试A1':[['指标代码','指标名称','行序号','余额人民币','上期-余额人民币','增减额-余额人民币','环比-余额人民币'],['I001']]})
    compare_headers=['业务类','数据日期','机构类代码','机构类名称','地区代码','地区名称','指标顺序码','指标代码','指标名称','数据属性','币种','频度','批次','数据值','上期值','增减额(亿元)','环比','备注','是否说明','说明内容','计算过程']
    book(comparison,{'2026年8月对比结果':[compare_headers,['人民币','2026-08-31','0001','机构','1100000','地区','I001','I001','指标','余额','人民币','月','1',120,100]]})
    originals={str(path):sha(path) for path in [target,feedback,common,forms,comparison]}
    with socket.socket() as sock: sock.bind(('127.0.0.1',0));port=sock.getsockname()[1]
    env=os.environ.copy();env['AUDIT_STATE_DIR']=str(state)
    stdout=(root/'server.stdout.log').open('w');stderr=(root/'server.stderr.log').open('w')
    server=subprocess.Popen([str(args.exe.resolve()),'serve','--port',str(port)],env=env,stdout=stdout,stderr=stderr,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    save(root/'server.json',{'pid':server.pid,'port':port,'exeSha256':sha(args.exe.resolve())})
    busy=False;results=[]
    try:
        for _ in range(200):
            try: api(port,'get_state');break
            except Exception: time.sleep(.1)
        for feature,method,call_args,paths in [
            ('S3-F03','import_central_explanation_feedback',[[str(feedback)]],{'centralExplainFile':target}),
            ('S3-F05','start_central_indicator_split',[],{'centralCompareFile':comparison,'centralCommonConfig':common,'centralFormsConfig':forms,'centralOutput':root/'forms-output'})]:
            directory=root/feature;directory.mkdir()
            for key,value in paths.items(): api(port,'choose_central',[key,str(value)])
            start_seq=api(port,'get_state').get('logVersion',0);started=time.monotonic()
            api(port,method,call_args);busy=True;progress=[];last=None
            while True:
                snapshot=api(port,'get_state');current=snapshot.get('currentWork')
                if current!=last: progress.append(current);last=current
                if not snapshot.get('busy'): busy=False;break
                if time.monotonic()-started>300: raise TimeoutError('live worker preserved')
                time.sleep(.02)
            save(directory/'final-state.json',snapshot);save(directory/'progress.json',progress)
            logs=[row for row in snapshot.get('log',[]) if row.get('seq',0)>start_seq]
            save(directory/'user-log.json',logs)
            (directory/'user-log.txt').write_text('\n'.join(row['text'] for row in logs),encoding='utf-8')
            time.sleep(.2);journal=pathlib.Path(snapshot['diagnosticLogPath'])
            (directory/'developer.jsonl').write_bytes(journal.read_bytes())
            stats=snapshot.get('runStatistics') or {};errors=[]
            if feature=='S3-F03':
                if stats.get('importedRows')!=1: errors.append('expected one imported feedback row')
                if openpyxl.load_workbook(target,data_only=True)['说明文件']['G2'].value!='核实后调整': errors.append('target feedback value missing')
                backups=list(root.glob('说明文件_导入前备份_*.xlsx'))
                if len(backups)!=1 or sha(backups[0])!=originals[str(target)]: errors.append('original target backup missing or changed')
            elif stats.get('outputFiles')!=1: errors.append('expected one generated financial workbook')
            result={'feature':feature,'statistics':stats,'runId':snapshot.get('runId'),'validationErrors':errors}
            results.append(result);save(root/'results.json',results);print(json.dumps(result,ensure_ascii=True),flush=True)
        unchanged={path:sha(pathlibpath)==digest for path,digest in originals.items() if (pathlibpath:=pathlib.Path(path))!=target}
        save(root/'input-safety.json',{'unchanged':unchanged,'targetIntentionallyUpdated':True,'targetOriginalSha256':originals[str(target)]})
        if not all(unchanged.values()) or any(row['validationErrors'] for row in results): raise SystemExit(1)
    finally:
        if not busy: server.terminate();server.wait(timeout=15)
        stdout.close();stderr.close()

if __name__=='__main__': main()
