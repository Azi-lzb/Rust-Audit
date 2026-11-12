"""Compare real before/after GUI result values and links without modifying files."""
import argparse, json, pathlib, urllib.parse
import openpyxl


def snapshot(path):
    workbook=openpyxl.load_workbook(path,data_only=False)
    result=[]
    for sheet in workbook:
        rows=[];links=[]
        for row in sheet:
            rows.append([cell.value for cell in row])
            for cell in row:
                if cell.hyperlink:
                    target=(cell.hyperlink.target or '').replace('\\','/')
                    # Only the run directory changes. Preserve filename, cell and fragment.
                    target=urllib.parse.unquote(target).rsplit('/',1)[-1]
                    links.append([cell.coordinate,target,cell.hyperlink.location])
        result.append({'sheet':sheet.title,'rows':rows,'links':links,'merges':list(map(str,sheet.merged_cells.ranges))})
    return result


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--before',type=pathlib.Path,required=True)
    parser.add_argument('--after',type=pathlib.Path,required=True);parser.add_argument('--report',type=pathlib.Path,required=True)
    args=parser.parse_args();before=list(args.before.rglob('汇总核查表校验.xlsx'));after=list(args.after.rglob('汇总核查表校验.xlsx'))
    if len(before)!=1 or len(after)!=1: raise RuntimeError('expected exactly one result in each isolated run')
    left,right=snapshot(before[0]),snapshot(after[0]);differences=[]
    if len(left)!=len(right): differences.append('sheet count')
    for index,(old,new) in enumerate(zip(left,right)):
        for key in ['sheet','rows','links','merges']:
            if old[key]!=new[key]: differences.append(f'sheet[{index}] {key}')
    report={'before':str(before[0]),'after':str(after[0]),'sheetCount':len(left),'resultRows':[len(s['rows'])-1 for s in left],
            'compared':['all result cell values/formulas','sheet names/order','hyperlinks except run directory','merged regions'],'differences':differences,'equal':not differences}
    args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=True))
    if differences: raise SystemExit(1)

if __name__=='__main__':main()
