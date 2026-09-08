import pandas as pd, numpy as np, math, json, shutil, re
from pathlib import Path
from collections import defaultdict, Counter
from copy import copy
from openpyxl import load_workbook

DAYS=[7]
REPORT_BUILD_DAYS=[7]
ALL_DAYS=range(1,8)
BASE=Path(r'D:\HY-日度数据监测\HY-日志导出月度汇总\9月视频审核操作日志')
XLSX=Path(r'D:\虎牙数据汇总\虎牙结算数据\26年结算\9月结算\9月人审量.xlsx')
ROOT=Path(r'C:\Users\guojing17\.easyclaw\workspace\虎牙数据看板')
HTML=ROOT/'index.html'
BACKUP=XLSX.with_name('9月人审量_追加9.7前备份.xlsx')
HTML_BACKUP=ROOT/'index_追加9.6前备份.html'
COLS=['vid','所属应用','视频时长','上传时间','工单生成时间','进入审核时间','审核完成时间','审核结果','审核人']
VALID_APPS=['虎牙视频','虎牙直播回放切片视频','虎牙AI礼物特效视频','虎牙一起看视频']


def dursec(s):
    try:
        p=[float(x) for x in str(s).split(':')]
        return p[0]*3600+p[1]*60+p[2] if len(p)==3 else np.nan
    except Exception:
        return np.nan


def norm_type(s):
    return s.fillna('').astype(str).str.replace('—','-',regex=False).str.replace('–','-',regex=False).str.strip()


def enrich(d):
    d=d.copy()
    d['duration_sec']=d['视频时长'].map(dursec)
    for c in ['工单生成时间','进入审核时间','审核完成时间']:
        d[c]=pd.to_datetime(d[c],errors='coerce')
    # 排队时长口径：进入审核时间 - 工单生成时间，剔除重新推审及>60分钟异常值。
    d['queue_min']=(d['进入审核时间']-d['工单生成时间']).dt.total_seconds()/60
    d['queue_excluded']=norm_type(d['视频类型']).eq('爬取视频-重新推审')
    d['audit_sec']=(d['审核完成时间']-d['进入审核时间']).dt.total_seconds()
    d['mult']=d['audit_sec']/d['duration_sec']
    d.loc[(d['duration_sec']<=0)|(d['audit_sec']<0)|(d['audit_sec']>3600),'mult']=np.nan
    d['weight']=np.where(d['duration_sec']>3600,1.5,0.78)
    d['bad']=d['审核结果'].isin(['不通过','已删除'])
    return d


def eligible(d):
    # “虎牙低质视频”不计算任何量级，仅在日报通道分布中作为异常通道反馈。
    return d.loc[~d['所属应用'].fillna('').astype(str).str.strip().eq('虎牙低质视频')].copy()


def metrics(d):
    d=eligible(d)
    q=d.loc[(~d.queue_excluded)&d.queue_min.between(0,60),'queue_min']
    actual=len(d)
    return {
        'actual':actual,
        'trans':math.floor(d.weight.sum()),
        'passed':int((d['审核结果']=='通过').sum()),
        'risk':int((d['审核结果']=='通过（风险）').sum()),
        'bad':int(d.bad.sum()),
        'viol':round(d.bad.sum()/actual*100,2) if actual else 0,
        'queue':round(q.mean(),2) if len(q) else np.nan,
        'queue_n':int(q.count()),
        'queue_sum':float(q.sum()),
        'mult':round(d.mult.mean(),2) if d.mult.notna().any() else np.nan
    }


def load_day(day):
    p=BASE/f'视频审核操作日志-9.{day}.csv'
    assert p.exists(),p
    raw=pd.read_csv(p,encoding='utf-8',low_memory=False,dtype={'vid':str})
    miss=[c for c in COLS+['视频类型','上传端'] if c not in raw.columns]
    assert not miss,f'{p.name} 缺少列: {miss}'
    # 全量纳入“易盾BPO”，只在 eligible() 中排除所属应用“虎牙低质视频”。
    b=raw[raw['审核人'].fillna('').astype(str).str.contains('易盾BPO',regex=False)].copy()
    assert len(b)>0,f'9.{day} 无易盾BPO数据'
    b['vid']=b['vid'].fillna('').astype(str).str.replace('\t','',regex=False).str.strip()
    assert b['vid'].ne('').all(),f'9.{day} 存在空VID'
    return raw,enrich(b)


def make_report(day,b,m):
    m=m.copy()
    e=eligible(b)
    m['people']=int(e['审核人'].nunique())
    m['percap']=math.floor(m['trans']/m['people'])
    persons=[]
    for reviewer,g in e.groupby('审核人'):
        x=metrics(g)
        persons.append({'name':reviewer.replace('-易盾BPO',''),'audit':x['actual'],'trans':x['trans'],'passed':x['passed'],'bad':x['bad'],'viol':x['viol'],'queue':x['queue'],'mult':x['mult']})
    persons.sort(key=lambda x:(-x['audit'],x['name']))
    apps={str(k):int(v) for k,v in b['所属应用'].value_counts().to_dict().items()}
    other={k:v for k,v in apps.items() if k not in VALID_APPS}
    high=[p['name'] for p in persons if p['viol']>15]
    cap='🟢' if m['percap']>=230 else '🔴'
    vs='🟢' if not high else '🟡'
    qs='🟢' if m['queue']<=2 else '🔴'
    iso=f'2026-09-{day:02d}'
    L=[f'📊 【虎牙视频审核日报】{iso}','','📈 整体业务情况','指标│数值',
       f"实际审核总量│{m['actual']} 条",f"转化后审核总量│{m['trans']} 条",f"通过量│{m['passed']} 条",f"通过（风险）量│{m['risk']} 条",f"不通过量│{m['bad']} 条",f"违规率│{m['viol']:.2f}%",f"平均排队时长│{m['queue']:.2f} 分钟",f"平均审核倍率│{m['mult']:.2f}",f"当班次人数│{m['people']} 人",f"人均转化后审核量│{m['percap']} 条/人",'',
       '👥 当班次人员情况','姓名│审核量│转化后│通过量│不通过量│违规率│排队时长│审核倍率']
    for p in persons:
        qv=f"{p['queue']:.2f}" if pd.notna(p['queue']) else '无有效数据'
        mv=f"{p['mult']:.2f}" if pd.notna(p['mult']) else '无有效数据'
        L.append(f"{p['name']}│{p['audit']}│{p['trans']}│{p['passed']}│{p['bad']}│{p['viol']:.2f}%│{qv}│{mv}")
    L += ['', '⚠️ 异常预警','预警类型│状态│详情',
          f"产能异常│{cap}│人均转化后审核量 {m['percap']} 条/人 {'≥' if m['percap']>=230 else '<'} 目标230",
          f"违规率异常│{vs}│"+('无人员超过15%' if not high else '、'.join(high)+' > 正常范围15%'),
          f"排队时长异常│{qs}│平均排队时长 {m['queue']:.2f} 分钟 {'≤' if m['queue']<=2 else '>'} 目标2分钟",'',
          '💡 建议措施',
          ('- 产能提升：当前人效达标，继续保持班次节奏。' if m['percap']>=230 else '- 产能提升：当日人效未达230条/人，建议结合低峰时段优化排班和任务分配。'),
          ('- 违规率控制：整体及个人违规率均在正常范围内。' if not high else '- 违规率控制：重点复核'+'、'.join(high)+'的高违规样本，确认口径与样本结构。'),
          f"- 审核倍率关注：当日平均审核倍率 {m['mult']:.2f}，持续关注高倍率人员及长视频审核耗时。",'',
          '📊 通道分布','通道名称│数量│状态']
    for a in VALID_APPS:
        L.append(f"{a}│{apps.get(a,0)}│"+('✅' if apps.get(a,0) else 'ℹ️ 当日无数据'))
    L.append(f"其他异常通道│{sum(other.values())}│"+('⚠️ '+'；'.join(f'{k}{v}条' for k,v in other.items()) if other else '✅ 无异常通道'))
    report=ROOT/f'虎牙视频审核日报-{iso}.txt'
    report.write_text('\n'.join(L),encoding='utf-8')
    summary={'day':f'9.{day}','metrics':m,'persons':persons,'apps':apps,'other_apps':other,'queue_excluded_records':int(b.queue_excluded.sum()),'report':str(report)}
    (ROOT/f'analysis_9.{day}.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    return summary,'\n'.join(L)


# 一次读取9.1-9.7源数据，生成9.7日报，看板从源CSV全量重建。
loaded={}
for day in ALL_DAYS:
    loaded[day]=load_day(day)
reports={}; report_texts={}
for day in REPORT_BUILD_DAYS:
    raw,b=loaded[day]
    reports[day],report_texts[day]=make_report(day,b,metrics(b))

# 月度结算表追加，既有9.1-9.3数据同步核验，9.7不重复追加。
old=pd.read_excel(XLSX,dtype={'vid':str,'数据日期':str})
assert list(old.columns[:11])==COLS+['数据日期','视频类型'],old.columns.tolist()
if not BACKUP.exists(): shutil.copy2(XLSX,BACKUP)
wb=load_workbook(XLSX); ws=wb[wb.sheetnames[0]]; last=ws.max_row
ws.cell(1,17).value='上传端'
if ws.cell(1,11).has_style: ws.cell(1,17)._style=copy(ws.cell(1,11)._style)
existing=old['数据日期'].astype(str).str.replace(r'\.0$','',regex=True)

# 回填已存在日期的Q列，按原追加顺序逐行核对VID，防止错位。
for day in ALL_DAYS:
    ds=f'9.{day}'
    positions=[i for i,v in enumerate(existing.tolist()) if v==ds]
    if not positions: continue
    src=eligible(loaded[day][1]).reset_index(drop=True)
    old_vid=old.iloc[positions]['vid'].fillna('').astype(str).str.replace('\t','',regex=False).str.strip().tolist()
    src_vid=src['vid'].fillna('').astype(str).str.strip().tolist()
    assert old_vid==src_vid,f'{ds} 既有结算数据顺序与源CSV不一致，禁止直接回填上传端'
    for pos,upload in zip(positions,src['上传端'].tolist()):
        cell=ws.cell(pos+2,17); cell.value=None if pd.isna(upload) else upload
        if ws.cell(pos+2,11).has_style: cell._style=copy(ws.cell(pos+2,11)._style)

added={}
for day in DAYS:
    b=eligible(loaded[day][1]); ds=f'9.{day}'
    if (existing==ds).any():
        added[day]=0
        continue
    for row in b[COLS+['视频类型','上传端']].itertuples(index=False,name=None):
        last+=1
        vals=list(row[:9])+[ds,row[9]]
        for ci,val in enumerate(vals,1):
            cell=ws.cell(last,ci); cell.value=None if pd.isna(val) else val
            src=ws.cell(last-1,ci)
            if src.has_style: cell._style=copy(src._style)
            if ci in (5,6,7): cell.number_format='yyyy-mm-dd hh:mm:ss'
        qcell=ws.cell(last,17); qcell.value=None if pd.isna(row[10]) else row[10]
        if ws.cell(last,11).has_style: qcell._style=copy(ws.cell(last,11)._style)
    added[day]=len(b)
wb.save(XLSX)

# 结算表强制回读核验：行数、日期、VID、视频类型、Q列上传端及重复追加。
post=pd.read_excel(XLSX,dtype={'vid':str,'数据日期':str})
assert post.columns[16]=='上传端',post.columns.tolist()
pdts=post['数据日期'].astype(str).str.replace(r'\.0$','',regex=True)
for day in ALL_DAYS:
    b=eligible(loaded[day][1]).reset_index(drop=True); chk=post.loc[pdts==f'9.{day}'].reset_index(drop=True)
    assert len(chk)==len(b),(day,len(chk),len(b))
    assert chk['vid'].fillna('').astype(str).str.strip().ne('').all(),f'9.{day} 结算表VID为空'
    assert chk['视频类型'].fillna('').astype(str).str.strip().ne('').all(),f'9.{day} 结算表视频类型为空'
    assert chk['vid'].fillna('').astype(str).str.strip().tolist()==b['vid'].tolist(),f'9.{day} VID顺序不一致'
    assert chk['上传端'].fillna('').astype(str).tolist()==b['上传端'].fillna('').astype(str).tolist(),f'9.{day} Q列上传端不一致'
    assert not b['所属应用'].fillna('').astype(str).str.strip().eq('虎牙低质视频').any()
assert len(post)-len(old)==sum(added.values()),(len(old),len(post),added)

# 从9.1-9.7源CSV重建看板全部当月指标。
labels=[]; audit=[]; trans=[]; queues=[]; viol=[]; zv=[]; zq=[]
people_total=0; bad_total=0; qsum=0.; qcount=0.; pw=defaultdict(float); pdays=defaultdict(set); hourly=Counter(); apps=Counter()
for day in ALL_DAYS:
    raw,b=loaded[day]; e=eligible(b); mm=metrics(e)
    labels.append(f'9.{day}'); audit.append(mm['actual']); trans.append(mm['trans']); queues.append(mm['queue']); viol.append(mm['viol'])
    people_total+=e['审核人'].nunique(); bad_total+=mm['bad']; qsum+=mm['queue_sum']; qcount+=mm['queue_n']
    hourly.update(e['审核完成时间'].dt.hour.dropna().astype(int).tolist()); apps.update(b['所属应用'].fillna('空值').tolist())
    for name,g in e.groupby('审核人'):
        n=name.replace('-易盾BPO',''); pw[n]+=g.weight.sum(); pdays[n].add(day)
    zr=raw[raw['审核人'].fillna('').astype(str).str.contains('证通BPO')].copy()
    z=enrich(zr); zm=metrics(z)
    zv.append(zm['actual']); zq.append(zm['queue'])
pr=sorted(((n,round(pw[n]/len(pdays[n]),1)) for n in pw),key=lambda x:(-x[1],x[0]))[:12]
cum_actual=sum(audit); cum_trans=sum(trans); eff=math.floor(cum_trans/people_total)
cum_viol=round(bad_total/cum_actual*100,2); cum_queue=round(qsum/qcount,2); hr=[hourly[i] for i in range(24)]
if not HTML_BACKUP.exists(): shutil.copy2(HTML,HTML_BACKUP)
text=HTML.read_text(encoding='utf-8')

def sub(p,r,count=1):
    global text
    text,n=re.subn(p,r,text,count=count,flags=re.S)
    assert n==count,(p,n,count)

def arr(x): return json.dumps(x,ensure_ascii=False,separators=(',',':'))
text=text.replace('易盾BPO 核心效能监控（8月）','易盾BPO 核心效能监控（9月）')
sub(r'(数据周期: )[^<]+',r'\g<1>2026-09-01 至 2026-09-07')
sub(r'(报告生成: )[^<]+',r'\g<1>2026-09-08 10:00:00')
sub(r'(<div class="label">累计实际审核量</div>\s*<div class="value blue">)[^<]+',rf'\g<1>{cum_actual:,} ')
sub(r'(<div class="label">累计转化后审核量</div>\s*<div class="value green">)[^<]+',rf'\g<1>{cum_trans:,} ')
sub(r'(<div class="label">日均人效\(转化后\)</div>\s*<div class="value purple">)[^<]+',rf'\g<1>{eff} ')
sub(r'(<div class="label">平均违规率</div>\s*<div class="value orange">)[^<]+',rf'\g<1>{cum_viol:.2f} ')
sub(r'(<div class="label">平均排队时长</div>\s*<div class="value red">)[^<]+',rf'\g<1>{cum_queue:.2f} ')
text=re.sub(r'9月累计\d+天','9月累计7天',text)
text=re.sub(r'[89]月在班总人力\d+人',f'9月在班总人力{people_total}人',text)
for var,val in [('dates',labels),('auditVolumes',audit),('transformedVolumes',trans),('queueTimes',queues),('personNames',[x[0] for x in pr]),('personValues',[x[1] for x in pr]),('hourlyData',hr),('julyLabels',labels),('yidunVolumes',audit),('zhengtongVolumes',zv),('yidunQueueTimes',queues),('zhengtongQueueTimes',zq)]:
    sub(rf'const {var} = \[[^;]*\];',f'const {var} = {arr(val)};')
text=re.sub(r'9月人员排行\(9\.1-9\.\d+日均转化\)','9月人员排行(9.1-9.7日均转化)',text)
text=re.sub(r'9月24小时时段数据（9\.1-9\.\d+累计）','9月24小时时段数据（9.1-9.7累计）',text)
sub(r'(new Chart\(document\.getElementById\(\'violationRateChart\'\).*?labels: )\[[^\]]*\]',rf'\g<1>{arr(labels)}')
sub(r'(new Chart\(document\.getElementById\(\'violationRateChart\'\).*?label: \'每日违规率\',\s*data: )\[[^\]]*\]',rf'\g<1>{arr(viol)}')
# 应用图按源数据累计，含异常通道，避免历史硬编码。
app_order=[a for a in VALID_APPS if apps.get(a,0)>0]+[a for a in apps if a not in VALID_APPS]
app_vals=[int(apps[a]) for a in app_order]
sub(r'(new Chart\(document\.getElementById\(\'julyChannelChart\'\).*?labels: )\[[^\]]*\]',rf'\g<1>{arr(app_order)}')
sub(r'(new Chart\(document\.getElementById\(\'julyChannelChart\'\).*?datasets: \[\{\s*data: )\[[^\]]*\]',rf'\g<1>{arr(app_vals)}')
HTML.write_text(text,encoding='utf-8')

dash={'through':'9.7','cum_actual':cum_actual,'cum_trans':cum_trans,'people_total':int(people_total),'eff':eff,'cum_viol':cum_viol,'cum_queue':cum_queue,'daily':{d:{'actual':a,'trans':t,'queue':q,'viol':v} for d,a,t,q,v in zip(labels,audit,trans,queues,viol)},'person_rank':pr,'apps':dict(apps),'hourly':hr,'zt_volumes':zv,'zt_queues':zq}
(ROOT/'dashboard_9.7_summary.json').write_text(json.dumps(dash,ensure_ascii=False,indent=2),encoding='utf-8')
(ROOT/'daily_reports_9.7.txt').write_text('\n\n==============================\n\n'.join(report_texts[d] for d in DAYS),encoding='utf-8')

# 看板内部硬编码和数组长度自查。
check=HTML.read_text(encoding='utf-8')
assert '2026-09-07' in check and '9月累计7天' in check and '9.3' in check
for x in [labels,audit,trans,queues,zv,zq]: assert len(x)==7
assert sum(audit)==cum_actual and sum(trans)==cum_trans and sum(hr)==cum_actual
# 日报整体数据独立核验：总量、结果分布、人员合计必须一致。
for day in REPORT_BUILD_DAYS:
    b=eligible(loaded[day][1]); s=reports[day]; m=s['metrics']
    assert m['actual']==len(b)==sum(p['audit'] for p in s['persons'])
    assert m['passed']==int((b['审核结果']=='通过').sum())
    assert m['risk']==int((b['审核结果']=='通过（风险）').sum())
    assert m['bad']==int(b['审核结果'].isin(['不通过','已删除']).sum())

result={'reports':{d:reports[d] for d in DAYS},'workbook_before':len(old),'workbook_after':len(post),'added':added,'dashboard':dash,'selfcheck':{'report':True,'monthly_workbook':True,'dashboard':True}}
print(json.dumps(result,ensure_ascii=False,indent=2))

