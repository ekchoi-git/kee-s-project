"""Build the new film timeline and extract genuine motion from take 2."""
from pathlib import Path
import json
import subprocess
import sys

import imageio_ffmpeg
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parent
FFMPEG=imageio_ffmpeg.get_ffmpeg_exe()
RAW=ROOT/"raw-take2.mkv"
CREATED=RAW.stat().st_birthtime
manifest=json.loads((ROOT/"capture.json").read_text(encoding="utf-8"))
events={event["label"]:event for event in manifest}


def run(arguments):
    subprocess.run([FFMPEG,"-hide_banner","-loglevel","error","-y",*arguments],check=True,timeout=60)


for label in ["t2-02-first-input","t2-09-totals-input","t2-13-format-input"]:
    event=events[label]
    clean=label+"-clean"
    time=event["time"]+4
    target=ROOT/"captures"/f"{clean}-desktop.png"
    run(["-ss",str(time-CREATED),"-i",str(RAW),"-frames:v","1","-update","1",str(target)])
    with Image.open(target) as source:
        for name,box in event["windows"].items():
            source.crop(box).save(ROOT/"captures"/f"{clean}-{name}.png")
    derived={**event,"label":clean,"time":time,"desktop":str(target),"source":"raw-take2.mkv","sourceSeconds":time-CREATED}
    manifest=[item for item in manifest if item["label"]!=clean]+[derived]

for name,label in [("first","t2-03-first-processing"),("totals","t2-10-totals-processing"),("format","t2-14-format-processing")]:
    folder=ROOT/"motion"/name
    folder.mkdir(parents=True,exist_ok=True)
    run(["-ss",str(events[label]["time"]-CREATED+.15),"-i",str(RAW),"-t","5.2","-vf","crop=515:270:1390:735,fps=12","-frames:v","60",str(folder/"%03d.png")])
    assert len(list(folder.glob("*.png")))==60

(ROOT/"capture.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding="utf-8")

wide_clips={
    "first_send":(1793685980.485-CREATED-3,4),
    "first_result":(80,6),
    "totals_send":(1793686355.646-CREATED-3,4),
    "totals_result":(415,6),
    "format_send":(1793686457.3-CREATED-3,4),
    "format_result":(515.2,6),
    "save":(556,6),
    "totals_wait":(events["t2-10-totals-processing"]["time"]-CREATED+.15,1),
    "format_wait":(events["t2-14-format-processing"]["time"]-CREATED+.15,1),
}
for name,(start,seconds) in wide_clips.items():
    folder=ROOT/"wide-motion"/name
    folder.mkdir(parents=True,exist_ok=True)
    metadata={"source":"raw-take2.mkv","startSeconds":start,"durationSeconds":seconds,"fps":12,"frames":seconds*12}
    info=folder/"clip.json"
    cached=info.exists() and json.loads(info.read_text(encoding="utf-8"))==metadata and len(list(folder.glob("*.png")))==seconds*12
    if not cached:
        run(["-ss",str(start),"-i",str(RAW),"-t",str(seconds+.2),"-vf","fps=12","-frames:v",str(seconds*12),"-compression_level","3",str(folder/"%03d.png")])
        assert len(list(folder.glob("*.png")))==seconds*12
        info.write_text(json.dumps(metadata,indent=2),encoding="utf-8")
roles={
    "connected":"t2-01-connected",
    "first_input":"t2-02-first-input-clean",
    "first_processing":"t2-03-first-processing",
    "first_complete":"t2-05-first-complete",
    "initial":"t2-08-same-sheet",
    "same_sheet":"t2-08-same-sheet",
    "totals_input":"t2-09-totals-input-clean",
    "totals_processing":"t2-10-totals-processing",
    "totals_complete":"t2-11-totals-complete",
    "totals_workspace":"t2-11-totals-complete",
    "totals":"t2-13-format-input-clean",
    "format_input":"t2-13-format-input-clean",
    "format_processing":"t2-14-format-processing",
    "format_complete":"t2-15-format-complete",
    "final":"t2-16-final-table",
    "saved":"t2-17-saved",
}
shots=[]


def card(kind,seconds,title,subtitle=""):
    shots.append(dict(kind=kind,seconds=seconds,title=title,subtitle=subtitle))


def panel(role,window,crop,box):
    return dict(role=role,window=window,crop=crop,box=box)


def screen(id,seconds,phase,caption,panels,**extra):
    shots.append(dict(id=id,kind="screen",seconds=seconds,phase=phase,caption=caption,panels=panels,**extra))


def desktop(id,seconds,phase,caption,role=None,live=None,**extra):
    shot=dict(id=id,kind="desktop",seconds=seconds,phase=phase,caption=caption,**extra)
    if live:
        shot["desktopLive"]=live
    else:
        shot["desktopRole"]=role
    shots.append(shot)


composer=[15,755,530,995]
input_box=[190,300,1540,540]
table=[9,208,801,536]
table_box=[170,225,1580,650]
formula=[17,105,440,151]

card("opening",7,"셀리","열린 Excel에 글로 요청하는 AI 도우미")
card("chapter",2,"1. 열린 Excel 연결")
desktop("workspace_start",5,"열린 Excel 연결","Excel 파일을 열어두고, 옆의 셀리에서 작업을 요청합니다.",role="connected")
screen("new_welcome",3,"열린 Excel 연결","원하는 작업을 한국어로 입력하는 AI 도우미, 셀리.",
       [panel("connected","celly",[25,200,530,485],[300,255,1320,550])],label="셀리 · 새로운 시작 화면")
screen("connection",2,"열린 Excel 연결","열린 파일과 작업할 시트를 확인합니다.",
       [panel("connected","celly",[15,790,530,866],[240,340,1440,245])],label="셀리 · 파일과 작업 시트")
card("chapter",2,"2. 회사별·연도별 집계")
desktop("first_context",2,"회사별·연도별 집계","왼쪽의 거래내역을 대상으로, 오른쪽 셀리에 요청을 입력합니다.",role="first_input")
screen("first_input",8,"회사별·연도별 집계","회사별·연도별 집계표를\n새 시트에 만들어 달라고 입력합니다.",
       [panel("first_input","celly",composer,input_box)],label="셀리 · 첫 번째 요청")
desktop("first_send",4,"회사별·연도별 집계","셀리의 보내기를 누르면, 열린 Excel 작업을 시작합니다.",live="first_send",click=True,clickAt=3.15)
screen("first_processing",2,"회사별·연도별 집계","셀리가 요청을 읽고 집계표를 만듭니다.",
       [panel("first_processing","celly",None,[190,285,1540,580])],live="first",note="처리 대기 일부 생략")
desktop("first_result_wide",6,"회사별·연도별 집계","완료 응답을 확인하고, Excel에서 만들어진 결과 시트를 엽니다.",live="first_result")
screen("first_reply",2,"회사별·연도별 집계","완료 응답에서 결과 시트를 확인합니다.",
       [panel("first_complete","celly",[25,200,520,545],[330,215,1260,650])],label="셀리 · 실제 작업 완료")
screen("initial",4,"회사별·연도별 집계","행은 회사명, 열은 연도. 집계표가 만들어졌습니다.",
       [panel("initial","excel",[9,208,650,515],[170,225,1580,650])],label="Excel · 회사별 연도별",unit="단위: 원")
card("chapter",2,"3. 첨부 없이 이어서 작업")
desktop("continuity",3,"첨부 없이 이어서 작업","같은 Excel 표와 같은 대화를 유지한 채, 합계를 추가합니다.",role="same_sheet")
screen("totals_input",4,"첨부 없이 이어서 작업","같은 대화에서 연도별·회사별 합계를 요청합니다.",
       [panel("totals_input","celly",composer,input_box)],label="셀리 · 두 번째 요청")
desktop("totals_send",4,"첨부 없이 이어서 작업","표를 그대로 열어둔 채, 셀리에 합계 추가를 요청합니다.",live="totals_send",click=True,clickAt=3.15)
screen("totals_processing",1,"첨부 없이 이어서 작업","같은 결과 시트에 합계를 추가합니다.",
       [panel("totals_processing","celly",None,[190,285,1540,580])],live="totals",note="처리 대기 일부 생략")
desktop("totals_result_wide",6,"연도별·회사별 합계","셀리가 작업하면, 열려 있는 같은 표에 합계 행과 열이 추가됩니다.",live="totals_result",note="처리 대기 일부 생략")
screen("annual_total",4,"연도별·회사별 합계","아래에는 연도별 합계.",
       [panel("totals","excel",table,table_box)],label="Excel · 연도별 합계 행",unit="단위: 원",highlights=[[238,817,1730,875]])
screen("company_total",4,"연도별·회사별 합계","오른쪽에는 회사별 합계와 전체 총합.",
       [panel("totals","excel",table,table_box)],label="Excel · 회사별 합계 열",unit="단위: 원",highlights=[[1458,267,1730,875]])
card("chapter",2,"4. 천원 단위로 표시")
desktop("format_context",3,"천원 단위로 표시","합계를 추가한 같은 표에서, 표시형식 변경까지 이어갑니다.",role="format_input")
screen("format_input",5,"천원 단위로 표시","같은 대화에서 천원 단위 표시도 요청합니다.",
       [panel("format_input","celly",composer,input_box)],label="셀리 · 세 번째 요청")
desktop("format_send",4,"천원 단위로 표시","셀리에 요청을 보내고, 같은 Excel 표의 변화를 확인합니다.",live="format_send",click=True,clickAt=3.15)
screen("format_processing",1,"천원 단위로 표시","금액과 합계 셀의 표시형식을 바꿉니다.",
       [panel("format_processing","celly",None,[190,285,1540,580])],live="format",note="처리 대기 일부 생략")
desktop("format_result_wide",6,"천원 단위로 표시","같은 표의 금액과 합계가 천원 단위로 표시됩니다.",live="format_result",note="처리 대기 일부 생략")
screen("format_before",2,"천원 단위로 표시","변경 전: 원 단위 금액과 합계.",
       [panel("totals","excel",table,[170,205,1580,590]),panel("totals","excel",formula,[480,833,960,85])],label="변경 전 · 같은 셀 F13",unit="단위: 원")
screen("format_after",2,"천원 단위로 표시","변경 후: 표시만 천원 단위로.",
       [panel("final","excel",table,[170,205,1580,590]),panel("final","excel",formula,[480,833,960,85])],label="변경 후 · 같은 셀 F13",unit="단위: 천원")
screen("final",5,"천원 단위로 표시","원래 금액과 수식은 그대로,\n화면 표시는 천원 단위로.",
       [panel("final","excel",table,table_box)],label="Excel · 최종 결과",unit="단위: 천원")
card("chapter",2,"5. 확인하고 저장")
desktop("save_wide",6,"확인하고 저장","작업 결과를 확인한 뒤, Excel에서 Ctrl+S로 저장합니다.",live="save")
screen("saved",3,"확인하고 저장","이어 작업한 결과를 Excel에서 저장하세요.",
       [panel("saved","excel",[285,1,910,62],[340,220,1240,125]),panel("saved","excel",table,[200,370,1520,510])],label="Excel · 저장 완료",shortcut="Ctrl + S",unit="단위: 천원",unitX=1640,unitY=315)
card("ending",5,"파일 첨부 없이, 말로 이어가는 Excel 작업")

# From the requested 38-second point onward, the two complete application windows
# stay in exactly the same frame. Magnification is a local loupe over blank cells.
def round_lens(focus,title,zoom=2.0):
    return {"shape":"circle","focus":focus,"zoom":zoom,"size":[360,360],"position":[860,530],"title":title}

def text_lens(source,title,zoom=1.65):
    return {"shape":"rect","source":source,"zoom":zoom,"position":[500,690],"title":title}

replacements={
    "first_reply":dict(desktopRole="initial",lens=text_lens([1415,224,1868,305],"작업 완료",1.5)),
    "initial":dict(desktopRole="initial",lens=round_lens([230,302],"회사별 · 연도별 집계")),
    "totals_input":dict(desktopRole="totals_input",lens=text_lens([1415,875,1870,932],"합계 추가 요청")),
    "totals_processing":dict(desktopLive="totals_wait"),
    "annual_total":dict(desktopRole="totals_workspace",lens=round_lens([230,532],"연도별 합계")),
    "company_total":dict(desktopRole="totals_workspace",lens=round_lens([726,302],"회사별 합계")),
    "format_context":dict(desktopRole="totals_workspace"),
    "format_input":dict(desktopRole="format_input",lens=text_lens([1415,875,1870,932],"표시형식 변경 요청")),
    "format_processing":dict(desktopLive="format_wait"),
    "format_before":dict(desktopRole="final",lens=round_lens([726,532],"전체 합계 · 단위: 천원"),caption="금액과 합계의 표시만 천원 단위로 바뀌었습니다."),
    "format_after":dict(desktopRole="final",lens=text_lens([285,122,450,159],"수식은 그대로",2.3),caption="실제 값과 수식은 그대로 유지됩니다."),
    "final":dict(desktopRole="final",caption="원래 금액과 수식은 그대로, 화면 표시는 천원 단위로."),
    "saved":dict(desktopRole="saved",lens=text_lens([375,22,840,59],"Excel에 저장됨",1.6)),
}
position=0
for n,shot in enumerate(shots):
    start=position
    position+=shot["seconds"]
    if start<38:
        continue
    if shot["kind"]=="chapter":
        phase=shot["title"].partition(". ")[2]
        role="initial" if phase.startswith("첨부") else "totals_workspace" if phase.startswith("천원") else "final"
        shots[n]=dict(id=f"phase_{n}",kind="desktop",seconds=shot["seconds"],phase=phase,caption=phase,desktopRole=role,phaseTransition=True)
        continue
    if shot["kind"]=="ending":
        shots[n]=dict(id="ending_workspace",kind="desktop",seconds=shot["seconds"],phase="셀리",caption="지금 열린 Excel에서, 바꾸고 싶은 일을 입력해 보세요.",desktopRole="saved",ending=True)
        continue
    replacement=replacements.get(shot.get("id"))
    if replacement:
        keep={k:v for k,v in shot.items() if k in ["id","seconds","phase","caption","note","click","clickAt"]}
        shots[n]={**keep,"kind":"desktop",**replacement}

assert position==123
position=0
for shot in shots:
    if position+shot["seconds"]>38:
        assert shot["kind"]=="desktop",(position,shot.get("id"))
    position+=shot["seconds"]

desktop("gif_final_workspace",2,"같은 파일에서 계속","같은 Excel 파일에서, 천원 단위 표시까지.",role="final",gifOnly=True)
index={shot.get("id"):i for i,shot in enumerate(shots) if shot.get("id")}
gif=[dict(shot=index[id],seconds=seconds,caption=caption) for id,seconds,caption in [
    ("continuity",2,"Excel은 그대로, 셀리에 이어서 요청"),
    ("totals_input",2,"첨부 없이, 합계 추가 요청"),
    ("company_total",2,"같은 표에 연도별·회사별 합계 추가"),
    ("format_input",2,"표시형식도 말로 변경"),
    ("gif_final_workspace",2,"같은 Excel 파일에서, 천원 단위 표시까지"),
]]
gif_only=shots.pop()
assert gif_only.get("gifOnly")
# Keep the GIF-only composition outside the film timeline.
gif[-1]["composition"]=gif_only
gif[-1].pop("shot")
assert sum(s["seconds"] for s in shots)==123
assert sum(s["seconds"] for s in gif)==10
config={"roles":roles,"shots":shots,"gif":gif,"take":2,"source":"new native PNG captures and raw-take2.mkv","revision":"After 38s, fixed full workspace only; local magnifier overlays; quiet-instrument visual direction","wideSeconds":sum(s["seconds"] for s in shots if s["kind"]=="desktop"),"fullWorkspaceAfterSeconds":38,"designSkill":"design-taste-frontend","designDirection":"quiet-instrument","directionSeed":259388649,"contentAndTimingBefore38Preserved":True}
(ROOT/"film-edit.json").write_text(json.dumps(config,ensure_ascii=False,indent=2),encoding="utf-8")
print(f"Prepared {len(shots)} shots, 123s video, {config['wideSeconds']}s full workspace, 10s GIF, genuine full-workspace send/result clips.")
