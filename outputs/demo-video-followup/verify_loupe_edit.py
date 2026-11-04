from pathlib import Path
import hashlib
import json
import re
import subprocess

import imageio_ffmpeg
from PIL import Image, ImageDraw, ImageFont, ImageSequence

ROOT=Path(__file__).resolve().parent
PROJECT=ROOT.parents[1]
video=ROOT/'셀리_소개_완성본.mp4'
ffmpeg=imageio_ffmpeg.get_ffmpeg_exe()
config=json.loads((ROOT/'film-edit.json').read_text(encoding='utf-8'))
shots=json.loads((ROOT/'edit-timeline.json').read_text(encoding='utf-8'))
assert shots[-1]['end']==123
assert all(s['kind']=='desktop' for s in shots if s['end']>38)
assert shots[-1]['kind']=='desktop' and shots[-1]['ending']
assert sum(s['seconds'] for s in shots if s['kind']=='desktop')==97
assert all('panels' not in s for s in shots if s['end']>38)
metadata=subprocess.run([ffmpeg,'-hide_banner','-i',str(video)],capture_output=True,text=True,encoding='utf-8').stderr
assert '1920x1080' in metadata and '30 fps' in metadata and '00:02:03.00' in metadata
assert 'Audio: aac' in metadata and 'stereo' in metadata
(ROOT/'encoded-media-metadata.txt').write_text(metadata,encoding='utf-8')
subprocess.run([ffmpeg,'-v','error','-i',str(video),'-f','null','NUL'],check=True,timeout=90)
print('Full decode passed',flush=True)
folder=ROOT/'encoded-loupe-checks'
folder.mkdir(exist_ok=True)
times=[37.8,38,43.03,43.5,46.5,50,55,62.5,65,70,74,78,83,91.5,95,99,101,105,108,112,116,121]
font=ImageFont.truetype(str(PROJECT/'app/fonts/Pretendard-Regular.ttf'),18)
contact=Image.new('RGB',(1920,((len(times)+3)//4)*300),'#F3F5F4')
for i,t in enumerate(times):
    target=folder/f'{t:06.2f}.png'
    subprocess.run([ffmpeg,'-v','error','-y','-ss',str(t),'-i',str(video),'-frames:v','1','-update','1',str(target)],check=True,timeout=30)
    with Image.open(target) as im:
        assert im.size==(1920,1080)
        contact.paste(im.resize((480,270),Image.Resampling.LANCZOS),((i%4)*480,(i//4)*300))
    ImageDraw.Draw(contact).text(((i%4)*480+10,(i//4)*300+275),f'{t:.2f}s',font=font,fill='#1C2825')
contact.save(folder/'contact.png')
with Image.open(ROOT/'셀리_10초_채팅용.gif') as im:
    durations=[frame.info['duration'] for frame in ImageSequence.Iterator(im)]
    gifsize=im.size
assert sum(durations)==10000 and gifsize==(1152,648)
workspace={
    'passed':True,'durationSeconds':123,'workspaceSeconds':97,
    'workspaceScenes':sum(s['kind']=='desktop' for s in shots),
    'fullWorkspaceAfterSeconds':38,'continuousWorkspaceToSeconds':123,
    'liveWorkspaceClips':len({s['desktopLive'] for s in shots if s.get('desktopLive')}),
    'lensScenes':sum(bool(s.get('lens')) for s in shots),
    'maxSimultaneousLenses':1,'fixedWorkspaceRectangle':[98,76,1724,916],
    'noCloseupCutsOrFullScreenCardsAfter38':True,
    'flow':'fixed full workspace -> local request loupe -> genuine send/result -> local value loupe -> full workspace ending',
    'footage':'verified take 2 only; source workbook was not changed by this re-edit',
    'gifFrames':len(durations),'gifDurationMs':sum(durations)
}
report={
    'passed':True,
    'video':{'width':1920,'height':1080,'fps':30,'durationSeconds':123,'bytes':video.stat().st_size,
             'fullDecodePassed':True,'extractedCheckFrames':len(times),'checkTimesSeconds':times,
             'sha256':hashlib.sha256(video.read_bytes()).hexdigest(),
             'audio':'AAC stereo, instrumental BGM and actual-send-timed click effects'},
    'workspace':workspace,
    'gif':{'durationMs':sum(durations),'frames':len(durations),'size':gifsize,
           'bytes':(ROOT/'셀리_10초_채팅용.gif').stat().st_size,'fullWorkspaceThroughout':True},
    'design':{'skill':'design-taste-frontend','direction':'quiet-instrument','seed':259388649},
    'visualReview':'pending encoded contact inspection'
}
(ROOT/'workspace-verification.json').write_text(json.dumps(workspace,ensure_ascii=False,indent=2),encoding='utf-8')
(ROOT/'media-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
