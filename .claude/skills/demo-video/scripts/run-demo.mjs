#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import YAML from 'yaml';
import { chromium } from 'playwright';
import ffmpegPath from 'ffmpeg-static';
import ffprobe from 'ffprobe-static';
import OpenAI from 'openai';
import dotenv from 'dotenv';

const here = path.dirname(fileURLToPath(import.meta.url));
const skillDir = path.dirname(here);
dotenv.config({ path: path.join(skillDir, '.env'), quiet: true });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const run = (bin, args) => new Promise((resolve, reject) => {
  const child = spawn(bin, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
  let log = '', out = ''; child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { log += d; }); child.on('error', reject);
  child.on('close', (code) => code === 0 ? resolve(out) : reject(new Error(`${path.basename(bin)} exited ${code}: ${log.slice(-1200)}`)));
});
// shortcut: keyword match on the style text, add entries when a new tone is needed
const pickVoice = (style = '') => /밝|활기|경쾌|발랄/.test(style) ? 'coral' : /중후|진중|권위|낮은/.test(style) ? 'onyx' : 'marin';
const sha = (v) => crypto.createHash('sha256').update(v).digest('hex').slice(0, 24);
const q = (v) => String(v).replaceAll("'", "''");
const overlay = String.raw`(() => {
  const setup = () => {
  const root = document.createElement('div'); root.id='__demo_overlay'; root.innerHTML='<div id="__demo_cursor"><svg width="36" height="36" viewBox="0 0 24 24"><path d="M3 2l7.5 19 2.7-8.1L21 10z" fill="#111" stroke="#fff" stroke-width="1.6" stroke-linejoin="round"/></svg></div><div id="__demo_ripple"></div><div id="__demo_highlight"><span></span></div><div id="__demo_caption"></div><div id="__demo_card"><h1></h1><p></p></div>';
  const css='#__demo_overlay{position:fixed;inset:0;z-index:2147483647;pointer-events:none;font-family:Arial,"Malgun Gothic",sans-serif}#__demo_cursor{position:fixed;left:0;top:0;width:36px;height:36px;transform:translate(-4.5px,-3px);filter:drop-shadow(0 2px 3px rgba(0,0,0,.4));transition:left .035s linear,top .035s linear}#__demo_ripple{position:fixed;width:12px;height:12px;border:4px solid #52d6ff;border-radius:50%;opacity:0;transform:translate(-50%,-50%)}#__demo_ripple.on{animation:__demo_pop .55s ease-out}@keyframes __demo_pop{0%{opacity:1;transform:translate(-50%,-50%) scale(.1)}100%{opacity:0;transform:translate(-50%,-50%) scale(5)}}#__demo_caption{position:fixed;bottom:54px;left:50%;transform:translateX(-50%);max-width:76%;padding:15px 26px;border-radius:12px;background:rgba(0,0,0,.78);color:white;font-size:32px;line-height:1.35;text-align:center;opacity:0;transition:opacity .2s}#__demo_caption.on{opacity:1}#__demo_card{position:fixed;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;color:#fff;background:linear-gradient(135deg,#121b38,#305aa5);opacity:0;transition:opacity .35s}#__demo_card.on{opacity:1}#__demo_card h1{font-size:72px;margin:0 0 18px}#__demo_card p{font-size:30px;margin:0;opacity:.85}#__demo_highlight{position:fixed;border:5px solid #54d9ff;border-radius:12px;box-shadow:0 0 0 9999px rgba(0,0,0,.18),0 0 25px #54d9ff;opacity:0;transition:opacity .2s}#__demo_highlight.on{opacity:1}#__demo_highlight span{position:absolute;top:-39px;left:-5px;background:#54d9ff;color:#07101f;padding:6px 12px;border-radius:7px;font-size:20px;font-weight:bold}';
  const style=document.createElement('style');style.textContent=css;root.append(style);document.documentElement.append(root);
  const $=(id)=>document.getElementById(id); window.__demo={
    cursor:({x,y})=>{$('__demo_cursor').style.left=x+'px';$('__demo_cursor').style.top=y+'px'},
    click:({x,y})=>{const e=$('__demo_ripple');e.style.left=x+'px';e.style.top=y+'px';e.classList.remove('on');void e.offsetWidth;e.classList.add('on')},
    caption:(text)=>{$('__demo_caption').textContent=text||'';$('__demo_caption').classList.toggle('on',!!text)},
    card:(title,subtitle)=>{$('__demo_card').querySelector('h1').textContent=title;$('__demo_card').querySelector('p').textContent=subtitle||'';$('__demo_card').classList.add('on')},
    hideCard:()=>$('__demo_card').classList.remove('on'), highlight:(b,label)=>{const e=$('__demo_highlight');Object.assign(e.style,{left:(b.x-8)+'px',top:(b.y-8)+'px',width:(b.width+16)+'px',height:(b.height+16)+'px'});e.querySelector('span').textContent=label||'';e.classList.add('on')}, clear:()=>$('__demo_highlight').classList.remove('on')
  };
  };
  if (document.documentElement) setup(); else document.addEventListener('DOMContentLoaded', setup, { once: true });
})()`;

async function duration(file) { const out = await run(ffprobe.path, ['-v','error','-show_entries','format=duration','-of','default=nw=1:nk=1',file]); return Math.ceil(Number(out.trim()) * 1000); }
async function makeVoice(text, voice, cache) {
  const mode = voice?.mode || 'none'; if (mode === 'none') return null;
  const key = sha(JSON.stringify({ text, mode, model: voice.model, name: voice.name, voice: voice.voice, style: voice.style }));
  const ext = mode === 'local' ? 'wav' : 'mp3'; const file = path.join(cache, `${key}.${ext}`);
  try { await fs.access(file); return { file, durationMs: await duration(file) }; } catch {}
  if (mode === 'openai') {
    if (!process.env.OPENAI_API_KEY) throw new Error('OpenAI 음성이 선택되었지만 .env의 OPENAI_API_KEY가 비어 있습니다.');
    const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
    const speech = await client.audio.speech.create({ model: voice.model || 'gpt-4o-mini-tts', voice: voice.voice || pickVoice(voice.style), input: text, instructions: voice.style || '차분하고 친근한 제품 소개 톤, 또렷한 한국어' });
    await fs.writeFile(file, Buffer.from(await speech.arrayBuffer()));
  } else if (mode === 'local') {
    const ps = path.join(cache, `${key}.ps1`); const voiceArg = voice.name ? `$s.SelectVoice('${q(voice.name)}');` : `$k=$s.GetInstalledVoices()|Where-Object{$_.VoiceInfo.Culture.Name -eq 'ko-KR'}|Select-Object -First 1; if(-not $k){throw '설치된 한국어 음성이 없습니다.'}; $s.SelectVoice($k.VoiceInfo.Name);`;
    await fs.writeFile(ps, `\ufeff$ErrorActionPreference='Stop'; Add-Type -AssemblyName System.Speech; $s=New-Object System.Speech.Synthesis.SpeechSynthesizer; ${voiceArg} $s.Rate=0; $s.SetOutputToWaveFile('${q(file)}'); $s.Speak('${q(text)}'); $s.Dispose()`);
    try { await run('powershell.exe', ['-NoProfile','-ExecutionPolicy','Bypass','-File',ps]); } catch (e) { await fs.rm(file, { force: true }); throw new Error(`로컬 음성 생성 실패: ${e.message}`); } finally { await fs.rm(ps, { force: true }); }
  } else throw new Error(`지원하지 않는 voice.mode: ${mode}`);
  return { file, durationMs: await duration(file) };
}
function sentences(text) { return String(text || '').split(/(?<=[.!?。！？])\s+/).map((v) => v.trim()).filter(Boolean); }
async function makeChapterVoice(text, voice, cache) {
  if ((voice?.mode || 'none') === 'none') return null;
  const parts = [];
  for (const sentence of sentences(text)) parts.push(await makeVoice(sentence, voice, cache));
  if (!parts.length) return null;
  if (parts.length === 1) return { ...parts[0], sentences: parts };
  const file = path.join(cache, `${sha(JSON.stringify({ text, voice, kind: 'chapter' }))}.m4a`);
  try { await fs.access(file); return { file, durationMs: await duration(file), sentences: parts }; } catch {}
  const args = ['-y']; for (const part of parts) args.push('-i', part.file);
  args.push('-filter_complex', `${parts.map((_, i) => `[${i}:a]`).join('')}concat=n=${parts.length}:v=0:a=1[a]`, '-map', '[a]', '-c:a', 'aac', file);
  await run(ffmpegPath, args); return { file, durationMs: await duration(file), sentences: parts };
}
async function main() {
  const scenarioPath = path.resolve(process.argv[2] || 'scenario.yml');
  const raw = await fs.readFile(scenarioPath, 'utf8'); const s = YAML.parse(raw);
  if (!s?.url || !Array.isArray(s.chapters) || !s.chapters.length) throw new Error('시나리오에는 url과 하나 이상의 chapters가 필요합니다.');
  const base = path.dirname(scenarioPath), name = String(s.name || path.basename(scenarioPath, path.extname(scenarioPath))).replace(/[^a-zA-Z0-9_-]/g, '-');
  const out = path.join(base, 'out', name), cache = path.join(out, 'voice-cache'); await fs.mkdir(cache, { recursive: true });
  const voice = s.voice || { mode: 'none' }; const prepared = [];
  // TTS is deliberately created and cached sentence-by-sentence before recording.
  for (const chapter of s.chapters) prepared.push(await makeChapterVoice(chapter.narration || '', voice, cache));
  const size = s.viewport || { width: 1920, height: 1080 }; if (size.width !== 1920 || size.height !== 1080) throw new Error('이 스킬의 녹화 규격은 1920x1080입니다.');
  const browser = await chromium.launch({ headless: true }); let page, context, video; const timeline = []; let started;
  try {
    context = await browser.newContext({ viewport: size, recordVideo: { dir: out, size }, deviceScaleFactor: 1 });
    await context.addInitScript(overlay); page = await context.newPage(); video = page.video(); started = Date.now(); await page.goto(s.url, { waitUntil: 'domcontentloaded' });
    let current = { x: 80, y: 80 }, caption = ''; await page.evaluate((p) => window.__demo.cursor(p), current);
    page.on('domcontentloaded', () => page.evaluate(([p, t]) => { window.__demo?.cursor(p); window.__demo?.caption(t); }, [current, caption]).catch(() => {}));
    const move = async (selector) => { const b = await page.locator(selector).first().boundingBox(); if (!b) throw new Error(`요소를 찾을 수 없습니다: ${selector}`); const end={x:b.x+b.width/2,y:b.y+b.height/2}, start={...current}, bend={x:(start.x+end.x)/2+(end.y-start.y)*.15,y:(start.y+end.y)/2-(end.x-start.x)*.15}; for(let i=1;i<=24;i++){const t=i/24,u=1-t;const p={x:u*u*start.x+2*u*t*bend.x+t*t*end.x,y:u*u*start.y+2*u*t*bend.y+t*t*end.y};await page.evaluate((v)=>window.__demo.cursor(v),p);await sleep(18);} current=end; return b; };
    for (let i=0;i<s.chapters.length;i++) { const c=s.chapters[i], startedAt=Date.now()-started, audio=prepared[i];
      await page.evaluate(([t,n])=>window.__demo.card(t,n), [c.title || `Chapter ${i+1}`, c.narration || '']); await sleep(c.cardMs ?? 1600); await page.evaluate(()=>window.__demo.hideCard());
      caption=c.narration||''; await page.evaluate((t)=>window.__demo.caption(t), caption); const actionsStart=Date.now();
      for (const a of c.actions || []) { if (a.do==='goto') await page.goto(new URL(a.url, s.url).href,{waitUntil:'domcontentloaded'}); else if(a.do==='wait') await sleep(a.ms||500); else if(a.do==='move') await move(a.selector); else if(a.do==='click'){await move(a.selector);await page.evaluate((p)=>window.__demo.click(p),current);await page.mouse.click(current.x,current.y);await sleep(a.afterMs??280)} else if(a.do==='type'){if(a.selector){await move(a.selector);await page.locator(a.selector).first().click();}await page.keyboard.type(a.text||'',{delay:a.delayMs??90});} else if(a.do==='highlight'){const b=await move(a.selector);await page.evaluate(([box,label])=>window.__demo.highlight(box,label),[b,a.label||'']);await sleep(a.ms??900)} else if(a.do==='screenshot') await page.screenshot({path:path.join(out,a.path||`chapter-${i+1}.png`),fullPage:false}); else throw new Error(`지원하지 않는 액션: ${a.do}`); }
      const actionMs=Date.now()-actionsStart, target=Math.max(actionMs, audio?.durationMs||0)+(voice.paddingMs??700); if(target>actionMs) await sleep(target-actionMs); caption=''; await page.evaluate(()=>{window.__demo.caption('');window.__demo.clear()});
      timeline.push({ chapter:i+1,title:c.title,startMs:startedAt,audioStartMs:startedAt+(c.cardMs??1600),durationMs:Date.now()-started-startedAt,actionMs,audio });
    }
  } catch (error) { if (page) { await page.screenshot({path:path.join(out,'failure.png'),fullPage:false}).catch(()=>{}); } await fs.writeFile(path.join(out,'failure.json'),JSON.stringify({message:error.message,at:new Date().toISOString()},null,2)); throw error; }
  finally { if(context) await context.close(); await browser.close(); }
  const webm = path.join(out,'recording.webm'); await fs.rename(await video.path(), webm); const hasAudio = timeline.some((x)=>x.audio); const final = path.join(out,'final.mp4'), audioFile=path.join(out,'narration.m4a');
  if (hasAudio) { const inputs=[]; const filters=[]; let n=0; for(const item of timeline){if(!item.audio)continue;inputs.push('-i',item.audio.file);filters.push(`[${n}:a]adelay=${item.audioStartMs}|${item.audioStartMs}[a${n}]`);n++;} filters.push(`${Array.from({length:n},(_,i)=>`[a${i}]`).join('')}amix=inputs=${n}:duration=longest:normalize=0[a]`); await run(ffmpegPath,['-y',...inputs,'-filter_complex',filters.join(';'),'-map','[a]','-c:a','aac',audioFile]); }
  const args=['-y','-i',webm]; if(hasAudio) args.push('-i',audioFile,'-map','0:v:0','-map','1:a:0'); args.push('-c:v','libx264','-pix_fmt','yuv420p','-movflags','+faststart','-c:a','aac',final); await run(ffmpegPath,args);
  await fs.writeFile(path.join(out,'timeline.json'),JSON.stringify(timeline.map(({audio,...r})=>({...r,audio:audio ? {file:path.basename(audio.file),durationMs:audio.durationMs} : null})),null,2));
  if(hasAudio) await fs.writeFile(path.join(out,'subtitles.srt'),timeline.filter((x)=>x.audio).map((x,i)=>`${i+1}\n${stamp(x.audioStartMs)} --> ${stamp(x.audioStartMs+x.audio.durationMs)}\n${s.chapters[x.chapter-1].narration}\n`).join('\n'));
  console.log(`완료: ${final}`);
}
function stamp(ms){const h=Math.floor(ms/3600000),m=Math.floor(ms%3600000/60000),s=Math.floor(ms%60000/1000),z=ms%1000;return [h,m,s].map(x=>String(x).padStart(2,'0')).join(':')+','+String(z).padStart(3,'0');}
main().catch((e)=>{const key=process.env.OPENAI_API_KEY;console.error(`실패: ${key?e.message.replaceAll(key,'***'):e.message}`); if (process.env.DEMO_VIDEO_DEBUG) console.error(e.stack); process.exitCode=1});
