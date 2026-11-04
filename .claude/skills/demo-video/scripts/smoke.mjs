#!/usr/bin/env node
// Self-check: node scripts/smoke.mjs records a two-page sample, then checks the failure paths.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import ffprobe from 'ffprobe-static';

const here = path.dirname(fileURLToPath(import.meta.url));
const dir = path.join(here, '..', 'out', 'smoke-test');
await fs.rm(dir, { recursive: true, force: true });
await fs.mkdir(dir, { recursive: true });
const html = path.join(dir, 'page.html');
await fs.writeFile(html, '<!doctype html><meta charset=utf-8><body style="font:40px sans-serif;padding:200px"><h1 id=t>시연 페이지</h1><input id=q style="font-size:32px"><a id=n href="page.html?2">다음</a>');
const url = pathToFileURL(html).href;
const chapters = [
  { title: '첫 장면', narration: '첫 번째 문장입니다. 비율은 3.5퍼센트입니다.', actions: [
    { do: 'click', selector: '#q' }, { do: 'type', text: '고객', delayMs: 60 }, { do: 'highlight', selector: '#t', label: '제목', ms: 300 }] },
  { title: '이동', narration: '다음 페이지로 이동합니다.', actions: [{ do: 'click', selector: '#n' }, { do: 'wait', ms: 300 }, { do: 'screenshot', path: 'after-nav.png' }] },
];
const run = async (name, voice, list = chapters, env = {}) => {
  const file = path.join(dir, `${name}.json`); // JSON is valid YAML
  await fs.writeFile(file, JSON.stringify({ name, url, voice, chapters: list }));
  return spawnSync(process.execPath, [path.join(here, 'run-demo.mjs'), file], { encoding: 'utf8', env: { ...process.env, ...env } });
};
const probe = (file) => JSON.parse(spawnSync(ffprobe.path, ['-v', 'error', '-show_streams', '-show_format', '-of', 'json', file], { encoding: 'utf8' }).stdout);

let r = await run('none', { mode: 'none' });
assert.equal(r.status, 0, r.stderr);
let info = probe(path.join(dir, 'out/none/final.mp4'));
const video = info.streams.find((s) => s.codec_type === 'video');
assert.deepEqual([video.codec_name, video.pix_fmt, video.width, video.height], ['h264', 'yuv420p', 1920, 1080]);
assert.ok(!info.streams.some((s) => s.codec_type === 'audio'));
await fs.access(path.join(dir, 'out/none/recording.webm'));

r = await run('bad-selector', { mode: 'none' }, [{ title: '실패', actions: [{ do: 'wait', ms: 100 }, { do: 'move', selector: '#missing' }] }]);
assert.notEqual(r.status, 0);
await fs.access(path.join(dir, 'out/bad-selector/failure.png'));

r = await run('no-key', { mode: 'openai' }, chapters, { OPENAI_API_KEY: '' });
assert.notEqual(r.status, 0);
assert.match(r.stderr, /OPENAI_API_KEY/);
await fs.access(path.join(dir, 'out/no-key/final.mp4')).then(() => assert.fail('must not fall back'), () => {});

r = await run('bad-voice', { mode: 'local', name: '없는 음성' });
assert.notEqual(r.status, 0);
assert.match(r.stderr, /로컬 음성 생성 실패/);

r = await run('local', { mode: 'local', paddingMs: 300 });
if (/한국어 음성이 없습니다/.test(r.stderr)) console.log('건너뜀: 한국어 로컬 음성 없음');
else {
  assert.equal(r.status, 0, r.stderr);
  const out = path.join(dir, 'out/local');
  const timeline = JSON.parse(await fs.readFile(path.join(out, 'timeline.json'), 'utf8'));
  for (const c of timeline) {
    assert.ok(c.audio.durationMs > 500, 'measured voice length');
    assert.ok(c.durationMs >= Math.max(c.actionMs, c.audio.durationMs) + 300, 'scene covers voice + padding');
  }
  const last = timeline.at(-1);
  info = probe(path.join(out, 'final.mp4'));
  assert.ok(info.streams.some((s) => s.codec_type === 'audio'));
  assert.ok(Number(info.format.duration) * 1000 >= last.audioStartMs + last.audio.durationMs, 'ending not cut');
  assert.match(await fs.readFile(path.join(out, 'subtitles.srt'), 'utf8'), /첫 번째 문장입니다/);
  const cached = path.join(out, 'voice-cache', timeline.at(-1).audio.file);
  const before = (await fs.stat(cached)).mtimeMs;
  assert.equal((await run('local', { mode: 'local', paddingMs: 300 })).status, 0);
  assert.equal((await fs.stat(cached)).mtimeMs, before, 'voice cache reused');
}
console.log('smoke 통과');
