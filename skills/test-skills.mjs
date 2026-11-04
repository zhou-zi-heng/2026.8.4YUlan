import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const mem = {};
let seq = 0;
const context = {
    console,
    crypto,
    DB: {
        async setSetting(key, value) { mem[key] = JSON.parse(JSON.stringify(value)); },
        async getSetting(key, fallback) {
            return Object.prototype.hasOwnProperty.call(mem, key) ? mem[key] : (fallback === undefined ? null : fallback);
        }
    },
    gId() { seq += 1; return 'skid' + String(seq).padStart(8, '0'); },
    toast() {}
};
context.globalThis = context;
vm.createContext(context);
const code = fs.readFileSync(new URL('../js/skills.js', import.meta.url), 'utf8');
vm.runInContext(code + '\nthis.Skills = Skills;\n', context);
const Skills = context.Skills;

async function reset() {
    for (const k of Object.keys(mem)) delete mem[k];
    await Skills.load();
}

const sampleFile = JSON.parse(fs.readFileSync(new URL('./example/skill.json', import.meta.url), 'utf8'));
const sampleMd = fs.readFileSync(new URL('./example/SKILL.md', import.meta.url), 'utf8');
const embedded = JSON.parse(JSON.stringify(Skills.sampleSource()));
assert.deepEqual(embedded.meta, sampleFile);
assert.equal(embedded.markdown.trim(), sampleMd.trim());

await reset();
const installed = await Skills.installSample();
assert.equal(installed.skill.enabled, false);
assert.equal(Skills.prepareForSend('请条理说明一下').injection, '');

await Skills.toggleEnabled(installed.skill.id, true);
const on = Skills.prepareForSend('请条理说明一下');
assert.ok(on.injection.includes('【示例技能：条理化回答】'));
assert.equal(on.used.length, 1);
assert.equal(on.used[0].reason, '匹配');
assert.equal(on.used[0].name, '条理化回答');
assert.ok(on.injection.length <= Skills.LIMITS.MAX_TOTAL_CHARS);

const off = Skills.prepareForSend('今天天气不错');
assert.equal(off.injection, '');
assert.equal(off.used.length, 0);

await Skills.toggleEnabled(installed.skill.id, false);
assert.equal(Skills.prepareForSend('请按步骤条理化').injection, '');
await Skills.toggleEnabled(installed.skill.id, true);

const saved = mem.feifan_skills;
assert.ok(saved && saved.skills && saved.skills[0].enabled === true);
await Skills.load();
assert.equal(Skills.list()[0].enabled, true);
assert.ok(Skills.prepareForSend('分点说说').injection.includes('【示例技能：条理化回答】'));

await assert.rejects(
    () => Skills.installPackage({ jsonText: '{"name":"坏","tools":["shell"]}', markdown: '不要执行' }),
    /不允许的字段/
);
assert.equal(Skills.list().length, 1);

await assert.rejects(
    () => Skills.installPackage({ jsonText: '{"name":"短","triggers":["a"]}', markdown: '正文' }),
    /2–40/
);

const manualOnly = await Skills.installPackage({
    jsonText: '{"name":"仅手动","description":"没有触发词","version":1}',
    markdown: '手动说明书MANUAL_ONLY'
});
await Skills.toggleEnabled(manualOnly.skill.id, true);
assert.equal(Skills.prepareForSend('仅手动也别自动带上').injection, '');
Skills.toggleManual(manualOnly.skill.id, true);
const manualPack = Skills.prepareForSend('随便写');
assert.ok(manualPack.injection.includes('MANUAL_ONLY'));
assert.equal(manualPack.used.find(u => u.name === '仅手动').reason, '手动');
Skills.consumeManual(manualPack.used.filter(u => u.name === '仅手动'));
assert.equal(Skills.prepareForSend('随便写').injection.includes('MANUAL_ONLY'), false);
assert.ok(Skills.prepareForSend('请条理说明').injection.includes('【示例技能：条理化回答】'));

const base = '你是助手';
assert.equal(Skills.attachToSystem(base, ''), base);
const joined = Skills.attachToSystem(base, on.injection);
assert.ok(joined.startsWith(base + '\n\n'));
assert.ok(joined.includes('【示例技能：条理化回答】'));
assert.equal(Skills.attachToSystem('', on.injection), on.injection);

await Skills.toggleEnabled(installed.skill.id, false);
assert.equal(Skills.replay([{ id: installed.skill.id, name: '条理化回答', reason: '匹配' }]).injection, '');
await Skills.toggleEnabled(installed.skill.id, true);
const replayed = Skills.replay([{ id: installed.skill.id, reason: '手动' }]);
assert.ok(replayed.injection.includes('【示例技能：条理化回答】'));
assert.equal(replayed.used[0].reason, '手动');

await reset();
const bigBody = '甲'.repeat(9000);
await Skills.installPackage({
    jsonText: '{"name":"很长","triggers":["预算"]}',
    markdown: bigBody + '\nSMALL_SHOULD_NOT_BE_THE_WHOLE'
});
await Skills.installPackage({
    jsonText: '{"name":"很短","triggers":["预算"]}',
    markdown: '短说明SMALL_SKILL_TOKEN'
});
const ids = Skills.list();
await Skills.toggleEnabled(ids[0].id, true);
await Skills.toggleEnabled(ids[1].id, true);
const budgeted = Skills.prepareForSend('请看预算');
assert.equal(budgeted.used.length, 2);
assert.equal(budgeted.used[0].truncated, true);
assert.equal(budgeted.used[1].truncated, false);
assert.ok(budgeted.injection.includes('已按长度预算截断'));
assert.ok(budgeted.injection.includes('SMALL_SKILL_TOKEN'));
assert.ok(budgeted.injection.length <= Skills.LIMITS.MAX_TOTAL_CHARS);
assert.ok(!budgeted.injection.includes('甲'.repeat(3000)));

await reset();
for (let n = 1; n <= 4; n++) {
    const made = await Skills.installPackage({
        jsonText: JSON.stringify({ name: '项' + n, triggers: ['全部'] }),
        markdown: '正文' + n
    });
    await Skills.toggleEnabled(made.skill.id, true);
    Skills.toggleManual(made.skill.id, true);
}
const capped = Skills.prepareForSend('');
assert.equal(capped.used.length, Skills.LIMITS.MAX_COUNT);
assert.equal(capped.omitted, 1);
assert.equal(Array.from(capped.used, u => u.name).join(','), '项1,项2,项3');

await reset();
const names = ['甲', '乙', '丙', '丁'];
const madeIds = [];
for (const name of names) {
    const triggers = name === '丁' ? [] : ['命中'];
    const made = await Skills.installPackage({
        jsonText: JSON.stringify({ name, triggers }),
        markdown: '正文-' + name
    });
    madeIds.push(made.skill.id);
    await Skills.toggleEnabled(made.skill.id, true);
}
Skills.toggleManual(madeIds[3], true);
const mixed = Skills.prepareForSend('请命中这个说明');
assert.equal(Array.from(mixed.used, u => u.name).join(','), '丁,甲,乙');
assert.equal(mixed.omitted, 1);
assert.equal(mixed.used[0].reason, '手动');

const picked = Skills.pickZipTexts([
    { path: '../skill.json', text: 'bad' },
    { path: 'SKILL.md', text: '# 根上的说明，目录不同' },
    { path: 'sub/skill.json', text: '{"name":"子目录"}' },
    { path: 'sub/SKILL.md', text: '# 同一目录' },
    { path: 'notes/SKILL.md', text: '# 其它' }
]);
assert.equal(picked.jsonText, '{"name":"子目录"}');
assert.equal(picked.markdown, '# 同一目录');

let history = [];
function outbound(text) {
    const pack = Skills.prepareForSend(text);
    history.push({ role: 'user', content: text });
    const system = Skills.attachToSystem('', pack.injection);
    const msgs = [];
    if (system) msgs.push({ role: 'system', content: system });
    history.forEach(m => msgs.push({ role: m.role, content: m.content }));
    return msgs;
}
await reset();
const again = await Skills.installSample();
await Skills.toggleEnabled(again.skill.id, true);
history = [];
const first = outbound('请条理说明');
assert.ok(first[0].content.includes('【示例技能：条理化回答】'));
assert.ok(!first.some(m => m.role === 'user' && String(m.content).includes('【示例技能：条理化回答】')));
const second = outbound('你好');
const blob = JSON.stringify(second);
assert.ok(!blob.includes('【示例技能：条理化回答】'));
assert.equal(second.filter(m => m.role === 'system').length, 0);

const app = fs.readFileSync(new URL('../js/app.js', import.meta.url), 'utf8');
assert.match(app, /Skills\.prepareForSend/);
assert.match(app, /Skills\.attachToSystem/);
assert.match(app, /userMsg\._skills=skillPack\.used\.map\(function\(u\)\{return \{id:u\.id,name:u\.name,reason:u\.reason\};\}\)/);
assert.match(app, /content:m\._actual\|\|m\.content/);
assert.match(app, /engineType==='image'/);

console.log('skills tests ok');
