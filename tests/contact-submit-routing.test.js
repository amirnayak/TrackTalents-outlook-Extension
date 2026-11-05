const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync(require.resolve('../public/taskpane.js'), 'utf8');
function extract(name) {
  const start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `${name} exists`);
  const end = source.indexOf('\nfunction ', start + 1);
  return source.slice(start, end < 0 ? undefined : end);
}
function route(actionId, attachments, authenticated = true) {
  const calls = [];
  const context = vm.createContext({
    state: { auth: authenticated ? { accessToken: 'test' } : null, currentItem: { attachments } },
    readCurrentItem: () => calls.push('read-item'),
    isAttachEmailAction: (id) => id === 'attach-email',
    isTokenExpired: () => false,
    openAttachEmailPanel: () => calls.push('link-email'),
    openImportModal: (id) => calls.push(['attachment-selector', id]),
    reserveActionWindow: () => { calls.push('reserve-window'); return 'window'; },
    handleDirectActionLaunch: (id, window) => calls.push(['direct', id, window])
  });
  vm.runInContext(extract('isResumeImportAction') + '\n' + extract('handleActionRequest'), context);
  context.handleActionRequest(actionId);
  return calls;
}
for (const attachments of [[], [{ name: 'resume.pdf' }], [{ name: 'resume.pdf' }, { name: 'resume.doc' }]]) {
  test(`contact submission opens ATS directly with ${attachments.length} attachments`, () => {
    assert.deepEqual(route('submit-resume-contact', attachments), [
      'read-item', 'reserve-window', ['direct', 'submit-resume-contact', 'window']
    ]);
  });
}
test('contact submission retains login handling', () => {
  assert.deepEqual(route('submit-resume-contact', [], false), [
    'read-item', ['direct', 'submit-resume-contact', null]
  ]);
});
for (const action of ['add-candidate', 'source-resume-job']) {
  test(`${action} retains attachment selection`, () => {
    assert.deepEqual(route(action, [{ name: 'resume.doc' }]), ['read-item', ['attachment-selector', action]]);
  });
}
for (const action of ['add-contact', 'add-job']) {
  test(`${action} retains direct launch`, () => {
    assert.deepEqual(route(action, []), ['read-item', 'reserve-window', ['direct', action, 'window']]);
  });
}
test('link email retains its panel', () => {
  assert.deepEqual(route('attach-email', []), ['read-item', 'link-email']);
});
