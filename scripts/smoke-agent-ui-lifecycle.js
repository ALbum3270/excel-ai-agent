const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const elements = new Map();
const rendered = [];
const sentMessages = [];
let restoreCalls = 0;

const document = {
    getElementById(id) {
        return elements.get(id) || null;
    },
    querySelector() { return null; },
    createElement() {
        return {
            id: '', className: '', dataset: {}, style: {}, innerHTML: '',
            appendChild() {}, querySelector() { return null; }, querySelectorAll() { return []; }
        };
    }
};

const window = {
    setInterval() { return 1; },
    clearInterval() {},
    agentProgressTimers: {},
    agentCardState: { active: true, session: { uuid: 'run-1' }, locked: false }
};

function createChatSection(sender, timestamp, uuid) {
    const chat = { id: 'chat-' + uuid, dataset: {}, style: {}, innerHTML: '' };
    const content = { id: 'content-' + uuid, textContent: '', innerHTML: '' };
    elements.set(chat.id, chat);
    elements.set(content.id, content);
    return uuid;
}

function appendRenderer(uuid, message) {
    rendered.push({ uuid, message });
}

const context = vm.createContext({
    window,
    document,
    console,
    Date,
    String,
    Number,
    Math,
    createChatSection,
    appendRenderer,
    escapeHtml: text => String(text),
    restoreAgentRequestUi() { restoreCalls += 1; },
    sendMessageToServer(message) { sentMessages.push(message); }
});

const root = path.join(__dirname, '..');
const agentProtocolPath = path.join(root, 'ShareRibbon', 'Resources', 'js', 'agent-protocol.js');
const agentCardPath = path.join(root, 'ShareRibbon', 'Resources', 'js', 'agent-card.js');
vm.runInContext(fs.readFileSync(agentProtocolPath, 'utf8'), context, { filename: agentProtocolPath });
vm.runInContext(fs.readFileSync(agentCardPath, 'utf8'), context, { filename: agentCardPath });
context.restoreAgentRequestUi = () => { restoreCalls += 1; };
window.restoreAgentRequestUi = context.restoreAgentRequestUi;

window.completeAgent('run-1', true, '共有 4 位销售人员。', '');
assert.deepStrictEqual(
    rendered,
    [{ uuid: 'agent-final-run-1', message: '共有 4 位销售人员。' }],
    'a successful Agent run must render its final model answer as an independent AI message');

window.completeAgent('run-1', true, '共有 4 位销售人员。', '');
assert.strictEqual(rendered.length, 1, 'duplicate terminal events must not duplicate the answer bubble');

const restoresBeforeAbort = restoreCalls;
window.abortAgent('run-2');
assert.strictEqual(sentMessages.length, 1, 'the Agent card stop control must send one backend cancellation request');
assert.strictEqual(sentMessages[0].type, 'abortAgent', 'the stop control must use the typed Agent protocol');
assert.strictEqual(sentMessages[0].payload.sessionId, 'run-2', 'cancellation must be bound to the card session');
assert.strictEqual(restoreCalls, restoresBeforeAbort,
    'the request UI must remain locked until the backend publishes a terminal cancellation state');
window.completeAgent('run-2', false, 'cancelled', '', 'cancelled');
assert.strictEqual(restoreCalls, restoresBeforeAbort + 1,
    'the backend cancellation terminal state must unlock the request UI exactly once');
assert.strictEqual(rendered.length, 1, 'a cancelled run must not render a failed answer bubble');

const baseChat = fs.readFileSync(path.join(root, 'ShareRibbon', 'Controls', 'BaseChatControl.vb'), 'utf8');
assert.ok(
    /HandleStopMessage[\s\S]*?HttpStreamSvc\.CancelRequest\(requestUuid\)[\s\S]*?_agentKernelService[\s\S]*?AbortAgent\(\)/.test(baseChat),
    'the shared bottom stop button must cancel both the HTTP stream and the active Agent loop');
assert.ok(/Register\("abortAgent", Sub\(jsonDoc\) HandleAbortAgent\(jsonDoc\)\)/.test(baseChat) &&
    /HandleAbortAgent\(jsonDoc As JObject\)[\s\S]*?payload[\s\S]*?sessionId[\s\S]*?AbortAgent\(expectedSessionId\)/.test(baseChat),
    'the backend cancellation route must preserve the expected Agent session id');
assert.ok(/SendAndGetStreamingResponseAsync\([\s\S]*?cancellationToken As CancellationToken[\s\S]*?CreateCancellationSource\(cancellationToken\)/.test(baseChat),
    'the active Agent cancellation token must reach the provider HTTP request');

const agentService = fs.readFileSync(path.join(root, 'ShareRibbon', 'Controls', 'Services', 'AgentKernelService.vb'), 'utf8');
assert.ok(/_officeHarness\.RunAsync\(turn, requestCancellation\.Token\)/.test(agentService),
    'AgentKernelService must pass a request cancellation token into the harness');
assert.ok(!/If String\.IsNullOrWhiteSpace\(sessionId\) Then Return[\s\S]{0,500}CancelActiveAgentRequest\(\)/.test(agentService),
    'AbortAgent must cancel the live request even when UI/session bookkeeping is temporarily missing');
assert.ok(/CurrentHarnessRunId\s*=\s*e\.RunId/.test(agentService),
    'the live harness run id must be published before RunAsync returns');
assert.ok(/_sendAiRequest\([\s\S]{0,300}GetActiveAgentCancellationToken\(\)/.test(agentService),
    'AgentKernelService must forward the active request token to the model transport');
assert.ok(/CodeCancelled[\s\S]{0,300}"cancelled"/.test(agentService) &&
    /completeAgent\([\s\S]{0,300}resolvedStatus/.test(agentService),
    'cancellation must be returned to the card as a distinct terminal state');

const harness = fs.readFileSync(path.join(root, 'ShareRibbon', 'Agent', 'Harness', 'OfficeHarness.vb'), 'utf8');
assert.ok(/pending\.CancellationSource\?\.Cancel\(\)/.test(harness),
    'OfficeHarness.CancelAsync must cancel the live run token');
assert.ok(!/CANCEL_NOT_AT_SAFE_POINT/.test(harness),
    'cancellation must no longer be rejected merely because the run is between approval points');

const loop = fs.readFileSync(path.join(root, 'ShareRibbon', 'Agent', 'LoopEngine.vb'), 'utf8');
assert.ok(/AwaitWithCancellation\([\s\S]*?ThinkAsync/.test(loop),
    'the adaptive model wait must observe request cancellation');
assert.ok(/Office COM calls are not interrupted halfway through/.test(loop),
    'host mutations must stop at a documented safe boundary');

const python = fs.readFileSync(path.join(root, 'ShareRibbon', 'Services', 'Python', 'PythonComputeService.vb'), 'utf8');
assert.ok(/process\.Kill\(\)[\s\S]*?ThrowIfCancellationRequested/.test(python),
    'PythonCompute must terminate its child process when the Agent request is cancelled');

console.log('PASS: Agent final answers render independently and cancellation reaches the live adaptive loop');
