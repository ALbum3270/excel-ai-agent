(() => {
  const messages = document.getElementById('messages');
  const emptyState = document.getElementById('emptyState');
  const input = document.getElementById('promptInput');
  const composer = document.getElementById('composer');
  const sendButton = document.getElementById('sendButton');
  const cancelButton = document.getElementById('cancelButton');
  const settingsButton = document.getElementById('settingsButton');
  const statusText = document.getElementById('statusText');
  const statusDot = document.getElementById('statusDot');
  const activity = document.getElementById('activity');
  const activityTitle = document.getElementById('activityTitle');
  const activityCount = document.getElementById('activityCount');
  const activitySteps = document.getElementById('activitySteps');
  const approval = document.getElementById('approval');
  const approvalMessage = document.getElementById('approvalMessage');
  const modeOptions = Array.from(document.querySelectorAll('.mode-option'));
  let mode = 'execute';

  const webview = window.chrome && window.chrome.webview;
  const post = (payload) => { if (webview) webview.postMessage(payload); };

  function setBusy(value) {
    sendButton.hidden = value;
    cancelButton.hidden = !value;
    input.disabled = value;
    modeOptions.forEach((option) => { option.disabled = value; });
    statusDot.classList.toggle('busy', value);
    if (!value && statusText.textContent !== '失败') statusText.textContent = '就绪';
  }

  function appendMessage(role, content, success) {
    emptyState.hidden = true;
    const node = document.createElement('div');
    node.className = `message ${role}`;
    if (role === 'assistant' && success === false) node.classList.add('failed');
    node.textContent = content || '';
    messages.appendChild(node);
    messages.scrollTop = messages.scrollHeight;
  }

  function setPlan(payload) {
    activity.hidden = false;
    activityTitle.textContent = payload.understanding || '执行计划';
    activitySteps.replaceChildren();
    (payload.steps || []).forEach((step) => {
      const item = document.createElement('li');
      item.dataset.index = String(step.index);
      item.textContent = step.description || `步骤 ${step.index}`;
      activitySteps.appendChild(item);
    });
    activityCount.textContent = `${(payload.steps || []).length} 步`;
  }

  function updateStep(payload) {
    activity.hidden = false;
    let item = activitySteps.querySelector(`[data-index="${payload.index}"]`);
    if (!item) {
      item = document.createElement('li');
      item.dataset.index = String(payload.index);
      activitySteps.appendChild(item);
    }
    if (payload.message) item.textContent = payload.message;
    item.classList.toggle('done', payload.status === 'completed');
    item.classList.toggle('failed', payload.status === 'failed');
  }

  function receive(payload) {
    switch (payload.type) {
      case 'appendMessage':
        appendMessage(payload.role, payload.content, payload.success);
        if (payload.role === 'error') {
          statusText.textContent = '失败';
          statusDot.classList.add('error');
        }
        break;
      case 'busy':
        setBusy(Boolean(payload.value));
        if (!payload.value) statusDot.classList.remove('error');
        break;
      case 'phase':
        statusText.textContent = payload.message || payload.phase || '执行中';
        break;
      case 'plan':
        setPlan(payload);
        break;
      case 'step':
        updateStep(payload);
        break;
      case 'iteration':
        activity.hidden = false;
        activityCount.textContent = `第 ${payload.index} 轮`;
        break;
      case 'approval':
        approval.hidden = !payload.visible;
        approvalMessage.textContent = payload.message || '';
        break;
      case 'focusInput':
        input.focus();
        break;
    }
  }

  if (webview) webview.addEventListener('message', (event) => receive(event.data || {}));

  composer.addEventListener('submit', (event) => {
    event.preventDefault();
    const text = input.value.trim();
    if (!text) return;
    input.value = '';
    post({ type: 'send', text, mode });
  });
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      composer.requestSubmit();
    }
  });
  cancelButton.addEventListener('click', () => post({ type: 'cancel' }));
  settingsButton.addEventListener('click', () => post({ type: 'settings' }));
  modeOptions.forEach((option) => {
    option.addEventListener('click', () => {
      mode = option.dataset.mode || 'read_only';
      modeOptions.forEach((item) => {
        const selected = item === option;
        item.classList.toggle('selected', selected);
        item.setAttribute('aria-checked', String(selected));
      });
    });
  });
  document.getElementById('approveButton').addEventListener('click', () => post({ type: 'approve' }));
  document.getElementById('rejectButton').addEventListener('click', () => post({ type: 'reject' }));

})();
