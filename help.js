const TUTORIAL_ACTIVE_KEY = 'tutorial-active';
const TUTORIAL_PROMPT_KEY = 'tutorial-prompt-dismissed';
const TUTORIAL_PADDING = 12;
const TUTORIAL_RETRY_LIMIT = 15;

function readStorage(key) {
  try {
    return window.localStorage.getItem(key);
  } catch (error) {
    console.warn('Failed to access localStorage', error);
    return null;
  }
}

function writeStorage(key, value) {
  try {
    window.localStorage.setItem(key, value);
  } catch (error) {
    console.warn('Failed to write localStorage', error);
  }
}

function removeStorage(key) {
  try {
    window.localStorage.removeItem(key);
  } catch (error) {
    console.warn('Failed to remove localStorage value', error);
  }
}

function escapeHTML(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function parseTutorialSteps(rawText) {
  if (typeof rawText !== 'string' || rawText.trim() === '') {
    return [];
  }
  const lines = rawText.split(/\r?\n/);
  const steps = [];
  let current = null;

  lines.forEach(line => {
    const header = line.match(/^\s*\[([^\]]+)](.*)$/);
    if (header) {
      if (current) {
        steps.push(current);
      }
      const id = header[1].trim();
      const rest = header[2] || '';
      current = { id, lines: [] };
      if (rest.trim() !== '') {
        current.lines.push(rest.trim());
      }
      return;
    }
    if (!current) {
      return;
    }
    current.lines.push(line);
  });

  if (current) {
    steps.push(current);
  }

  return steps
    .map(step => {
      const joined = step.lines.join('\n').replace(/\(br\)/g, '\n');
      const normalized = joined
        .split('\n')
        .map(part => part.trim())
        .join('\n')
        .trim();
      return { id: step.id, text: normalized };
    })
    .filter(step => step.id && step.text !== undefined);
}

function formatStepText(text) {
  if (!text) {
    return '';
  }
  const escaped = escapeHTML(text);
  return escaped.replace(/\n/g, '<br>');
}

function createHelpButton() {
  const button = document.createElement('button');
  button.id = 'help-button';
  button.type = 'button';
  button.textContent = 'ヘルプ';
  document.body.appendChild(button);
  return button;
}

function createTutorialOverlay() {
  const overlay = document.createElement('div');
  overlay.id = 'tutorial-overlay';
  overlay.setAttribute('aria-hidden', 'true');

  const blocker = document.createElement('div');
  blocker.id = 'tutorial-overlay-blocker';

  const highlight = document.createElement('div');
  highlight.id = 'tutorial-highlight';

  const bubble = document.createElement('div');
  bubble.id = 'tutorial-bubble';

  const bubbleText = document.createElement('div');
  bubbleText.id = 'tutorial-bubble-text';

  const controls = document.createElement('div');
  controls.id = 'tutorial-controls';

  const prevBtn = document.createElement('button');
  prevBtn.id = 'tutorial-prev';
  prevBtn.type = 'button';
  prevBtn.textContent = '戻る';

  const nextBtn = document.createElement('button');
  nextBtn.id = 'tutorial-next';
  nextBtn.type = 'button';
  nextBtn.textContent = '進む';

  controls.appendChild(prevBtn);
  controls.appendChild(nextBtn);

  bubble.appendChild(bubbleText);
  bubble.appendChild(controls);

  overlay.appendChild(blocker);
  overlay.appendChild(highlight);
  overlay.appendChild(bubble);

  document.body.appendChild(overlay);

  let visible = false;

  function show() {
    overlay.classList.add('is-visible');
    overlay.setAttribute('aria-hidden', 'false');
    visible = true;
  }

  function hide() {
    overlay.classList.remove('is-visible');
    overlay.setAttribute('aria-hidden', 'true');
    visible = false;
  }

  function setHighlightRect(rect) {
    if (!rect) {
      highlight.style.display = 'none';
      return;
    }
    highlight.style.display = 'block';
    highlight.style.width = `${rect.width + TUTORIAL_PADDING * 2}px`;
    highlight.style.height = `${rect.height + TUTORIAL_PADDING * 2}px`;
    const top = Math.max(0, rect.top - TUTORIAL_PADDING);
    const left = Math.max(0, rect.left - TUTORIAL_PADDING);
    highlight.style.top = `${top}px`;
    highlight.style.left = `${left}px`;
  }

  function placeBubble(rect) {
    bubble.classList.remove('tutorial-bubble--center');
    bubble.classList.remove('tutorial-bubble--above');
    bubble.classList.remove('tutorial-bubble--below');
    bubble.style.visibility = 'hidden';
    requestAnimationFrame(() => {
      const margin = 16;
      const bubbleWidth = bubble.offsetWidth;
      const bubbleHeight = bubble.offsetHeight;
      let top = rect.bottom + margin;
      let positionClass = 'tutorial-bubble--below';
      if (top + bubbleHeight > window.innerHeight - margin) {
        top = Math.max(margin, rect.top - bubbleHeight - margin);
        positionClass = 'tutorial-bubble--above';
      }
      let left = rect.left + rect.width / 2 - bubbleWidth / 2;
      left = Math.max(margin, Math.min(left, window.innerWidth - bubbleWidth - margin));
      bubble.style.top = `${top}px`;
      bubble.style.left = `${left}px`;
      bubble.classList.add(positionClass);
      const arrowCenter = rect.left + rect.width / 2 - left;
      const clampedArrow = Math.max(24, Math.min(arrowCenter, bubbleWidth - 24));
      bubble.style.setProperty('--arrow-left', `${clampedArrow}px`);
      bubble.style.visibility = 'visible';
    });
  }

  function centerBubble() {
    highlight.style.display = 'none';
    bubble.classList.remove('tutorial-bubble--above');
    bubble.classList.remove('tutorial-bubble--below');
    bubble.classList.add('tutorial-bubble--center');
    bubble.style.visibility = 'hidden';
    requestAnimationFrame(() => {
      const margin = 16;
      const bubbleWidth = bubble.offsetWidth;
      const bubbleHeight = bubble.offsetHeight;
      const top = Math.max(margin, (window.innerHeight - bubbleHeight) / 2);
      const left = Math.max(margin, (window.innerWidth - bubbleWidth) / 2);
      bubble.style.top = `${top}px`;
      bubble.style.left = `${left}px`;
      bubble.style.removeProperty('--arrow-left');
      bubble.style.visibility = 'visible';
    });
  }

  return {
    container: overlay,
    blocker,
    highlight,
    bubble,
    bubbleText,
    prevBtn,
    nextBtn,
    show,
    hide,
    isVisible: () => visible,
    setHighlightRect,
    placeBubble,
    centerBubble
  };
}

function createModalOverlay() {
  const overlay = document.createElement('div');
  overlay.className = 'tutorial-modal-overlay';
  const modal = document.createElement('div');
  modal.className = 'tutorial-modal';
  overlay.appendChild(modal);
  document.body.appendChild(overlay);
  return { overlay, modal };
}

function isTutorialActive() {
  return readStorage(TUTORIAL_ACTIVE_KEY) === '1';
}

function setTutorialActive(flag) {
  if (flag) {
    writeStorage(TUTORIAL_ACTIVE_KEY, '1');
  } else {
    removeStorage(TUTORIAL_ACTIVE_KEY);
  }
}

function isPromptDismissed() {
  return readStorage(TUTORIAL_PROMPT_KEY) === '1';
}

function setPromptDismissed() {
  writeStorage(TUTORIAL_PROMPT_KEY, '1');
}

function initializeHelp(path, options = {}) {
  if (!path || typeof document === 'undefined' || !document.body) {
    return;
  }

  const stepConfigs = options.steps || {};
  const helpButton = createHelpButton();
  const overlay = createTutorialOverlay();

  let steps = [];
  let stepsReady = false;
  let loadingPromise = null;
  let pendingStart = false;
  let currentIndex = 0;
  let currentStep = null;
  let currentTarget = null;
  let currentConfig = null;
  let retryTimer = null;
  let listenersAttached = false;

  const resolveStepConfig = id => {
    const config = stepConfigs[id];
    if (!config) {
      return null;
    }
    if (typeof config === 'string') {
      return { selector: config };
    }
    if (typeof config === 'function') {
      return { getElement: config };
    }
    return config;
  };

  const resolveStepElement = id => {
    const config = resolveStepConfig(id);
    if (!config) {
      return null;
    }
    if (typeof config.getElement === 'function') {
      try {
        const result = config.getElement();
        if (result instanceof Element) {
          return result;
        }
      } catch (error) {
        console.warn('Failed to resolve tutorial target via getElement', error);
      }
    }
    if (config.selector) {
      try {
        const element = document.querySelector(config.selector);
        if (element instanceof Element) {
          return element;
        }
      } catch (error) {
        console.warn('Invalid tutorial selector', config.selector, error);
      }
    }
    if (Array.isArray(config.selectors)) {
      for (const selector of config.selectors) {
        try {
          const element = document.querySelector(selector);
          if (element instanceof Element) {
            return element;
          }
        } catch (error) {
          console.warn('Invalid tutorial selector', selector, error);
        }
      }
    }
    return null;
  };

  const cleanupRetryTimer = () => {
    if (retryTimer !== null) {
      window.clearTimeout(retryTimer);
      retryTimer = null;
    }
  };

  const cleanupCurrentTarget = () => {
    if (currentTarget) {
      currentTarget.classList.remove('tutorial-highlight-target');
      currentTarget = null;
    }
  };

  const reposition = () => {
    if (!overlay.isVisible()) {
      return;
    }
    if (currentTarget) {
      const rect = currentTarget.getBoundingClientRect();
      overlay.setHighlightRect(rect);
      overlay.placeBubble(rect);
    } else {
      overlay.centerBubble();
    }
  };

  const attachListeners = () => {
    if (listenersAttached) {
      return;
    }
    listenersAttached = true;
    window.addEventListener('resize', reposition);
    window.addEventListener('scroll', reposition, true);
    window.addEventListener('keydown', handleKeyDown, true);
  };

  const detachListeners = () => {
    if (!listenersAttached) {
      return;
    }
    listenersAttached = false;
    window.removeEventListener('resize', reposition);
    window.removeEventListener('scroll', reposition, true);
    window.removeEventListener('keydown', handleKeyDown, true);
  };

  const handleKeyDown = event => {
    if (event.key === 'Escape') {
      event.preventDefault();
      finishTutorial();
    }
  };

  const highlightElement = element => {
    if (!(element instanceof Element)) {
      overlay.centerBubble();
      return;
    }
    if (currentTarget !== element) {
      cleanupCurrentTarget();
      currentTarget = element;
      currentTarget.classList.add('tutorial-highlight-target');
    }
    const rect = element.getBoundingClientRect();
    if (rect.top < 0 || rect.bottom > window.innerHeight || rect.left < 0 || rect.right > window.innerWidth) {
      element.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
      setTimeout(reposition, 350);
    }
    overlay.setHighlightRect(element.getBoundingClientRect());
    overlay.placeBubble(element.getBoundingClientRect());
  };

  const attemptResolveTarget = (stepId, retryCount = 0) => {
    cleanupRetryTimer();
    const step = currentStep;
    if (!step || step.id !== stepId) {
      return;
    }
    const target = resolveStepElement(stepId);
    if (target) {
      highlightElement(target);
      return;
    }
    if (retryCount >= TUTORIAL_RETRY_LIMIT) {
      cleanupCurrentTarget();
      overlay.centerBubble();
      return;
    }
    retryTimer = window.setTimeout(() => {
      attemptResolveTarget(stepId, retryCount + 1);
    }, 250);
  };

  const showStep = index => {
    if (!Array.isArray(steps) || steps.length === 0) {
      return;
    }
    const nextIndex = Math.min(Math.max(index, 0), steps.length - 1);
    if (currentConfig && typeof currentConfig.onExit === 'function') {
      try {
        currentConfig.onExit();
      } catch (error) {
        console.warn('tutorial step onExit failed', error);
      }
    }
    cleanupRetryTimer();
    cleanupCurrentTarget();
    currentConfig = null;

    currentIndex = nextIndex;
    currentStep = steps[currentIndex];
    overlay.bubbleText.innerHTML = formatStepText(currentStep.text);
    overlay.prevBtn.disabled = currentIndex === 0;
    overlay.nextBtn.textContent = currentIndex === steps.length - 1 ? '終了' : '進む';

    currentConfig = resolveStepConfig(currentStep.id);
    if (currentConfig && typeof currentConfig.onEnter === 'function') {
      try {
        currentConfig.onEnter();
      } catch (error) {
        console.warn('tutorial step onEnter failed', error);
      }
    }

    requestAnimationFrame(() => {
      attemptResolveTarget(currentStep.id);
      reposition();
    });
  };

  const finishTutorial = () => {
    pendingStart = false;
    setTutorialActive(false);
    cleanupRetryTimer();
    if (currentConfig && typeof currentConfig.onExit === 'function') {
      try {
        currentConfig.onExit();
      } catch (error) {
        console.warn('tutorial step onExit failed', error);
      }
    }
    currentConfig = null;
    cleanupCurrentTarget();
    overlay.hide();
    detachListeners();
  };

  const startTutorialInternal = () => {
    if (!stepsReady) {
      pendingStart = true;
      loadSteps();
      return;
    }
    if (!Array.isArray(steps) || steps.length === 0) {
      pendingStart = false;
      console.warn('No tutorial steps were loaded for', path);
      return;
    }
    pendingStart = false;
    setTutorialActive(true);
    overlay.show();
    attachListeners();
    showStep(0);
  };

  const loadSteps = () => {
    if (!loadingPromise) {
      loadingPromise = fetch(path)
        .then(response => response.text())
        .then(text => {
          steps = parseTutorialSteps(text);
        })
        .catch(error => {
          console.error('Failed to load tutorial content', error);
          steps = [];
        })
        .finally(() => {
          stepsReady = true;
          if (pendingStart || isTutorialActive()) {
            startTutorialInternal();
          }
        });
    }
    return loadingPromise;
  };

  const startTutorial = () => {
    startTutorialInternal();
  };

  overlay.prevBtn.addEventListener('click', () => {
    if (currentIndex > 0) {
      showStep(currentIndex - 1);
    }
  });

  overlay.nextBtn.addEventListener('click', () => {
    if (!Array.isArray(steps) || steps.length === 0) {
      finishTutorial();
      return;
    }
    if (currentIndex >= steps.length - 1) {
      finishTutorial();
    } else {
      showStep(currentIndex + 1);
    }
  });

  helpButton.addEventListener('click', () => {
    startTutorial();
  });

  const showReminder = () => {
    const { overlay: reminderOverlay, modal } = createModalOverlay();
    const message = document.createElement('p');
    message.className = 'tutorial-modal-message';
    message.textContent = '右下のヘルプボタンからいつでも確認できます。';
    const actions = document.createElement('div');
    actions.className = 'tutorial-modal-actions';
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = '閉じる';
    actions.appendChild(closeBtn);
    modal.appendChild(message);
    modal.appendChild(actions);

    const cleanup = () => {
      reminderOverlay.remove();
      window.removeEventListener('keydown', onKeyDown, true);
    };

    const onKeyDown = event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        cleanup();
      }
    };

    closeBtn.addEventListener('click', cleanup);
    reminderOverlay.addEventListener('click', event => {
      if (event.target === reminderOverlay) {
        cleanup();
      }
    });
    window.addEventListener('keydown', onKeyDown, true);
    setTimeout(() => closeBtn.focus(), 0);
  };

  const showTutorialPrompt = () => {
    const { overlay: promptOverlay, modal } = createModalOverlay();

    const message = document.createElement('p');
    message.className = 'tutorial-modal-message';
    message.textContent = 'チュートリアルを行いますか？';

    const checkboxLabel = document.createElement('label');
    checkboxLabel.className = 'tutorial-modal-checkbox';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'tutorial-prompt-checkbox';
    const checkboxText = document.createElement('span');
    checkboxText.textContent = '再度表示しない';
    checkboxLabel.appendChild(checkbox);
    checkboxLabel.appendChild(checkboxText);

    const actions = document.createElement('div');
    actions.className = 'tutorial-modal-actions';

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = '閉じる';

    const startBtn = document.createElement('button');
    startBtn.type = 'button';
    startBtn.className = 'tutorial-modal-primary';
    startBtn.textContent = '始める';

    actions.appendChild(closeBtn);
    actions.appendChild(startBtn);

    modal.appendChild(message);
    modal.appendChild(checkboxLabel);
    modal.appendChild(actions);

    const cleanup = () => {
      promptOverlay.remove();
      window.removeEventListener('keydown', onKeyDown, true);
    };

    const onKeyDown = event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        handleClose();
      }
    };

    const handleClose = () => {
      if (checkbox.checked) {
        setPromptDismissed();
      }
      cleanup();
      showReminder();
    };

    const handleStart = () => {
      if (checkbox.checked) {
        setPromptDismissed();
      }
      cleanup();
      startTutorial();
    };

    closeBtn.addEventListener('click', handleClose);
    startBtn.addEventListener('click', handleStart);
    promptOverlay.addEventListener('click', event => {
      if (event.target === promptOverlay) {
        handleClose();
      }
    });
    window.addEventListener('keydown', onKeyDown, true);
    setTimeout(() => startBtn.focus(), 0);
  };

  const maybeShowPrompt = () => {
    if (isTutorialActive()) {
      startTutorial();
      return;
    }
    if (!isPromptDismissed()) {
      showTutorialPrompt();
    }
  };

  loadSteps();
  window.setTimeout(maybeShowPrompt, 200);
}
