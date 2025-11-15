(function setupRouter() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const SPA_STATE_KEY = '__yoPayrollSpa__';
  const viewTemplates = {
    home: 'view-home',
    sheets: 'view-sheets',
    payroll: 'view-payroll',
    settings: 'view-settings'
  };

  const viewRoot = document.getElementById('view-root');
  const headerLeftSlot = document.getElementById('header-left-slot');
  const headerRightSlot = document.getElementById('header-right-slot');
  const versionLabel = document.getElementById('version');

  if (!viewRoot) {
    console.warn('View root not found; router initialization skipped.');
    return;
  }

  const clearChildren = node => {
    if (!node) {
      return;
    }
    while (node.firstChild) {
      node.removeChild(node.firstChild);
    }
  };

  const createHeaderButton = config => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = config.className || 'back-btn';
    button.textContent = config.text || '';
    if (config.id) {
      button.id = config.id;
    }
    button.addEventListener('click', event => {
      event.preventDefault();
      if (typeof config.onClick === 'function') {
        config.onClick(event);
      }
    });
    return button;
  };

  const headerConfigs = {
    home: () => ({
      showVersion: true,
      leftButtons: [],
      rightButtons: [
        {
          id: 'settings',
          text: '設定',
          className: 'back-btn',
          onClick: () => navigateTo('settings')
        }
      ]
    }),
    sheets: () => ({
      showVersion: false,
      leftButtons: [
        {
          id: 'sheets-back',
          text: '＜店舗選択',
          className: 'back-btn',
          onClick: () => navigateBack()
        }
      ],
      rightButtons: [
        {
          id: 'sheets-home',
          text: 'はじめから',
          className: 'back-btn',
          onClick: () => navigateTo('home')
        }
      ]
    }),
    payroll: () => ({
      showVersion: false,
      leftButtons: [
        {
          id: 'payroll-back',
          text: '＜月選択',
          className: 'back-btn',
          onClick: () => navigateBack()
        }
      ],
      rightButtons: [
        {
          id: 'payroll-home',
          text: 'はじめから',
          className: 'back-btn',
          onClick: () => navigateTo('home')
        }
      ]
    }),
    settings: () => ({
      showVersion: false,
      leftButtons: [
        {
          id: 'settings-back',
          text: '＜店舗選択',
          className: 'back-btn',
          onClick: () => navigateBack()
        }
      ],
      rightButtons: [
        {
          id: 'settings-home',
          text: 'はじめから',
          className: 'back-btn',
          onClick: () => navigateTo('home')
        }
      ]
    })
  };

  const sanitizeView = view => (viewTemplates[view] ? view : 'home');

  const applyHeader = (viewName, detail) => {
    const configFactory = headerConfigs[viewName] || headerConfigs.home;
    const config = typeof configFactory === 'function' ? configFactory(detail) : configFactory;

    clearChildren(headerLeftSlot);
    clearChildren(headerRightSlot);

    if (headerLeftSlot) {
      (config.leftButtons || []).forEach(buttonConfig => {
        headerLeftSlot.appendChild(createHeaderButton(buttonConfig));
      });
    }

    if (headerRightSlot) {
      (config.rightButtons || []).forEach(buttonConfig => {
        headerRightSlot.appendChild(createHeaderButton(buttonConfig));
      });
    }

    if (versionLabel) {
      if (config.showVersion === false) {
        versionLabel.style.visibility = 'hidden';
      } else {
        versionLabel.style.visibility = '';
      }
    }
  };

  const createUrl = (view, params) => {
    const search = new URLSearchParams();
    search.set('view', view);
    Object.keys(params || {}).forEach(key => {
      const value = params[key];
      if (value !== undefined && value !== null) {
        search.set(key, String(value));
      }
    });
    const query = search.toString();
    const base = `${window.location.pathname}`;
    const hash = window.location.hash || '';
    return query ? `${base}?${query}${hash}` : `${base}${hash}`;
  };

  const createState = (view, params, index) => ({
    [SPA_STATE_KEY]: true,
    view,
    params,
    index
  });

  const parseStateFromLocation = () => {
    const search = new URLSearchParams(window.location.search);
    const rawView = search.get('view');
    const view = sanitizeView(rawView || 'home');
    const params = {};
    search.forEach((value, key) => {
      if (key !== 'view') {
        params[key] = value;
      }
    });
    return createState(view, params, 0);
  };

  let stateCounter = 0;
  let activeView = null;

  const activateState = (state, options = {}) => {
    const viewName = sanitizeView(state.view);
    const templateId = viewTemplates[viewName];
    const template = document.getElementById(templateId);
    if (!template) {
      console.warn(`Template for view "${viewName}" is missing.`);
      return;
    }

    const previous = activeView;
    if (previous) {
      const destroyDetail = {
        view: previous.name,
        params: { ...previous.params },
        element: previous.element
      };
      document.dispatchEvent(new CustomEvent('yo:view:destroy', { detail: destroyDetail }));
    }

    viewRoot.textContent = '';
    const fragment = template.content.cloneNode(true);
    viewRoot.appendChild(fragment);
    const insertedElement = viewRoot.firstElementChild || null;

    const params = {};
    Object.entries(state.params || {}).forEach(([key, value]) => {
      if (value !== undefined && value !== null) {
        params[key] = String(value);
      }
    });

    const searchParams = new URLSearchParams();
    Object.entries(params).forEach(([key, value]) => {
      searchParams.set(key, value);
    });

    const detail = {
      view: viewName,
      element: insertedElement,
      params,
      searchParams,
      navigationType: options.navigationType || 'push',
      previousView: previous ? previous.name : null
    };

    document.body.setAttribute('data-view', viewName);
    applyHeader(viewName, detail);
    document.dispatchEvent(new CustomEvent('yo:view:init', { detail }));
    activeView = {
      name: viewName,
      element: insertedElement,
      params
    };

    if (typeof requestAnimationFrame === 'function' && typeof viewRoot.focus === 'function') {
      requestAnimationFrame(() => {
        try {
          viewRoot.focus({ preventScroll: true });
        } catch (error) {
          try {
            viewRoot.focus();
          } catch (focusError) {
            // Ignore focus failures.
          }
        }
      });
    }
  };

  const navigateTo = (view, params = {}, options = {}) => {
    const normalizedView = sanitizeView(view);
    const normalizedParams = {};
    Object.keys(params || {}).forEach(key => {
      const value = params[key];
      if (value !== undefined && value !== null) {
        normalizedParams[key] = String(value);
      }
    });

    const index = options && options.replace ? stateCounter : stateCounter + 1;
    const state = createState(normalizedView, normalizedParams, index);
    const url = createUrl(normalizedView, normalizedParams);

    if (options && options.replace) {
      if (typeof history.replaceState === 'function') {
        history.replaceState(state, '', url);
      }
      stateCounter = index;
      activateState(state, { navigationType: 'replace' });
      return;
    }

    if (typeof history.pushState === 'function') {
      history.pushState(state, '', url);
    }
    stateCounter = index;
    activateState(state, { navigationType: 'push' });
  };

  const navigateBack = () => {
    if (stateCounter > 0 && typeof history.back === 'function') {
      history.back();
      return;
    }
    navigateTo('home', {}, { replace: true });
  };

  window.navigateTo = navigateTo;
  window.navigateBack = navigateBack;

  window.addEventListener('popstate', event => {
    const state = event.state && event.state[SPA_STATE_KEY]
      ? event.state
      : parseStateFromLocation();
    stateCounter = typeof state.index === 'number' ? state.index : 0;
    activateState(state, { navigationType: 'pop' });
  });

  const initialState = (() => {
    const parsed = parseStateFromLocation();
    const normalizedView = sanitizeView(parsed.view);
    const normalizedParams = {};
    Object.keys(parsed.params || {}).forEach(key => {
      const value = parsed.params[key];
      if (value !== undefined && value !== null) {
        normalizedParams[key] = String(value);
      }
    });
    const initial = createState(normalizedView, normalizedParams, 0);
    const url = createUrl(normalizedView, normalizedParams);
    if (typeof history.replaceState === 'function') {
      history.replaceState(initial, '', url);
    }
    stateCounter = 0;
    return initial;
  })();

  activateState(initialState, { navigationType: 'replace' });
})();
