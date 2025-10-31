function initializeHelp(path) {
  const button = document.createElement('button');
  button.id = 'help-button';
  button.textContent = 'ヘルプ';
  document.body.appendChild(button);

  const overlay = document.createElement('div');
  overlay.id = 'help-overlay';
  const popup = document.createElement('div');
  popup.id = 'help-popup';
  const content = document.createElement('div');
  content.id = 'help-content';
  const closeBtn = document.createElement('button');
  closeBtn.id = 'help-close';
  closeBtn.textContent = '閉じる';
  popup.appendChild(content);
  popup.appendChild(closeBtn);
  overlay.appendChild(popup);
  document.body.appendChild(overlay);

  const escapeHTML = str =>
    str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

  const renderContent = rawText => {
    const escaped = escapeHTML(rawText);
    const withBold = escaped.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    return withBold.replace(/\n/g, '<br>');
  };

  button.addEventListener('click', async () => {
    try {
      const res = await fetch(path);
      const text = await res.text();
      content.innerHTML = renderContent(text);
    } catch (e) {
      content.textContent = 'ヘルプを読み込めませんでした。';
    }
    overlay.style.display = 'flex';
  });

  function hide() {
    overlay.style.display = 'none';
  }

  closeBtn.addEventListener('click', hide);
  overlay.addEventListener('click', e => {
    if (e.target === overlay) hide();
  });
}
