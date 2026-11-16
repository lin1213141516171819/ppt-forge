window.__ModuleLoader__.load({
  id: '@local/ppt-forge',
  factory(require) {
    const React = require('react');
    const h = React.createElement;
    const { useEffect, useState } = React;

    const THEME_SWATCHES = [
      { id: 'navy', label: '海军蓝', color: '#1D4ED8' },
      { id: 'teal', label: '松石绿', color: '#0F766E' },
      { id: 'wine', label: '酒红', color: '#9F1239' },
      { id: 'graphite', label: '石墨', color: '#3F3F46' },
    ];

    const styles = {
      page: {
        display: 'flex',
        flexDirection: 'column',
        gap: '16px',
        maxWidth: '780px',
        padding: '4px 0 28px',
        color: 'var(--dsw-alias-label-primary)',
      },
      intro: {
        margin: 0,
        fontSize: '14px',
        lineHeight: 1.7,
        color: 'var(--dsw-alias-label-secondary)',
      },
      card: {
        border: '1px solid var(--dsw-alias-border-l1)',
        borderRadius: '10px',
        background: 'var(--dsw-alias-bg-layer-1)',
        padding: '14px 18px',
        display: 'flex',
        flexDirection: 'column',
        gap: '10px',
      },
      row: {
        display: 'flex',
        alignItems: 'center',
        gap: '12px',
        fontSize: '13px',
      },
      rowLabel: {
        width: '110px',
        flexShrink: 0,
        color: 'var(--dsw-alias-label-secondary)',
      },
      sectionTitle: {
        margin: '10px 0 0',
        fontSize: '15px',
        fontWeight: 600,
      },
      steps: {
        margin: 0,
        paddingLeft: '20px',
        display: 'flex',
        flexDirection: 'column',
        gap: '8px',
        fontSize: '14px',
        lineHeight: 1.7,
        color: 'var(--dsw-alias-label-secondary)',
      },
      hint: {
        margin: 0,
        fontSize: '12px',
        color: 'var(--dsw-alias-label-secondary)',
        opacity: 0.85,
      },
      swatchRow: {
        display: 'flex',
        flexWrap: 'wrap',
        gap: '10px',
      },
      swatch: {
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        fontSize: '13px',
        color: 'var(--dsw-alias-label-secondary)',
      },
      dot: {
        width: '14px',
        height: '14px',
        borderRadius: '4px',
        flexShrink: 0,
      },
      code: {
        fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Consolas, monospace',
        fontSize: '12px',
        padding: '1px 6px',
        borderRadius: '5px',
        border: '1px solid var(--dsw-alias-border-l1)',
        background: 'var(--dsw-alias-bg-layer-2)',
      },
    };

    function statusTone(ok) {
      return ok ? 'var(--dsw-alias-state-success-primary)' : 'var(--dsw-alias-state-error-primary)';
    }

    function PptForgeSettings() {
      const [status, setStatus] = useState(null);
      const [failed, setFailed] = useState(null);

      useEffect(() => {
        let cancelled = false;
        fetch('/ppt-forge/status')
          .then((response) => response.json())
          .then((data) => {
            if (!cancelled) setStatus(data);
          })
          .catch((error) => {
            if (!cancelled) setFailed(String(error && error.message ? error.message : error));
          });
        return () => {
          cancelled = true;
        };
      }, []);

      const ready = status && status.ok;
      let statusText = '正在检测运行环境…';
      if (failed) statusText = `检测失败：${failed}`;
      else if (status) {
        statusText = ready
          ? `Python ${status.python.version} · python-pptx ${status.python.pptxVersion}`
          : '未找到可用的 Python 运行时';
      }

      return h(
        'div',
        { style: styles.page },
        h(
          'p',
          { style: styles.intro },
          '把现有的 .pptx 交给它：读取每页的文字、表格与演讲备注，按你的指令重新组织，生成一份干净的 16:9 新演示文稿。旧文件的版式和图片不会被沿用。',
        ),
        h(
          'div',
          { style: styles.card },
          h(
            'div',
            { style: styles.row },
            h('span', { style: styles.rowLabel }, '运行环境'),
            h(
              'span',
              { style: { display: 'flex', alignItems: 'center', gap: '8px' } },
              h('span', {
                style: {
                  width: '8px',
                  height: '8px',
                  borderRadius: '50%',
                  background: ready ? statusTone(true) : statusTone(false),
                },
              }),
              h('span', null, statusText),
            ),
          ),
          h(
            'div',
            { style: styles.row },
            h('span', { style: styles.rowLabel }, '工具名'),
            h('code', { style: styles.code }, 'ppt_forge'),
          ),
          h(
            'div',
            { style: styles.row },
            h('span', { style: styles.rowLabel }, '可用主题'),
            h(
              'span',
              { style: styles.swatchRow },
              THEME_SWATCHES.map((theme) =>
                h(
                  'span',
                  { key: theme.id, style: styles.swatch },
                  h('span', { style: { ...styles.dot, background: theme.color } }),
                  theme.label,
                ),
              ),
            ),
          ),
        ),
        h('h3', { style: styles.sectionTitle }, '使用方法'),
        h(
          'ol',
          { style: styles.steps },
          h('li', null, '在对话里上传一个 .pptx 文件，或直接告诉我它的路径。'),
          h(
            'li',
            null,
            '说明你想怎么重组——例如「保留内容，按新结构重排成 10 页」「压缩成汇报提纲」「按章节整理成教学课件」。',
          ),
          h('li', null, '我会调用 ppt_forge 读取原文件内容，生成新的 .pptx 并交付给你；生成的文件保存在当前工作目录的 ppt-forge-output/ 下。'),
        ),
        h('h3', { style: styles.sectionTitle }, '多份样本学习'),
        h(
          'p',
          { style: styles.intro },
          '把几十份 .pptx 放进同一个文件夹告诉我路径，我会批量分析它们的结构习惯、版式、语言与配色，生成可复用的风格档案（存于 ppt-forge-library/，含 JSON 与可读 Markdown）。之后做新主题时按这份档案生成，风格与样本保持一致。',
        ),
        h(
          'p',
          { style: styles.hint },
          '提示：也可以在对话里直接说「用 ppt_forge 生成一份关于 XX 的 PPT」，我会先拟大纲、再生成。',
        ),
      );
    }

    return {
      inject: ['slots'],
      apply(ctx) {
        ctx.slots.inject('settings.section', () =>
          ctx.slots.register(
            {
              name: 'settings.section',
              id: 'ppt-forge',
              order: 60,
              label: 'PPT 工坊',
            },
            PptForgeSettings,
          ),
        );
      },
    };
  },
});
