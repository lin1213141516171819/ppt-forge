import { spawn, spawnSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const PYTHON_DIR = join(HERE, 'python');
const ROUTE_PREFIX = '/ppt-forge';
const DEFAULT_THEMES = ['navy', 'teal', 'wine', 'graphite'];
const HISTORY_LIMIT = 200;

export const name = 'ppt-forge';
export const inject = ['tools', 'workingDirectory'];

/** Windows-safe timestamp for default output file names. */
function stamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function safeFileStem(text) {
  const cleaned = String(text || 'presentation')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 60);
  return cleaned.length > 0 ? cleaned : 'presentation';
}

/**
 * Locate a Python interpreter that can actually import python-pptx.
 * Order: explicit config, DSH_PYTHON, the bundled runtime next to the App,
 * the DSH home runtime, then PATH names.
 */
export function resolvePython(config = {}) {
  const candidates = [];
  const seen = new Set();
  const push = (value) => {
    if (!value || seen.has(value)) return;
    seen.add(value);
    candidates.push(value);
  };

  push(config.pythonPath);
  push(process.env.DSH_PYTHON);

  const execDir = (() => {
    try {
      return dirname(process.execPath);
    } catch {
      return undefined;
    }
  })();
  const homes = [];
  if (process.env.DSH_HOME) homes.push(process.env.DSH_HOME);
  if (execDir) homes.push(execDir);

  const executable = process.platform === 'win32' ? 'python.exe' : 'python3';
  for (const home of homes) {
    push(join(home, 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', executable));
    push(join(home, 'runtime', 'primary-runtime', 'dependencies', 'python', executable));
    push(join(home, 'resources', 'runtime', 'primary-runtime', 'dependencies', 'python', executable));
  }
  push('python');
  push('python3');
  push('py');

  for (const candidate of candidates) {
    const hasSeparator = candidate.includes('\\') || candidate.includes('/');
    if (hasSeparator && !existsSync(candidate)) continue;
    try {
      const probe = spawnSync(
        candidate,
        ['-c', 'import sys, pptx; print(sys.version.split()[0], pptx.__version__)'],
        {
          timeout: 20000,
          windowsHide: true,
          encoding: 'utf8',
        },
      );
      if (probe.status === 0) {
        const [version = '', pptxVersion = ''] = (probe.stdout || '').trim().split(/\s+/);
        return { command: candidate, version, pptxVersion };
      }
    } catch {
      // try the next candidate
    }
  }
  return undefined;
}

/** Run one bundled python script and parse its JSON stdout. */
function runPython(python, script, args, { stdin, signal } = {}) {
  const scriptPath = join(PYTHON_DIR, script);
  return new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(python, [scriptPath, ...args], {
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
      signal,
    });
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => (stdout += chunk));
    child.stderr.on('data', (chunk) => (stderr += chunk));
    child.stdin.on('error', (error) => rejectPromise(error));
    child.on('error', (error) => rejectPromise(error));
    child.on('close', (code) => {
      let parsed;
      try {
        parsed = JSON.parse(stdout);
      } catch {
        parsed = undefined;
      }
      if (parsed && typeof parsed === 'object' && typeof parsed.error === 'string') {
        const error = new Error(parsed.error);
        error.code = parsed.code;
        error.issues = parsed.issues;
        rejectPromise(error);
        return;
      }
      if (code !== 0) {
        rejectPromise(
          new Error(
            (stderr || stdout || `python exited with code ${code}`).slice(0, 2000).trim(),
          ),
        );
        return;
      }
      if (parsed === undefined) {
        rejectPromise(new Error(`python returned no JSON: ${(stdout || stderr).slice(0, 800)}`));
        return;
      }
      resolvePromise(parsed);
    });
    if (stdin !== undefined) {
      child.stdin.write(stdin);
      child.stdin.end();
    }
  });
}

/** Default project folder inside the session working directory. */
function defaultOutputPath(cwd, outline) {
  const stem = safeFileStem(outline?.title || outline?.subtitle || 'presentation');
  return join(cwd, 'ppt-forge-output', `${stem}-${stamp()}-${randomUUID().slice(0, 8)}`, 'deck.pptx');
}

/** List generated decks (only inside the requested working directory). */
function listHistory(cwd) {
  const historyDir = join(cwd, 'ppt-forge-output');
  if (!existsSync(historyDir)) return [];
  try {
    const files = [];
    const collect = (dir, prefix = '') => {
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = join(dir, entry.name);
        const relative = join(prefix, entry.name);
        // Do not follow directory symlinks outside the output tree.
        if (entry.isDirectory()) collect(full, relative);
        else if (entry.isFile() && entry.name.toLowerCase().endsWith('.pptx')) {
          const info = statSync(full);
          files.push({ name: relative, size: info.size, modified: info.mtimeMs });
        }
      }
    };
    collect(historyDir);
    return files
      .sort((left, right) => right.modified - left.modified)
      .slice(0, HISTORY_LIMIT);
  } catch {
    return [];
  }
}

/** Read and parse a JSON request body (bounded). */
function readJsonBody(req, limit = 8 * 1024 * 1024) {
  return new Promise((resolvePromise, rejectPromise) => {
    let raw = '';
    let size = 0;
    req.setEncoding('utf8');
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > limit) {
        rejectPromise(new Error('request body too large'));
        req.destroy();
        return;
      }
      raw += chunk;
    });
    req.on('end', () => {
      try {
        resolvePromise(raw.length > 0 ? JSON.parse(raw) : {});
      } catch {
        rejectPromise(new Error('invalid JSON body'));
      }
    });
    req.on('error', (error) => rejectPromise(error));
  });
}

function sendJson(res, status, value) {
  const body = JSON.stringify(value);
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
  });
  res.end(body);
}

function sendText(res, status, text) {
  res.writeHead(status, { 'content-type': 'text/plain; charset=utf-8' });
  res.end(text);
}

/** Resolve relative paths inside the outline before the Python renderer runs. */
function resolveImagePaths(outline, cwd) {
  if (typeof outline.background === 'string' && outline.background.trim()) {
    outline.background = resolve(cwd, outline.background);
  }
  if (typeof outline.template === 'string' && outline.template.trim()) {
    outline.template = resolve(cwd, outline.template);
  }
  const slides = Array.isArray(outline.slides) ? outline.slides : [];
  for (const slide of slides) {
    if (!slide || typeof slide !== 'object') continue;
    if (typeof slide.image === 'string' && slide.image.trim()) {
      slide.image = resolve(cwd, slide.image);
    }
    if (Array.isArray(slide.images)) {
      slide.images = slide.images.map((entry) =>
        typeof entry === 'string' && entry.trim() ? resolve(cwd, entry) : entry,
      );
    }
  }
}

/** Template library location inside the working directory. */
function templateLibrary(cwd) {
  const dir = join(cwd, 'ppt-forge-library', 'templates');
  const registry = join(dir, 'registry.json');
  return { dir, registry };
}

function listTemplates(cwd) {
  const { registry } = templateLibrary(cwd);
  if (!existsSync(registry)) return [];
  try {
    const data = JSON.parse(readFileSync(registry, 'utf8'));
    return Array.isArray(data.templates) ? data.templates : [];
  } catch {
    return [];
  }
}

/** Import one .pptx into the template library and register it by name. */
function addTemplate({ sourcePath, name, cwd }) {
  if (!sourcePath || !existsSync(sourcePath)) throw new Error(`template source not found: ${sourcePath}`);
  const { dir, registry } = templateLibrary(cwd);
  mkdirSync(dir, { recursive: true });
  const stem = safeFileStem(name || basename(sourcePath, '.pptx'));
  const target = join(dir, `${stem}.pptx`);
  copyFileSync(sourcePath, target);
  const templates = listTemplates(cwd).filter((entry) => entry.name !== stem);
  templates.push({
    name: stem,
    file: `${stem}.pptx`,
    addedAt: new Date().toISOString(),
    source: sourcePath,
  });
  writeFileSync(registry, JSON.stringify({ templates }, null, 2), 'utf8');
  return { name: stem, path: target, count: templates.length };
}

/** Resolve a library template name to its absolute file path. */
function resolveTemplateName(cwd, name) {
  if (!name) return undefined;
  const entry = listTemplates(cwd).find((item) => item.name === String(name));
  if (!entry) return undefined;
  const { dir } = templateLibrary(cwd);
  const path = join(dir, entry.file);
  return existsSync(path) ? path : undefined;
}

export function apply(ctx, config = {}) {
  const runtime = resolvePython(config);
  const python = runtime ? runtime.command : undefined;
  const defaultTheme = config.defaultTheme || 'navy';

  /** Locate the bundled LibreOffice Kit CLI used for render review. */
  function findLibreOfficeKit() {
    try {
      const execDir = dirname(process.execPath);
      const nodeExe = join(execDir, 'resources', 'runtime', 'primary-runtime', 'dependencies', 'node', 'bin', 'node.exe');
      const cli = join(
        execDir,
        'resources',
        'app.asar.unpacked',
        'dsh',
        'node_modules',
        '@deepseek-ai',
        'libreoffice-kit',
        'lib',
        'cli.js',
      );
      if (existsSync(nodeExe) && existsSync(cli)) return { node: nodeExe, cli };
    } catch {
      // ignore and report missing below
    }
    return undefined;
  }

  /** Run the render CLI and parse its JSON manifest. */
  function runRender({ deckPath, outDir, pages, signal }) {
    const kit = findLibreOfficeKit();
    if (!kit) return Promise.reject(new Error('LibreOffice Kit not found in this installation'));
    return new Promise((resolvePromise, rejectPromise) => {
      const args = [kit.cli, 'render', '--input', deckPath, '--output-dir', outDir];
      if (pages) args.push('--pages', String(pages));
      const child = spawn(kit.node, args, { windowsHide: true, signal });
      let stdout = '';
      let stderr = '';
      child.stdout.setEncoding('utf8');
      child.stderr.setEncoding('utf8');
      child.stdout.on('data', (chunk) => (stdout += chunk));
      child.stderr.on('data', (chunk) => (stderr += chunk));
      child.on('error', (error) => rejectPromise(error));
      child.on('close', (code) => {
        let parsed;
        try {
          parsed = JSON.parse(stdout);
        } catch {
          parsed = undefined;
        }
        if (code === 0 && parsed && typeof parsed === 'object' && !parsed.error) {
          resolvePromise(parsed);
          return;
        }
        rejectPromise(new Error((stderr || stdout || `render exited with code ${code}`).slice(0, 2000)));
      });
    });
  }

  /** Render a deck into its project preview/ folder and summarize the result. */
  async function renderDeck({ path, cwd, pages, signal }) {
    if (!python) throw new Error('Python runtime not found. Set pythonPath or DSH_PYTHON.');
    if (!path || typeof path !== 'string') throw new Error('path is required for action=render');
    const deckPath = resolve(cwd, path);
    if (!existsSync(deckPath)) throw new Error(`deck not found: ${deckPath}`);
    const projectDir = dirname(deckPath);
    const deckSha256 = createHash('sha256').update(readFileSync(deckPath)).digest('hex');
    const nonce = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const stagingDir = join(projectDir, `.preview-${nonce}`);
    const previewDir = join(projectDir, 'preview', `${safeFileStem(basename(deckPath))}-${deckSha256.slice(0, 12)}-${nonce}`);
    try {
      const manifest = await runRender({ deckPath, outDir: stagingDir, pages, signal });
      mkdirSync(dirname(previewDir), { recursive: true });
      renameSync(stagingDir, previewDir);
      const images = (manifest.images || []).map((image) => ({ page: image.page, path: join(previewDir, basename(image.path)) }));
      const quality = await runPython(python, 'quality.py', ['render', deckPath], {
        stdin: JSON.stringify({ deckSha256, images, missingFonts: manifest.missingFonts || [] }), signal,
      });
      return {
        deck: deckPath,
        previewDir,
        pageCount: manifest.pageCount,
        images,
        missingFonts: manifest.missingFonts || [],
        quality,
        reviewToken: quality.reviewToken,
      };
    } catch (error) {
      if (python && !signal?.aborted) {
        try {
          await runPython(python, 'quality.py', ['render', deckPath], {
            stdin: JSON.stringify({ deckSha256, error: String(error.message || error), images: [] }), signal,
          });
        } catch { /* preserve the original render error */ }
      }
      throw error;
    } finally {
      if (dirname(stagingDir) === projectDir && basename(stagingDir).startsWith('.preview-') && existsSync(stagingDir)) {
        rmSync(stagingDir, { recursive: true, force: true });
      }
    }
  }

  /** Shared build operation used by both the tool and the HTTP route. */
  async function buildDeck({ outline, cwd, output, theme, style, signal, render = false }) {
    if (!python) throw new Error('Python runtime not found. Set pythonPath in the plugin config or DSH_PYTHON.');
    if (!outline || typeof outline !== 'object' || Array.isArray(outline)) {
      throw new Error('outline must be a JSON object');
    }
    const resolved = { ...outline };
    if (Array.isArray(outline.slides)) {
      // Tool arguments are frozen; build writable slide objects before path rewrites.
      resolved.slides = outline.slides.map((slide) =>
        slide && typeof slide === 'object' ? { ...slide } : slide,
      );
    }
    if (theme) resolved.theme = theme;
    if (!resolved.theme) resolved.theme = defaultTheme;
    if (style && typeof style === 'object' && !Array.isArray(style)) resolved.style = style;
    resolveImagePaths(resolved, cwd);
    if (!Array.isArray(resolved.slides) || resolved.slides.length === 0) {
      throw new Error('outline.slides must be a non-empty array');
    }
    const outputPath = output ? resolve(cwd, output) : defaultOutputPath(cwd, resolved);
    mkdirSync(dirname(outputPath), { recursive: true });
    const tempInput = join(tmpdir(), `ppt-forge-outline-${Date.now()}-${Math.random().toString(36).slice(2)}.json`);
    writeFileSync(tempInput, JSON.stringify(resolved), 'utf8');
    try {
      const result = await runPython(python, 'build.py', [tempInput, outputPath], { signal });
      const summary = {
        output: outputPath,
        projectDir: dirname(outputPath),
        slideCount: result.slideCount,
        theme: result.theme,
        mode: result.mode || 'free',
        qualityReport: result.qualityReport,
        quality: result.quality,
        readyForDelivery: false,
        ...(result.pageTypes ? { pageTypes: result.pageTypes } : {}),
        ...(result.samples ? { samples: result.samples } : {}),
      };
      if (render) {
        try {
          const preview = await renderDeck({ path: outputPath, cwd, signal });
          summary.preview = {
            dir: preview.previewDir,
            pageCount: preview.pageCount,
            images: preview.images,
            missingFonts: preview.missingFonts,
          };
          summary.quality = preview.quality;
          summary.reviewToken = preview.reviewToken;
        } catch (error) {
          summary.previewError = error instanceof Error ? error.message : String(error);
          if (result.qualityReport && existsSync(result.qualityReport)) {
            summary.quality = JSON.parse(readFileSync(result.qualityReport, 'utf8'));
          }
        }
      }
      return summary;
    } finally {
      try {
        unlinkSync(tempInput);
      } catch {
        // ignore cleanup failures
      }
    }
  }

  async function extractDeck({ path, cwd, signal }) {
    if (!python) throw new Error('Python runtime not found. Set pythonPath in the plugin config or DSH_PYTHON.');
    if (!path || typeof path !== 'string') throw new Error('path is required for action=extract');
    const sourcePath = resolve(cwd, path);
    if (!existsSync(sourcePath)) throw new Error(`source not found: ${sourcePath}`);
    const result = await runPython(python, 'extract.py', [sourcePath], { signal });
    return { source: sourcePath, ...result };
  }

  async function editDeck({ path, cwd, output, operations, signal }) {
    if (!python) throw new Error('Python runtime not found. Set pythonPath or DSH_PYTHON.');
    if (!path || typeof path !== 'string') throw new Error('path is required for action=edit');
    if (!operations || typeof operations !== 'object' || Array.isArray(operations)) throw new Error('operations must be an object');
    const sourcePath = resolve(cwd, path);
    if (!existsSync(sourcePath)) throw new Error(`deck not found: ${sourcePath}`);
    const outputPath = resolve(cwd, output || path);
    const operationsPath = join(tmpdir(), `ppt-forge-edit-${Date.now()}-${Math.random().toString(36).slice(2)}.json`);
    writeFileSync(operationsPath, JSON.stringify(operations), 'utf8');
    try {
      const edited = await runPython(python, 'edit.py', [sourcePath, outputPath, operationsPath], { signal });
      const quality = await runPython(python, 'quality.py', ['check', outputPath], { signal });
      return { ...edited, qualityReport: quality.reportPath, quality };
    } finally {
      try { unlinkSync(operationsPath); } catch { /* cleanup */ }
    }
  }

  async function extractAssets({ path, cwd, name, signal }) {
    if (!python) throw new Error('Python runtime not found. Set pythonPath in the plugin config or DSH_PYTHON.');
    if (!path || typeof path !== 'string') throw new Error('path is required for action=assets');
    const inputPath = resolve(cwd, path);
    if (!existsSync(inputPath)) throw new Error(`input not found: ${inputPath}`);
    const profileName = safeFileStem(name || `assets-${stamp()}`);
    const libraryDir = join(cwd, 'ppt-forge-library');
    mkdirSync(libraryDir, { recursive: true });
    const result = await runPython(
      python,
      'extract_assets.py',
      [inputPath, libraryDir, '--name', profileName],
      { signal },
    );
    if (result && typeof result === 'object' && typeof result.error === 'string') {
      throw new Error(result.error);
    }
    return result;
  }

  async function learnLibrary({ path, cwd, name, signal }) {
    if (!python) throw new Error('Python runtime not found. Set pythonPath in the plugin config or DSH_PYTHON.');
    if (!path || typeof path !== 'string') throw new Error('path is required for action=learn');
    const inputPath = resolve(cwd, path);
    if (!existsSync(inputPath)) throw new Error(`input not found: ${inputPath}`);
    const profileName = safeFileStem(name || `style-${stamp()}`);
    const base = join(cwd, 'ppt-forge-library', profileName);
    mkdirSync(dirname(base), { recursive: true });
    const result = await runPython(
      python,
      'learn.py',
      [inputPath, base, '--name', profileName],
      { signal },
    );
    if (result && typeof result === 'object' && typeof result.error === 'string') {
      throw new Error(result.error);
    }
    return result;
  }

  ctx.effect(() => {
    const dispose = ctx.tools.register({
      name: 'ppt_forge',
      description:
        'Build, inspect, edit, or learn from PowerPoint decks. extract returns content; inspect returns template shape IDs, names, bounds and estimated capacity; edit applies atomic page deletion/addition/reordering, targeted text edits, TOC synchronization and page-number updates; learn analyzes samples; assets extracts images; template-add/template-list manage templates. build validates a structured outline, supports explicit fieldMap bindings and clears unbound template prose unless preserveText is supplied. render produces versioned previews and a reviewToken. check returns current quality state. Visual and content reviews are separate; only readyForDelivery=true indicates both passed for the current deck.',
      parameters: {
        action: {
          type: 'string',
          required: true,
          description:
            'edit | extract | inspect | learn | assets | build | render | check | review | template-add | template-list',
        },
        path: {
          type: 'string',
          description:
            'Path to a source .pptx for action=extract or action=template-add, a deck for action=render, or a folder/file for action=learn and action=assets. Relative paths resolve against the session working directory.',
        },
        outline: {
          type: 'object',
          additionalProperties: true,
          description:
            'Deck outline for action=build: {title, subtitle, theme, style, background, slides:[...]}. Slide types: cover {title,subtitle,tag,image}; toc {title,items[]}; section {number,title,subtitle,image}; content {title,bullets[],notes}; two-column {title,leftTitle,leftBullets,rightTitle,rightBullets}; table {title,headers[],rows[[]]}; image {title,image,caption}; image-text {title,image,caption,layout:left|right,bullets[]}; gallery {title,images[],captions[]}; stats {title,items:[{value,label,note}]}; end {title,subtitle}. background is an optional deck-wide image path. Missing image paths render a dashed "图片占位" placeholder box.',
        },
        style: {
          type: 'object',
          additionalProperties: true,
          description:
            'Optional style overrides for action=build, usually mapped from a learned profile: {titleAlign:"left"|"center", showPageNumber:boolean, showFooter:boolean, footerText:string, coverEyebrow:string, coverBadge:string, imageFit:"contain"|"cover", imageRadius:number, autoTitleSize:boolean, bulletChar:string, titleSize:number, bodySize:number, palette:{dark,primary,accent,ink,body,muted,line,soft}}. footerText/coverEyebrow/coverBadge default empty so no English boilerplate is emitted; imageFit:"cover" center-crops images to their frame and rounds corners; autoTitleSize shrinks over-long titles instead of wrapping.',
        },
        output: {
          type: 'string',
          description:
            'Optional output .pptx path for action=build. Defaults to <working directory>/ppt-forge-output/<title>-<timestamp>-<unique-id>/deck.pptx, with an adjacent deck.quality.json report.',
        },
        theme: {
          type: 'string',
          description: `Color theme: ${DEFAULT_THEMES.join(' | ')}. Defaults to the profile setting (${defaultTheme}).`,
        },
        template: {
          type: 'string',
          description:
            'Optional .pptx template path for action=build. When provided, pages are duplicated from the template and their text rewritten, so the deck keeps the template design. Prefer samplePages to choose which template pages to clone.',
        },
        templateName: {
          type: 'string',
          description:
            'Optional template name from the local template library (imported via action=template-add). Alternative to the template path.',
        },
        samplePages: {
          type: 'object',
          additionalProperties: true,
          description:
            'Optional 1-based template page overrides for template mode, e.g. {cover: 1, section: 4, content: 6, end: 16}. Automatic page classification is heuristic; explicit samples are more reliable.',
        },
        fieldMap: {
          type: 'object', additionalProperties: true,
          description: 'Template bindings keyed by slide type, e.g. {content:{title:{name:"Title"},bullets:[{id:5},{id:6}]}}. Use inspect first. Per-slide outline.slides[].fieldMap overrides this. Every supplied content field must be mapped. Bind table to a table shape; image/images to pictures or placeholders. Unmapped decorative objects are preserved.',
        },
        qualityMode: {
          type: 'string',
          description: 'warn (default) records estimated layout issues; strict refuses to replace the output when static checks have findings. Both modes require full render and explicit visual review before readyForDelivery becomes true.',
        },
        review: {
          type: 'object', additionalProperties: true,
          description: 'For action=review: {reviewType:"visual"|"content", reviewToken:string, findings:[{page:1,severity:"warning"|"error",message:string}], reviewer?:string}. Submit only after reviewing the current deck. Visual and content approval are separate.',
        },
        operations: {
          type: 'object', additionalProperties: true,
          description: 'For action=edit: {deletePages:[1], addPages:[{after:2,templatePage:3,replace:{Title:"New title"}}], pageOrder:[1,3,2], replace:[{page:2,selector:{name:"Body"},text:"New text"}], syncToc:{page:2,selector:{name:"TOC"}}, pageNumbers:{selector:{name:"PageNumber"},format:"{page:02d}/{total:02d}"}}.',
        },
        render: {
          type: 'boolean',
          description: 'For action=build: render into a new versioned subdirectory under <projectDir>/preview/, returning images, missingFonts, quality and reviewToken. Rendered files still require explicit visual review.',
        },
        pages: {
          type: 'string',
          description: 'For action=render: page selection such as "1,3,5" or "1-4". Omit to render every page.',
        },
        name: {
          type: 'string',
          description:
            'Profile/asset name for action=learn or action=assets, used as the folder name under <working directory>/ppt-forge-library/. Defaults to a timestamp.',
        },
      },
      output: {
        schema: { type: 'object', additionalProperties: true },
        render: (_args, value) => [{ type: 'text', text: JSON.stringify(value) }],
      },
      async execute(args, exec) {
        if (exec.agent === undefined) throw new Error('ppt_forge requires an agent Session');
        const cwd = await ctx.workingDirectory.ensure(exec.agent, exec.signal);
        const action = String(args?.action || '');
        if (action === 'edit') {
          const edited = await editDeck({
            path: args?.path,
            cwd,
            output: args?.output,
            operations: args?.operations,
            signal: exec.signal,
          });
          return { action, ...edited };
        }
        if (action === 'inspect' || action === 'check' || action === 'review') {
          if (!python) throw new Error('Python runtime not found. Set pythonPath or DSH_PYTHON.');
          if (!args?.path || typeof args.path !== 'string') throw new Error('path is required');
          const source = resolve(cwd, args.path);
          const result = action === 'inspect'
            ? await runPython(python, 'inspection.py', [source], { signal: exec.signal })
            : await runPython(python, 'quality.py', [action, source], {
              ...(action === 'review' ? { stdin: JSON.stringify(args.review || {}) } : {}), signal: exec.signal,
            });
          return { action, ...result };
        }
        if (action === 'extract') {
          const extraction = await extractDeck({ path: args?.path, cwd, signal: exec.signal });
          return { action, ...extraction };
        }
        if (action === 'learn') {
          const learned = await learnLibrary({
            path: args?.path,
            cwd,
            name: args?.name,
            signal: exec.signal,
          });
          return { action, ...learned };
        }
        if (action === 'assets') {
          const assets = await extractAssets({
            path: args?.path,
            cwd,
            name: args?.name,
            signal: exec.signal,
          });
          return { action, ...assets };
        }
        if (action === 'build') {
          let outline = args?.outline;
          let templatePath = typeof args?.template === 'string' ? args.template : undefined;
          if (!templatePath && args?.templateName) {
            templatePath = resolveTemplateName(cwd, args.templateName);
            if (!templatePath) {
              throw new Error(`template "${args.templateName}" not found in the library; import it with action=template-add`);
            }
          }
          if (templatePath) {
            outline = { ...(outline && typeof outline === 'object' ? outline : {}), template: templatePath };
          }
          if (args?.samplePages && typeof args.samplePages === 'object') {
            outline = { ...(outline || {}), samplePages: args.samplePages };
          }
          if (args?.fieldMap !== undefined) outline = { ...(outline || {}), fieldMap: args.fieldMap };
          if (args?.qualityMode !== undefined) outline = { ...(outline || {}), qualityMode: args.qualityMode };
          const built = await buildDeck({
            outline,
            cwd,
            output: args?.output,
            theme: args?.theme,
            style: args?.style,
            render: args?.render === true,
            signal: exec.signal,
          });
          return { action, ...built };
        }
        if (action === 'template-add') {
          const sourcePath = resolve(cwd, String(args?.path || ''));
          const added = addTemplate({ sourcePath, name: args?.name, cwd });
          return { action, ...added };
        }
        if (action === 'template-list') {
          return { action, templates: listTemplates(cwd) };
        }
        if (action === 'render') {
          const rendered = await renderDeck({
            path: args?.path,
            cwd,
            pages: args?.pages,
            signal: exec.signal,
          });
          return { action, ...rendered };
        }
        throw new Error('action must be edit, extract, inspect, learn, assets, build, render, check, review, template-add, or template-list');
      },
    });
    return () => dispose();
  });

  // Optional browser carrier: present in the web profile, absent in headless ones.
  ctx.inject(['webServer'], (child) => {
    child.effect(() => {
      const dispose = child.webServer.register({
        kind: 'prefix',
        path: ROUTE_PREFIX,
        async handler(req, res) {
          try {
            const url = new URL(req.url || '/', 'http://localhost');
            const pathname = url.pathname;

            if (pathname === `${ROUTE_PREFIX}/status` && req.method === 'GET') {
              sendJson(res, 200, {
                ok: python !== undefined,
                python: runtime
                  ? { command: runtime.command, version: runtime.version, pptxVersion: runtime.pptxVersion }
                  : null,
                themes: DEFAULT_THEMES,
                defaultTheme,
              });
              return;
            }

            if (pathname === `${ROUTE_PREFIX}/history` && req.method === 'GET') {
              const cwd = url.searchParams.get('cwd') || process.cwd();
              sendJson(res, 200, { cwd, files: listHistory(cwd) });
              return;
            }

            if (pathname === `${ROUTE_PREFIX}/build` && req.method === 'POST') {
              const body = await readJsonBody(req);
              if (!python) throw new Error('Python runtime not found on the Host');
              const targetDir = body.cwd ? resolve(String(body.cwd)) : process.cwd();
              const built = await buildDeck({
                outline: { ...body.outline, ...(body.fieldMap !== undefined ? {fieldMap: body.fieldMap} : {}), ...(body.qualityMode !== undefined ? {qualityMode: body.qualityMode} : {}) },
                cwd: targetDir,
                output: body.output,
                theme: body.theme,
                style: body.style,
                render: body.render === true,
              });
              sendJson(res, 200, built);
              return;
            }

            sendText(res, 404, 'not found');
          } catch (error) {
            sendJson(res, 400, { error: error instanceof Error ? error.message : String(error), code: error.code, issues: error.issues });
          }
        },
      });
      return () => dispose();
    });
  });
}
