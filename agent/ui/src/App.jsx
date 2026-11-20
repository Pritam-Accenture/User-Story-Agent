import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  Bar, BusyIndicator, Button, Card, CardHeader, Dialog, FlexBox, Label, MessageStrip, MultiComboBox, MultiComboBoxItem,
  Option, Page, ProgressIndicator, RadioButton, Select, Text, TextArea, Title,
} from '@ui5/webcomponents-react';

const api = async (path, opts) => {
  const r = await fetch('/api' + path, opts);
  const body = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(body.detail || r.statusText);
  return body;
};
const post = (path, json) => api(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(json || {}) });
const grey = { color: '#6a6d70' };
const W = { width: '100%', maxWidth: 1000 };

function Section({ title, subtitle, children }) {
  return (
    <Card header={<CardHeader titleText={title} subtitleText={subtitle} />} style={W}>
      <div style={{ padding: '0.75rem 1rem 1rem' }}>{children}</div>
    </Card>
  );
}

function Slot({ slot, info, accept, onChange, setError, hint }) {
  const ref = useRef();
  const [busy, setBusy] = useState(false);
  const pick = async (e) => {
    setBusy(true);
    try {
      for (const f of e.target.files) {
        const fd = new FormData(); fd.append('file', f);
        const r = await fetch(`/api/files/${slot}`, { method: 'POST', body: fd });
        if (!r.ok) throw new Error((await r.json()).detail || 'Upload failed');
      }
      onChange();
    } catch (err) { setError(err.message); } finally { setBusy(false); e.target.value = ''; }
  };
  const remove = async (name) => { await api(`/files/${slot}/${encodeURIComponent(name)}`, { method: 'DELETE' }); onChange(); };
  const has = info.files.length > 0;
  return (
    <div style={{ padding: '0.4rem 0', borderBottom: '1px solid #eee' }}>
      <FlexBox alignItems="Center" style={{ gap: '1rem' }}>
        <input ref={ref} type="file" multiple={info.multiple} accept={accept} hidden onChange={pick} />
        <Button design={has ? 'Default' : info.mandatory ? 'Emphasized' : 'Transparent'} icon="upload" onClick={() => ref.current.click()} style={{ minWidth: 330, justifyContent: 'flex-start' }}>
          {info.label}{info.mandatory ? ' *' : ' (optional)'}
        </Button>
        {busy && <BusyIndicator active size="S" />}
        {!has && !busy && <Text style={grey}>{hint}</Text>}
      </FlexBox>
      {info.files.map((f) => (
        <FlexBox key={f.name} alignItems="Center" style={{ gap: '0.5rem', marginLeft: '0.5rem' }}>
          <Text>✔ {f.name}</Text><Text style={grey}>{Math.max(1, Math.round(f.size / 1024))} KB</Text>
          <Button design="Transparent" icon="decline" tooltip="Remove" onClick={() => remove(f.name)} />
        </FlexBox>
      ))}
    </div>
  );
}

export default function App() {
  const [st, setSt] = useState(null);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');
  const [help, setHelp] = useState(false);
  const [freeText, setFreeText] = useState(null);
  const timer = useRef();

  const load = useCallback(() => api('/state').then(setSt).catch((e) => setError('Backend not reachable: ' + e.message)), []);
  useEffect(() => { load(); const t = setInterval(load, 2000); return () => clearInterval(t); }, [load]);
  const save = (patch) => { clearTimeout(timer.current); timer.current = setTimeout(() => post('/config', patch).then(load), 400); };

  if (!st) return <BusyIndicator active style={{ margin: '4rem' }} />;
  const { config, files, claude, models, extract, job, countries, detected } = st;
  const detIds = detected.items.map((d) => d.id), detCodes = detected.countries.map((d) => d.code);
  const byId = new Map(st.known_scope.map((o) => [o.id, o]));
  detected.items.forEach((d) => byId.set(d.id, { id: d.id, name: d.name || (byId.get(d.id) || {}).name || '' }));
  const scopeOptions = [...byId.values()].sort((x, y) => x.id.localeCompare(y.id));
  const fTxt = freeText ?? config.free_text;
  const ex = extract.status || {}, meta = extract.meta;
  const exBusy = ['waiting_login', 'fetching', 'downloading'].includes(ex.state) && extract.running;
  const run = async (fn) => { setError(''); setInfo(''); try { await fn(); await load(); } catch (e) { setError(e.message); } };

  const missing = [];
  if (!st.scope_ids.length) missing.push('scope item IDs');
  if (!st.country_ids.length) missing.push('a country');
  Object.entries(files).forEach(([k, v]) => { if (v.mandatory && !v.files.length) missing.push(v.label); });
  if (config.output_mode === 'template' && !files.template.files.length) missing.push('output template');

  const mandatory = ['workshop', 'kdd', 'gap', 'l4', 'bpd', 'backlog'];
  return (
    <Page header={<Bar startContent={<Title level="H3">SAP S/4HANA Public Cloud – User Story Agent</Title>} endContent={<Button icon="hint" design="Transparent" onClick={() => setHelp(true)}>How it works</Button>} />} style={{ height: '100vh' }} backgroundDesign="Solid">
      <Dialog open={help} onClose={() => setHelp(false)} headerText="How it works and how to get the best result"
        footer={<Bar endContent={<Button design="Emphasized" onClick={() => setHelp(false)}>Close</Button>} />} style={{ width: 'min(720px, 95vw)' }}>
        <div style={{ padding: '0.5rem 1rem', lineHeight: 1.5, fontFamily: 'var(--sapFontFamily)', fontSize: 'var(--sapFontSize)', color: 'var(--sapTextColor)' }}>
          <Title level="H5">What the agent does</Title>
          <Text>It turns your project material into sprint-ready L4 activities, user stories, tasks and sub-tasks, each traced back to a scope item and to the document it came from. The result is a new Excel file – in the sample layout or in your own template. Your existing files are never changed.</Text>
          <Title level="H5" style={{ marginTop: 16 }}>Steps</Title>
          <ol style={{ margin: '4px 0', paddingLeft: '1.2rem' }}>
            <li>Sign in to Claude (top section). Pick a model – the strongest is selected by default.</li>
            <li>Optional: press <b>Refresh</b> to update the stored SAP Process Navigator data (needs your SAP sign-in in Edge). Generation itself never contacts SAP.</li>
            <li>Upload your files. The <b>KDD register</b> and the <b>L4 plan</b> are required; everything else is optional but improves the result.</li>
            <li>Check the detected <b>scope items</b> and <b>countries</b>. Remove the ones you don't want and add any that are missing.</li>
            <li>Choose the output (sample format or your own template) and press <b>Generate</b>. Then download the Excel file.</li>
          </ol>
          <Title level="H5" style={{ marginTop: 16 }}>What gives the best result</Title>
          <ul style={{ margin: '4px 0', paddingLeft: '1.2rem' }}>
            <li><b>More and better material.</b> Every file you add gives Claude more real project detail to write from. Add workshop notes or minutes, the Gap / WRICEF list, BPDs, the product backlog and a traceability matrix.</li>
            <li><b>Specific, decided content.</b> Complete KDD entries with the decision and its rationale, and workshop notes that name the scope item, beat short headings or empty rows.</li>
            <li><b>Scope item IDs written out</b> (for example BD6, J45) in your files, so they are detected and each requirement lands on the right scope item.</li>
            <li><b>Name the countries</b> in your material. They drive the country-specific test steps and the Country column.</li>
            <li><b>Use the free-text box</b> for rules the files don't hold: naming conventions, priorities, sprint length, what to leave out.</li>
            <li><b>Keep the Process Navigator data fresh.</b> Refresh it before a big run, so activities and test steps match the latest SAP release.</li>
            <li><b>Use your own template when you have one.</b> Claude reads its headers and dropdowns and fills them in; a clearly labelled template is mapped more reliably.</li>
            <li><b>Check the run notes.</b> After generating, they list what was unmapped, skipped or approximated.</li>
          </ul>
          <Title level="H5" style={{ marginTop: 16 }}>Good to know</Title>
          <ul style={{ margin: '4px 0', paddingLeft: '1.2rem' }}>
            <li>Files can be Word, PowerPoint, PDF, Excel, CSV or text. They stay on your computer; the text is sent to Claude under your own login to write the stories.</li>
            <li>Without material the output is mostly standard SAP content. Requirements that don't match any scope item are grouped under an "unmapped" activity for you to place.</li>
            <li>The owner column is left blank. Review the generated stories before loading them into your tool – Claude can make mistakes.</li>
            <li>Every launch starts blank. Nothing you entered is kept between sessions.</li>
          </ul>
        </div>
      </Dialog>
      <FlexBox direction="Column" alignItems="Center" style={{ gap: '1rem', padding: '1rem' }}>
        {error && <MessageStrip design="Negative" onClose={() => setError('')} style={W}>{error}</MessageStrip>}
        {info && <MessageStrip design="Information" onClose={() => setInfo('')} style={W}>{info}</MessageStrip>}

        <Section title="Claude" subtitle="Story wording and reading of your attachments use your own Claude login – no API key needed">
          <FlexBox alignItems="Center" style={{ gap: '1rem' }}>
            <MessageStrip hideCloseButton design={claude.loggedIn ? 'Positive' : 'Critical'} style={{ flex: 1 }}>{claude.detail}</MessageStrip>
            {!claude.loggedIn && <Button design="Emphasized" onClick={() => run(async () => setInfo((await post('/claude/login')).message))}>Sign in to Claude</Button>}
            <Button onClick={() => run(() => api('/claude/status'))}>Check</Button>
          </FlexBox>
          {claude.loggedIn && (
            <FlexBox alignItems="Center" style={{ gap: '1rem', marginTop: 10 }}>
              <Label showColon>Claude model</Label>
              {models.list.length === 0 ? <Text style={grey}>Loading the models available to you…</Text> : (
                <Select style={{ minWidth: 320 }} onChange={(e) => post('/config', { model: e.detail.selectedOption.dataset.value }).then(load)}>
                  {models.list.map((m) => (
                    <Option key={m.value} data-value={m.value} selected={(config.model || models.best || 'default') === m.value}>
                      {m.label}{m.value === models.best && !config.model ? ' – strongest, selected by default' : ''}
                    </Option>
                  ))}
                </Select>
              )}
              <Text style={grey}>Only models available to your Claude login are listed</Text>
            </FlexBox>
          )}
        </Section>

        <Section title="SAP Process Navigator data" subtitle="Stored on this computer and reused – the app does not contact SAP during generation">
          <FlexBox alignItems="Center" style={{ gap: '1rem' }}>
            <Text style={{ flex: 1 }}>
              {meta ? <>Last refreshed: <b>{new Date(meta.last_refreshed).toLocaleString()}</b></> : 'Last refreshed: never'}
            </Text>
            <Button design="Emphasized" icon="refresh" disabled={exBusy} onClick={() => run(async () => {
              await post('/extract/refresh');
              setInfo('An Edge window opens – sign in to SAP there and leave it open until the status shows Done. The refresh covers the scope item IDs and countries entered below.');
            })}>Refresh</Button>
          </FlexBox>
          {ex.state && ex.state !== 'never' && ex.state !== 'done' && (
            <div style={{ marginTop: 8 }}>
              <MessageStrip hideCloseButton design={ex.state === 'error' ? 'Negative' : 'Information'}>{ex.message}</MessageStrip>
              {exBusy && ex.total > 0 && <ProgressIndicator value={Math.round((100 * ex.done) / ex.total)} style={{ marginTop: 8 }} />}
            </div>
          )}
          {ex.state === 'done' && ex.notes?.length > 0 && (
            <details style={{ marginTop: 8 }}><summary>{ex.notes.length} notes from the last refresh</summary>
              <ul>{ex.notes.map((n, i) => <li key={i}><Text>{n}</Text></li>)}</ul></details>
          )}
        </Section>

        <Section title="Inputs" subtitle="Files stay on this computer. * = mandatory">
          {mandatory.map((k) => <Slot key={k} slot={k} info={files[k]} onChange={load} setError={setError}
            accept={k === 'l4' ? '.xlsx,.xlsm' : undefined} hint={k === 'l4' ? 'Excel (the L4 plan format)' : undefined} />)}
          <Slot slot="trace" info={files.trace} onChange={load} setError={setError} />
          <Slot slot="other" info={files.other} onChange={load} setError={setError} />
          <div style={{ height: 14 }} />
          <Label showColon>Scope items</Label>
          <MultiComboBox style={{ width: '100%' }} placeholder={detected.loading ? 'Detecting scope items in your files…' : 'Detected from your files – type an ID or name to add more'}
            onSelectionChange={(e) => {
              const sel = e.detail.items.map((i) => i.dataset.id);
              post('/config', { scope_excluded: detIds.filter((x) => !sel.includes(x)), scope_added: sel.filter((x) => !detIds.includes(x)) }).then(load);
            }}>
            {scopeOptions.map((o) => <MultiComboBoxItem key={o.id} data-id={o.id} text={o.name ? `${o.id} – ${o.name}` : o.id} selected={st.scope_ids.includes(o.id)} />)}
          </MultiComboBox>
          <Text style={grey}>{detected.loading ? 'Detecting…' : `${detIds.length} detected in your files · ${st.scope_ids.length} selected`}. Remove any you don't want, or search to add others.</Text>
          <div style={{ height: 10 }} />
          <Label showColon>Countries</Label>
          <MultiComboBox style={{ width: '100%' }} placeholder={detected.loading ? 'Detecting countries in your files…' : 'Detected from your files – type a code or name to add more'}
            onSelectionChange={(e) => {
              const sel = e.detail.items.map((i) => i.dataset.code);
              post('/config', { country_excluded: detCodes.filter((x) => !sel.includes(x)), country_added: sel.filter((x) => !detCodes.includes(x)) }).then(load);
            }}>
            {countries.map((c) => <MultiComboBoxItem key={c.code} data-code={c.code} text={`${c.code} – ${c.name}`} selected={st.country_ids.includes(c.code)} />)}
          </MultiComboBox>
          <Text style={grey}>{detected.loading ? 'Detecting…' : `${detCodes.length} detected in your files · ${st.country_ids.length} selected`}</Text>
          <div style={{ height: 10 }} />
          <Label showColon>Anything else the agent should know</Label>
          <TextArea rows={3} style={{ width: '100%' }} placeholder="Free text – constraints, naming rules, priorities, context…"
            value={fTxt} onInput={(e) => { setFreeText(e.target.value); save({ free_text: e.target.value }); }} />
        </Section>

        <Section title="Output" subtitle="Choose how the result is written. A new file is always created">
          <FlexBox direction="Column" style={{ gap: '0.25rem' }}>
            <RadioButton name="out" text="Sample output format (L4 Activity › User Story › Task › Sub-task)" checked={config.output_mode === 'default'}
              onChange={() => post('/config', { output_mode: 'default' }).then(load)} />
            <Text style={{ ...grey, marginLeft: '2rem' }}><a href="/api/template/sample">Download the sample</a></Text>
            <RadioButton name="out" text="Based on my own template (Cloud ALM, ADM tool, ADO, Jira, anything in Excel)" checked={config.output_mode === 'template'}
              onChange={() => post('/config', { output_mode: 'template' }).then(load)} />
          </FlexBox>
          {config.output_mode === 'template' && (
            <div style={{ marginLeft: '2rem', marginTop: 6 }}>
              <Slot slot="template" info={{ ...files.template, mandatory: true }} onChange={load} setError={setError} accept=".xlsx,.xlsm" hint="Excel template – Claude works out which column gets what" />
            </div>
          )}
        </Section>

        <Section title="Generate" subtitle="Creates a new Excel file – existing files are never changed">
          <FlexBox alignItems="Center" style={{ gap: '1rem' }} wrap="Wrap">
            <Button design="Emphasized" icon="play" disabled={missing.length > 0 || job.state === 'running' || exBusy}
              onClick={() => run(() => post('/generate'))}>Generate L4 plan and user stories</Button>
          </FlexBox>
          {missing.length > 0 && <Text style={{ ...grey, display: 'block', marginTop: 6 }}>Still needed: {missing.join(' · ')}</Text>}
          {job.state === 'running' && <FlexBox alignItems="Center" style={{ gap: 8, marginTop: 8 }}><BusyIndicator active size="S" /><Text>{job.message}</Text></FlexBox>}
          {job.state === 'error' && <MessageStrip hideCloseButton design="Negative" style={{ marginTop: 8 }}>{job.message}</MessageStrip>}
          {job.state === 'done' && job.result && (
            <div style={{ marginTop: 12 }}>
              <MessageStrip hideCloseButton design="Positive">
                {job.result.stories} stories for {job.result.scope_items} scope items ({job.result.rows} rows){job.result.used_llm ? ' – wording by Claude' : ' – template wording'}
                {job.result.mode === 'template' ? ' – written into your template' : ''}
              </MessageStrip>
              <div style={{ margin: '0.75rem 0' }}>
                <Button icon="download" design="Emphasized" onClick={() => (window.location.href = `/api/download/${job.result.file}`)}>Download {job.result.file}</Button>
              </div>
              {job.result.log?.length > 0 && (
                <details><summary>{job.result.log.length} run notes</summary>
                  <ul>{job.result.log.map((n, i) => <li key={i}><Text>{n}</Text></li>)}</ul></details>
              )}
            </div>
          )}
        </Section>
      </FlexBox>
    </Page>
  );
}
