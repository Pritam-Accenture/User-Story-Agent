import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  Bar, BusyIndicator, Button, Card, CardHeader, FlexBox, Label, MessageStrip, MultiComboBox, MultiComboBoxItem,
  Page, ProgressIndicator, RadioButton, Text, TextArea, Title,
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
        {!has && !busy && <Text style={grey}>{hint || 'Any format: Word, PowerPoint, PDF, Excel, CSV, text'}</Text>}
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
  const [scopeText, setScopeText] = useState(null);
  const [freeText, setFreeText] = useState(null);
  const timer = useRef();

  const load = useCallback(() => api('/state').then(setSt).catch((e) => setError('Backend not reachable: ' + e.message)), []);
  useEffect(() => { load(); const t = setInterval(load, 2000); return () => clearInterval(t); }, [load]);
  const save = (patch) => { clearTimeout(timer.current); timer.current = setTimeout(() => post('/config', patch).then(load), 400); };

  if (!st) return <BusyIndicator active style={{ margin: '4rem' }} />;
  const { config, files, claude, extract, job, countries } = st;
  const sTxt = scopeText ?? config.scope_text, fTxt = freeText ?? config.free_text;
  const ex = extract.status || {}, meta = extract.meta;
  const exBusy = ['waiting_login', 'fetching', 'downloading'].includes(ex.state) && extract.running;
  const run = async (fn) => { setError(''); setInfo(''); try { await fn(); await load(); } catch (e) { setError(e.message); } };

  const missing = [];
  if (!st.scope_ids.length) missing.push('scope item IDs');
  if (!config.countries.length) missing.push('a country');
  Object.entries(files).forEach(([k, v]) => { if (v.mandatory && !v.files.length) missing.push(v.label); });
  if (config.output_mode === 'template' && !files.template.files.length) missing.push('output template');

  const mandatory = ['workshop', 'kdd', 'gap', 'l4', 'bpd', 'backlog'];
  return (
    <Page header={<Bar startContent={<Title level="H3">SAP S/4HANA Public Cloud – User Story Agent</Title>} />} style={{ height: '100vh' }} backgroundDesign="Solid">
      <FlexBox direction="Column" alignItems="Center" style={{ gap: '1rem', padding: '1rem' }}>
        {error && <MessageStrip design="Negative" onClose={() => setError('')} style={W}>{error}</MessageStrip>}
        {info && <MessageStrip design="Information" onClose={() => setInfo('')} style={W}>{info}</MessageStrip>}

        <Section title="Claude" subtitle="Story wording and reading of your attachments use your own Claude login – no API key needed">
          <FlexBox alignItems="Center" style={{ gap: '1rem' }}>
            <MessageStrip hideCloseButton design={claude.loggedIn ? 'Positive' : 'Critical'} style={{ flex: 1 }}>{claude.detail}</MessageStrip>
            {!claude.loggedIn && <Button design="Emphasized" onClick={() => run(async () => setInfo((await post('/claude/login')).message))}>Sign in to Claude</Button>}
            <Button onClick={() => run(() => api('/claude/status'))}>Check</Button>
          </FlexBox>
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
          <Label showColon>Scope item IDs</Label>
          <TextArea rows={2} style={{ width: '100%' }} placeholder="e.g. BD6, J45, 1EM (from the DDA) – separate with commas, spaces or new lines"
            value={sTxt} onInput={(e) => { setScopeText(e.target.value); save({ scope_text: e.target.value }); }} />
          <Text style={grey}>{st.scope_ids.length} scope item{st.scope_ids.length === 1 ? '' : 's'} recognised{st.scope_ids.length ? ': ' + st.scope_ids.join(', ') : ''}</Text>
          <div style={{ height: 10 }} />
          <Label showColon>Countries</Label>
          <MultiComboBox style={{ width: '100%' }} placeholder="Select one or several countries"
            onSelectionChange={(e) => { const c = e.detail.items.map((i) => i.dataset.code); post('/config', { countries: c }).then(load); }}>
            {countries.map((c) => <MultiComboBoxItem key={c.code} data-code={c.code} text={`${c.name} (${c.code})`} selected={config.countries.includes(c.code)} />)}
          </MultiComboBox>
          <div style={{ height: 14 }} />
          {mandatory.map((k) => <Slot key={k} slot={k} info={files[k]} onChange={load} setError={setError}
            accept={k === 'l4' ? '.xlsx,.xlsm' : undefined} hint={k === 'l4' ? 'Excel (the L4 plan format)' : undefined} />)}
          <Slot slot="trace" info={files.trace} onChange={load} setError={setError} />
          <Slot slot="other" info={files.other} onChange={load} setError={setError} />
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
