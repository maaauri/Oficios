// V5 Deep — Prototipo completo con navegación entre pantallas
// Pantallas: Inicio (bandeja), Estadísticas, Revaluación, Informe multa

const PALETTES = {
  light: {
    bg: '#f7f8fa', panel: '#ffffff', soft: '#eff2f6', softer: '#f4f6f9',
    border: '#e2e6ec', borderStrong: '#d4dae3',
    text: '#1c2633', subtext: '#6b7684', dim: '#9aa3b0',
    accent: '#0B3D6B', blue: '#1E6FB8', blueSoft: '#e8f0fa',
    success: '#1a7f5a', successSoft: '#e4f3ec',
    warn: '#c47a00', warnSoft: '#fdf2e0',
    danger: '#b42d2d', dangerSoft: '#fbe7e7',
    lilac: '#8a4fb5', lilacSoft: '#f2ebf8',
    teal: '#0e8a82', tealSoft: '#dff2f0',
    neutral: '#556270', neutralSoft: '#eaedf1',
  },
  dark: {
    bg: '#0f1722', panel: '#17212f', soft: '#1c2838', softer: '#1a2433',
    border: '#253244', borderStrong: '#2d3d52',
    text: '#e4eaf2', subtext: '#9aa8bc', dim: '#6b7a8f',
    accent: '#3d8bd9', blue: '#5aa9e6', blueSoft: '#1c3550',
    success: '#4db87d', successSoft: '#1a3a2a',
    warn: '#e0a54a', warnSoft: '#3d2e15',
    danger: '#e06464', dangerSoft: '#3d1e1e',
    lilac: '#a88be8', lilacSoft: '#2d2540',
    teal: '#4ec9b0', tealSoft: '#1a3530',
    neutral: '#8798b0', neutralSoft: '#1f2a3a',
  },
};

const AREA_COLORS = {
  'Conexiones': 'blue',
  'PMGD': 'warn',
  'Servicio al Cliente': 'success',
  'Pérdidas': 'lilac',
  'Sin área': 'neutral',
  'Cobranza': 'danger',
  'Lectura': 'teal',
};

function V5App({ tweaks = {} }) {
  const [screen, setScreen] = React.useState('inicio');
  const [selectedOficio, setSelectedOficio] = React.useState(null);
  const [multaFor, setMultaFor] = React.useState(null);

  const theme = tweaks.dark ? 'dark' : 'light';
  const pal = PALETTES[theme];
  const density = tweaks.density || 'normal'; // compact | normal | comfy

  const go = (s, payload) => {
    if (s === 'revaluar') setSelectedOficio(payload || null);
    if (s === 'multa') setMultaFor(payload || null);
    setScreen(s);
  };

  return (
    <div style={{ width: '100%', height: '100%', background: pal.bg, color: pal.text,
      fontFamily: 'Inter, system-ui, sans-serif', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <AppHeader pal={pal} screen={screen} go={go} density={density} />
      <div style={{ flex: 1, overflow: 'auto' }}>
        {screen === 'inicio' && <Inicio pal={pal} go={go} density={density} />}
        {screen === 'stats' && <Stats pal={pal} density={density} />}
        {screen === 'revaluar' && <Revaluar pal={pal} go={go} oficio={selectedOficio} density={density} />}
        {screen === 'multa' && <InformeMulta pal={pal} go={go} oficio={multaFor} density={density} />}
      </div>
      <AppFooter pal={pal} />
    </div>
  );
}

// ─── HEADER ────────────────────────────────────────────────────────────
function AppHeader({ pal, screen, go, density }) {
  const tabs = [
    { id: 'inicio', l: 'Bandeja' },
    { id: 'stats', l: 'Estadísticas' },
  ];
  return (
    <header style={{ background: pal.panel, borderBottom: `1px solid ${pal.border}`, padding: density === 'compact' ? '10px 24px' : '14px 28px' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
        <div style={{ width: 32, height: 32, background: pal.accent, color: '#fff', borderRadius: 7, display: 'grid', placeItems: 'center', fontWeight: 700, fontSize: 12, letterSpacing: 0.5 }}>CGE</div>
        <div>
          <div style={{ fontSize: 15, fontWeight: 600, letterSpacing: -0.3 }}>Gestión de Oficios</div>
          <div style={{ fontSize: 11, color: pal.subtext }}>Comercial · Servicio al Cliente</div>
        </div>
        <div style={{ flex: 1 }} />
        <nav style={{ display: 'flex', gap: 2, background: pal.soft, padding: 3, borderRadius: 8 }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => go(t.id)}
              style={{
                background: screen === t.id || (t.id === 'inicio' && (screen === 'revaluar' || screen === 'multa')) ? pal.panel : 'transparent',
                boxShadow: screen === t.id || (t.id === 'inicio' && (screen === 'revaluar' || screen === 'multa')) ? '0 1px 2px rgba(0,0,0,0.06)' : 'none',
                border: 'none', padding: '7px 14px', borderRadius: 6, fontSize: 13, color: pal.text,
                fontWeight: screen === t.id ? 600 : 500, cursor: 'pointer', fontFamily: 'inherit',
              }}>
              {t.l}
            </button>
          ))}
        </nav>
        <button style={btn(pal, 'primary')} onClick={() => alert('Ejecutando análisis de 3 PDFs nuevos…')}>
          <span style={{ fontSize: 10 }}>▶</span> Ejecutar análisis
        </button>
      </div>
    </header>
  );
}

function AppFooter({ pal }) {
  return (
    <footer style={{ padding: '8px 28px', borderTop: `1px solid ${pal.border}`, background: pal.panel, display: 'flex', alignItems: 'center', gap: 10 }}>
      <span style={{ display: 'inline-block', width: 7, height: 7, borderRadius: 4, background: pal.success }} />
      <span style={{ fontSize: 11.5, color: pal.subtext }}>Completado: 6 PDFs procesados · hace 12 min</span>
      <div style={{ flex: 1 }} />
      <span style={{ fontSize: 10.5, color: pal.dim }}>v2.4.0</span>
    </footer>
  );
}

// ─── INICIO (bandeja) ──────────────────────────────────────────────────
const SAMPLE_OFICIOS = [
  { nro: 'S17172', tipo: 'Ordinario', area: 'Conexiones', plazo: '23-04-2026', diasRest: 2, asunto: 'Verifica cumplimiento de plazos de conexión cliente libre', multa: false, conf: 94, fecha: '17-02-2026', corregido: false },
  { nro: 'S16985', tipo: 'Ordinario', area: 'Lectura', plazo: '24-04-2026', diasRest: 3, asunto: 'Formula cargos a CGE por incumplimiento lectura medidores', multa: true, conf: 71, fecha: '18-02-2026', multaMonto: 250, corregido: false },
  { nro: 'S17170', tipo: 'Ordinario', area: 'Conexiones', plazo: '25-04-2026', diasRest: 4, asunto: 'Verifica cumplimiento instrucción reposición suministro', multa: true, conf: 78, fecha: '17-02-2026', multaMonto: 120, corregido: false },
  { nro: 'S16181', tipo: 'Ordinario', area: 'Conexiones', plazo: '24-04-2026', diasRest: 3, asunto: 'Otorga a CGE prórroga para entrega de antecedentes', multa: false, conf: 98, fecha: '15-02-2026', corregido: false },
  { nro: 'S16981', tipo: 'Ordinario', area: 'PMGD', plazo: '27-04-2026', diasRest: 6, asunto: 'Responde consulta pública sobre reglamento PMGD', multa: false, conf: 99, fecha: '19-02-2026', corregido: false },
  { nro: 'S17171', tipo: 'Ordinario', area: 'Conexiones', plazo: '28-04-2026', diasRest: 7, asunto: 'Informa la admisibilidad del reclamo y solicita respuesta', multa: false, conf: 97, fecha: '17-02-2026', corregido: false },
  { nro: 'S17205', tipo: 'Circular', area: 'Conexiones', plazo: '28-04-2026', diasRest: 7, asunto: 'Instruye a las empresas distribuidoras sobre procedimiento', multa: false, conf: 96, fecha: '20-02-2026', corregido: false },
  { nro: 'S14925', tipo: 'Ordinario', area: 'Conexiones', plazo: '02-05-2026', diasRest: 11, asunto: 'Da respuesta a denuncia cliente por calidad de servicio', multa: false, conf: 99, fecha: '05-02-2026', corregido: false },
];

function Inicio({ pal, go, density }) {
  const [tab, setTab] = React.useState('hoy');
  const [q, setQ] = React.useState('');

  const filtered = React.useMemo(() => {
    let list = SAMPLE_OFICIOS;
    if (tab === 'vencer') list = list.filter(o => o.diasRest <= 5);
    if (tab === 'multas') list = list.filter(o => o.multa);
    if (q) list = list.filter(o => (o.nro + o.asunto + o.area).toLowerCase().includes(q.toLowerCase()));
    return list;
  }, [tab, q]);

  const tabs = [
    { id: 'hoy', l: 'Hoy', n: SAMPLE_OFICIOS.length },
    { id: 'vencer', l: 'Por vencer', n: SAMPLE_OFICIOS.filter(o => o.diasRest <= 5).length, tone: 'warn' },
    { id: 'multas', l: 'Multas', n: SAMPLE_OFICIOS.filter(o => o.multa).length, tone: 'danger' },
    { id: 'todos', l: 'Histórico', n: 128 },
  ];

  return (
    <div style={{ padding: density === 'compact' ? '16px 24px 24px' : '20px 28px 28px' }}>
      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 18 }}>
        <Kpi pal={pal} v="128" l="Oficios totales" sub="+6 hoy" icon="▦" />
        <Kpi pal={pal} v="100%" l="Accuracy agente" sub="0 correcciones este mes" color={pal.success} icon="✓" />
        <Kpi pal={pal} v="10" l="Multas detectadas" sub="2 esta semana" color={pal.warn} icon="◆" />
        <Kpi pal={pal} v="3" l="Plazos críticos" sub="menos de 5 días" color={pal.danger} icon="⏱" />
      </div>

      {/* Alert */}
      <div style={{ background: pal.warnSoft, border: `1px solid ${pal.warn}33`, borderRadius: 10, padding: '12px 16px', display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
        <div style={{ width: 28, height: 28, borderRadius: 14, background: pal.warn, color: '#fff', display: 'grid', placeItems: 'center', fontSize: 14, fontWeight: 700 }}>!</div>
        <div style={{ flex: 1, fontSize: 13 }}>
          <div style={{ fontWeight: 600, color: pal.text }}>3 oficios con plazo en menos de 5 días</div>
          <div style={{ fontSize: 11.5, color: pal.subtext, marginTop: 2 }}>S17172 vence en 2d · S16985 en 3d (multa) · S16181 en 3d</div>
        </div>
        <button style={btn(pal, 'ghost')} onClick={() => setTab('vencer')}>Revisar</button>
      </div>

      {/* Tabs + search */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 4, borderBottom: `1px solid ${pal.border}`, marginBottom: 14 }}>
        {tabs.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)}
            style={{
              background: 'transparent', border: 'none', padding: '10px 14px',
              borderBottom: tab === t.id ? `2px solid ${pal.accent}` : '2px solid transparent',
              fontSize: 13, color: tab === t.id ? pal.text : pal.subtext,
              fontWeight: tab === t.id ? 600 : 500, cursor: 'pointer', display: 'flex', gap: 8, alignItems: 'center', fontFamily: 'inherit', marginBottom: -1,
            }}>
            {t.l}
            <span style={{
              background: t.tone === 'warn' ? pal.warnSoft : t.tone === 'danger' ? pal.dangerSoft : pal.soft,
              color: t.tone === 'warn' ? pal.warn : t.tone === 'danger' ? pal.danger : pal.subtext,
              fontSize: 10.5, padding: '1px 7px', borderRadius: 10, fontWeight: 600,
            }}>{t.n}</span>
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <div style={{ position: 'relative' }}>
          <input
            value={q} onChange={e => setQ(e.target.value)}
            placeholder="Buscar por nro, asunto, área…"
            style={{ border: `1px solid ${pal.border}`, background: pal.panel, color: pal.text,
              borderRadius: 6, padding: '6px 12px 6px 28px', fontSize: 12, width: 240, fontFamily: 'inherit', outline: 'none' }} />
          <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: pal.subtext, fontSize: 12 }}>⌕</span>
        </div>
      </div>

      {/* Grid */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 12 }}>
        {filtered.map(o => <OficioCard key={o.nro} pal={pal} o={o} go={go} density={density} />)}
        {filtered.length === 0 && (
          <div style={{ gridColumn: '1 / -1', padding: 40, textAlign: 'center', color: pal.subtext, fontSize: 13, border: `1px dashed ${pal.border}`, borderRadius: 10 }}>
            Sin oficios para este filtro.
          </div>
        )}
      </div>
    </div>
  );
}

function OficioCard({ pal, o, go, density }) {
  const tc = o.diasRest <= 3 ? pal.danger : o.diasRest <= 5 ? pal.warn : pal.success;
  const bg = o.diasRest <= 3 ? pal.dangerSoft : o.diasRest <= 5 ? pal.warnSoft : pal.successSoft;
  const areaKey = AREA_COLORS[o.area] || 'blue';
  const areaColor = pal[areaKey];
  const pad = density === 'compact' ? '12px 14px' : '14px 16px';

  return (
    <div style={{ background: pal.panel, border: `1px solid ${pal.border}`, borderRadius: 10, padding: pad,
      display: 'flex', flexDirection: 'column', gap: 8, transition: 'border-color 0.15s' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12, fontWeight: 600, color: pal.text }}>{o.nro}</div>
        <div style={{ fontSize: 11, color: pal.subtext }}>· {o.tipo}</div>
        <div style={{ flex: 1 }} />
        {o.multa && <span style={{ fontSize: 10, color: pal.warn, background: pal.warnSoft, padding: '2px 8px', borderRadius: 4, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.3 }}>Multa</span>}
        <span style={{ fontSize: 10.5, color: tc, background: bg, padding: '2px 8px', borderRadius: 4, fontWeight: 600 }}>{o.diasRest}d</span>
      </div>
      <div style={{ fontSize: 13, color: pal.text, lineHeight: 1.4, minHeight: 36 }}>{o.asunto}</div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, paddingTop: 8, borderTop: `1px solid ${pal.border}` }}>
        <span style={{ display: 'inline-block', width: 8, height: 8, borderRadius: 4, background: areaColor }} />
        <span style={{ fontSize: 11.5, color: pal.text }}>{o.area}</span>
        <span style={{ fontSize: 11, color: pal.subtext }}>· {o.plazo}</span>
        <div style={{ flex: 1 }} />
        <button style={btn(pal, 'ghost', 'sm')} onClick={() => go('revaluar', o)}>Revaluar</button>
        {o.multa && <button style={btn(pal, 'ghost', 'sm')} onClick={() => go('multa', o)}>Informe</button>}
      </div>
    </div>
  );
}

// ─── ESTADÍSTICAS ──────────────────────────────────────────────────────
function Stats({ pal, density }) {
  const areas = [
    { name: 'Conexiones', count: 56, pct: 43.8, tone: 'blue' },
    { name: 'PMGD', count: 23, pct: 18.0, tone: 'warn' },
    { name: 'Servicio al Cliente', count: 19, pct: 14.8, tone: 'success' },
    { name: 'Pérdidas', count: 13, pct: 10.2, tone: 'lilac' },
    { name: 'Sin área', count: 8, pct: 6.3, tone: 'neutral' },
    { name: 'Cobranza', count: 6, pct: 4.7, tone: 'danger' },
    { name: 'Lectura', count: 3, pct: 2.2, tone: 'teal' },
  ];
  const tipos = [
    { name: 'Oficio ordinario', count: 113, pct: 88.3 },
    { name: 'Oficio circular', count: 11, pct: 8.6 },
    { name: 'Resolución', count: 4, pct: 3.1 },
  ];
  const monthly = [
    { m: 'Oct', n: 18, multas: 1 }, { m: 'Nov', n: 22, multas: 2 }, { m: 'Dic', n: 15, multas: 1 },
    { m: 'Ene', n: 26, multas: 2 }, { m: 'Feb', n: 31, multas: 3 }, { m: 'Mar', n: 10, multas: 0 },
    { m: 'Abr', n: 6, multas: 1 },
  ];

  return (
    <div style={{ padding: density === 'compact' ? '16px 24px 24px' : '20px 28px 28px' }}>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 14, marginBottom: 18 }}>
        <h2 style={{ margin: 0, fontSize: 18, fontWeight: 600, letterSpacing: -0.3 }}>Estadísticas</h2>
        <div style={{ fontSize: 12, color: pal.subtext }}>Total 128 oficios · últimos 7 meses</div>
        <div style={{ flex: 1 }} />
        <select style={{ border: `1px solid ${pal.border}`, background: pal.panel, color: pal.text, borderRadius: 6, padding: '6px 10px', fontSize: 12, fontFamily: 'inherit' }}>
          <option>Últimos 7 meses</option><option>Último año</option><option>Todo el histórico</option>
        </select>
        <button style={btn(pal, 'ghost')}>⇩ Exportar</button>
      </div>

      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 16 }}>
        <Kpi pal={pal} v="128" l="Oficios totales" sub="+8% vs periodo anterior" />
        <Kpi pal={pal} v="100%" l="Accuracy agente" sub="0 correcciones" color={pal.success} />
        <Kpi pal={pal} v="7" l="Categorías activas" sub="1 nueva: Pérdidas" />
        <Kpi pal={pal} v="10" l="Multas detectadas" sub="Monto total: UTM 1.840" color={pal.warn} />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.3fr 1fr', gap: 12 }}>
        {/* Monthly trend */}
        <Panel pal={pal} title="Volumen mensual" subtitle="Oficios y multas por mes">
          <MonthlyChart pal={pal} data={monthly} />
        </Panel>

        {/* Tipos */}
        <Panel pal={pal} title="Por tipo de oficio">
          {tipos.map(t => (
            <div key={t.name} style={{ padding: '8px 0', borderBottom: `1px solid ${pal.border}` }}>
              <div style={{ display: 'flex', alignItems: 'baseline', fontSize: 12.5 }}>
                <div style={{ flex: 1 }}>{t.name}</div>
                <div style={{ fontVariantNumeric: 'tabular-nums', color: pal.subtext, width: 30, textAlign: 'right' }}>{t.count}</div>
                <div style={{ fontVariantNumeric: 'tabular-nums', color: pal.subtext, width: 48, textAlign: 'right' }}>{t.pct}%</div>
              </div>
              <div style={{ height: 4, background: pal.soft, borderRadius: 2, marginTop: 6, overflow: 'hidden' }}>
                <div style={{ width: `${t.pct}%`, height: '100%', background: pal.blue }} />
              </div>
            </div>
          ))}
        </Panel>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.3fr', gap: 12, marginTop: 12 }}>
        {/* Donut */}
        <Panel pal={pal} title="Distribución por área" subtitle="Responsables asignados por el agente">
          <div style={{ display: 'flex', alignItems: 'center', gap: 20, padding: '12px 0' }}>
            <Donut pal={pal} data={areas} />
            <div style={{ flex: 1 }}>
              {areas.map(a => (
                <div key={a.name} style={{ display: 'flex', alignItems: 'center', padding: '4px 0', fontSize: 11.5 }}>
                  <span style={{ width: 9, height: 9, borderRadius: 2, background: pal[a.tone], marginRight: 8 }} />
                  <div style={{ flex: 1 }}>{a.name}</div>
                  <div style={{ color: pal.subtext, width: 30, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{a.count}</div>
                  <div style={{ color: pal.subtext, width: 42, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{a.pct}%</div>
                </div>
              ))}
            </div>
          </div>
        </Panel>

        {/* Agent performance */}
        <Panel pal={pal} title="Desempeño del agente" subtitle="Últimos 30 días">
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10, marginBottom: 14 }}>
            <MiniStat pal={pal} v="2.3s" l="Tiempo promedio" />
            <MiniStat pal={pal} v="100%" l="Clasificación correcta" color={pal.success} />
            <MiniStat pal={pal} v="0" l="Correcciones manuales" color={pal.success} />
          </div>
          <div style={{ fontSize: 11, color: pal.subtext, marginBottom: 8, textTransform: 'uppercase', letterSpacing: 0.4 }}>Actividad reciente</div>
          {[
            { t: 'hace 12 min', m: 'Ejecutó análisis de 6 PDFs nuevos', c: pal.success },
            { t: 'hace 1 día', m: 'Detectó posible multa en S16985', c: pal.warn },
            { t: 'hace 2 días', m: 'Clasificó 4 oficios → Conexiones', c: pal.blue },
            { t: 'hace 3 días', m: 'Nueva categoría detectada: Pérdidas', c: pal.lilac },
          ].map((e, i) => (
            <div key={i} style={{ display: 'flex', gap: 10, padding: '6px 0', fontSize: 12, borderTop: i === 0 ? 'none' : `1px solid ${pal.border}` }}>
              <span style={{ width: 6, height: 6, borderRadius: 3, background: e.c, marginTop: 7 }} />
              <div style={{ flex: 1 }}>{e.m}</div>
              <div style={{ fontSize: 11, color: pal.subtext }}>{e.t}</div>
            </div>
          ))}
        </Panel>
      </div>
    </div>
  );
}

function MonthlyChart({ pal, data }) {
  const max = Math.max(...data.map(d => d.n));
  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'flex-end', gap: 10, height: 140, padding: '10px 0' }}>
        {data.map(d => (
          <div key={d.m} style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6 }}>
            <div style={{ fontSize: 10.5, color: pal.subtext, fontVariantNumeric: 'tabular-nums' }}>{d.n}</div>
            <div style={{ width: '100%', flex: 1, display: 'flex', alignItems: 'flex-end', position: 'relative' }}>
              <div style={{ width: '100%', height: `${(d.n / max) * 100}%`, background: pal.blue, borderRadius: '3px 3px 0 0', position: 'relative' }}>
                {d.multas > 0 && (
                  <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: `${(d.multas / d.n) * 100}%`, background: pal.warn, borderRadius: d.multas === d.n ? '3px 3px 0 0' : 0 }} />
                )}
              </div>
            </div>
            <div style={{ fontSize: 11, color: pal.subtext }}>{d.m}</div>
          </div>
        ))}
      </div>
      <div style={{ display: 'flex', gap: 14, fontSize: 11, color: pal.subtext, paddingTop: 8, borderTop: `1px solid ${pal.border}` }}>
        <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}><span style={{ width: 10, height: 10, background: pal.blue, borderRadius: 2 }} />Oficios</span>
        <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}><span style={{ width: 10, height: 10, background: pal.warn, borderRadius: 2 }} />Con multa</span>
      </div>
    </div>
  );
}

function Donut({ pal, data }) {
  const size = 150, r = 55, stroke = 22;
  const cx = size / 2, cy = size / 2;
  const C = 2 * Math.PI * r;
  let off = 0;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={pal.soft} strokeWidth={stroke} />
      {data.map((d, i) => {
        const len = (d.pct / 100) * C;
        const el = (
          <circle key={i} cx={cx} cy={cy} r={r} fill="none"
            stroke={pal[d.tone]} strokeWidth={stroke}
            strokeDasharray={`${len} ${C - len}`}
            strokeDashoffset={-off}
            transform={`rotate(-90 ${cx} ${cy})`} />
        );
        off += len;
        return el;
      })}
      <text x={cx} y={cy - 4} textAnchor="middle" fontSize="18" fontWeight="600" fill={pal.text} fontFamily="Inter">128</text>
      <text x={cx} y={cy + 12} textAnchor="middle" fontSize="10" fill={pal.subtext} fontFamily="Inter">oficios</text>
    </svg>
  );
}

// ─── REVALUAR ──────────────────────────────────────────────────────────
function Revaluar({ pal, go, oficio, density }) {
  const [sel, setSel] = React.useState(oficio || SAMPLE_OFICIOS[0]);
  const [newArea, setNewArea] = React.useState('');
  const [newPlazo, setNewPlazo] = React.useState('');
  const [multaCorrection, setMultaCorrection] = React.useState('');
  const [filter, setFilter] = React.useState('');

  const areas = Object.keys(AREA_COLORS);
  const filtered = SAMPLE_OFICIOS.filter(o => !filter || (o.nro + o.asunto).toLowerCase().includes(filter.toLowerCase()));

  const hasChanges = newArea || newPlazo || multaCorrection;

  return (
    <div style={{ padding: density === 'compact' ? '16px 24px' : '20px 28px' }}>
      <Breadcrumb pal={pal} go={go} current="Revaluar oficio" />

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.4fr', gap: 14 }}>
        {/* Lista */}
        <Panel pal={pal} title="Oficios registrados" subtitle="Selecciona uno para corregir">
          <input
            value={filter} onChange={e => setFilter(e.target.value)}
            placeholder="Filtrar…"
            style={{ border: `1px solid ${pal.border}`, background: pal.softer, color: pal.text,
              borderRadius: 6, padding: '6px 10px', fontSize: 12, width: '100%', marginBottom: 10, fontFamily: 'inherit', outline: 'none' }} />
          <div style={{ display: 'flex', flexDirection: 'column', gap: 1, maxHeight: 460, overflow: 'auto' }}>
            {filtered.map(o => {
              const active = sel.nro === o.nro;
              const areaKey = AREA_COLORS[o.area] || 'blue';
              return (
                <button key={o.nro} onClick={() => { setSel(o); setNewArea(''); setNewPlazo(''); setMultaCorrection(''); }}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 10, padding: '8px 10px',
                    background: active ? pal.blueSoft : 'transparent',
                    border: 'none', borderLeft: active ? `3px solid ${pal.accent}` : '3px solid transparent',
                    borderRadius: 4, cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit',
                    color: pal.text,
                  }}>
                  <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, fontWeight: 600, width: 60 }}>{o.nro}</span>
                  <span style={{ width: 6, height: 6, borderRadius: 3, background: pal[areaKey] }} />
                  <span style={{ flex: 1, fontSize: 11.5, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: pal.subtext }}>
                    {o.area}
                  </span>
                  <span style={{ fontSize: 10.5, color: o.diasRest <= 3 ? pal.danger : pal.subtext, fontWeight: o.diasRest <= 3 ? 600 : 400 }}>{o.diasRest}d</span>
                </button>
              );
            })}
          </div>
        </Panel>

        {/* Detalle + correcciones */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <Panel pal={pal} title={`Oficio ${sel.nro}`} subtitle={sel.tipo + ' · ' + sel.fecha}>
            <div style={{ fontSize: 14, lineHeight: 1.45, color: pal.text, marginBottom: 14 }}>{sel.asunto}</div>

            <div style={{ fontSize: 10.5, color: pal.subtext, textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 8 }}>Clasificación actual (agente)</div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
              <ReadOnly pal={pal} l="Área asignada" v={sel.area} color={pal[AREA_COLORS[sel.area] || 'blue']} />
              <ReadOnly pal={pal} l="Plazo respuesta" v={sel.plazo} sub={sel.diasRest + ' días restantes'} color={sel.diasRest <= 3 ? pal.danger : pal.text} />
              <ReadOnly pal={pal} l="Multa / cargos" v={sel.multa ? 'Sí, detectada' : 'No detectada'} color={sel.multa ? pal.warn : pal.success} />
            </div>
            <div style={{ marginTop: 10, padding: '8px 12px', background: pal.softer, borderRadius: 6, fontSize: 11.5, color: pal.subtext, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ color: pal.accent }}>◈</span>
              Confianza del agente: <b style={{ color: sel.conf >= 90 ? pal.success : sel.conf >= 80 ? pal.blue : pal.warn }}>{sel.conf}%</b>
              <span>·</span>
              <span>palabras clave: <i>{sel.area === 'Conexiones' ? 'conexión, plazo, cumplimiento' : sel.area === 'Lectura' ? 'lectura, medidor, cargos' : 'consulta, reglamento'}</i></span>
            </div>
          </Panel>

          <Panel pal={pal} title="Corregir asignación" subtitle="Los cambios se guardan en el Excel y reentrenan al agente.">
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field pal={pal} label="Nueva área responsable">
                <select value={newArea} onChange={e => setNewArea(e.target.value)} style={inputStyle(pal)}>
                  <option value="">(sin cambio)</option>
                  {areas.map(a => <option key={a} value={a}>{a}</option>)}
                </select>
              </Field>
              <Field pal={pal} label="Nuevo plazo (DD-MM-YYYY)">
                <input value={newPlazo} onChange={e => setNewPlazo(e.target.value)} placeholder="DD-MM-YYYY" style={inputStyle(pal)} />
              </Field>
            </div>

            <Field pal={pal} label="¿Es multa / formulación de cargos?" style={{ marginTop: 14 }}>
              <div style={{ display: 'flex', gap: 8 }}>
                {[['', '(sin cambio)'], ['si', 'Sí, es multa'], ['no', 'No es multa']].map(([v, l]) => (
                  <button key={v} onClick={() => setMultaCorrection(v)}
                    style={{
                      flex: 1, padding: '9px 12px', border: `1px solid ${multaCorrection === v ? pal.accent : pal.border}`,
                      background: multaCorrection === v ? pal.blueSoft : pal.panel, color: pal.text,
                      borderRadius: 6, fontSize: 12, cursor: 'pointer', fontWeight: multaCorrection === v ? 600 : 400, fontFamily: 'inherit',
                    }}>{l}</button>
                ))}
              </div>
            </Field>

            <div style={{ display: 'flex', gap: 10, marginTop: 18, paddingTop: 14, borderTop: `1px solid ${pal.border}` }}>
              <button style={btn(pal, 'ghost')} onClick={() => go('inicio')}>Cancelar</button>
              <div style={{ flex: 1 }} />
              {hasChanges && <span style={{ fontSize: 11.5, color: pal.warn, alignSelf: 'center' }}>● Cambios sin guardar</span>}
              <button style={{ ...btn(pal, 'primary'), opacity: hasChanges ? 1 : 0.5 }} disabled={!hasChanges} onClick={() => { alert('Corrección guardada'); go('inicio'); }}>
                ✓ Guardar corrección
              </button>
            </div>
          </Panel>
        </div>
      </div>
    </div>
  );
}

// ─── INFORME DE MULTA ──────────────────────────────────────────────────
function InformeMulta({ pal, go, oficio, density }) {
  const o = oficio || SAMPLE_OFICIOS.find(x => x.multa) || SAMPLE_OFICIOS[1];
  const monto = o.multaMonto || 250;

  return (
    <div style={{ padding: density === 'compact' ? '16px 24px' : '20px 28px' }}>
      <Breadcrumb pal={pal} go={go} current={`Informe de multa · ${o.nro}`} />

      {/* Banner */}
      <div style={{ background: `linear-gradient(135deg, ${pal.warn}15, ${pal.danger}15)`,
        border: `1px solid ${pal.warn}40`, borderRadius: 10, padding: '16px 20px',
        display: 'flex', alignItems: 'center', gap: 16, marginBottom: 16 }}>
        <div style={{ width: 42, height: 42, borderRadius: 10, background: pal.warn, color: '#fff', display: 'grid', placeItems: 'center', fontSize: 20 }}>◆</div>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 10.5, color: pal.subtext, textTransform: 'uppercase', letterSpacing: 0.8 }}>Formulación de cargos detectada</div>
          <div style={{ fontSize: 17, fontWeight: 600, marginTop: 2, letterSpacing: -0.3 }}>Oficio {o.nro} · {o.asunto}</div>
        </div>
        <div style={{ textAlign: 'right' }}>
          <div style={{ fontSize: 10.5, color: pal.subtext, textTransform: 'uppercase', letterSpacing: 0.8 }}>Monto estimado</div>
          <div style={{ fontSize: 22, fontWeight: 600, color: pal.warn, fontVariantNumeric: 'tabular-nums' }}>UTM {monto}</div>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: 14 }}>
        {/* Informe */}
        <Panel pal={pal} title="Informe generado" subtitle="Revísalo antes de exportar a PDF / Word.">
          <Section title="1. Identificación del oficio" pal={pal}>
            <GridKv pal={pal} items={[
              ['Número oficio', o.nro],
              ['Tipo', o.tipo],
              ['Fecha recepción', o.fecha],
              ['Plazo respuesta', `${o.plazo} (${o.diasRest} días)`],
              ['Área asignada', o.area],
              ['Remitente', 'SEC'],
            ]} />
          </Section>

          <Section title="2. Hechos imputados" pal={pal}>
            <div style={{ fontSize: 12.5, lineHeight: 1.5, color: pal.text }}>
              Se imputa a CGE el incumplimiento en la toma de lectura de medidores eléctricos correspondientes al ciclo facturación febrero 2026, en las comunas de San Bernardo, Rancagua y Temuco. Se detectaron <b>1.247 medidores</b> sin lectura dentro del plazo reglamentario establecido en el artículo 117° del DS 327.
            </div>
          </Section>

          <Section title="3. Normativa infringida" pal={pal}>
            <ul style={{ margin: 0, paddingLeft: 18, fontSize: 12.5, lineHeight: 1.6 }}>
              <li>Artículo 117° DS 327 — Obligación de lectura de medidores</li>
              <li>Artículo 225° DFL 4/2006 — Calidad de servicio</li>
              <li>Resolución Exenta SEC Nº 2841/2023</li>
            </ul>
          </Section>

          <Section title="4. Propuesta de descargos" pal={pal}>
            <div style={{ fontSize: 12.5, lineHeight: 1.5, color: pal.text }}>
              El agente propone: <b>(a)</b> adjuntar registros SAP con evidencia de rutas bloqueadas por contingencia climática en febrero, <b>(b)</b> demostrar cumplimiento posterior dentro de 5 días hábiles, <b>(c)</b> solicitar atenuante por ausencia de reincidencia en los últimos 24 meses.
            </div>
          </Section>
        </Panel>

        {/* Meta + acciones */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <Panel pal={pal} title="Detalles">
            <Kv pal={pal} l="Severidad estimada">
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontWeight: 600, color: pal.warn }}>
                <span style={{ width: 8, height: 8, borderRadius: 4, background: pal.warn }} /> Media
              </span>
            </Kv>
            <Kv pal={pal} l="Monto base">UTM 200</Kv>
            <Kv pal={pal} l="Agravantes">+25% (1.247 afectados)</Kv>
            <Kv pal={pal} l="Atenuantes">−0% (sin reincidencia)</Kv>
            <Kv pal={pal} l="Monto estimado"><b style={{ color: pal.warn, fontSize: 14 }}>UTM {monto}</b> ≈ $16.5M CLP</Kv>
          </Panel>

          <Panel pal={pal} title="Deadline">
            <div style={{ textAlign: 'center', padding: '14px 0' }}>
              <div style={{ fontSize: 44, fontWeight: 600, color: o.diasRest <= 3 ? pal.danger : pal.warn, letterSpacing: -1 }}>{o.diasRest}</div>
              <div style={{ fontSize: 12, color: pal.subtext, marginTop: 2 }}>días para responder a SEC</div>
              <div style={{ fontSize: 11, color: pal.subtext, marginTop: 8, fontFamily: 'JetBrains Mono, monospace' }}>Vence {o.plazo}</div>
            </div>
          </Panel>

          <Panel pal={pal}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              <button style={btn(pal, 'primary')}>⇩ Exportar informe (Word)</button>
              <button style={btn(pal, 'ghost')}>⇩ Exportar informe (PDF)</button>
              <button style={btn(pal, 'ghost')}>✉ Enviar al responsable</button>
              <button style={{ ...btn(pal, 'ghost'), color: pal.danger }} onClick={() => go('revaluar', o)}>✎ No es multa · revaluar</button>
            </div>
          </Panel>
        </div>
      </div>
    </div>
  );
}

// ─── HELPERS ───────────────────────────────────────────────────────────
function Panel({ pal, title, subtitle, children }) {
  return (
    <div style={{ background: pal.panel, border: `1px solid ${pal.border}`, borderRadius: 10, padding: 16 }}>
      {title && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 13, fontWeight: 600 }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11.5, color: pal.subtext, marginTop: 2 }}>{subtitle}</div>}
        </div>
      )}
      {children}
    </div>
  );
}

function Kpi({ v, l, sub, color, pal, icon }) {
  return (
    <div style={{ background: pal.panel, border: `1px solid ${pal.border}`, borderRadius: 10, padding: '14px 16px', position: 'relative' }}>
      {icon && <div style={{ position: 'absolute', top: 14, right: 14, fontSize: 18, color: pal.dim, opacity: 0.5 }}>{icon}</div>}
      <div style={{ fontSize: 11.5, color: pal.subtext }}>{l}</div>
      <div style={{ fontSize: 28, fontWeight: 600, color: color || pal.text, letterSpacing: -0.5, fontVariantNumeric: 'tabular-nums', marginTop: 2 }}>{v}</div>
      <div style={{ fontSize: 11, color: color || pal.subtext, marginTop: 2 }}>{sub}</div>
    </div>
  );
}

function MiniStat({ v, l, color, pal }) {
  return (
    <div style={{ padding: '10px 12px', background: pal.softer, borderRadius: 6 }}>
      <div style={{ fontSize: 18, fontWeight: 600, color: color || pal.text, fontVariantNumeric: 'tabular-nums' }}>{v}</div>
      <div style={{ fontSize: 10.5, color: pal.subtext, marginTop: 2 }}>{l}</div>
    </div>
  );
}

function ReadOnly({ pal, l, v, sub, color }) {
  return (
    <div style={{ padding: '10px 12px', background: pal.softer, borderRadius: 6 }}>
      <div style={{ fontSize: 10.5, color: pal.subtext, textTransform: 'uppercase', letterSpacing: 0.4 }}>{l}</div>
      <div style={{ fontSize: 13.5, fontWeight: 600, color: color || pal.text, marginTop: 3 }}>{v}</div>
      {sub && <div style={{ fontSize: 10.5, color: pal.subtext, marginTop: 1 }}>{sub}</div>}
    </div>
  );
}

function Field({ pal, label, children, style }) {
  return (
    <div style={style}>
      <div style={{ fontSize: 11, color: pal.subtext, marginBottom: 5, fontWeight: 500 }}>{label}</div>
      {children}
    </div>
  );
}

function Section({ title, pal, children }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ fontSize: 11, color: pal.subtext, textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 8, fontWeight: 600 }}>{title}</div>
      {children}
    </div>
  );
}

function GridKv({ pal, items }) {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '10px 20px' }}>
      {items.map(([l, v]) => (
        <div key={l} style={{ display: 'flex', gap: 8, fontSize: 12.5 }}>
          <div style={{ color: pal.subtext, minWidth: 110 }}>{l}</div>
          <div style={{ color: pal.text, fontWeight: 500 }}>{v}</div>
        </div>
      ))}
    </div>
  );
}

function Kv({ pal, l, children }) {
  return (
    <div style={{ display: 'flex', alignItems: 'baseline', padding: '7px 0', borderBottom: `1px solid ${pal.border}`, fontSize: 12.5 }}>
      <div style={{ color: pal.subtext, flex: 1 }}>{l}</div>
      <div>{children}</div>
    </div>
  );
}

function Breadcrumb({ pal, go, current }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, marginBottom: 14 }}>
      <button onClick={() => go('inicio')} style={{ background: 'transparent', border: 'none', color: pal.blue, cursor: 'pointer', padding: 0, fontFamily: 'inherit', fontSize: 12 }}>
        ← Bandeja
      </button>
      <span style={{ color: pal.dim }}>/</span>
      <span style={{ color: pal.text, fontWeight: 500 }}>{current}</span>
    </div>
  );
}

function inputStyle(pal) {
  return {
    border: `1px solid ${pal.border}`, background: pal.panel, color: pal.text,
    borderRadius: 6, padding: '8px 10px', fontSize: 12.5, width: '100%',
    fontFamily: 'inherit', outline: 'none',
  };
}

function btn(pal, kind = 'primary', size = 'md') {
  const pad = size === 'sm' ? '5px 10px' : '8px 14px';
  const fs = size === 'sm' ? 11.5 : 12.5;
  if (kind === 'primary') {
    return { background: pal.accent, color: '#fff', border: 'none', padding: pad, borderRadius: 6, fontSize: fs, fontWeight: 500, cursor: 'pointer', fontFamily: 'inherit', display: 'inline-flex', alignItems: 'center', gap: 6 };
  }
  return { background: pal.panel, color: pal.text, border: `1px solid ${pal.border}`, padding: pad, borderRadius: 6, fontSize: fs, fontWeight: 500, cursor: 'pointer', fontFamily: 'inherit', display: 'inline-flex', alignItems: 'center', gap: 6 };
}

window.V5App = V5App;
