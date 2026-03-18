import React, { useState, useRef, useCallback } from 'react'
import * as XLSX from 'xlsx'
import {
  ComposedChart, BarChart, LineChart, AreaChart,
  Bar, Line, Area, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer
} from 'recharts'

// ==================== 定数 ====================
const COLORS = {
  bg: '#0f1117',
  card: '#1a1d27',
  border: '#2a2e3d',
  text: '#e8eaed',
  sub: '#8b8fa3',
  accent1: '#4ecdc4',
  accent2: '#ff6b6b',
  accent3: '#ffd93d',
  accent4: '#6c5ce7',
  up: '#00c853',
  down: '#ff5252',
  tooltip: '#1e2130',
}

const COOP_NAMES = { 2: 'アイチョイス', 6: '岐阜', 7: '一宮' }
const COOP_COLORS = { 2: '#4ecdc4', 6: '#ffd93d', 7: '#6c5ce7' }

// ==================== ユーティリティ ====================
function fmtNum(n) {
  if (n == null || isNaN(n)) return '-'
  if (Math.abs(n) >= 1e8) return (n / 1e8).toFixed(1) + '億'
  if (Math.abs(n) >= 1e4) return Math.round(n / 1e4) + '万'
  return Math.round(n).toLocaleString()
}

function fmtPct(val, base) {
  if (!base || base === 0) return null
  const pct = ((val - base) / base) * 100
  return pct
}

function YoY({ cur, prev, isPt }) {
  if (prev == null || cur == null) return null
  const diff = isPt ? (cur - prev) : fmtPct(cur, prev)
  if (diff == null) return null
  const up = diff >= 0
  const color = up ? COLORS.up : COLORS.down
  const label = isPt
    ? (up ? '+' : '') + diff.toFixed(1) + 'pt'
    : (up ? '+' : '') + diff.toFixed(1) + '%'
  return <span style={{ color, fontSize: 12, fontWeight: 700 }}>{label}</span>
}

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null
  return (
    <div style={{
      background: COLORS.tooltip, border: `1px solid ${COLORS.border}`,
      borderRadius: 8, padding: '10px 14px', fontSize: 12, color: COLORS.text,
      boxShadow: '0 8px 32px rgba(0,0,0,0.4)'
    }}>
      <div style={{ marginBottom: 4, fontWeight: 700 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color }}>
          {p.name}: {typeof p.value === 'number' ? (p.value % 1 === 0 ? p.value.toLocaleString() : p.value.toFixed(2)) : p.value}
        </div>
      ))}
    </div>
  )
}

// ==================== データ処理 ====================
function processData(rows) {
  const filtered = rows.filter(r => r['区分'] === 1)

  // yearlyData
  const yearMap = {}
  filtered.forEach(r => {
    const y = r['対象年度']
    if (!yearMap[y]) yearMap[y] = { rows: [], 食材供給高Sum: 0, 実質GPSum: 0 }
    yearMap[y].rows.push(r)
    yearMap[y].食材供給高Sum += (r['食材供給高'] || 0)
    yearMap[y].実質GPSum += (r['実質GP'] || 0)
  })

  const yearlyData = Object.keys(yearMap).map(y => {
    const d = yearMap[y]
    const rs = d.rows
    const mean = (key) => rs.reduce((a, r) => a + (r[key] || 0), 0) / rs.length
    return {
      年度: Number(y),
      食材供給高: d.食材供給高Sum,
      実質GP: d.実質GPSum,
      食材利用人数: mean('食材利用人数'),
      食材金額割合: mean('食材金額割合'),
      食材受注人数割合: mean('食材受注人数割合'),
      食材一人当利用高: mean('食材一人当利用高'),
      食材一点当単価: mean('食材一点当単価'),
      食材一人当点数: mean('食材一人当点数'),
      rowCount: rs.length,
    }
  }).sort((a, b) => a.年度 - b.年度)

  // coopYearlyData
  const coopYearMap = {}
  filtered.forEach(r => {
    const key = `${r['対象年度']}_${r['生協コード']}`
    if (!coopYearMap[key]) coopYearMap[key] = {
      年度: r['対象年度'], 生協コード: r['生協コード'], rows: [],
      食材供給高Sum: 0, 実質GPSum: 0
    }
    coopYearMap[key].rows.push(r)
    coopYearMap[key].食材供給高Sum += (r['食材供給高'] || 0)
    coopYearMap[key].実質GPSum += (r['実質GP'] || 0)
  })

  const coopYearlyData = Object.values(coopYearMap).map(d => {
    const rs = d.rows
    const mean = (key) => rs.reduce((a, r) => a + (r[key] || 0), 0) / rs.length
    return {
      年度: d.年度,
      生協コード: d.生協コード,
      生協名: COOP_NAMES[d.生協コード] || `生協${d.生協コード}`,
      食材供給高: d.食材供給高Sum,
      実質GP: d.実質GPSum,
      食材利用人数: mean('食材利用人数'),
      食材金額割合: mean('食材金額割合'),
      食材受注人数割合: mean('食材受注人数割合'),
      食材一人当利用高: mean('食材一人当利用高'),
      食材一点当単価: mean('食材一点当単価'),
      食材一人当点数: mean('食材一人当点数'),
    }
  }).sort((a, b) => a.年度 - b.年度 || a.生協コード - b.生協コード)

  // weeklyData
  const weeklyData = filtered.map(r => ({ ...r }))

  // latestFullYear判定
  const maxRows = Math.max(...yearlyData.map(d => d.rowCount))
  const sortedYears = yearlyData.map(d => d.年度).sort((a, b) => a - b)
  const latestYear = sortedYears[sortedYears.length - 1]
  const latestData = yearlyData.find(d => d.年度 === latestYear)
  const latestFullYear = (latestData && latestData.rowCount < maxRows / 2)
    ? sortedYears[sortedYears.length - 2]
    : latestYear

  return { yearlyData, coopYearlyData, weeklyData, latestFullYear }
}

// ==================== コンポーネント ====================
function KPICard({ label, value, sub, color, prev, prevVal, isPt }) {
  const cur = typeof value === 'number' ? value : null
  return (
    <div style={{
      background: COLORS.card, borderRadius: 12, padding: '20px 22px',
      borderTop: `3px solid ${color}`, flex: '1 1 200px', minWidth: 180,
    }}>
      <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: COLORS.text, marginBottom: 4 }}>
        {value}
      </div>
      <div style={{ fontSize: 11, color: COLORS.sub, marginBottom: 6 }}>{sub}</div>
      {prev != null && cur != null && (
        <YoY cur={cur} prev={prevVal} isPt={isPt} />
      )}
    </div>
  )
}

function ChartCard({ title, children }) {
  return (
    <div style={{
      background: COLORS.card, borderRadius: 12,
      border: `1px solid ${COLORS.border}`, padding: 20, marginBottom: 24
    }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.sub, marginBottom: 16 }}>{title}</div>
      {children}
    </div>
  )
}

function SectionTitle({ icon, title }) {
  return (
    <div style={{ fontSize: 17, fontWeight: 700, color: COLORS.text, margin: '32px 0 16px', display: 'flex', alignItems: 'center', gap: 8 }}>
      <span>{icon}</span>{title}
    </div>
  )
}

// ==================== タブ1: 概況 ====================
function Tab1Overview({ yearlyData, latestFullYear }) {
  const cur = yearlyData.find(d => d.年度 === latestFullYear)
  const years = yearlyData.map(d => d.年度).sort((a, b) => a - b)
  const prevYear = years[years.indexOf(latestFullYear) - 1]
  const prev = yearlyData.find(d => d.年度 === prevYear)

  if (!cur) return <div style={{ color: COLORS.sub }}>データなし</div>

  const firstYear = years[0]
  const firstData = yearlyData.find(d => d.年度 === firstYear)

  const penetrationGrowth = firstData && firstData.食材金額割合 > 0
    ? (cur.食材金額割合 / firstData.食材金額割合).toFixed(1)
    : null
  const userGrowth = firstData && firstData.食材利用人数 > 0
    ? Math.round(((cur.食材利用人数 - firstData.食材利用人数) / firstData.食材利用人数) * 100)
    : null
  const priceGrowth = firstData && firstData.食材一人当利用高 > 0
    ? Math.round(((cur.食材一人当利用高 - firstData.食材一人当利用高) / firstData.食材一人当利用高) * 100)
    : null

  return (
    <div>
      <SectionTitle icon="📊" title={`主要KPI（${latestFullYear}年度）`} />

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 24 }}>
        <KPICard
          label="食材供給高" value={fmtNum(cur.食材供給高)} sub="年間合計" color={COLORS.accent1}
          prev={prev} prevVal={prev?.食材供給高} isPt={false}
        />
        <KPICard
          label="食材利用人数（平均）" value={Math.round(cur.食材利用人数).toLocaleString() + '人'} sub="週平均" color={COLORS.accent3}
          prev={prev} prevVal={prev?.食材利用人数} isPt={false}
        />
        <KPICard
          label="食材金額割合" value={cur.食材金額割合.toFixed(1) + '%'} sub="総受注高に占める比率" color={COLORS.accent4}
          prev={prev} prevVal={prev?.食材金額割合} isPt={true}
        />
        <KPICard
          label="実質GP" value={fmtNum(cur.実質GP)} sub="年間合計" color={COLORS.accent2}
          prev={prev} prevVal={prev?.実質GP} isPt={false}
        />
      </div>

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 32 }}>
        <KPICard
          label="食材一人当利用高" value={Math.round(cur.食材一人当利用高).toLocaleString() + '円'} sub="客単価" color={COLORS.accent1}
          prev={prev} prevVal={prev?.食材一人当利用高} isPt={false}
        />
        <KPICard
          label="食材一点当単価" value={Math.round(cur.食材一点当単価).toLocaleString() + '円'} sub="商品単価" color={COLORS.accent3}
          prev={prev} prevVal={prev?.食材一点当単価} isPt={false}
        />
        <KPICard
          label="食材一人当点数" value={cur.食材一人当点数.toFixed(2) + '点'} sub="購入点数/人" color={COLORS.accent4}
          prev={prev} prevVal={prev?.食材一人当点数} isPt={false}
        />
        <KPICard
          label="食材受注人数割合" value={cur.食材受注人数割合.toFixed(1) + '%'} sub="総利用者に占める比率" color={COLORS.accent2}
          prev={prev} prevVal={prev?.食材受注人数割合} isPt={true}
        />
      </div>

      <div style={{
        background: 'linear-gradient(135deg, #1e2a35, #1a1d27)',
        borderRadius: 12, padding: '24px 28px', border: `1px solid ${COLORS.border}`
      }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: COLORS.accent1, marginBottom: 16 }}>
          📋 マーケター注目ポイント
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {penetrationGrowth && (
            <div style={{ fontSize: 14, color: COLORS.text }}>
              📈 <strong>浸透率の成長</strong>：{firstYear}年度から{latestFullYear}年度にかけて食材金額割合は
              <span style={{ color: COLORS.accent1 }}> {penetrationGrowth}倍</span>に成長
              （{firstData.食材金額割合.toFixed(1)}% → {cur.食材金額割合.toFixed(1)}%）
            </div>
          )}
          {userGrowth != null && (
            <div style={{ fontSize: 14, color: COLORS.text }}>
              👥 <strong>利用人数の変化</strong>：{firstYear}年度比で週平均利用人数が
              <span style={{ color: userGrowth >= 0 ? COLORS.up : COLORS.down }}> {userGrowth >= 0 ? '+' : ''}{userGrowth}%</span> 変化
              （{Math.round(firstData.食材利用人数).toLocaleString()}人 → {Math.round(cur.食材利用人数).toLocaleString()}人）
            </div>
          )}
          {priceGrowth != null && (
            <div style={{ fontSize: 14, color: COLORS.text }}>
              💰 <strong>客単価の変化</strong>：一人当利用高が{firstYear}年度比
              <span style={{ color: priceGrowth >= 0 ? COLORS.up : COLORS.down }}> {priceGrowth >= 0 ? '+' : ''}{priceGrowth}%</span>
              （¥{Math.round(firstData.食材一人当利用高).toLocaleString()} → ¥{Math.round(cur.食材一人当利用高).toLocaleString()}）
            </div>
          )}
          <div style={{ fontSize: 14, color: COLORS.text }}>
            🎯 <strong>次のアクション提案</strong>：
            {cur.食材金額割合 < 10
              ? '食材浸透率がまだ低水準のため、ターゲット層への積極的なプロモーションが有効です'
              : cur.食材金額割合 < 20
              ? '浸透率が成長段階にあります。リピーター育成と新規開拓を並行して進めましょう'
              : '高い浸透率を維持しながら、客単価向上（高付加価値メニュー展開）を検討する段階です'}
          </div>
        </div>
      </div>
    </div>
  )
}

// ==================== タブ2: 成長推移 ====================
function Tab2Growth({ yearlyData }) {
  return (
    <div>
      <SectionTitle icon="📈" title="食材供給高 & 実質GP 年度推移" />
      <ChartCard title="年度別 食材供給高・実質GP">
        <ResponsiveContainer width="100%" height={320}>
          <ComposedChart data={yearlyData} margin={{ top: 10, right: 30, left: 20, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="年度" stroke={COLORS.sub} tick={{ fontSize: 12, fill: COLORS.sub }} />
            <YAxis yAxisId="left" stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }}
              tickFormatter={v => fmtNum(v)} />
            <YAxis yAxisId="right" orientation="right" stroke={COLORS.sub}
              tick={{ fontSize: 11, fill: COLORS.sub }} tickFormatter={v => fmtNum(v)} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="食材供給高" name="食材供給高" fill={COLORS.accent1} opacity={0.85} />
            <Bar yAxisId="right" dataKey="実質GP" name="実質GP" fill={COLORS.accent2} opacity={0.85} />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📈" title="浸透率（食材金額割合・受注人数割合）推移" />
      <ChartCard title="年度別 浸透率推移">
        <ResponsiveContainer width="100%" height={300}>
          <LineChart data={yearlyData} margin={{ top: 10, right: 30, left: 0, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="年度" stroke={COLORS.sub} tick={{ fontSize: 12, fill: COLORS.sub }} />
            <YAxis stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }}
              tickFormatter={v => v.toFixed(1) + '%'} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />
            <Line dataKey="食材金額割合" name="食材金額割合(%)" stroke={COLORS.accent4}
              strokeWidth={2} dot={{ r: 4 }} />
            <Line dataKey="食材受注人数割合" name="食材受注人数割合(%)" stroke={COLORS.accent3}
              strokeWidth={2} dot={{ r: 4 }} />
          </LineChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  )
}

// ==================== タブ3: 生協比較 ====================
function Tab3Coop({ coopYearlyData, latestFullYear }) {
  const years = [...new Set(coopYearlyData.map(d => d.年度))].sort((a, b) => a - b)
  const coops = [...new Set(coopYearlyData.map(d => d.生協コード))]

  const latestCoopData = coopYearlyData.filter(d => d.年度 === latestFullYear && d.食材供給高 > 0)

  // グラフ3用データ（年度×生協）
  const supplyChartData = years.map(y => {
    const obj = { 年度: y }
    coops.forEach(c => {
      const found = coopYearlyData.find(d => d.年度 === y && d.生協コード === c)
      obj[COOP_NAMES[c] || `生協${c}`] = found ? found.食材供給高 : 0
    })
    return obj
  })

  // グラフ4用データ（浸透率）
  const penetrationChartData = years.map(y => {
    const obj = { 年度: y }
    coops.forEach(c => {
      const found = coopYearlyData.find(d => d.年度 === y && d.生協コード === c)
      obj[COOP_NAMES[c] || `生協${c}`] = found && found.食材供給高 > 0 ? found.食材金額割合 : null
    })
    return obj
  })

  return (
    <div>
      <SectionTitle icon="🏢" title={`生協別KPI（${latestFullYear}年度）`} />
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 32 }}>
        {latestCoopData.map(d => (
          <div key={d.生協コード} style={{
            background: COLORS.card, borderRadius: 12, padding: '20px 24px',
            border: `1px solid ${COLORS.border}`, flex: '1 1 200px', minWidth: 220,
            borderTop: `3px solid ${COOP_COLORS[d.生協コード] || COLORS.accent1}`
          }}>
            <div style={{ fontSize: 15, fontWeight: 700, color: COOP_COLORS[d.生協コード] || COLORS.accent1, marginBottom: 12 }}>
              {d.生協名}
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              {[
                ['食材供給高', fmtNum(d.食材供給高)],
                ['食材利用人数', Math.round(d.食材利用人数).toLocaleString() + '人'],
                ['食材金額割合', d.食材金額割合.toFixed(1) + '%'],
                ['一人当利用高', '¥' + Math.round(d.食材一人当利用高).toLocaleString()],
              ].map(([k, v]) => (
                <div key={k}>
                  <div style={{ fontSize: 10, color: COLORS.sub }}>{k}</div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: COLORS.text }}>{v}</div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      <SectionTitle icon="📊" title="生協別 食材供給高 年度推移" />
      <ChartCard title="生協別 食材供給高（年度別）">
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={supplyChartData} margin={{ top: 10, right: 30, left: 20, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="年度" stroke={COLORS.sub} tick={{ fontSize: 12, fill: COLORS.sub }} />
            <YAxis stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} tickFormatter={v => fmtNum(v)} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />
            {coops.map(c => (
              <Bar key={c} dataKey={COOP_NAMES[c] || `生協${c}`}
                fill={COOP_COLORS[c] || COLORS.accent1} opacity={0.85} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📈" title="生協別 食材金額割合（浸透率）推移" />
      <ChartCard title="生協別 食材金額割合（%）">
        <ResponsiveContainer width="100%" height={300}>
          <LineChart data={penetrationChartData} margin={{ top: 10, right: 30, left: 0, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="年度" stroke={COLORS.sub} tick={{ fontSize: 12, fill: COLORS.sub }} />
            <YAxis stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }}
              tickFormatter={v => v != null ? v.toFixed(1) + '%' : ''} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />
            {coops.map(c => (
              <Line key={c} dataKey={COOP_NAMES[c] || `生協${c}`}
                stroke={COOP_COLORS[c] || COLORS.accent1}
                strokeWidth={2} dot={{ r: 4 }} connectNulls={false} />
            ))}
          </LineChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  )
}

// ==================== タブ4: 週次分析 ====================
function Tab4Weekly({ weeklyData, yearlyData }) {
  const years = [...new Set(weeklyData.map(r => r['対象年度']))].sort((a, b) => a - b)
  const coopCodes = [...new Set(weeklyData.map(r => r['生協コード']))]

  const [selYear, setSelYear] = useState(years[years.length - 1])
  const [selCoop, setSelCoop] = useState(2)

  const filtered = weeklyData
    .filter(r => r['対象年度'] === selYear && r['生協コード'] === selCoop)
    .sort((a, b) => a['SEQ'] - b['SEQ'])
    .map(r => ({
      週: r['配送週'] || `SEQ${r['SEQ']}`,
      SEQ: r['SEQ'],
      食材供給高: r['食材供給高'] || 0,
      食材利用人数: r['食材利用人数'] || 0,
      食材金額割合: r['食材金額割合'] || 0,
      食材受注人数割合: r['食材受注人数割合'] || 0,
    }))

  const btnStyle = (active) => ({
    padding: '6px 16px', borderRadius: 6, border: 'none', cursor: 'pointer',
    fontSize: 13, fontWeight: 600,
    background: active ? COLORS.accent1 : COLORS.card,
    color: active ? '#000' : COLORS.sub,
  })

  return (
    <div>
      <SectionTitle icon="🔍" title="フィルター" />
      <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap', marginBottom: 24 }}>
        <div>
          <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 8 }}>年度</div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {years.map(y => (
              <button key={y} style={btnStyle(selYear === y)} onClick={() => setSelYear(y)}>{y}年度</button>
            ))}
          </div>
        </div>
        <div>
          <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 8 }}>生協</div>
          <div style={{ display: 'flex', gap: 8 }}>
            {coopCodes.map(c => (
              <button key={c} style={btnStyle(selCoop === c)} onClick={() => setSelCoop(c)}>
                {COOP_NAMES[c] || `生協${c}`}
              </button>
            ))}
          </div>
        </div>
      </div>

      <SectionTitle icon="📈" title="週次 食材供給高" />
      <ChartCard title={`${selYear}年度 ${COOP_NAMES[selCoop] || `生協${selCoop}`} 週次供給高`}>
        <ResponsiveContainer width="100%" height={300}>
          <ComposedChart data={filtered} margin={{ top: 10, right: 30, left: 20, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="SEQ" stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} />
            <YAxis stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} tickFormatter={v => fmtNum(v)} />
            <Tooltip content={<CustomTooltip />} />
            <Area dataKey="食材供給高" name="食材供給高" fill={COLORS.accent1} stroke={COLORS.accent1}
              fillOpacity={0.3} strokeWidth={2} />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="👥" title="週次 食材利用人数" />
      <ChartCard title={`${selYear}年度 週次利用人数`}>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={filtered} margin={{ top: 10, right: 30, left: 0, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="SEQ" stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} />
            <YAxis stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }}
              tickFormatter={v => v.toLocaleString()} />
            <Tooltip content={<CustomTooltip />} />
            <Line dataKey="食材利用人数" name="食材利用人数（人）" stroke={COLORS.accent3}
              strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📉" title="週次 浸透率" />
      <ChartCard title={`${selYear}年度 週次浸透率`}>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={filtered} margin={{ top: 10, right: 30, left: 0, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="SEQ" stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} />
            <YAxis stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }}
              tickFormatter={v => v.toFixed(1) + '%'} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />
            <Line dataKey="食材金額割合" name="食材金額割合(%)" stroke={COLORS.accent4}
              strokeWidth={2} dot={false} />
            <Line dataKey="食材受注人数割合" name="食材受注人数割合(%)" stroke={COLORS.accent2}
              strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  )
}

// ==================== タブ5: KPI深掘り ====================
function Tab5KPI({ yearlyData }) {
  const tableYears = [...yearlyData].sort((a, b) => a.年度 - b.年度)

  return (
    <div>
      <SectionTitle icon="🎯" title="客単価（一人当利用高）の分解" />
      <div style={{ background: COLORS.card, borderRadius: 12, padding: '16px 20px', marginBottom: 24, color: COLORS.text, fontSize: 14, lineHeight: 1.8 }}>
        <strong style={{ color: COLORS.accent1 }}>一人当利用高</strong> ＝
        <strong style={{ color: COLORS.accent3 }}> 一点当単価</strong> ×
        <strong style={{ color: COLORS.accent4 }}> 一人当点数</strong>
        <br />
        客単価を向上させるには「高価格帯メニューの展開（単価↑）」または「購入点数の増加（クロスセル↑）」が有効です。
      </div>

      <ChartCard title="年度別 一人当利用高・一点当単価・一人当点数">
        <ResponsiveContainer width="100%" height={320}>
          <ComposedChart data={yearlyData} margin={{ top: 10, right: 60, left: 20, bottom: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
            <XAxis dataKey="年度" stroke={COLORS.sub} tick={{ fontSize: 12, fill: COLORS.sub }} />
            <YAxis yAxisId="left" stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }}
              tickFormatter={v => '¥' + v.toLocaleString()} />
            <YAxis yAxisId="right" orientation="right" stroke={COLORS.sub}
              tick={{ fontSize: 11, fill: COLORS.sub }} tickFormatter={v => v.toFixed(2) + '点'} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="食材一人当利用高" name="一人当利用高(円)" fill={COLORS.accent1} opacity={0.85} />
            <Line yAxisId="left" dataKey="食材一点当単価" name="一点当単価(円)" stroke={COLORS.accent3}
              strokeWidth={2} dot={{ r: 4 }} />
            <Line yAxisId="right" dataKey="食材一人当点数" name="一人当点数" stroke={COLORS.accent4}
              strokeWidth={2} dot={{ r: 4 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📋" title="年度別KPI一覧" />
      <div style={{ overflowX: 'auto', marginBottom: 24 }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
              {['年度', '食材供給高', '食材利用人数\n(週平均)', '金額割合(%)', '受注人数割合(%)', '一人当利用高', '一点当単価', '一人当点数', '実質GP'].map(h => (
                <th key={h} style={{ padding: '10px 12px', color: COLORS.sub, textAlign: 'right', whiteSpace: 'pre' }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tableYears.map((d, i) => {
              const prev = tableYears[i - 1]
              return (
                <tr key={d.年度} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                  <td style={{ padding: '10px 12px', color: COLORS.accent1, fontWeight: 700 }}>{d.年度}</td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    {fmtNum(d.食材供給高)}
                    {prev && <div><YoY cur={d.食材供給高} prev={prev.食材供給高} /></div>}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    {Math.round(d.食材利用人数).toLocaleString()}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    {d.食材金額割合.toFixed(1)}
                    {prev && <div><YoY cur={d.食材金額割合} prev={prev.食材金額割合} isPt /></div>}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    {d.食材受注人数割合.toFixed(1)}
                    {prev && <div><YoY cur={d.食材受注人数割合} prev={prev.食材受注人数割合} isPt /></div>}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    ¥{Math.round(d.食材一人当利用高).toLocaleString()}
                    {prev && <div><YoY cur={d.食材一人当利用高} prev={prev.食材一人当利用高} /></div>}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    ¥{Math.round(d.食材一点当単価).toLocaleString()}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    {d.食材一人当点数.toFixed(2)}
                  </td>
                  <td style={{ padding: '10px 12px', color: COLORS.text, textAlign: 'right' }}>
                    {fmtNum(d.実質GP)}
                    {prev && <div><YoY cur={d.実質GP} prev={prev.実質GP} /></div>}
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      <SectionTitle icon="💡" title="分析インサイト" />
      <div style={{ background: COLORS.card, borderRadius: 12, padding: '16px 20px', border: `1px solid ${COLORS.border}` }}>
        {tableYears.length >= 2 && (() => {
          const first = tableYears[0]
          const last = tableYears[tableYears.length - 1]
          const priceChange = ((last.食材一人当利用高 - first.食材一人当利用高) / first.食材一人当利用高* 100).toFixed(1)
          const unitChange = ((last.食材一点当単価 - first.食材一点当単価) / first.食材一点当単価 * 100).toFixed(1)
          const countChange = (last.食材一人当点数 - first.食材一人当点数).toFixed(2)
          return (
            <div style={{ fontSize: 14, color: COLORS.text, lineHeight: 2 }}>
              <div>• {first.年度}→{last.年度}年度の客単価変化: <strong style={{ color: COLORS.accent1 }}>
                {priceChange >= 0 ? '+' : ''}{priceChange}%</strong></div>
              <div>• 一点当単価変化: <strong style={{ color: COLORS.accent3 }}>
                {unitChange >= 0 ? '+' : ''}{unitChange}%</strong>（商品価格帯の変化）</div>
              <div>• 一人当点数変化: <strong style={{ color: COLORS.accent4 }}>
                {countChange >= 0 ? '+' : ''}{countChange}点</strong>（購買行動の変化）</div>
              <div style={{ marginTop: 8, color: COLORS.sub }}>
                {Number(countChange) > 0
                  ? '→ 購入点数の増加がクロスセル成功を示しています'
                  : '→ 購入点数の向上余地があります。セット提案を強化しましょう'}
              </div>
            </div>
          )
        })()}
      </div>
    </div>
  )
}

// ==================== タブ6: AI分析チャット ====================
function DynamicChart({ chartData }) {
  if (!chartData) return null
  const { chartType, title, xKey, series, data } = chartData

  const renderChart = () => {
    const commonProps = {
      data,
      margin: { top: 10, right: 30, left: 20, bottom: 10 }
    }
    const commonChildren = [
      <CartesianGrid key="grid" strokeDasharray="3 3" stroke={COLORS.border} />,
      <XAxis key="x" dataKey={xKey} stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} />,
      <YAxis key="y" stroke={COLORS.sub} tick={{ fontSize: 11, fill: COLORS.sub }} />,
      <Tooltip key="tooltip" content={<CustomTooltip />} />,
      <Legend key="legend" wrapperStyle={{ color: COLORS.sub, fontSize: 12 }} />,
      ...series.map((s, i) => {
        if (s.type === 'bar' || chartType === 'bar') {
          return <Bar key={i} dataKey={s.key} name={s.name} fill={s.color || COLORS.accent1} opacity={0.85} />
        }
        return <Line key={i} dataKey={s.key} name={s.name} stroke={s.color || COLORS.accent1}
          strokeWidth={2} dot={{ r: 3 }} />
      })
    ]

    if (chartType === 'bar') {
      return <BarChart {...commonProps}>{commonChildren}</BarChart>
    } else if (chartType === 'line') {
      return <LineChart {...commonProps}>{commonChildren}</LineChart>
    } else {
      return <ComposedChart {...commonProps}>{commonChildren}</ComposedChart>
    }
  }

  return (
    <div style={{ marginTop: 16 }}>
      <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 8 }}>{title}</div>
      <ResponsiveContainer width="100%" height={300}>
        {renderChart()}
      </ResponsiveContainer>
    </div>
  )
}

const SAMPLE_QUESTIONS = [
  'アイチョイスの年度別成長率を教えて',
  '一番供給高が伸びた時期はいつ？',
  '生協別の一人当利用高を比較して',
  '季節トレンドを分析して',
]

function Tab6AIChat({ yearlyData, coopYearlyData, weeklyData }) {
  const [messages, setMessages] = useState([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const chatRef = useRef(null)

  const buildSystemPrompt = useCallback(() => {
    // 週次データサマリー（年度×生協の統計）
    const weeklyGroups = {}
    weeklyData.forEach(r => {
      const key = `${r['対象年度']}_${r['生協コード']}`
      if (!weeklyGroups[key]) weeklyGroups[key] = {
        年度: r['対象年度'], 生協: COOP_NAMES[r['生協コード']] || `生協${r['生協コード']}`,
        supply: [], users: [], penetration: []
      }
      weeklyGroups[key].supply.push(r['食材供給高'] || 0)
      weeklyGroups[key].users.push(r['食材利用人数'] || 0)
      weeklyGroups[key].penetration.push(r['食材金額割合'] || 0)
    })
    const weeklySummary = Object.values(weeklyGroups).map(g => ({
      年度: g.年度, 生協: g.生協,
      供給高_平均: Math.round(g.supply.reduce((a, b) => a + b, 0) / g.supply.length),
      供給高_最大: Math.max(...g.supply),
      供給高_最小: Math.min(...g.supply),
      利用人数_平均: Math.round(g.users.reduce((a, b) => a + b, 0) / g.users.length),
      浸透率_平均: (g.penetration.reduce((a, b) => a + b, 0) / g.penetration.length).toFixed(2),
      週数: g.supply.length,
    }))

    return `あなたは食材セット事業のマーケティングアナリストです。
以下のデータを元に質問に回答してください。

【データの説明】
- 生協の食材セット（ミールキット）の週次販売データ
- 生協名: アイチョイス（コード2）、岐阜（コード6）、一宮（コード7）
- 主要指標: 食材供給高（売上）、食材利用人数、食材金額割合（浸透率）、食材一人当利用高（客単価）、食材一点当単価、食材一人当点数、実質GP（粗利）

【年度別集計データ】
${JSON.stringify(yearlyData, null, 2)}

【生協×年度別集計データ】
${JSON.stringify(coopYearlyData, null, 2)}

【週次データサマリー】
${JSON.stringify(weeklySummary, null, 2)}

【回答ルール】
1. 日本語で回答
2. 数値は具体的に示す（「約X億円」「前年比+XX%」など）
3. マーケターとしてのインサイトや提案も含める
4. グラフ化すべきデータがある場合、回答の最後に以下形式で含める:

---CHART_DATA---
{
  "chartType": "bar" または "line" または "composed",
  "title": "グラフタイトル",
  "xKey": "X軸のキー名",
  "series": [
    {"key": "データキー", "name": "表示名", "type": "bar"または"line", "color": "#4ecdc4"}
  ],
  "data": [
    {"X軸キー": "値", "データキー": 数値}
  ]
}
---END_CHART_DATA---

グラフ不要な質問ではCHART_DATAセクションを含めない。`
  }, [yearlyData, coopYearlyData, weeklyData])

  const sendMessage = useCallback(async (question) => {
    if (!question.trim() || loading) return
    setInput('')
    setLoading(true)
    setMessages(prev => [...prev, { role: 'user', text: question }])

    try {
      const systemPrompt = buildSystemPrompt()
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1000,
          messages: [
            { role: 'user', content: systemPrompt + '\n\n質問: ' + question }
          ]
        })
      })
      const data = await response.json()
      const fullText = data?.content?.[0]?.text || '回答を取得できませんでした。'

      let text = fullText
      let chartData = null
      const chartMatch = fullText.match(/---CHART_DATA---\s*([\s\S]*?)\s*---END_CHART_DATA---/)
      if (chartMatch) {
        text = fullText.replace(/---CHART_DATA---[\s\S]*?---END_CHART_DATA---/, '').trim()
        try {
          chartData = JSON.parse(chartMatch[1])
        } catch {
          // JSONパース失敗時はテキストのみ
        }
      }

      setMessages(prev => [...prev, { role: 'ai', text, chartData }])
    } catch {
      setMessages(prev => [...prev, { role: 'ai', text: '分析に失敗しました。もう一度お試しください。' }])
    } finally {
      setLoading(false)
      setTimeout(() => {
        if (chatRef.current) chatRef.current.scrollTop = chatRef.current.scrollHeight
      }, 100)
    }
  }, [loading, buildSystemPrompt])

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: 'calc(100vh - 200px)', minHeight: 500 }}>
      {/* チャット履歴 */}
      <div ref={chatRef} style={{
        flex: 1, overflowY: 'auto', padding: '16px 0', marginBottom: 16
      }}>
        {messages.length === 0 && !loading && (
          <div style={{ textAlign: 'center', color: COLORS.sub, marginTop: 40 }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>💬</div>
            <div style={{ fontSize: 16, marginBottom: 8, color: COLORS.text }}>AI分析チャット</div>
            <div style={{ fontSize: 13, marginBottom: 32 }}>データについて何でも質問してください</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, justifyContent: 'center' }}>
              {SAMPLE_QUESTIONS.map(q => (
                <button key={q} onClick={() => sendMessage(q)} style={{
                  background: COLORS.card, border: `1px solid ${COLORS.border}`,
                  borderRadius: 20, padding: '8px 16px', color: COLORS.text,
                  fontSize: 13, cursor: 'pointer'
                }}>{q}</button>
              ))}
            </div>
          </div>
        )}

        {messages.map((m, i) => (
          <div key={i} style={{
            display: 'flex', justifyContent: m.role === 'user' ? 'flex-end' : 'flex-start',
            marginBottom: 16
          }}>
            <div style={{
              maxWidth: '80%',
              background: m.role === 'user' ? '#2a2e3d' : COLORS.card,
              borderRadius: m.role === 'user' ? '12px 12px 2px 12px' : '12px 12px 12px 2px',
              padding: '12px 16px',
            }}>
              <div style={{ fontSize: 14, color: COLORS.text, lineHeight: 1.7, whiteSpace: 'pre-wrap' }}>
                {m.text}
              </div>
              {m.chartData && <DynamicChart chartData={m.chartData} />}
            </div>
          </div>
        ))}

        {loading && (
          <div style={{ display: 'flex', justifyContent: 'flex-start', marginBottom: 16 }}>
            <div style={{
              background: COLORS.card, borderRadius: '12px 12px 12px 2px',
              padding: '12px 20px', color: COLORS.sub, fontSize: 14,
              display: 'flex', alignItems: 'center', gap: 8
            }}>
              <span style={{ animation: 'pulse 1.5s infinite' }}>●</span> 分析中...
            </div>
          </div>
        )}
      </div>

      {/* 入力エリア */}
      <div style={{ display: 'flex', gap: 10 }}>
        <input
          value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMessage(input) } }}
          placeholder="質問を入力してください..."
          style={{
            flex: 1, background: COLORS.card, border: `1px solid ${COLORS.border}`,
            borderRadius: 8, padding: '10px 14px', color: COLORS.text,
            fontSize: 14, outline: 'none'
          }}
        />
        <button
          onClick={() => sendMessage(input)}
          disabled={loading || !input.trim()}
          style={{
            background: loading || !input.trim() ? COLORS.border : COLORS.accent1,
            color: loading || !input.trim() ? COLORS.sub : '#000',
            border: 'none', borderRadius: 8, padding: '10px 20px',
            fontSize: 14, fontWeight: 700, cursor: loading || !input.trim() ? 'not-allowed' : 'pointer'
          }}
        >
          送信
        </button>
      </div>
    </div>
  )
}

// ==================== メインApp ====================
const TABS = [
  { id: 'overview', label: '📊 概況' },
  { id: 'growth', label: '📈 成長推移' },
  { id: 'coop', label: '🏢 生協比較' },
  { id: 'weekly', label: '📅 週次分析' },
  { id: 'kpi', label: '🎯 KPI深掘り' },
  { id: 'ai', label: '💬 AI分析' },
]

export default function App() {
  const [activeTab, setActiveTab] = useState('overview')
  const [dashData, setDashData] = useState(null)
  const [uploadedAt, setUploadedAt] = useState(null)
  const [fileName, setFileName] = useState('')
  const [dragging, setDragging] = useState(false)
  const fileInputRef = useRef(null)

  const handleFile = useCallback((file) => {
    if (!file) return
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(sheet)
        const processed = processData(rows)
        setDashData(processed)
        setUploadedAt(new Date().toLocaleString('ja-JP'))
      } catch (err) {
        alert('ファイルの読み込みに失敗しました: ' + err.message)
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    setDragging(false)
    const file = e.dataTransfer.files[0]
    if (file) handleFile(file)
  }, [handleFile])

  if (!dashData) {
    return (
      <div style={{
        minHeight: '100vh', background: COLORS.bg, display: 'flex',
        alignItems: 'center', justifyContent: 'center', fontFamily: 'sans-serif'
      }}>
        <div
          onDragOver={e => { e.preventDefault(); setDragging(true) }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          style={{
            background: COLORS.card, borderRadius: 16, padding: '48px 56px',
            border: `2px dashed ${dragging ? COLORS.accent1 : COLORS.border}`,
            textAlign: 'center', maxWidth: 480, transition: 'border-color 0.2s'
          }}
        >
          <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
          <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.text, marginBottom: 8 }}>
            食材セット マーケティングダッシュボード
          </div>
          <div style={{ fontSize: 14, color: COLORS.sub, marginBottom: 24 }}>
            syokuzai.xlsx をアップロードしてください
          </div>
          <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 28,
            background: '#0f1117', padding: '8px 14px', borderRadius: 6, fontFamily: 'monospace' }}>
            C:\Users\n-harada\Desktop\syokuzai.xlsx
          </div>
          <button
            onClick={() => fileInputRef.current?.click()}
            style={{
              background: COLORS.accent1, color: '#000', border: 'none',
              borderRadius: 8, padding: '12px 28px', fontSize: 15, fontWeight: 700,
              cursor: 'pointer', marginBottom: 16
            }}
          >
            ファイルを選択
          </button>
          <div style={{ fontSize: 12, color: COLORS.sub }}>
            またはここにファイルをドラッグ＆ドロップ
          </div>
          <input
            ref={fileInputRef} type="file" accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={e => handleFile(e.target.files[0])}
          />
        </div>
      </div>
    )
  }

  const { yearlyData, coopYearlyData, weeklyData, latestFullYear } = dashData

  return (
    <div style={{ minHeight: '100vh', background: COLORS.bg, fontFamily: 'sans-serif' }}>
      {/* ヘッダー */}
      <div style={{
        background: COLORS.card, borderBottom: `1px solid ${COLORS.border}`,
        padding: '16px 32px', display: 'flex', alignItems: 'center', justifyContent: 'space-between'
      }}>
        <div>
          <div style={{ fontSize: 20, fontWeight: 700, color: COLORS.text }}>
            <span style={{ color: COLORS.accent1 }}>食材セット</span> マーケティングダッシュボード
          </div>
          <div style={{ fontSize: 12, color: COLORS.sub, marginTop: 2 }}>
            更新: {uploadedAt} ／ {fileName}
          </div>
        </div>
        <button
          onClick={() => fileInputRef.current?.click()}
          style={{
            background: 'transparent', border: `1px solid ${COLORS.border}`,
            borderRadius: 8, padding: '8px 16px', color: COLORS.sub,
            fontSize: 13, cursor: 'pointer'
          }}
        >
          📁 ファイル更新
        </button>
        <input
          ref={fileInputRef} type="file" accept=".xlsx,.xls"
          style={{ display: 'none' }}
          onChange={e => handleFile(e.target.files[0])}
        />
      </div>

      {/* タブ */}
      <div style={{
        background: COLORS.card, borderBottom: `1px solid ${COLORS.border}`,
        padding: '0 32px', display: 'flex', gap: 4
      }}>
        {TABS.map(t => (
          <button key={t.id} onClick={() => setActiveTab(t.id)} style={{
            background: 'transparent', border: 'none', borderBottom: `2px solid ${activeTab === t.id ? COLORS.accent1 : 'transparent'}`,
            padding: '14px 18px', color: activeTab === t.id ? COLORS.accent1 : COLORS.sub,
            fontSize: 14, fontWeight: 600, cursor: 'pointer', transition: 'all 0.2s'
          }}>
            {t.label}
          </button>
        ))}
      </div>

      {/* コンテンツ */}
      <div style={{ padding: '24px 32px', maxWidth: 1200, margin: '0 auto' }}>
        {activeTab === 'overview' && (
          <Tab1Overview yearlyData={yearlyData} latestFullYear={latestFullYear} />
        )}
        {activeTab === 'growth' && (
          <Tab2Growth yearlyData={yearlyData} />
        )}
        {activeTab === 'coop' && (
          <Tab3Coop coopYearlyData={coopYearlyData} latestFullYear={latestFullYear} />
        )}
        {activeTab === 'weekly' && (
          <Tab4Weekly weeklyData={weeklyData} yearlyData={yearlyData} />
        )}
        {activeTab === 'kpi' && (
          <Tab5KPI yearlyData={yearlyData} />
        )}
        {activeTab === 'ai' && (
          dashData ? (
            <Tab6AIChat
              yearlyData={yearlyData}
              coopYearlyData={coopYearlyData}
              weeklyData={weeklyData}
            />
          ) : (
            <div style={{ color: COLORS.sub, textAlign: 'center', marginTop: 60 }}>
              syokuzai.xlsx をアップロードしてからご利用ください
            </div>
          )
        )}
      </div>
    </div>
  )
}
