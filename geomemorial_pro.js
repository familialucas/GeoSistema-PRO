// GeoSistema PRO - v4.0
// ============================================================
// NOVIDADES v4.0:
//
//  1. PERSISTÊNCIA (localStorage)
//     - Pontos, seleção, ordem e dados do imóvel salvos
//       automaticamente a cada alteração.
//     - Restauração completa ao reabrir a página.
//     - Botão "Nova Sessão" para começar do zero.
//
//  2. CONVERSÃO UTM → GEOGRÁFICO
//     - CSV/TXT com colunas E/N convertidos automaticamente.
//     - Painel manual "UTM → Geográfico" na sidebar.
//     - Detecção automática de fuso por heurística.
//
//  3. VALIDAÇÃO DE POLIGONAL
//     - Detecta segmentos cruzados (auto-interseção).
//     - Alerta visual no mapa e na sidebar.
//
//  4. EXPORTAÇÃO CSV e KML
//     - Exporta vértices selecionados em CSV (reimportável)
//       ou KML (Google Earth / QGIS).
// ============================================================

'use strict';

let map, poligonal;
let pontosBase       = [];
let selecionados     = [];
let labelsDistancia  = [];
let labelsAzimute    = [];
let marcadoresBase   = [];
let marcadoresSel    = [];
let alertaCruzamento = null;
let draggingIndex    = null;

// ─────────────────────────────────────────────────────────────
// UTILITÁRIOS
// ─────────────────────────────────────────────────────────────
function atualizarElementos(id, valor) {
    document.querySelectorAll('#' + id).forEach(el => { el.innerText = valor; });
}
function _val(id) {
    return document.getElementById(id)?.value?.trim() || '';
}

// ─────────────────────────────────────────────────────────────
// 1. INICIALIZAÇÃO DO MAPA
// ─────────────────────────────────────────────────────────────
function initMap() {
    if (!document.getElementById('map')) return;

    map = L.map('map', { maxZoom: 22 }).setView([-12.5315, -40.3054], 14);

    const osm = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 22, maxNativeZoom: 19, attribution: '© OpenStreetMap'
    });
    const sat = L.tileLayer(
        'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        { maxZoom: 22, maxNativeZoom: 19, attribution: '© Esri' }
    );
    osm.addTo(map);
    L.control.layers({ 'Mapa': osm, 'Satélite': sat }).addTo(map);

    poligonal = L.polygon([], { color: '#1a73e8', weight: 3, fillOpacity: 0.15 }).addTo(map);
    setTimeout(() => map.invalidateSize(), 500);

    _injetarEstilos();
    _criarBotaoMobile();
    _criarPainelUTM();
    _criarBotoesExtra();
    _sessaoRestaurar();
}

// ─────────────────────────────────────────────────────────────
// 2. PERSISTÊNCIA — localStorage
// ─────────────────────────────────────────────────────────────
const _CHAVE = 'geosistema_pro_v4';

function _sessaoSalvar() {
    try {
        const dados = {
            pontosBase,
            selecionadosNomes: selecionados.map(p => p.nome),
            confrontantes: Array.from(document.querySelectorAll('.input-vizinho')).map(el => el.value),
            campos: {
                nomeProprietario: _val('nomeProprietario'),
                matriculaImovel:  _val('matriculaImovel'),
                nomeMunicipio:    _val('nomeMunicipio'),
                nomeComarca:      _val('nomeComarca'),
                nomeCartorio:     _val('nomeCartorio'),
                nomeDatum:        _val('nomeDatum'),
                nomeRT:           _val('nomeRT'),
                numCREA:          _val('numCREA'),
            }
        };
        localStorage.setItem(_CHAVE, JSON.stringify(dados));
    } catch(e) { console.warn('Sessão não salva:', e); }
}

function _sessaoRestaurar() {
    try {
        const raw = localStorage.getItem(_CHAVE);
        if (!raw) return;
        const dados = JSON.parse(raw);

        if (dados.pontosBase?.length) {
            pontosBase = dados.pontosBase;
            _garantirContainerOrdenavel();
            renderizarLista();
            if (map) {
                map.fitBounds(L.latLngBounds(pontosBase.map(p => [p.lat, p.lon])), { padding: [40, 40] });
                _renderizarMarcadoresBase();
            }
            atualizarTabelaAltimetria();
            mostrarStatus(`✓ Sessão restaurada — ${pontosBase.length} vértice(s)`, 'ok');
        }

        if (dados.selecionadosNomes?.length) {
            selecionados = dados.selecionadosNomes
                .map(nome => pontosBase.find(p => p.nome === nome))
                .filter(Boolean);
            sincronizarAtivos();
            atualizarMapa();
            renderizarOrdenavel();
            setTimeout(() => {
                const inputs = document.querySelectorAll('.input-vizinho');
                (dados.confrontantes || []).forEach((v, i) => { if (inputs[i]) inputs[i].value = v; });
            }, 150);
        }

        if (dados.campos) {
            Object.entries(dados.campos).forEach(([id, val]) => {
                const el = document.getElementById(id);
                if (el && val) el.value = val;
            });
        }
    } catch(e) { console.warn('Erro ao restaurar sessão:', e); }
}

function novaSessao() {
    if (!confirm('Iniciar nova sessão? Todos os dados serão apagados.')) return;
    localStorage.removeItem(_CHAVE);
    pontosBase = []; selecionados = [];
    limparSelecao();
    const ld = document.getElementById('listaPontos') || document.getElementById('botoesPontos');
    if (ld) ld.innerHTML = '';
    const lo = document.getElementById('listaOrdenavel');
    if (lo) lo.innerHTML = '<p style="font-size:11px;color:#aaa;padding:4px 0;">Nenhum vértice selecionado.</p>';
    marcadoresBase.forEach(m => map?.removeLayer(m)); marcadoresBase = [];
    mostrarStatus('Nova sessão iniciada.', 'info');
}

function _autoSave() { _sessaoSalvar(); }

// ─────────────────────────────────────────────────────────────
// 3. CONVERSÃO UTM → GEOGRÁFICO (WGS84 / SIRGAS 2000)
// ─────────────────────────────────────────────────────────────
function utmParaGeo(E, N, fuso, hem) {
    hem = (hem || 'S').toUpperCase();
    const a  = 6378137.0, f = 1 / 298.257223563;
    const b  = a * (1 - f);
    const e2 = 1 - (b * b) / (a * a);
    const k0 = 0.9996;
    const lon0 = (fuso * 6 - 183) * Math.PI / 180;

    const x = E - 500000;
    const y = hem === 'S' ? N - 10000000 : N;

    const e1 = (1 - Math.sqrt(1 - e2)) / (1 + Math.sqrt(1 - e2));
    const M  = y / k0;
    const mu = M / (a * (1 - e2 / 4 - 3 * e2 * e2 / 64 - 5 * e2 * e2 * e2 / 256));

    const fp = mu
        + (3 * e1 / 2 - 27 * e1 * e1 * e1 / 32) * Math.sin(2 * mu)
        + (21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32) * Math.sin(4 * mu)
        + (151 * e1 * e1 * e1 / 96) * Math.sin(6 * mu)
        + (1097 * e1 * e1 * e1 * e1 / 512) * Math.sin(8 * mu);

    const sinFp = Math.sin(fp), cosFp = Math.cos(fp), tanFp = sinFp / cosFp;
    const N1 = a / Math.sqrt(1 - e2 * sinFp * sinFp);
    const T1 = tanFp * tanFp;
    const C1 = e2 / (1 - e2) * cosFp * cosFp;
    const R1 = a * (1 - e2) / Math.pow(1 - e2 * sinFp * sinFp, 1.5);
    const D  = x / (N1 * k0);

    const lat = fp - (N1 * tanFp / R1) * (
        D * D / 2
        - (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * e2 / (1 - e2)) * D * D * D * D / 24
        + (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * e2 / (1 - e2) - 3 * C1 * C1) * Math.pow(D, 6) / 720
    );
    const lon = lon0 + (
        D - (1 + 2 * T1 + C1) * D * D * D / 6
        + (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * e2 / (1 - e2) + 24 * T1 * T1) * Math.pow(D, 5) / 120
    ) / cosFp;

    return { lat: lat * 180 / Math.PI, lon: lon * 180 / Math.PI };
}

// ─────────────────────────────────────────────────────────────
// 4. LEITURA DE ARQUIVO
// ─────────────────────────────────────────────────────────────
function lerArquivo(input) {
    const file = input.files[0];
    if (!file) return;
    if (file.size > 10 * 1024 * 1024) { mostrarStatus('Arquivo muito grande (máx 10MB).', 'erro'); return; }

    pontosBase = []; selecionados = [];
    mostrarStatus('Lendo arquivo...', 'info');

    const nome = file.name.toLowerCase();
    if (nome.endsWith('.kmz')) {
        const r = new FileReader();
        r.onload = e => processarKMZ(e.target.result);
        r.onerror = () => mostrarStatus('Erro ao ler KMZ.', 'erro');
        r.readAsArrayBuffer(file); return;
    }
    const r = new FileReader();
    r.onload = e => {
        const c = e.target.result; let erros = [];
        if      (nome.endsWith('.kml'))                           erros = processarKML(c);
        else if (nome.endsWith('.gpx'))                           erros = processarGPX(c);
        else if (nome.endsWith('.dxf'))                           erros = processarDXF(c);
        else if (nome.endsWith('.html') || nome.endsWith('.htm')) erros = processarHTML(c);
        else if (nome.endsWith('.csv') || nome.endsWith('.txt'))  erros = processarCSV(c);
        else { mostrarStatus('Formato não reconhecido.', 'erro'); return; }
        finalizarLeitura(file.name, erros);
    };
    r.onerror = () => mostrarStatus('Erro ao ler arquivo.', 'erro');
    r.readAsText(file, 'UTF-8');
}

function mostrarStatus(msg, tipo) {
    document.querySelectorAll('#statusImportacao').forEach(el => {
        el.textContent = msg;
        el.style.color = tipo === 'erro' ? '#c0392b' : tipo === 'aviso' ? '#e67e00' : tipo === 'ok' ? '#1a7a40' : '#1a73e8';
    });
}

function finalizarLeitura(nomeArq, erros = []) {
    if (!pontosBase.length) { mostrarStatus(`⚠ Nenhum ponto em "${nomeArq}".`, 'erro'); return; }
    mostrarStatus(
        `✓ ${pontosBase.length} vértice(s) de "${nomeArq}"` + (erros.length ? ` | ${erros.length} ignorada(s)` : ''),
        erros.length ? 'aviso' : 'ok'
    );
    _garantirContainerOrdenavel();
    renderizarLista();
    if (map && pontosBase.length) {
        map.fitBounds(L.latLngBounds(pontosBase.map(p => [p.lat, p.lon])), { padding: [40, 40] });
        _renderizarMarcadoresBase();
    }
    atualizarTabelaAltimetria();
    _atualizarBotoesExportar();
    _autoSave();
}

// ─────────────────────────────────────────────────────────────
// 5. PARSERS
// ─────────────────────────────────────────────────────────────

// ── CSV / TXT (aceita geográfico e UTM) ──────────────────────
function processarCSV(texto) {
    const erros = [];
    const linhas = texto.split(/\r?\n/).filter(l => l.trim());
    if (linhas.length < 2) { mostrarStatus('CSV vazio.', 'erro'); return erros; }

    const sep = (linhas[0].match(/;/g)||[]).length > (linhas[0].match(/,/g)||[]).length ? ';' : ',';
    const cab = linhas[0].split(sep).map(c => c.trim().replace(/^["']|["']$/g,'').toLowerCase());

    const iN    = cab.findIndex(c => /^(nome|id|ponto|v[eé]rtice|pt|vertex|name)/i.test(c));
    const iLbl  = cab.findIndex(c => /^(label|r[oó]tulo)/i.test(c));
    const iGrp  = cab.findIndex(c => /^(grupo|group|tipo|type)/i.test(c));
    const iLat = cab.findIndex(c => /^(lat|latitude|y_geo|geo_y|lat_dd|lat_dms)/i.test(c));
    const iLon = cab.findIndex(c => /^(lon|long|longitude|x_geo|geo_x|lon_dd|lon_dms)/i.test(c));
    const iAlt = cab.findIndex(c => /^(alt|altitude|z|elev|elevation|cota)/i.test(c));
    const iE   = cab.findIndex(c => /^(e$|east|easting|utm_e|x_utm|este)/i.test(c));
    const iNu  = cab.findIndex(c => /^(n$|north|northing|utm_n|y_utm|norte)/i.test(c));
    const iFus = cab.findIndex(c => /^(fuso|zone|fus)/i.test(c));
    const iHem = cab.findIndex(c => /^(hem|hemisf)/i.test(c));

    const modoUTM = (iLat === -1 || iLon === -1) && iE !== -1 && iNu !== -1;

    for (let i = 1; i < linhas.length; i++) {
        const col = linhas[i].split(sep).map(c => c.trim().replace(/^["']|["']$/g,''));
        if (col.length < 2) continue;
        let lat, lon, latDMS, lonDMS;

        if (modoUTM) {
            const Ev  = parseFloat(col[iE]?.replace(',','.'));
            const Nv  = parseFloat(col[iNu]?.replace(',','.'));
            const fus = iFus >= 0 ? parseInt(col[iFus]) : 24;
            const hem = iHem >= 0 ? col[iHem] : 'S';
            if (isNaN(Ev) || isNaN(Nv)) { erros.push(`Linha ${i+1}: UTM inválido`); continue; }
            const g = utmParaGeo(Ev, Nv, fus, hem);
            lat = g.lat; lon = g.lon;
            latDMS = decimalParaDMSLegivel(lat); lonDMS = decimalParaDMSLegivel(lon);
        } else {
            const latStr = col[iLat] || ''; const lonStr = col[iLon] || '';
            lat = converterDMS(latStr); lon = converterDMS(lonStr);
            latDMS = latStr; lonDMS = lonStr;
            if (!latStr || isNaN(lat) || lat < -90 || lat > 90)   { erros.push(`Linha ${i+1}: latitude inválida`);  continue; }
            if (!lonStr || isNaN(lon) || lon < -180 || lon > 180) { erros.push(`Linha ${i+1}: longitude inválida`); continue; }
        }
        const nomeBase = (iN >= 0 ? col[iN] : null) || ('P' + i);
        const labelVal = (iLbl >= 0 ? col[iLbl] : null) || nomeBase;
        const grupoVal = (iGrp >= 0 && col[iGrp] && GRUPOS_CORES[col[iGrp]]) ? col[iGrp] : 'perimetro';
        pontosBase.push({ nome: nomeBase, label: labelVal, grupo: grupoVal, latDMS, lonDMS, alt: iAlt >= 0 ? parseFloat(col[iAlt]) || 0 : 0, lat, lon });
    }
    return erros;
}

// ── HTML ──────────────────────────────────────────────────────
function processarHTML(html) {
    const erros = [];
    const m = html.match(/const\s+pontos\s*=\s*(\[[\s\S]*?\]);/);
    if (m) {
        try {
            const arr = Function('"use strict"; return ' + m[1])();
            arr.forEach((p, i) => {
                const ls = p.latDMS || p.lat?.toString() || '';
                const lo = p.lonDMS || p.lon?.toString() || '';
                const nomeBase = p.id || p.nome || ('P'+(i+1));
                pontosBase.push({ nome: nomeBase, label: p.label || nomeBase, grupo: p.grupo || 'perimetro', latDMS: ls, lonDMS: lo, alt: p.alt||0, lat: p.latDec ?? converterDMS(ls), lon: p.lonDec ?? converterDMS(lo) });
            });
            if (pontosBase.length) return erros;
        } catch(e) { console.warn('Fallback para tabela HTML:', e); }
    }
    const doc = new DOMParser().parseFromString(html, 'text/html');
    doc.querySelectorAll('table').forEach(tab => {
        const ths = Array.from(tab.querySelectorAll('th')).map(t => t.textContent.trim().toLowerCase());
        const iNm = ths.findIndex(t => /nome|name|ponto|v[eé]rtice/i.test(t));
        const iLa = ths.findIndex(t => /lat/i.test(t));
        const iLo = ths.findIndex(t => /lon/i.test(t));
        const iAl = ths.findIndex(t => /alt/i.test(t));
        if (iLa === -1 || iLo === -1) return;
        tab.querySelectorAll('tbody tr').forEach((row, ri) => {
            const c = row.querySelectorAll('td');
            if (c.length < 2) return;
            const ls = c[iLa]?.textContent.trim() || ''; const lo = c[iLo]?.textContent.trim() || '';
            if (!ls) return;
            const nomeBase2 = (iNm >= 0 ? c[iNm]?.textContent.trim() : null) || ('P'+(ri+1));
            pontosBase.push({ nome: nomeBase2, label: nomeBase2, grupo: 'perimetro', latDMS: ls, lonDMS: lo, alt: iAl >= 0 ? parseFloat(c[iAl]?.textContent)||0 : 0, lat: converterDMS(ls), lon: converterDMS(lo) });
        });
    });
    return erros;
}

// ── KML ───────────────────────────────────────────────────────
function processarKML(txt) {
    const erros = [];
    let doc;
    try { doc = new DOMParser().parseFromString(txt, 'text/xml'); if (doc.querySelector('parsererror')) { mostrarStatus('KML inválido.','erro'); return erros; } }
    catch(e) { mostrarStatus('Erro KML.','erro'); return erros; }

    function getAlt(pm) { for (const d of pm.querySelectorAll('ExtendedData SimpleData')) if (/alt|cota|elev|z/i.test(d.getAttribute('name')||'')) return parseFloat(d.textContent)||0; return null; }

    doc.querySelectorAll('Placemark').forEach((pm, i) => {
        const nome = pm.querySelector('name')?.textContent.trim() || ('P'+(i+1));
        const pc   = pm.querySelector('Point coordinates'); if (!pc) return;
        const pts  = pc.textContent.trim().split(','); if (pts.length < 2) return;
        const lon = parseFloat(pts[0]), lat = parseFloat(pts[1]), alt = getAlt(pm) ?? (parseFloat(pts[2])||0);
        if (isNaN(lat)||isNaN(lon)) { erros.push(`"${nome}": inválido`); return; }
        pontosBase.push({ nome, label: nome, grupo: 'perimetro', alt, lat, lon, latDMS: decimalParaDMSLegivel(lat), lonDMS: decimalParaDMSLegivel(lon) });
    });

    if (!pontosBase.length) {
        let idx = 1;
        for (const sel of ['Polygon outerBoundaryIs coordinates','LinearRing coordinates','LineString coordinates']) {
            doc.querySelectorAll(sel).forEach(el => {
                el.textContent.trim().split(/\s+/).forEach(par => {
                    const p = par.split(','); if (p.length < 2) return;
                    const lon = parseFloat(p[0]), lat = parseFloat(p[1]);
                    if (!isNaN(lat)&&!isNaN(lon)) pontosBase.push({ nome:'V'+idx, label:'V'+idx, grupo:'perimetro', alt:parseFloat(p[2])||0, lat, lon, latDMS:decimalParaDMSLegivel(lat), lonDMS:decimalParaDMSLegivel(lon) });
                    idx++;
                });
                if (pontosBase.length > 1) { const p0=pontosBase[0],pN=pontosBase[pontosBase.length-1]; if (Math.abs(p0.lat-pN.lat)<1e-7&&Math.abs(p0.lon-pN.lon)<1e-7) pontosBase.pop(); }
            });
            if (pontosBase.length) break;
        }
    }
    return erros;
}

function processarKMZ(buf) {
    const carregar = () => JSZip.loadAsync(buf).then(zip => {
        const kml = Object.values(zip.files).find(f => f.name.toLowerCase().endsWith('.kml'));
        if (!kml) { mostrarStatus('KMZ: sem KML interno.','erro'); return; }
        kml.async('string').then(txt => finalizarLeitura('arquivo.kmz', processarKML(txt)));
    }).catch(() => mostrarStatus('Erro ao descompactar KMZ.','erro'));

    if (typeof JSZip === 'undefined') {
        const s = document.createElement('script');
        s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
        s.onload = carregar; s.onerror = () => mostrarStatus('JSZip falhou.','erro');
        document.head.appendChild(s);
    } else carregar();
}

// ── GPX ───────────────────────────────────────────────────────
// Suporta: <wpt> waypoints, <trkpt> track points, <rtept> route points
// Exportado por: Garmin, Leica, coletores Android/iOS, Google Maps
function processarGPX(txt) {
    const erros = [];
    let doc;
    try {
        doc = new DOMParser().parseFromString(txt, 'text/xml');
        if (doc.querySelector('parsererror')) { mostrarStatus('GPX inválido.', 'erro'); return erros; }
    } catch(e) { mostrarStatus('Erro ao ler GPX.', 'erro'); return erros; }

    let idx = 1;

    // Função auxiliar — extrai um ponto de qualquer elemento GPX com lat/lon
    function extrairPonto(el, tipoDefault) {
        const lat = parseFloat(el.getAttribute('lat'));
        const lon = parseFloat(el.getAttribute('lon'));
        if (isNaN(lat) || isNaN(lon)) return null;

        // Altitude: <ele> para track/route, <ele> ou <extensions> para waypoints
        const eleEl = el.querySelector('ele');
        const alt   = eleEl ? parseFloat(eleEl.textContent) || 0 : 0;

        // Nome: <name> > <cmt> > <desc> > tipo+índice
        const nomeEl = el.querySelector('name') || el.querySelector('cmt') || el.querySelector('desc');
        const nome   = nomeEl?.textContent?.trim() || (tipoDefault + String(idx).padStart(2,'0'));

        return { nome, label: nome, grupo: 'perimetro', alt, lat, lon,
            latDMS: decimalParaDMSLegivel(lat), lonDMS: decimalParaDMSLegivel(lon) };
    }

    // 1. Waypoints <wpt> — pontos avulsos, ideal para vértices GNSS
    doc.querySelectorAll('wpt').forEach(el => {
        const p = extrairPonto(el, 'WPT');
        if (p) { pontosBase.push(p); idx++; }
    });

    // 2. Track points <trkpt> dentro de <trkseg> — trajeto gravado
    if (!pontosBase.length) {
        doc.querySelectorAll('trkseg').forEach(seg => {
            seg.querySelectorAll('trkpt').forEach(el => {
                const p = extrairPonto(el, 'T');
                if (p) { pontosBase.push(p); idx++; }
            });
        });
        // Remove ponto final duplicado se fechar o polígono
        if (pontosBase.length > 1) {
            const p0 = pontosBase[0], pN = pontosBase[pontosBase.length - 1];
            if (Math.abs(p0.lat - pN.lat) < 1e-7 && Math.abs(p0.lon - pN.lon) < 1e-7) pontosBase.pop();
        }
    }

    // 3. Route points <rtept> — rota planejada
    if (!pontosBase.length) {
        doc.querySelectorAll('rtept').forEach(el => {
            const p = extrairPonto(el, 'R');
            if (p) { pontosBase.push(p); idx++; }
        });
    }

    if (!pontosBase.length) erros.push('GPX sem waypoints, track points ou route points válidos.');
    return erros;
}

// ── DXF ───────────────────────────────────────────────────────
// Suporta: entidades POINT e LWPOLYLINE / POLYLINE com coordenadas XY
// Exportado por: AutoCAD, QGIS ("DXF only points"), Leica Geo Office,
//                Topcon Tools, South Office
// Datum: o DXF não carrega datum — assume-se que XY = lon/lat geográfico
//        OU UTM. Se |X| > 180 trata como UTM (fuso detectado automaticamente).
function processarDXF(txt) {
    const erros = [];
    const linhas = txt.split(/\r?\n/).map(l => l.trim());

    // O DXF é um arquivo de pares (código de grupo / valor)
    // Lemos sequencialmente procurando as entidades que nos interessam
    let i = 0;
    const total = linhas.length;

    function proximoGrupo() {
        if (i >= total - 1) return null;
        const codigo = parseInt(linhas[i]);
        const valor  = linhas[i + 1];
        i += 2;
        return { codigo, valor };
    }

    // Acumula pontos de uma entidade LWPOLYLINE/POLYLINE
    function _criarPonto(x, y, z, nome) {
        const lat = Math.abs(x) > 180 ? null : y;  // se |X|>180 provavelmente UTM
        const lon = Math.abs(x) > 180 ? null : x;

        if (lat !== null) {
            // Coordenadas geográficas diretas
            if (isNaN(lat) || lat < -90 || lat > 90) return null;
            if (isNaN(lon) || lon < -180 || lon > 180) return null;
            return { nome, label: nome, grupo: 'perimetro', alt: parseFloat(z) || 0,
                lat, lon, latDMS: decimalParaDMSLegivel(lat), lonDMS: decimalParaDMSLegivel(lon) };
        } else {
            // Tenta interpretar como UTM — fuso detectado por E (X)
            const E = parseFloat(x), N = parseFloat(y);
            if (isNaN(E) || isNaN(N)) return null;
            // Heurística de fuso: E ÷ 100000 arredondado dá o fuso aproximado
            const fusoEstimado = Math.round((E - 500000) / 6 / 111319 + 24) || 24;
            // Mais simples: fuso por longitude central = (fuso*6 - 183)
            // Como não sabemos o fuso exato, tentamos 23 e 24 (Bahia)
            const g = utmParaGeo(E, N, fusoEstimado, 'S');
            if (isNaN(g.lat) || g.lat < -90 || g.lat > 90) return null;
            return { nome, label: nome, grupo: 'perimetro', alt: parseFloat(z) || 0,
                lat: g.lat, lon: g.lon,
                latDMS: decimalParaDMSLegivel(g.lat), lonDMS: decimalParaDMSLegivel(g.lon) };
        }
    }

    let idx = 1;
    let emSecaoEntidades = false;

    while (i < total - 1) {
        const g = proximoGrupo();
        if (!g) break;

        // Entrada na seção ENTITIES
        if (g.codigo === 2 && g.valor === 'ENTITIES') { emSecaoEntidades = true; continue; }
        if (g.codigo === 2 && g.valor === 'ENDSEC')   { emSecaoEntidades = false; continue; }
        if (!emSecaoEntidades) continue;

        // ── Entidade POINT ──────────────────────────────────────
        if (g.codigo === 0 && g.valor === 'POINT') {
            let x = 0, y = 0, z = 0, nome = 'P' + String(idx).padStart(2,'0');
            // Lê grupos até próxima entidade (código 0)
            while (i < total - 1) {
                const sub = proximoGrupo();
                if (!sub) break;
                if (sub.codigo === 0) { i -= 2; break; }  // devolve o cursor
                if (sub.codigo === 10) x = parseFloat(sub.valor) || 0;
                if (sub.codigo === 20) y = parseFloat(sub.valor) || 0;
                if (sub.codigo === 30) z = parseFloat(sub.valor) || 0;
                if (sub.codigo === 2 || sub.codigo === 3) nome = sub.valor.trim() || nome;
            }
            const p = _criarPonto(x, y, z, nome);
            if (p) { pontosBase.push(p); idx++; }
            else erros.push(`POINT "${nome}": coordenadas inválidas (${x}, ${y})`);
        }

        // ── Entidade LWPOLYLINE ─────────────────────────────────
        // (leve poligonal 2D — formato mais comum em exports de campo)
        if (g.codigo === 0 && (g.valor === 'LWPOLYLINE' || g.valor === 'POLYLINE')) {
            const tipo = g.valor;
            let verticesLW = [], xAtu = null, yAtu = null;
            let nomeLW = tipo + String(idx).padStart(2,'0');

            while (i < total - 1) {
                const sub = proximoGrupo();
                if (!sub) break;
                if (sub.codigo === 0 && sub.valor !== 'VERTEX') { i -= 2; break; }
                if (sub.codigo === 2 || sub.codigo === 3) nomeLW = sub.valor.trim() || nomeLW;

                // LWPOLYLINE: grupos 10/20 alternam X e Y no mesmo bloco
                if (sub.codigo === 10) {
                    if (xAtu !== null && yAtu !== null) { verticesLW.push([xAtu, yAtu, 0]); yAtu = null; }
                    xAtu = parseFloat(sub.valor);
                }
                if (sub.codigo === 20 && xAtu !== null) { yAtu = parseFloat(sub.valor); }
                if (sub.codigo === 30 && verticesLW.length) {
                    verticesLW[verticesLW.length - 1][2] = parseFloat(sub.valor) || 0;
                }
                // POLYLINE legacy: vértices em entidades VERTEX
                if (sub.codigo === 0 && sub.valor === 'VERTEX') {
                    let vx = 0, vy = 0, vz = 0;
                    while (i < total - 1) {
                        const vs = proximoGrupo(); if (!vs) break;
                        if (vs.codigo === 0) { i -= 2; break; }
                        if (vs.codigo === 10) vx = parseFloat(vs.valor) || 0;
                        if (vs.codigo === 20) vy = parseFloat(vs.valor) || 0;
                        if (vs.codigo === 30) vz = parseFloat(vs.valor) || 0;
                    }
                    verticesLW.push([vx, vy, vz]);
                }
            }
            // Último par pendente (LWPOLYLINE)
            if (xAtu !== null && yAtu !== null) verticesLW.push([xAtu, yAtu, 0]);

            // Remove vértice final duplicado se polígono fechado
            if (verticesLW.length > 1) {
                const v0 = verticesLW[0], vN = verticesLW[verticesLW.length - 1];
                if (Math.abs(v0[0]-vN[0]) < 1e-7 && Math.abs(v0[1]-vN[1]) < 1e-7) verticesLW.pop();
            }

            verticesLW.forEach(([vx, vy, vz], vi) => {
                const nomeV = `V${String(idx).padStart(2,'0')}`;
                const p = _criarPonto(vx, vy, vz, nomeV);
                if (p) { pontosBase.push(p); idx++; }
                else erros.push(`${tipo} vértice ${vi+1}: coordenadas inválidas (${vx}, ${vy})`);
            });
        }
    }

    if (!pontosBase.length) erros.push('DXF sem entidades POINT ou LWPOLYLINE/POLYLINE válidas.');
    return erros;
}

// ─────────────────────────────────────────────────────────────
// 6. LISTA DE VÉRTICES
// ─────────────────────────────────────────────────────────────

const GRUPOS_CORES = {
    perimetro:   { cor: '#1a73e8', icone: '⬟', label: 'Perímetro' },
    edificacao:  { cor: '#e67e00', icone: '⬛', label: 'Edificação' },
    referencia:  { cor: '#8e44ad', icone: '✦',  label: 'Referência' },
    outro:       { cor: '#555',    icone: '●',   label: 'Outro'      },
};

function _grupoInfo(g) { return GRUPOS_CORES[g] || GRUPOS_CORES['outro']; }

function renderizarLista() {
    const div = document.getElementById('listaPontos') || document.getElementById('botoesPontos');
    if (!div) return;
    div.innerHTML = '';
    if (!pontosBase.length) { div.innerHTML = '<p style="font-size:11px;color:#aaa;padding:6px;">Nenhum vértice carregado.</p>'; return; }

    // Agrupa pontos por grupo
    const grupos = {};
    pontosBase.forEach(p => {
        const g = p.grupo || 'perimetro';
        if (!grupos[g]) grupos[g] = [];
        grupos[g].push(p);
    });

    Object.entries(grupos).forEach(([g, pontos]) => {
        const info = _grupoInfo(g);
        // Cabeçalho do grupo
        const hdr = document.createElement('div');
        hdr.style.cssText = `font-size:10px;font-weight:bold;color:${info.cor};margin:6px 0 2px;padding:2px 4px;background:${info.cor}18;border-left:3px solid ${info.cor};border-radius:2px;display:flex;align-items:center;justify-content:space-between;`;
        hdr.innerHTML = `<span>${info.icone} ${info.label} (${pontos.length})</span>`;
        div.appendChild(hdr);

        pontos.forEach(p => {
            const item = document.createElement('div');
            item.className = 'ponto-item'; item.dataset.nome = p.nome;
            item.style.cssText = 'display:flex;align-items:center;gap:4px;padding:5px 6px;';

            // Ícone do grupo
            const badge = document.createElement('span');
            badge.style.cssText = `font-size:9px;color:${info.cor};min-width:12px;`;
            badge.textContent = info.icone;

            // Nome / label (clicável para renomear)
            const nome = document.createElement('span');
            nome.style.cssText = 'flex:1;font-size:12px;';
            nome.innerHTML = `<strong>${p.label || p.nome}</strong> <small style="color:#999;">${p.nome !== (p.label||p.nome) ? '('+p.nome+')' : ''}</small>`;

            // Seletor de grupo
            const sel = document.createElement('select');
            sel.style.cssText = 'font-size:9px;border:1px solid #ddd;border-radius:3px;padding:1px 2px;color:#555;max-width:72px;';
            sel.title = 'Mudar grupo';
            Object.entries(GRUPOS_CORES).forEach(([k, v]) => {
                const o = document.createElement('option');
                o.value = k; o.textContent = v.label;
                if (k === (p.grupo||'perimetro')) o.selected = true;
                sel.appendChild(o);
            });
            sel.addEventListener('change', e => {
                e.stopPropagation();
                p.grupo = e.target.value;
                // Se não é perímetro, remove da seleção
                if (p.grupo !== 'perimetro') {
                    const i = selecionados.findIndex(s => s.nome === p.nome);
                    if (i !== -1) selecionados.splice(i, 1);
                }
                renderizarLista(); atualizarMapa(); renderizarOrdenavel(); _autoSave();
            });

            // Botão renomear
            const btnR = document.createElement('button');
            btnR.title = 'Renomear rótulo';
            btnR.style.cssText = 'background:none;border:none;cursor:pointer;font-size:11px;padding:0 2px;width:auto;margin:0;color:#888;';
            btnR.textContent = '✏';
            btnR.addEventListener('click', e => {
                e.stopPropagation();
                const novo = prompt(`Novo rótulo para "${p.nome}":`, p.label || p.nome);
                if (novo !== null && novo.trim()) {
                    p.label = novo.trim();
                    renderizarLista(); atualizarMapa(); renderizarOrdenavel(); _autoSave();
                }
            });

            item.appendChild(badge);
            item.appendChild(nome);
            item.appendChild(sel);
            item.appendChild(btnR);

            // Clique para selecionar (só perímetro vai para poligonal diretamente)
            item.addEventListener('click', function(e) {
                if (e.target === sel || e.target === btnR) return;
                if ((p.grupo||'perimetro') !== 'perimetro') {
                    alert(`"${p.label||p.nome}" está no grupo "${_grupoInfo(p.grupo).label}". Mude para "Perímetro" para incluir na poligonal.`);
                    return;
                }
                const idx = selecionados.findIndex(s => s.nome === p.nome);
                if (idx === -1) { selecionados.push(p); this.classList.add('active'); }
                else            { selecionados.splice(idx,1); this.classList.remove('active'); }
                atualizarMapa(); renderizarOrdenavel(); _autoSave();
            });
            div.appendChild(item);
        });
    });

    _garantirContainerOrdenavel();
    renderizarOrdenavel();
}

// ─────────────────────────────────────────────────────────────
// 7. CONTAINER ORDENÁVEL
// ─────────────────────────────────────────────────────────────
function _garantirContainerOrdenavel() {
    if (document.getElementById('listaOrdenavel')) return;
    let anchor = document.getElementById('anchorOrdenavel');
    if (!anchor) {
        const ref = document.getElementById('listaPontos') || document.getElementById('botoesPontos');
        if (!ref?.parentNode) return;
        anchor = document.createElement('div'); anchor.id = 'anchorOrdenavel';
        ref.parentNode.insertBefore(anchor, ref.nextSibling);
    }
    anchor.innerHTML = `
        <div style="font-size:12px;font-weight:bold;color:#333;margin:10px 0 4px;border-bottom:1px solid #eee;padding-bottom:3px;">
            Ordem da Poligonal <small style="font-weight:normal;color:#999;">(arraste para reordenar)</small>
        </div>
        <div id="listaOrdenavel"></div>`;
}

function renderizarOrdenavel() {
    const c = document.getElementById('listaOrdenavel');
    if (!c) return;
    c.innerHTML = '';
    if (!selecionados.length) { c.innerHTML = '<p style="font-size:11px;color:#aaa;padding:4px 0;">Nenhum vértice selecionado.</p>'; return; }
    selecionados.forEach((p, idx) => {
        const item = document.createElement('div');
        item.className = 'ponto-sel'; item.draggable = true;
        item.innerHTML = `<span class="handle">⠿</span><span style="flex:1;"><strong>${p.label||p.nome}</strong>${p.label && p.label!==p.nome ? '<small style="color:#aaa;font-size:9px;"> ('+p.nome+')</small>' : ''}</span><span style="font-size:10px;color:#555;">${idx+1}°</span><button class="btn-rem" title="Remover">✕</button>`;
        item.querySelector('.btn-rem').addEventListener('click', e => { e.stopPropagation(); removerSelecionado(idx); });
        item.addEventListener('dragstart', e => { draggingIndex = idx; item.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; });
        item.addEventListener('dragend',   () => { item.classList.remove('dragging'); draggingIndex = null; });
        item.addEventListener('dragover',  e => e.preventDefault());
        item.addEventListener('drop', e => {
            e.preventDefault(); if (draggingIndex === null || draggingIndex === idx) return;
            const mv = selecionados.splice(draggingIndex,1)[0]; selecionados.splice(idx,0,mv);
            atualizarMapa(); renderizarOrdenavel(); sincronizarAtivos(); _autoSave();
        });
        c.appendChild(item);
    });
}

function removerSelecionado(idx) { selecionados.splice(idx,1); atualizarMapa(); renderizarOrdenavel(); sincronizarAtivos(); _autoSave(); }
function sincronizarAtivos() { document.querySelectorAll('.ponto-item').forEach(el => el.classList.toggle('active', !!selecionados.find(s => s.nome === el.dataset.nome))); }

// ─────────────────────────────────────────────────────────────
// 8. MARCADORES NO MAPA
// ─────────────────────────────────────────────────────────────
function _renderizarMarcadoresBase() {
    if (!map) return;
    marcadoresBase.forEach(m => map.removeLayer(m)); marcadoresBase = [];
    pontosBase.forEach(p => {
        if (selecionados.find(s => s.nome === p.nome)) return;
        const info = _grupoInfo(p.grupo || 'perimetro');
        const isEdif = (p.grupo || 'perimetro') !== 'perimetro';
        // Edificações: ícone quadrado laranja; perímetro/outros: círculo normal
        const shape = isEdif
            ? `<div style="background:${info.cor};border:2px solid white;border-radius:3px;width:22px;height:22px;display:flex;align-items:center;justify-content:center;font-size:8px;font-weight:bold;color:#fff;cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,0.4);">${(p.label||p.nome).substring(0,4)}</div>`
            : `<div style="background:#fff;border:2px solid ${info.cor};border-radius:50%;width:22px;height:22px;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:bold;color:${info.cor};cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,0.3);">${(p.label||p.nome).substring(0,4)}</div>`;
        const m = L.marker([p.lat, p.lon], { icon: L.divIcon({ className: '',
            html: shape, iconSize:[22,22], iconAnchor:[11,11] }) }).addTo(map);
        m.bindTooltip(`<b>${p.label||p.nome}</b>${p.label&&p.label!==p.nome?' <small>('+p.nome+')</small>':''}<br>${formatarDMSLegivel(p.latDMS)}<br><i style="color:${info.cor}">${info.label}</i>`, { permanent:false, direction:'top', offset:[0,-14] });
        m.on('click', () => {
            if ((p.grupo||'perimetro') !== 'perimetro') return; // não seleciona edificações pelo mapa
            const i=selecionados.findIndex(s=>s.nome===p.nome);
            if(i===-1) selecionados.push(p); else selecionados.splice(i,1);
            sincronizarAtivos(); atualizarMapa(); renderizarOrdenavel(); _autoSave();
        });
        marcadoresBase.push(m);
    });
}

// ─────────────────────────────────────────────────────────────
// 9. ATUALIZAÇÃO DO MAPA + VALIDAÇÃO
// ─────────────────────────────────────────────────────────────
function atualizarMapa() {
    if (!map || !poligonal) return;
    poligonal.setLatLngs(selecionados.map(p => [p.lat, p.lon]));
    [...labelsDistancia, ...labelsAzimute, ...marcadoresSel].forEach(l => map.removeLayer(l));
    labelsDistancia = []; labelsAzimute = []; marcadoresSel = [];
    if (alertaCruzamento) { map.removeLayer(alertaCruzamento); alertaCruzamento = null; }

    const corpoConf = document.getElementById('corpoConfrontantes');
    if (corpoConf) corpoConf.innerHTML = '';

    // Marcadores numerados dos selecionados
    selecionados.forEach((p, idx) => {
        const m = L.marker([p.lat, p.lon], {
            icon: L.divIcon({ className: '',
                html: `<div style="background:#1a73e8;border:2px solid #fff;border-radius:50%;width:26px;height:26px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:bold;color:#fff;box-shadow:0 2px 6px rgba(0,0,0,0.4);cursor:pointer;">${idx+1}</div>`,
                iconSize:[26,26], iconAnchor:[13,13] }), zIndexOffset: 1000
        }).addTo(map);
        m.bindTooltip(`<b>${p.label||p.nome}</b>${p.label&&p.label!==p.nome?' <small>('+p.nome+')</small>':''}<br>${formatarDMSLegivel(p.latDMS)}<br>${formatarDMSLegivel(p.lonDMS)}`, { permanent:false, direction:'top', offset:[0,-14] });
        m.on('click', () => { const i=selecionados.findIndex(s=>s.nome===p.nome); if(i!==-1) selecionados.splice(i,1); sincronizarAtivos(); atualizarMapa(); renderizarOrdenavel(); _autoSave(); });
        marcadoresSel.push(m);
    });

    _renderizarMarcadoresBase();

    if (selecionados.length > 1) {
        for (let i = 0; i < selecionados.length; i++) {
            const p1 = selecionados[i], p2 = selecionados[(i+1) % selecionados.length];
            if (i === selecionados.length - 1 && selecionados.length < 3) continue;
            const dist   = L.latLng(p1.lat,p1.lon).distanceTo(L.latLng(p2.lat,p2.lon));
            const azGrau = (Math.atan2(p2.lon-p1.lon, p2.lat-p1.lat) * 180/Math.PI + 360) % 360;
            const mid    = [(p1.lat+p2.lat)/2, (p1.lon+p2.lon)/2];

            labelsDistancia.push(L.marker(mid, { icon: L.divIcon({ className:'label-distancia', html: dist.toFixed(2)+'m', iconAnchor:[0,-8] }) }).addTo(map));
            labelsAzimute.push(L.marker([(p1.lat*.4+p2.lat*.6),(p1.lon*.4+p2.lon*.6)], { icon: L.divIcon({ className:'label-azimute', html: azimutePorExtenso(azGrau), iconAnchor:[0,10] }) }).addTo(map));

            if (corpoConf) {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td style="border:1px solid #ddd;padding:4px;font-size:11px;white-space:nowrap;">${p1.nome}–${p2.nome}</td>
                    <td style="border:1px solid #ddd;padding:4px;"><input type="text" class="input-vizinho" style="width:98%;border:1px solid #ccc;padding:2px;font-size:11px;" placeholder="Nome do confrontante" oninput="_autoSave()"></td>`;
                corpoConf.appendChild(tr);
            }
        }
        calcularArea();
        if (selecionados.length > 2) { map.fitBounds(poligonal.getBounds(), { padding:[40,40] }); _validarCruzamento(); }
    } else {
        atualizarElementos('txtArea', '0.00 m²'); atualizarElementos('txtPeri', '0.00 m');
    }
    atualizarTabelaAltimetria();
    _atualizarBotoesExportar();
}

// ─────────────────────────────────────────────────────────────
// 10. VALIDAÇÃO DE CRUZAMENTO
// ─────────────────────────────────────────────────────────────
function _validarCruzamento() {
    const n = selecionados.length; if (n < 4) return;
    const seg = selecionados.map((p,i) => [p, selecionados[(i+1)%n]]);
    function ccw(A,B,C) { return (C.lat-A.lat)*(B.lon-A.lon) > (B.lat-A.lat)*(C.lon-A.lon); }
    function intersecta(a,b,c,d) { return ccw(a,c,d)!==ccw(b,c,d) && ccw(a,b,c)!==ccw(a,b,d); }

    let cruzou = false;
    outer: for (let i=0;i<seg.length;i++) for (let j=i+2;j<seg.length;j++) {
        if (i===0 && j===seg.length-1) continue;
        if (intersecta(seg[i][0],seg[i][1],seg[j][0],seg[j][1])) { cruzou=true; break outer; }
    }

    const aviso = document.getElementById('avisoCruzamento');
    if (cruzou) {
        const centro = poligonal.getBounds().getCenter();
        alertaCruzamento = L.marker(centro, { icon: L.divIcon({ className:'',
            html:`<div style="background:#c0392b;color:white;padding:5px 10px;border-radius:4px;font-size:11px;font-weight:bold;white-space:nowrap;box-shadow:0 2px 6px rgba(0,0,0,0.4);">⚠ Poligonal auto-intersectante</div>`,
            iconAnchor:[90,10] }), zIndexOffset:2000 }).addTo(map);
        if (aviso) aviso.style.display = 'block';
    } else {
        if (aviso) aviso.style.display = 'none';
    }
}

// ─────────────────────────────────────────────────────────────
// 11. EXPORTAÇÃO CSV e KML
// ─────────────────────────────────────────────────────────────
function exportarCSV() {
    if (!selecionados.length) { alert('Selecione vértices antes de exportar.'); return; }
    const linhas = ['nome,label,grupo,latDMS,lonDMS,lat,lon,alt', ...selecionados.map(p => `${p.nome},${p.label||p.nome},${p.grupo||'perimetro'},${p.latDMS},${p.lonDMS},${p.lat},${p.lon},${p.alt??0}`)];
    _download(linhas.join('\n'), 'vertices_geosistema.csv', 'text/csv');
}

function exportarKML() {
    if (!selecionados.length) { alert('Selecione vértices antes de exportar.'); return; }
    const marks = selecionados.map(p => `  <Placemark><name>${p.nome}</name><Point><coordinates>${p.lon},${p.lat},${p.alt??0}</coordinates></Point></Placemark>`).join('\n');
    const coords = [...selecionados, selecionados[0]].map(p => `${p.lon},${p.lat},0`).join(' ');
    const kml = `<?xml version="1.0" encoding="UTF-8"?>\n<kml xmlns="http://www.opengis.net/kml/2.2">\n<Document>\n  <name>GeoSistema PRO</name>\n${marks}\n  <Placemark><name>Poligonal</name><Polygon><outerBoundaryIs><LinearRing><coordinates>${coords}</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>\n</Document>\n</kml>`;
    _download(kml, 'poligonal_geosistema.kml', 'application/vnd.google-earth.kml+xml');
}

function _download(conteudo, nome, tipo) {
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob([conteudo],{type:tipo})), download: nome });
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(a.href); }, 100);
}

// ─────────────────────────────────────────────────────────────
// 12. ALTIMETRIA + GRÁFICO
// ─────────────────────────────────────────────────────────────
function atualizarTabelaAltimetria() {
    const corpo = document.getElementById('corpoAltimetria'); if (!corpo) return;
    corpo.innerHTML = '';
    const lista = selecionados.length ? selecionados : pontosBase;
    lista.forEach(p => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td><b>${p.label||p.nome}</b>${p.label&&p.label!==p.nome?'<br><small style="color:#aaa;font-size:8px;">'+p.nome+'</small>':''}</td><td>${formatarDMSLegivel(p.latDMS)}</td><td>${formatarDMSLegivel(p.lonDMS)}</td><td>${p.alt!=null?parseFloat(p.alt).toFixed(2):'—'}</td>`;
        corpo.appendChild(tr);
    });
    _atualizarGraficoRelevo(lista);
}

function _atualizarGraficoRelevo(lista) {
    const canvas = document.getElementById('graficoRelevo'); if (!canvas || typeof Chart === 'undefined') return;
    if (canvas._chartInstance) canvas._chartInstance.destroy();
    canvas._chartInstance = new Chart(canvas, { type:'line', data:{ labels:lista.map(p=>p.label||p.nome), datasets:[{ label:'Altitude (m)', data:lista.map(p=>parseFloat(p.alt)||0), borderColor:'#1a73e8', backgroundColor:'rgba(26,115,232,0.1)', fill:true, tension:0.3, pointRadius:4, pointBackgroundColor:'#1a73e8' }] }, options:{ responsive:true, maintainAspectRatio:false, scales:{ y:{title:{display:true,text:'Altitude (m)'}}, x:{title:{display:true,text:'Vértice'}} } } });
}

// ─────────────────────────────────────────────────────────────
// 13. CÁLCULOS
// ─────────────────────────────────────────────────────────────
function converterDMS(dms) {
    if (!dms||typeof dms!=='string') return 0;
    if (/^-?[\d.]+$/.test(dms.trim())) return parseFloat(dms);
    const n = dms.replace(/[°d'"ms]/gi,' ').trim().match(/-?[\d.]+/g);
    if (!n||n.length<2) return parseFloat(dms)||0;
    const r = Math.abs(parseFloat(n[0])) + (parseFloat(n[1])/60||0) + (parseFloat(n[2])/3600||0);
    return (dms.trim().startsWith('-')||/[SWsw]/.test(dms)) ? -r : r;
}

function formatarDMSLegivel(dms) {
    if (!dms||typeof dms!=='string') return String(dms||'');
    if (/^-?[\d.]+$/.test(dms.trim())) return decimalParaDMSLegivel(parseFloat(dms));
    const n = dms.replace(/[°d'"ms]/gi,' ').trim().match(/-?[\d.]+/g);
    if (!n||n.length<2) return dms;
    return `${dms.trim().startsWith('-')?'-':''}${Math.abs(parseInt(n[0]))}°${String(parseInt(n[1])).padStart(2,'0')}'${parseFloat(n[2]||0).toFixed(2)}"`;
}

function decimalParaDMSLegivel(dec) {
    const s=dec<0?'-':'', a=Math.abs(dec), g=Math.floor(a), mF=(a-g)*60, m=Math.floor(mF), sg=((mF-m)*60).toFixed(2);
    return `${s}${g}°${String(m).padStart(2,'0')}'${sg}"`;
}

function calcularArea() {
    if (selecionados.length < 3) return;
    let a=0, p=0;
    for (let i=0;i<selecionados.length;i++) {
        const j=(i+1)%selecionados.length;
        a += (selecionados[i].lon*selecionados[j].lat) - (selecionados[j].lon*selecionados[i].lat);
        p += L.latLng(selecionados[i].lat,selecionados[i].lon).distanceTo(L.latLng(selecionados[j].lat,selecionados[j].lon));
    }
    const fL=111319.9, fLo=111319.9*Math.cos(selecionados[0].lat*Math.PI/180);
    const areaMq=Math.abs(a/2)*fL*fLo, areaHa=areaMq/10000;
    const exibir = areaMq>=10000 ? `${areaHa.toFixed(4)} ha (${areaMq.toFixed(2)} m²)` : `${areaMq.toFixed(2)} m²`;
    atualizarElementos('txtArea', exibir);
    atualizarElementos('txtPeri', p.toFixed(2)+' m');
    atualizarElementos('areaFinal', areaMq.toFixed(2));
    atualizarElementos('periFinal', p.toFixed(2));
}

function azimutePorExtenso(az) {
    let base, quad;
    if      (az<90)  { base=az;       quad=['N','E']; }
    else if (az<180) { base=180-az;   quad=['S','E']; }
    else if (az<270) { base=az-180;   quad=['S','O']; }
    else             { base=360-az;   quad=['N','O']; }
    const g=Math.floor(base), mF=(base-g)*60, m=Math.floor(mF), s=Math.round((mF-m)*60);
    return `${quad[0]} ${g}°${String(m).padStart(2,'0')}'${String(s).padStart(2,'0')}" ${quad[1]}`;
}

// ─────────────────────────────────────────────────────────────
// 14. GERAÇÃO DO MEMORIAL
// ─────────────────────────────────────────────────────────────
function prepararImpressao() {
    if (selecionados.length < 3) { alert('Selecione pelo menos 3 vértices.'); return; }

    const proprietario = _val('nomeProprietario') || 'NÃO INFORMADO';
    const matricula    = _val('matriculaImovel')  || 'NÃO INFORMADA';
    const municipio    = _val('nomeMunicipio')    || '___________';
    const comarca      = _val('nomeComarca')      || '___________';
    const cartorio     = _val('nomeCartorio')     || '___________';
    const datum        = _val('nomeDatum')        || 'SIRGAS 2000';
    const rt           = _val('nomeRT')           || 'EVERMONDO LUCAS';
    const crea         = _val('numCREA')          || '___________';
    const hoje         = new Date().toLocaleDateString('pt-BR');
    const areaStr      = document.getElementById('txtArea')?.innerText || '–';
    const periStr      = document.getElementById('txtPeri')?.innerText || '–';

    let memorial = `
    <div class="print-header">
        <div><h2 style="margin:0;color:#1a73e8;font-size:16pt;">GEOSISTEMA PRO</h2>
        <p style="margin:2pt 0;font-size:9pt;color:#555;">Engenharia e Agrimensura | Georreferenciamento de Imóveis</p></div>
        <div style="text-align:right;font-size:9pt;color:#555;"><p style="margin:0;">Emitido em: <b>${hoje}</b></p><p style="margin:0;">Datum: <b>${datum}</b></p></div>
    </div>
    <h3 style="text-align:center;margin:0 0 6pt;font-size:14pt;">MEMORIAL DESCRITIVO</h3>
    <p style="text-align:center;font-size:9pt;color:#555;margin:0 0 12pt;">Município: <b>${municipio}</b> &nbsp;|&nbsp; Comarca: <b>${comarca}</b> &nbsp;|&nbsp; Cartório: <b>${cartorio}</b></p>
    <table class="print-table" style="margin-bottom:12pt;">
        <tr><th style="width:30%;">Proprietário</th><td colspan="3">${proprietario}</td></tr>
        <tr><th>Matrícula</th><td>${matricula}</td><th style="width:20%;">Área Total</th><td>${areaStr}</td></tr>
        <tr><th>Perímetro</th><td>${periStr}</td><th>Sistema Geodésico</th><td>${datum}</td></tr>
    </table>`;

    const p0 = selecionados[0];
    memorial += `<p style="text-align:justify;line-height:1.7;margin-bottom:8pt;">Inicia-se a descrição deste perímetro no vértice <b>${p0.label||p0.nome}</b>, de coordenadas Latitude <b>${formatarDMSLegivel(p0.latDMS)}</b> e Longitude <b>${formatarDMSLegivel(p0.lonDMS)}</b>, altitude <b>${p0.alt!=null?parseFloat(p0.alt).toFixed(2)+'m':'—'}</b>, referenciadas ao Sistema Geodésico Brasileiro (<b>${datum}</b>), Município de <b>${municipio}</b>, Comarca de <b>${comarca}</b>, registrado no Cartório de <b>${cartorio}</b> sob matrícula nº <b>${matricula}</b>.</p>`;

    const inputs = document.querySelectorAll('.input-vizinho');
    let corpo = '';
    for (let i=0;i<selecionados.length;i++) {
        const p1=selecionados[i], p2=selecionados[(i+1)%selecionados.length];
        const viz  = inputs[i]?.value || 'CONFRONTANTE A DESIGNAR';
        const dist = L.latLng(p1.lat,p1.lon).distanceTo(L.latLng(p2.lat,p2.lon));
        const az   = azimutePorExtenso((Math.atan2(p2.lon-p1.lon,p2.lat-p1.lat)*180/Math.PI+360)%360);
        corpo += `Deste, segue confrontando com <b>${viz}</b>, azimute <b>${az}</b>, distância <b>${dist.toFixed(2)}m</b>, até o vértice <b>${p2.label||p2.nome}</b>, Latitude <b>${formatarDMSLegivel(p2.latDMS)}</b>, Longitude <b>${formatarDMSLegivel(p2.lonDMS)}</b>, altitude <b>${p2.alt!=null?parseFloat(p2.alt).toFixed(2)+'m':'—'}</b> (${datum}); `;
    }
    memorial += `<p style="text-align:justify;line-height:1.7;">${corpo}</p>`;

    memorial += `<div class="print-section"><h4 style="margin:12pt 0 6pt;">QUADRO DE COORDENADAS — ${datum}</h4>
    <table class="print-table"><thead><tr><th>Vértice</th><th>Latitude</th><th>Longitude</th><th>Alt (m)</th><th>Azimute → próximo</th><th>Dist (m)</th></tr></thead><tbody>`;
    selecionados.forEach((p,i) => {
        const nx=selecionados[(i+1)%selecionados.length];
        const dist=L.latLng(p.lat,p.lon).distanceTo(L.latLng(nx.lat,nx.lon));
        const az=azimutePorExtenso((Math.atan2(nx.lon-p.lon,nx.lat-p.lat)*180/Math.PI+360)%360);
        memorial+=`<tr><td><b>${p.label||p.nome}</b>${p.label&&p.label!==p.nome?'<br><small style="color:#aaa;font-size:8pt;">(ID: '+p.nome+')</small>':''}</td><td>${formatarDMSLegivel(p.latDMS)}</td><td>${formatarDMSLegivel(p.lonDMS)}</td><td>${p.alt!=null?parseFloat(p.alt).toFixed(2):'—'}</td><td>${az}</td><td>${dist.toFixed(2)}</td></tr>`;
    });
    memorial += `</tbody></table></div>`;

    // Seção de edificações e pontos internos (não participam da poligonal)
    const pontosInternos = pontosBase.filter(p => (p.grupo||'perimetro') !== 'perimetro');
    if (pontosInternos.length) {
        const porGrupo = {};
        pontosInternos.forEach(p => { const g=p.grupo||'outro'; if(!porGrupo[g])porGrupo[g]=[]; porGrupo[g].push(p); });
        memorial += `<div class="print-section"><h4 style="margin:12pt 0 6pt;">ESTRUTURAS E PONTOS INTERNOS</h4>`;
        Object.entries(porGrupo).forEach(([g, pts]) => {
            const info = _grupoInfo(g);
            memorial += `<p style="font-size:10pt;font-weight:bold;margin:8pt 0 4pt;">${info.label} — ${pts.length} ponto(s)</p>`;
            memorial += `<table class="print-table"><thead><tr><th>Ponto</th><th>Latitude</th><th>Longitude</th><th>Alt (m)</th></tr></thead><tbody>`;
            pts.forEach(p => {
                memorial += `<tr><td><b>${p.label||p.nome}</b>${p.label&&p.label!==p.nome?'<br><small style="color:#aaa;font-size:8pt;">(ID: '+p.nome+')</small>':''}</td><td>${formatarDMSLegivel(p.latDMS)}</td><td>${formatarDMSLegivel(p.lonDMS)}</td><td>${p.alt!=null?parseFloat(p.alt).toFixed(2):'—'}</td></tr>`;
            });
            memorial += `</tbody></table>`;
        });
        memorial += `</div>`;
    }

    memorial += `
    <div class="print-assinaturas">
        <div class="print-assinatura"><div class="linha"></div><b>${rt}</b><br>Responsável Técnico<br>CREA/CFT: ${crea}</div>
        <div class="print-assinatura"><div class="linha"></div><b>${proprietario}</b><br>Proprietário / Posseiro<br>CPF: ____________________</div>
    </div>
    <p style="text-align:center;font-size:8pt;color:#888;margin-top:20pt;">GeoSistema PRO | ${hoje} | ${datum}</p>`;

    const alvo = document.getElementById('textoNarrativo');
    const ap   = document.getElementById('printMemorial');
    if (alvo) alvo.innerHTML = memorial;
    if (ap)   ap.innerHTML   = memorial;
    if (map) map.invalidateSize();
    setTimeout(() => window.print(), 800);
}

// ─────────────────────────────────────────────────────────────
// 15. LIMPAR SELEÇÃO
// ─────────────────────────────────────────────────────────────
function limparSelecao() {
    selecionados = [];
    if (map && poligonal) {
        poligonal.setLatLngs([]);
        [...labelsDistancia,...labelsAzimute,...marcadoresSel].forEach(l=>map.removeLayer(l));
        labelsDistancia=[]; labelsAzimute=[]; marcadoresSel=[];
        if (alertaCruzamento) { map.removeLayer(alertaCruzamento); alertaCruzamento=null; }
        _renderizarMarcadoresBase();
    }
    sincronizarAtivos();
    atualizarElementos('txtArea','0.00 m²'); atualizarElementos('txtPeri','0.00 m');
    atualizarElementos('areaFinal','0');     atualizarElementos('periFinal','0');
    const cc=document.getElementById('corpoConfrontantes'); if(cc) cc.innerHTML='';
    const av=document.getElementById('avisoCruzamento');    if(av) av.style.display='none';
    renderizarOrdenavel(); atualizarTabelaAltimetria(); _autoSave();
}
function limpar() { limparSelecao(); }

function _criarPainelUTM() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar || document.getElementById('painelUTM')) return;

    const div = document.createElement('div');
    div.id = 'painelUTM';
    div.style.cssText = 'display:none;background:#f0f7ff;border:1px solid #1a73e8;border-radius:6px;padding:10px;margin-top:8px;';
    div.innerHTML = `
        <div style="font-size:12px;font-weight:bold;color:#1a73e8;margin-bottom:6px;">Conversor UTM → Geográfico</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-bottom:5px;">
            <div><div style="font-size:10px;color:#555;">Easting (E)</div>
                <input id="utmE" type="number" placeholder="ex: 456789" style="width:100%;padding:4px;border:1px solid #ccc;border-radius:3px;font-size:12px;"></div>
            <div><div style="font-size:10px;color:#555;">Northing (N)</div>
                <input id="utmN" type="number" placeholder="ex: 8765432" style="width:100%;padding:4px;border:1px solid #ccc;border-radius:3px;font-size:12px;"></div>
            <div><div style="font-size:10px;color:#555;">Fuso</div>
                <input id="utmFuso" type="number" value="24" min="1" max="60" style="width:100%;padding:4px;border:1px solid #ccc;border-radius:3px;font-size:12px;"></div>
            <div><div style="font-size:10px;color:#555;">Hemisfério</div>
                <select id="utmHem" style="width:100%;padding:4px;border:1px solid #ccc;border-radius:3px;font-size:12px;">
                    <option value="S" selected>S — Sul</option><option value="N">N — Norte</option>
                </select></div>
        </div>
        <button onclick="converterUTMManual()" style="width:100%;background:#1a73e8;color:white;border:none;border-radius:4px;padding:6px;font-size:12px;cursor:pointer;font-weight:bold;">Converter</button>
        <div id="resultadoUTM" style="font-size:11px;margin-top:6px;color:#333;min-height:16px;"></div>`;

    const ref = sidebar.querySelector('.section-title');
    if (ref) sidebar.insertBefore(div, ref); else sidebar.appendChild(div);
}

function togglePainelUTM() {
    const p = document.getElementById('painelUTM');
    if (p) p.style.display = p.style.display === 'none' ? 'block' : 'none';
}

function converterUTMManual() {
    const E   = parseFloat(document.getElementById('utmE')?.value);
    const N   = parseFloat(document.getElementById('utmN')?.value);
    const fus = parseInt(document.getElementById('utmFuso')?.value) || 24;
    const hem = document.getElementById('utmHem')?.value || 'S';
    const res = document.getElementById('resultadoUTM');
    if (!res) return;
    if (isNaN(E)||isNaN(N)) { res.textContent='⚠ Informe E e N válidos.'; res.style.color='#c00'; return; }
    const g = utmParaGeo(E, N, fus, hem);
    res.style.color = '#1a7a40';
    res.innerHTML = `<b>Lat:</b> ${g.lat.toFixed(8)}° &nbsp;(${decimalParaDMSLegivel(g.lat)})<br><b>Lon:</b> ${g.lon.toFixed(8)}° &nbsp;(${decimalParaDMSLegivel(g.lon)})`;
}

// ─────────────────────────────────────────────────────────────
// 17. BOTÕES EXTRAS
// ─────────────────────────────────────────────────────────────
function _criarBotoesExtra() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar || document.getElementById('botoesExtra')) return;

    const aviso = document.createElement('div');
    aviso.id = 'avisoCruzamento';
    aviso.style.cssText = 'display:none;background:#fde8e8;border:1px solid #c0392b;border-radius:4px;padding:7px 10px;font-size:12px;color:#c0392b;font-weight:bold;margin-top:6px;';
    aviso.textContent = '⚠ Poligonal auto-intersectante! Verifique a ordem dos vértices.';

    const bar = document.createElement('div');
    bar.id = 'botoesExtra';
    bar.style.cssText = 'display:flex;flex-direction:column;gap:5px;margin-top:8px;';
    bar.innerHTML = `
        <div style="background:#f8f9fa;border:1px solid #ddd;border-radius:4px;padding:7px;font-size:11px;">
            <div style="font-weight:bold;color:#555;margin-bottom:4px;">Legenda dos Grupos</div>
            ${Object.entries(GRUPOS_CORES).map(([k,v])=>`<div style="display:flex;align-items:center;gap:5px;margin-bottom:2px;"><span style="color:${v.cor};font-size:13px;">${v.icone}</span><span style="color:${v.cor};font-weight:bold;">${v.label}</span></div>`).join('')}
            <div style="color:#999;font-size:10px;margin-top:4px;">Use ✏ na lista para renomear. Use o seletor de grupo para classificar.</div>
        </div>
        <button onclick="togglePainelUTM()" style="background:#e8f0fe;color:#1a73e8;border:1px solid #1a73e8;border-radius:4px;padding:7px;font-size:12px;cursor:pointer;font-weight:bold;">📐 UTM → Geográfico</button>
        <div style="display:flex;gap:5px;">
            <button id="btnExportCSV" onclick="exportarCSV()" disabled style="flex:1;background:#e8f5e9;color:#1a7a40;border:1px solid #1a7a40;border-radius:4px;padding:7px;font-size:12px;cursor:pointer;font-weight:bold;opacity:0.4;">⬇ CSV</button>
            <button id="btnExportKML" onclick="exportarKML()" disabled style="flex:1;background:#fff3e0;color:#e67e00;border:1px solid #e67e00;border-radius:4px;padding:7px;font-size:12px;cursor:pointer;font-weight:bold;opacity:0.4;">⬇ KML</button>
        </div>
        <button onclick="novaSessao()" style="background:#fde8e8;color:#c0392b;border:1px solid #c0392b;border-radius:4px;padding:7px;font-size:12px;cursor:pointer;font-weight:bold;">🗑 Nova Sessão</button>`;

    const btnPrint = sidebar.querySelector('.btn-print');
    if (btnPrint) { btnPrint.parentNode.insertBefore(aviso, btnPrint.nextSibling); btnPrint.parentNode.insertBefore(bar, aviso.nextSibling); }
    else { sidebar.appendChild(aviso); sidebar.appendChild(bar); }
}

function _atualizarBotoesExportar() {
    ['btnExportCSV','btnExportKML'].forEach(id => {
        const btn = document.getElementById(id); if (!btn) return;
        const ativo = selecionados.length > 0;
        btn.disabled = !ativo; btn.style.opacity = ativo ? '1' : '0.4';
    });
}

// ─────────────────────────────────────────────────────────────
// 18. MOBILE + ESTILOS
// ─────────────────────────────────────────────────────────────
function _criarBotaoMobile() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar || document.getElementById('btnToggleSidebar')) return;
    const btn = document.createElement('button');
    btn.id = 'btnToggleSidebar'; btn.innerHTML = '☰';
    btn.style.cssText = 'display:none;position:fixed;top:10px;left:10px;z-index:3000;background:#1a73e8;color:white;border:none;border-radius:6px;width:40px;height:40px;font-size:20px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,0.4);';
    btn.onclick = () => sidebar.classList.toggle('sidebar-open');
    document.body.appendChild(btn);
}

function _injetarEstilos() {
    if (document.getElementById('_geoEstilos')) return;
    const s = document.createElement('style');
    s.id = '_geoEstilos';
    s.textContent = `
        @media (max-width:768px) {
            #btnToggleSidebar { display:flex !important; align-items:center; justify-content:center; }
            #sidebar { position:fixed !important; top:0; left:-340px; height:100vh; z-index:2500; transition:left 0.3s; box-shadow:4px 0 12px rgba(0,0,0,0.3); overflow-y:auto; }
            #sidebar.sidebar-open { left:0 !important; }
            .main-container { flex-direction:column; }
            #map-container { height:calc(100vh - 56px); }
        }
        .label-azimute { background:rgba(26,115,232,0.85); color:white; border-radius:3px; padding:1px 5px; font-size:9px !important; font-weight:bold; white-space:nowrap; border:none; }
        .label-distancia { background:white; border:1px solid #1a73e8; border-radius:3px; padding:1px 4px; font-size:10px !important; font-weight:bold; color:#1a73e8; white-space:nowrap; }
        #listaOrdenavel .ponto-sel { display:flex; align-items:center; gap:6px; padding:6px 8px; margin-bottom:3px; background:#d1e7dd; border:1px solid #0f5132; border-radius:4px; font-size:12px; cursor:grab; user-select:none; }
        #listaOrdenavel .ponto-sel.dragging { opacity:0.4; }
        #listaOrdenavel .ponto-sel .handle { font-size:14px; color:#666; }
        #listaOrdenavel .ponto-sel .btn-rem { margin-left:auto; background:none; border:none; color:#c00; cursor:pointer; font-size:14px; padding:0 2px; width:auto; margin-top:0; }
        @media print {
            @page { size:A4; margin:2cm; }
            body { font-family:'Times New Roman',serif; font-size:11pt; }
            .no-print, header, #sidebar, #btnToggleSidebar { display:none !important; }
            #areaImpressao { display:block !important; }
            .print-section { page-break-before:always; }
            .print-header { display:flex; justify-content:space-between; border-bottom:2pt solid #000; margin-bottom:12pt; padding-bottom:8pt; }
            .print-assinaturas { display:flex; justify-content:space-around; margin-top:60pt; page-break-inside:avoid; }
            .print-assinatura { width:40%; text-align:center; }
            .print-assinatura .linha { border-top:1pt solid #000; margin-bottom:4pt; }
            table.print-table { width:100%; border-collapse:collapse; font-size:9pt; }
            table.print-table th, table.print-table td { border:1pt solid #000; padding:3pt 5pt; }
            table.print-table th { background:#f0f0f0 !important; font-weight:bold; }
        }`;
    document.head.appendChild(s);
}

// ─────────────────────────────────────────────────────────────
// INICIALIZAÇÃO
// ─────────────────────────────────────────────────────────────
window.addEventListener('load', initMap);