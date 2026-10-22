/* Painel do Inventário: lê a planilha na mesma pasta (ou escolhida pelo botão) e desenha o painel. */
(function () {
  'use strict';

  var ARQUIVO = 'Inventario_Carrinhos_conferido.xlsx';
  var LS_KEY = 'painel-inventario-resolvidos';
  var $ = function (s) { return document.querySelector(s); };
  var NS = 'http://www.w3.org/2000/svg';
  var fmtN = new Intl.NumberFormat('pt-BR');
  var fmtP = new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 0 });

  var D = null; // dados lidos da planilha
  var st = { g: null, cart: '', q: '', s: 'open', sel: null, at: '', acart: '' };
  var local = new Set();

  try { JSON.parse(localStorage.getItem(LS_KEY) || '[]').forEach(function (k) { local.add(k); }); } catch (e) { /* sem armazenamento local */ }
  function saveLocal() { try { localStorage.setItem(LS_KEY, JSON.stringify(Array.from(local))); } catch (e) { /* ignora */ } }

  /* ---------- leitura da planilha ---------- */
  function txt(v) { return v == null ? '' : String(v).trim(); }
  function num(v) { var n = Number(v); return isFinite(n) ? n : 0; }

  function rowsOf(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws) throw new Error('A aba "' + name + '" não foi encontrada na planilha.');
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }

  function parse(wb) {
    var res = rowsOf(wb, 'Resumo'), dv = rowsOf(wb, 'Divergências'), cf = rowsOf(wb, 'Conferência');
    var carts = [], i;
    var lido = false;
    for (i = 0; i < res.length; i++) {
      var a = txt(res[i][0]);
      if (!lido) { if (a === 'Carrinho') lido = true; continue; }
      if (!a) continue;
      if (a === 'Total') break;
      carts.push({ nome: a, word: num(res[i][1]), confere: num(res[i][2]), local: num(res[i][3]), letra: num(res[i][4]), nao: num(res[i][5]), sistema: num(res[i][6]), sobras: num(res[i][7]) });
    }
    var cases = [], att = [], mode = 0;
    for (i = 0; i < dv.length; i++) {
      var r = dv[i], c0 = txt(r[0]);
      if (!c0) continue;
      if (c0.indexOf('Para aten') === 0) { mode = 1; continue; }
      if (c0 === 'Caso' || c0 === '#') continue;
      if (isNaN(Number(c0))) continue;
      if (mode === 0) {
        cases.push({ n: Number(c0), cart: txt(r[1]), item: txt(r[2]), tipo: txt(r[3]), snWord: txt(r[4]), snSis: txt(r[6]), local: txt(r[7]), porque: txt(r[8]), fazer: txt(r[9]), ok: txt(r[10]).toLowerCase() === 'sim' });
      } else {
        att.push({ id: Number(c0), cart: txt(r[1]), item: txt(r[2]), tipo: txt(r[3]), sn: txt(r[4]), inf: txt(r[8]) });
      }
    }
    cases.forEach(function (c) { c.key = [c.cart, c.item, c.tipo, c.snWord, c.snSis].join('|'); });

    // Conferência: linhas a partir da primeira com cabeçalho "Carrinho (Word)"
    var items = [], started = false;
    for (i = 0; i < cf.length; i++) {
      var q = cf[i];
      if (!started) { if (txt(q[0]).indexOf('Carrinho (Word)') === 0) started = true; continue; }
      if (!txt(q[0])) continue;
      items.push({ cart: txt(q[0]), n: txt(q[1]), key: txt(q[3]), status: txt(q[10]), obs: txt(q[12]) });
    }
    if (!carts.length || !items.length) throw new Error('A planilha não tem o formato esperado (abas Resumo, Divergências e Conferência).');
    return { carts: carts, cases: cases, att: att, items: items };
  }

  /* ---------- estado de uso, a partir das anotações do Word ---------- */
  var RE_FORA = /n[ãa]o liga|n[ãa]o est[áa] ligando|n[ãa]o carrega|tela quebrada|imposs[íi]vel de usar|totalmente quebrado|com defeito|n[ãa]o est[áa] pr[óo]prio|problema na tela|tela piscando sozinha|tela est[áa] com problema|tremendo/i;
  var RE_RESSALVA = /dobradi[çc]a|abre e fecha|soltando|v[áa]rias teclas|n[ãa]o est[ãa]o funcionando|power|piscando/i;
  function uso(obs) { if (RE_FORA.test(obs)) return 'fora'; if (RE_RESSALVA.test(obs)) return 'ressalva'; return 'ok'; }

  var RANK = { ok: 0, ressalva: 1, fora: 2 };
  function resumoUso(items) {
    var porCart = {}, vistos = new Map(), tot = { contados: 0, ok: 0, ressalva: 0, fora: 0, confere: 0, linhas: items.length };
    items.forEach(function (it) {
      var u = uso(it.obs);
      var o = porCart[it.cart] || (porCart[it.cart] = { ok: 0, ressalva: 0, fora: 0 });
      o[u]++;
      if (it.status === 'Confere') tot.confere++;
      var ant = vistos.get(it.key);
      if (!ant || RANK[u] > RANK[ant]) vistos.set(it.key, u); // N/S repetido em dois Words: vale o pior estado anotado
    });
    vistos.forEach(function (u) { tot.contados++; tot[u]++; });
    return { porCart: porCart, tot: tot };
  }

  /* ---------- nomes ---------- */
  function rotulo(nome) {
    var m = /Carrinho\s+(\d+)/i.exec(nome);
    if (/RESERVA/i.test(nome)) return 'Carrinho Reserva';
    if (m && m[1] === '13') return 'Carrinho 13 (quebrados)';
    return m ? 'Carrinho ' + m[1] : nome;
  }
  function grupo(tipo) {
    if (tipo === 'Local diferente no sistema') return 'sistema';
    if (tipo.indexOf('Sobra') === 0) return 'procurar';
    return 'etiqueta';
  }
  function feito(c) { return c.ok || local.has(c.key); }

  /* ---------- helpers de DOM ---------- */
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function sv(tag, attrs) { var e = document.createElementNS(NS, tag); Object.keys(attrs || {}).forEach(function (k) { e.setAttribute(k, attrs[k]); }); return e; }
  var COR = ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)'];

  function legend(sel, parts) {
    var box = $(sel); box.replaceChildren();
    parts.forEach(function (p) { var s = el('span'); var i = el('i'); i.style.background = p.c; s.append(i, p.l); box.append(s); });
  }

  function stacked(sel, rows, parts) {
    var box = $(sel); box.replaceChildren();
    var w = Math.max(300, box.clientWidth || 520), rowH = 30, top = 6, bottom = 26;
    var h = top + rows.length * rowH + bottom;
    var svg = sv('svg', { viewBox: '0 0 ' + w + ' ' + h, role: 'img' });
    var probe = sv('text', { style: 'font-size:12px' }); svg.append(probe); box.append(svg);
    var lw = 0;
    rows.forEach(function (r) { probe.textContent = r.rot; lw = Math.max(lw, probe.getComputedTextLength()); });
    probe.remove();
    var left = Math.ceil(lw) + 12, right = 34;
    var maxv = Math.max.apply(null, rows.map(function (r) { return parts.reduce(function (s, p) { return s + (r[p.k] || 0); }, 0); }).concat([1]));
    var x = function (v) { return left + (w - left - right) * v / maxv; };
    var step = maxv > 30 ? 10 : 5;
    for (var t = 0; t <= maxv; t += step) {
      svg.append(sv('line', { x1: x(t), x2: x(t), y1: top, y2: h - bottom, stroke: 'var(--grid)' }));
      var tl = sv('text', { x: x(t), y: h - 8, 'text-anchor': 'middle', fill: 'var(--muted)', style: 'font-size:11px' }); tl.textContent = t; svg.append(tl);
    }
    rows.forEach(function (r, i) {
      var y = top + i * rowH + 4, bh = rowH - 9;
      var lb = sv('text', { x: left - 8, y: y + bh / 2, dy: '0.35em', 'text-anchor': 'end', fill: 'var(--fg)', style: 'font-size:12px' }); lb.textContent = r.rot; svg.append(lb);
      var acc = 0;
      parts.forEach(function (p) {
        var v = r[p.k] || 0; if (!v) return;
        var rc = sv('rect', { x: x(acc), y: y, width: Math.max(1, x(acc + v) - x(acc) - 1), height: bh, rx: 3, fill: p.c });
        var tt = sv('title'); tt.textContent = r.rot + ' · ' + p.l + ': ' + v; rc.append(tt); svg.append(rc);
        acc += v;
      });
      var tot = sv('text', { x: x(acc) + 5, y: y + bh / 2, dy: '0.35em', fill: 'var(--muted)', style: 'font-size:11px' }); tot.textContent = acc; svg.append(tot);
    });
  }

  /* ---------- desenho ---------- */
  function drawAll() {
    var U = resumoUso(D.items), tot = U.tot;
    var abertos = D.cases.filter(function (c) { return !feito(c); });
    var g = { sistema: 0, etiqueta: 0, procurar: 0 };
    abertos.forEach(function (c) { g[grupo(c.tipo)]++; });

    $('#k-contados').textContent = fmtN.format(tot.contados);
    $('#k-pct').textContent = fmtP.format(tot.linhas ? tot.confere / tot.linhas : 0);
    $('#k-conf').textContent = fmtN.format(tot.confere) + ' de ' + fmtN.format(tot.linhas) + ' linhas';
    $('#k-casos').textContent = fmtN.format(abertos.length);
    $('#k-casos-n').textContent = 'de ' + D.cases.length + ' no total';
    $('#k-usaveis').textContent = fmtN.format(tot.ok + tot.ressalva);
    $('#k-inutil').textContent = fmtN.format(tot.fora);

    var tk = $('#takeaway'); tk.replaceChildren();
    if (!abertos.length) { tk.append('Nenhum caso aberto. Tudo resolvido!'); }
    else {
      var b = function (n) { return el('b', '', fmtN.format(n)); };
      tk.append(b(abertos.length), ' casos para resolver: ', b(g.sistema), ' pedem só uma mudança de local no sistema, ', b(g.etiqueta), ' pedem para olhar a etiqueta e ', b(g.procurar), ' são aparelhos do sistema que ainda não apareceram em nenhum Word.');
    }
    ['sistema', 'etiqueta', 'procurar'].forEach(function (k, i) {
      $('#g-' + k).textContent = fmtN.format(g[k]);
      var bar = $('#b-' + k); bar.style.width = (abertos.length ? 100 * g[k] / abertos.length : 0) + '%'; bar.style.background = COR[i];
    });
    document.querySelectorAll('.grp').forEach(function (b) { b.setAttribute('aria-pressed', String(st.g === b.dataset.g)); });

    var rows1 = D.carts.map(function (c) { return { rot: rotulo(c.nome), confere: c.confere, local: c.local, letra: c.letra, nao: c.nao }; });
    var p1 = [{ k: 'confere', l: 'Confere', c: COR[0] }, { k: 'local', l: 'Local diferente', c: COR[1] }, { k: 'letra', l: 'Letra final diferente', c: COR[2] }, { k: 'nao', l: 'Não encontrado', c: COR[3] }];
    legend('#leg1', p1); stacked('#chart1', rows1, p1);
    var rows2 = D.carts.map(function (c) { var u = U.porCart[c.nome] || { ok: 0, ressalva: 0, fora: 0 }; return { rot: rotulo(c.nome), ok: u.ok, ressalva: u.ressalva, fora: u.fora }; });
    var p2 = [{ k: 'ok', l: 'Usável', c: COR[0] }, { k: 'ressalva', l: 'Usável, com defeito sério', c: COR[1] }, { k: 'fora', l: 'Fora de uso', c: COR[3] }];
    legend('#leg2', p2); stacked('#chart2', rows2, p2);

    fillSelect('#f-cart', uniq(D.cases.map(function (c) { return c.cart; })), 'Todos os carrinhos', st.cart);
    fillSelect('#a-cart', uniq(D.att.map(function (c) { return c.cart; })), 'Todos os carrinhos', st.acart);
    drawCasos(); drawAtencao();
  }

  function uniq(a) { return Array.from(new Set(a)); }
  function fillSelect(sel, vals, all, cur) {
    var s = $(sel); s.replaceChildren(new Option(all, ''));
    vals.forEach(function (v) { s.append(new Option(v, v)); });
    s.value = cur;
  }

  function drawCasos() {
    var q = st.q.trim().toLowerCase();
    var rows = D.cases.filter(function (c) {
      return (!st.g || grupo(c.tipo) === st.g) && (!st.cart || c.cart === st.cart) &&
        (st.s === 'all' || (st.s === 'done') === feito(c)) &&
        (!q || [c.snWord, c.snSis, c.local, c.tipo, c.cart].join(' ').toLowerCase().indexOf(q) >= 0);
    });
    var tb = $('#casos-table tbody'); tb.replaceChildren();
    rows.forEach(function (c) {
      var tr = el('tr', (feito(c) ? 'done ' : '') + (st.sel === c.key ? 'sel' : ''));
      tr.dataset.key = c.key;
      var td = el('td', 'ck'); var cb = el('input'); cb.type = 'checkbox'; cb.checked = feito(c);
      cb.setAttribute('aria-label', 'Resolvido: caso ' + c.n);
      if (c.ok) { cb.disabled = true; cb.title = 'Marcado como verificado na planilha'; }
      cb.addEventListener('click', function (e) { e.stopPropagation(); });
      cb.addEventListener('change', function () { if (cb.checked) local.add(c.key); else local.delete(c.key); saveLocal(); drawAll(); });
      td.append(cb); tr.append(td);
      [c.n, c.cart, c.item || '—', null, c.snWord, c.snSis, c.local].forEach(function (v, i) {
        var t = el('td');
        if (i === 3) t.append(el('span', 'pill', c.tipo)); else t.textContent = v == null ? '—' : String(v).trim();
        if (i === 4 || i === 5) t.className = 'mono';
        tr.append(t);
      });
      tb.append(tr);
    });
    $('#f-count').textContent = rows.length + ' de ' + D.cases.length + ' casos';
    document.querySelectorAll('[data-s]').forEach(function (b) { b.setAttribute('aria-pressed', String(st.s === b.dataset.s)); });
    showDetail(D.cases.filter(function (c) { return c.key === st.sel; })[0]);
  }

  function showDetail(c) {
    $('#d-empty').hidden = !!c; $('#d-body').hidden = !c;
    if (!c) return;
    $('#d-head').textContent = 'Caso ' + c.n + ' · ' + c.cart + (c.item && c.item !== '—' ? ' · computador ' + c.item : '') + ' · ' + c.tipo;
    $('#d-porque').textContent = c.porque; $('#d-fazer').textContent = c.fazer;
  }

  function drawAtencao() {
    var rows = D.att.filter(function (a) {
      return (!st.acart || a.cart === st.acart) && (!st.at || (st.at === 'pat' ? a.tipo.indexOf('Patrim') === 0 : a.tipo.indexOf('Defeito') === 0));
    });
    var tb = $('#atencao-table tbody'); tb.replaceChildren();
    rows.forEach(function (a) {
      var tr = el('tr');
      [a.id, a.cart, a.item || '—', a.tipo, a.sn, a.inf].forEach(function (v, i) { var t = el('td', i === 4 ? 'mono' : '', String(v)); tr.append(t); });
      tb.append(tr);
    });
    $('#a-count').textContent = rows.length + ' de ' + D.att.length;
    document.querySelectorAll('[data-at]').forEach(function (b) { b.setAttribute('aria-pressed', String(st.at === b.dataset.at)); });
  }

  /* ---------- carregar ---------- */
  function mostrar(msgHtml) {
    var m = $('#msg'); m.replaceChildren(); m.hidden = !msgHtml;
    if (msgHtml) msgHtml.forEach(function (n) { m.append(n); });
  }
  function pasta() {
    var m = el('div');
    m.append(el('b', '', 'Não consegui ler a planilha automaticamente.'),
      el('p', '', 'Navegadores bloqueiam a leitura de arquivos quando a página é aberta direto da pasta. Você pode escolher o arquivo no botão "Escolher planilha", ou abrir o painel por um servidor local (veja o LEIAME.txt).'));
    return m;
  }

  function usar(buffer, origem) {
    try {
      var wb = XLSX.read(buffer, { type: 'array' });
      D = parse(wb);
      mostrar(null); $('#main').hidden = false;
      $('#src-info').textContent = origem;
      drawAll();
    } catch (e) {
      $('#main').hidden = true;
      var m = el('div'); m.append(el('b', '', 'Não consegui usar esta planilha.'), el('p', '', e.message));
      mostrar([m]);
      $('#src-info').textContent = 'Planilha não carregada';
    }
  }

  function buscar() {
    $('#src-info').textContent = 'Lendo ' + ARQUIVO + '…';
    fetch(ARQUIVO + '?t=' + Date.now()).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.arrayBuffer();
    }).then(function (buf) {
      usar(buf, 'Planilha: ' + ARQUIVO + ' (lida da pasta)');
    }).catch(function () {
      $('#src-info').textContent = 'Planilha não carregada';
      if (!D) { $('#main').hidden = true; mostrar([pasta()]); }
    });
  }

  $('#btn-reload').addEventListener('click', buscar);
  $('#file-in').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0]; if (!f) return;
    var rd = new FileReader();
    rd.onload = function () { usar(rd.result, 'Planilha: ' + f.name + ' (escolhida por você)'); };
    rd.readAsArrayBuffer(f);
    e.target.value = '';
  });

  document.querySelectorAll('.grp').forEach(function (b) {
    b.addEventListener('click', function () { st.g = st.g === b.dataset.g ? null : b.dataset.g; drawAll(); });
  });
  document.querySelectorAll('[data-s]').forEach(function (b) { b.addEventListener('click', function () { st.s = b.dataset.s; drawCasos(); }); });
  $('#f-cart').addEventListener('change', function (e) { st.cart = e.target.value; drawCasos(); });
  $('#f-q').addEventListener('input', function (e) { st.q = e.target.value; drawCasos(); });
  $('#casos-table').addEventListener('click', function (e) {
    var tr = e.target.closest('tr[data-key]'); if (!tr) return;
    st.sel = tr.dataset.key; drawCasos();
  });
  $('#a-cart').addEventListener('change', function (e) { st.acart = e.target.value; drawAtencao(); });
  document.querySelectorAll('[data-at]').forEach(function (b) {
    b.addEventListener('click', function () { st.at = st.at === b.dataset.at ? '' : b.dataset.at; drawAtencao(); });
  });
  var rz; window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (D) drawAll(); }, 150); });

  buscar();
})();
