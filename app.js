const months = [
  "114/01","114/02","114/03","114/04","114/05","114/06",
  "114/07","114/08","114/09","114/10","114/11","114/12"
];

const el = (id) => document.getElementById(id);

function n(v){ // to number
  const x = Number(v);
  return Number.isFinite(x) ? x : 0;
}

function fmt(v){
  const d = Number(el("decimals").value || 1);
  if (!Number.isFinite(v)) return "—";
  return v.toFixed(d);
}

function ratios(){
  const b = n(el("ratioB").value) / 100;
  const c = n(el("ratioC").value) / 100;
  return { b, c };
}

function buildRows(){
  const tbody = el("rows");
  tbody.innerHTML = "";

  months.forEach((m, idx) => {
    const tr = document.createElement("tr");

    const tdM = document.createElement("td");
    tdM.textContent = m;
    tr.appendChild(tdM);

    const makeInput = (key) => {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.type = "number";
      inp.step = "1";
      inp.min = "0";
      inp.dataset.key = key;
      inp.dataset.idx = String(idx);
      inp.placeholder = "0";
      inp.addEventListener("input", recalc);
      td.appendChild(inp);
      return td;
    };

    tr.appendChild(makeInput("D"));
    tr.appendChild(makeInput("E"));
    tr.appendChild(makeInput("F"));
    tr.appendChild(makeInput("H"));

    // G, I, J (computed)
    ["G","I","J"].forEach(k => {
      const td = document.createElement("td");
      td.id = `${k}-${idx}`;
      td.textContent = "—";
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

function getMonthData(){
  const inputs = Array.from(document.querySelectorAll("#rows input"));
  const data = months.map(()=>({D:0,E:0,F:0,H:0}));

  for (const inp of inputs){
    const idx = Number(inp.dataset.idx);
    const key = inp.dataset.key;
    data[idx][key] = n(inp.value);
  }
  return data;
}

function avg(arr){
  const vals = arr.filter(x => Number.isFinite(x));
  if (vals.length === 0) return NaN;
  return vals.reduce((a,b)=>a+b,0) / vals.length;
}

function recalc(){
  const { b, c } = ratios();
  const data = getMonthData();

  // per-month G/I/J
  data.forEach((r, idx) => {
    const G = r.D + r.E + r.F + r.H;            // 投保總人數
    const I = b * (r.D + r.E);                  // 原額最多可用
    const J = (b + c) * (r.D + r.E + r.F);      // 增額最多可用

    el(`G-${idx}`).textContent = fmt(G);
    el(`I-${idx}`).textContent = fmt(I);
    el(`J-${idx}`).textContent = fmt(J);
  });

  // quarterly
  const qTbody = el("qRows");
  qTbody.innerHTML = "";
  const quarters = [
    { label:"1~3月", idxs:[0,1,2] },
    { label:"4~6月", idxs:[3,4,5] },
    { label:"7~9月", idxs:[6,7,8] },
    { label:"10~12月", idxs:[9,10,11] },
  ];

  quarters.forEach(q => {
    const de = q.idxs.map(i => data[i].D + data[i].E);
    const def = q.idxs.map(i => data[i].D + data[i].E + data[i].F);

    const avgDE = avg(de);
    const avgDEF = avg(def);

    const K = avgDE * b;
    const L = avgDEF * (b + c);

    const tr = document.createElement("tr");
    const cells = [
      q.label,
      fmt(avgDE),
      fmt(K),
      fmt(avgDEF),
      fmt(L)
    ];

    cells.forEach((t, i) => {
      const td = document.createElement("td");
      td.textContent = t;
      if (i===0) td.style.textAlign = "left";
      tr.appendChild(td);
    });
    qTbody.appendChild(tr);
  });

  // annual
  const allDE = data.map(r => r.D + r.E);
  const allDEF = data.map(r => r.D + r.E + r.F);
  const avgDE12 = avg(allDE);
  const avgDEF12 = avg(allDEF);

  el("avgDE").textContent = fmt(avgDE12);
  el("avgDEF").textContent = fmt(avgDEF12);
  el("annualM").textContent = fmt(avgDE12 * b);
  el("annualN").textContent = fmt(avgDEF12 * (b + c));
}

function setMonth(idx, D, E, F, H){
  const rowInputs = Array.from(document.querySelectorAll(`#rows input[data-idx="${idx}"]`));
  const map = {D,E,F,H};
  rowInputs.forEach(inp => {
    inp.value = map[inp.dataset.key] ?? 0;
  });
}

function demo(){
  // 只示範 114/10~114/12（其餘留 0）
  el("ratioB").value = 10;
  el("ratioC").value = 15;

  // 你範例： (D+E) = 191,191,192；(D+E+F)=211,212,212
  // 這裡用 E 代表原額移工、F 代表增額移工，示範拆分：
  // 假設 114/10: D=167, E=24 => 191；再加 F=20 => 211
  // 114/11: D=167, E=24 => 191；再加 F=21 => 212
  // 114/12: D=168, E=24 => 192；再加 F=20 => 212
  setMonth(9, 167, 24, 20, 0);   // 114/10
  setMonth(10,167, 24, 21, 0);   // 114/11
  setMonth(11,168, 24, 20, 0);   // 114/12

  recalc();
}

function clearAll(){
  Array.from(document.querySelectorAll("#rows input")).forEach(i => i.value = "");
  el("company").value = "";
  el("ratioB").value = "";
  el("ratioC").value = "";
  recalc();
}

document.addEventListener("DOMContentLoaded", () => {
  buildRows();
  ["ratioB","ratioC","decimals"].forEach(id => el(id).addEventListener("input", recalc));
  el("fillDemo").addEventListener("click", demo);
  el("clearAll").addEventListener("click", clearAll);
  recalc();
});
function exportToExcel(){
  if (typeof XLSX === "undefined"){
    alert("Excel 匯出模組尚未載入，請確認已加入 xlsx.full.min.js");
    return;
  }

  const company = (el("company")?.value || "").trim() || "未命名";
  const B = n(el("ratioB").value); // %
  const C = n(el("ratioC").value); // %

  const data = getMonthData();

  // 1) 月資料工作表：含輸入 + 計算結果
  const ws1 = [];
  ws1.push([
    "月份","本籍(D)","移工(原額)(E)","移工(增額)(F)","外國技術人力(H)",
    "投保總人數(G)","單月I 原額最多可用","單月J 增額最多可用"
  ]);

  data.forEach((r, idx) => {
    const G = Number(el(`G-${idx}`)?.textContent || 0);
    const I = Number(el(`I-${idx}`)?.textContent || 0);
    const J = Number(el(`J-${idx}`)?.textContent || 0);

    ws1.push([
      months[idx],
      r.D, r.E, r.F, r.H,
      G, I, J
    ]);
  });

  // 2) 每季工作表：由畫面上 qRows 直接讀（避免重算差異）
  const ws2 = [];
  ws2.push(["區間","本籍+原額 三個月平均","原額可用(K)","本籍+原額+增額 三個月平均","增額可用(L)"]);

  const qtrs = Array.from(document.querySelectorAll("#qRows tr"));
  qtrs.forEach(tr => {
    const tds = Array.from(tr.querySelectorAll("td")).map(td => td.textContent.trim());
    // [區間, avgDE, K, avgDEF, L]
    ws2.push(tds);
  });

  // 3) 全年工作表：KPI
  const ws3 = [];
  ws3.push(["項目","數值"]);
  ws3.push(["公司名稱(A)", company]);
  ws3.push(["原額比率(B) (%)", B]);
  ws3.push(["目前使用最大增額比率(C) (%)", C]);
  ws3.push(["本籍+原額 全年平均人數", el("avgDE")?.textContent || "—"]);
  ws3.push(["全年 原額可用(M)", el("annualM")?.textContent || "—"]);
  ws3.push(["本籍+原額+增額 全年平均人數", el("avgDEF")?.textContent || "—"]);
  ws3.push(["全年 增額可用(N)", el("annualN")?.textContent || "—"]);

  // 4) 組成 xlsx
  const wb = XLSX.utils.book_new();

  const sheet1 = XLSX.utils.aoa_to_sheet(ws1);
  const sheet2 = XLSX.utils.aoa_to_sheet(ws2);
  const sheet3 = XLSX.utils.aoa_to_sheet(ws3);

  // 欄寬（讓 Excel 開起來好讀）
  sheet1["!cols"] = [
    { wch: 8 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 16 },
    { wch: 14 }, { wch: 18 }, { wch: 18 }
  ];
  sheet2["!cols"] = [{ wch: 10 }, { wch: 20 }, { wch: 14 }, { wch: 22 }, { wch: 14 }];
  sheet3["!cols"] = [{ wch: 22 }, { wch: 18 }];

  XLSX.utils.book_append_sheet(wb, sheet1, "每月明細");
  XLSX.utils.book_append_sheet(wb, sheet2, "每季平均");
  XLSX.utils.book_append_sheet(wb, sheet3, "全年摘要");

  // 檔名：公司_日期
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth()+1).padStart(2,"0");
  const d = String(now.getDate()).padStart(2,"0");
  const filename = `${company}_移工配額試算_${y}${m}${d}.xlsx`;

  XLSX.writeFile(wb, filename);
}

document.addEventListener("DOMContentLoaded", () => {
  // 你原本已有的監聽保留
  const btn = el("exportXlsx");
  if (btn) btn.addEventListener("click", exportToExcel);
});
