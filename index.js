"use strict";

const { Dialog, Plugin, showMessage } = require("siyuan");
const NAME = "siyuan-excel-database-import";
const VERSION = "1.1.0";
const DB = '[data-type="NodeAttributeView"][data-av-id], .av[data-av-id]';
const TYPES = new Set(["block", "text", "number", "select", "mSelect", "checkbox", "url", "email", "phone"]);

const html = (value) =>
  String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const post = async (path, data) => {
  const response = await fetch(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
  const result = await response.json();
  if (result.code) throw new Error(result.msg || path);
  return result.data;
};

const id = () => {
  const d = new Date();
  const time = [d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()]
    .map((item, index) => (index ? String(item).padStart(2, "0") : item))
    .join("");
  return `${time}-${Math.random().toString(36).slice(2, 9).padEnd(7, "0")}`;
};

const read = (file, method) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("文件读取失败"));
    reader[method](file);
  });

function parseCSV(text) {
  const rows = [];
  let row = [], cell = "", quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (char === '"' && quoted && text[i + 1] === '"') cell += text[i++];
    else if (char === '"') quoted = !quoted;
    else if (char === "," && !quoted) row.push(cell), (cell = "");
    else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && text[i + 1] === "\n") i += 1;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += char;
  }
  if (cell || row.length) row.push(cell), rows.push(row);
  return rows;
}

async function getWorkbook(file) {
  if (/\.csv$/i.test(file.name)) return { sheets: ["CSV"], rows: { CSV: parseCSV(await read(file, "readAsText")) } };
  const XLSX = await import(`/plugins/${NAME}/xlsx.mjs?v=${VERSION}`);
  const book = XLSX.read(await read(file, "readAsArrayBuffer"), { type: "array" });
  const rows = {};
  book.SheetNames.forEach((name) => {
    rows[name] = XLSX.utils.sheet_to_json(book.Sheets[name], { header: 1, defval: "", raw: false, blankrows: false });
  });
  return { sheets: book.SheetNames, rows };
}

function table(grid) {
  const width = Math.max(0, ...grid.map((row) => row.length));
  const headers = Array.from({ length: width }, (_, i) => String(grid[0]?.[i] ?? "").trim() || `未命名列 ${i + 1}`);
  const rows = grid.slice(1).map((row) => headers.map((_, i) => row[i] ?? "")).filter((row) => row.some((cell) => String(cell).trim()));
  return { headers, rows };
}

function databases() {
  const result = new Map();
  document.querySelectorAll(DB).forEach((element) => result.set(element.dataset.avId, element));
  return [...result].map(([avID, element]) => ({ avID, element }));
}

function options(columns, selected) {
  return `<option value="">跳过</option><option value="new"${selected === "new" ? " selected" : ""}>新建文本字段</option>${columns
    .filter((column) => TYPES.has(column.type))
    .map((column) => `<option value="${column.id}"${selected === column.id ? " selected" : ""}>${html(column.name)} · ${column.type}</option>`)
    .join("")}`;
}

function value(column, raw) {
  const content = String(raw ?? "").trim();
  if (column.type === "block") return { keyID: column.id, block: { content } };
  if (column.type === "number") return { keyID: column.id, number: { content: Number(content) || 0, isNotEmpty: content !== "" } };
  if (column.type === "checkbox") return { keyID: column.id, checkbox: { checked: /^(1|true|yes|是|√)$/i.test(content) } };
  if (column.type === "select" || column.type === "mSelect") return { keyID: column.id, mSelect: content ? [{ content }] : [] };
  if (column.type === "url") return { keyID: column.id, url: { content } };
  if (column.type === "email") return { keyID: column.id, email: { content } };
  if (column.type === "phone") return { keyID: column.id, phone: { content } };
  return { keyID: column.id, text: { content } };
}

class ExcelImport extends Plugin {
  onload() {
    this.addTopBar({ icon: "iconDownload", title: "Excel 导入数据库", callback: () => this.open() });
    this.observer = new MutationObserver(() => this.decorate());
    this.observer.observe(document.body, { childList: true, subtree: true });
    this.decorate();
  }

  onunload() {
    this.observer?.disconnect();
    document.querySelectorAll(".moon-excel-launcher").forEach((button) => button.remove());
  }

  decorate() {
    databases().forEach(({ avID, element }) => {
      const root = element.closest('[data-type="NodeAttributeView"]') || element;
      if (root.querySelector(":scope > .moon-excel-launcher")) return;
      root.classList.add("moon-excel-host");
      const button = document.createElement("button");
      button.className = "moon-excel-launcher b3-button b3-button--outline";
      button.textContent = "⇩ Excel 导入";
      button.onclick = (event) => (event.stopPropagation(), this.panel(avID));
      root.appendChild(button);
    });
  }

  open() {
    const list = databases();
    if (list.length === 1) return this.panel(list[0].avID);
    showMessage(list.length ? "请点击目标数据库右上角的 Excel 导入" : "请先打开包含数据库的文档");
  }

  async panel(avID) {
    const database = await post("/api/av/renderAttributeView", { id: avID, pageSize: -1 });
    const dialog = new Dialog({
      title: "Excel 导入数据库",
      width: "900px",
      content: `<div class="b3-dialog__content moon-excel">
        <strong>${html(database.name || "未命名数据库")}</strong>
        <input type="file" accept=".csv,.xlsx,.xls" data-file>
        <select class="b3-select fn__block" data-sheet disabled><option>请先选择文件</option></select>
        <span data-status></span><div data-mapping></div><div data-preview></div>
        <button class="b3-button b3-button--text fn__none" data-import>开始导入</button>
      </div>`,
    });
    const root = dialog.element.querySelector(".moon-excel");
    const $ = (selector) => root.querySelector(selector);
    let book, data;
    const status = (text, error = false) => (($("[data-status]").textContent = text), $("[data-status]").className = error ? "error" : "");

    const render = () => {
      data = table(book.rows[$("[data-sheet]").value]);
      if (!data.headers.length) throw new Error("工作表为空");
      $("[data-mapping]").innerHTML = data.headers.map((header, index) => {
        const match = database.view.columns.find((column) => column.name === header);
        const selected = match?.id || (index ? "new" : database.view.columns[0]?.id);
        return `<label class="moon-excel-map">${html(header)} → <select data-map>${options(database.view.columns, selected)}</select>
          <input data-name value="${html(header)}"${selected === "new" ? "" : " hidden"}></label>`;
      }).join("");
      root.querySelectorAll("[data-map]").forEach((select) => (select.onchange = () => (select.nextElementSibling.hidden = select.value !== "new")));
      $("[data-preview]").textContent = `${data.rows.length} 行数据待导入`;
      $("[data-import]").classList.remove("fn__none");
    };

    $("[data-file]").onchange = async (event) => {
      try {
        status("正在读取...");
        book = await getWorkbook(event.target.files[0]);
        $("[data-sheet]").innerHTML = book.sheets.map((name) => `<option>${html(name)}</option>`).join("");
        $("[data-sheet]").disabled = false;
        render();
        status("读取成功");
      } catch (error) {
        status(`读取失败：${error.message}`, true);
      }
    };
    $("[data-sheet]").onchange = render;
    $("[data-import]").onclick = async () => {
      try {
        const columns = [...database.view.columns];
        const maps = [];
        for (const [index, label] of [...root.querySelectorAll(".moon-excel-map")].entries()) {
          const target = label.querySelector("[data-map]").value;
          if (!target) continue;
          if (target === "new") {
            const name = label.querySelector("[data-name]").value.trim();
            const column = { id: id(), name, type: "text" };
            await post("/api/av/addAttributeViewKey", { avID, keyID: column.id, keyName: name, keyType: "text", keyIcon: "", previousKeyID: columns.at(-1)?.id || "" });
            columns.push(column); maps.push({ index, column });
          } else maps.push({ index, column: columns.find((column) => column.id === target) });
        }
        if (!maps.some((map) => map.column.id === columns[0].id)) throw new Error("必须映射主键");
        const rows = data.rows
          .map((row) => maps.map((map) => value(map.column, row[map.index])))
          .filter((row) => row.find((cell) => cell.keyID === columns[0].id)?.block?.content);
        for (let i = 0; i < rows.length; i += 100) {
          status(`正在导入 ${Math.min(i + 100, rows.length)} / ${rows.length}`);
          await post("/api/av/appendAttributeViewDetachedBlocksWithValues", { avID, blocksValues: rows.slice(i, i + 100) });
        }
        status(`完成：已导入 ${rows.length} 行`);
        showMessage(`已导入 ${rows.length} 行`);
      } catch (error) {
        status(`导入失败：${error.message}`, true);
      }
    };
  }
}

module.exports = ExcelImport;
