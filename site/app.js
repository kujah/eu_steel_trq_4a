const numberFmt = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const EXPORTER_ORDER = [
  "Other countries",
  "Korea",
  "Viet Nam",
  "Taiwan",
  "Turkiye",
  "India",
  "FTA Quota - CSQ",
  "United Kingdom",
  "Japan",
  "Egypt",
  "South Africa",
  "FTA Quota - Other countries",
];

let baseItems = [];

function formatQuantity(quantity) {
  if (!quantity || quantity.value == null) return "-";
  return numberFmt.format(quantity.value);
}

function formatPercent(value) {
  return value == null ? "-" : `${value.toFixed(2)}%`;
}

function parseSearchTarget(item) {
  return [
    item.order_number,
    item.exporter,
    item.product_group,
    item.origins_summary,
    item.origin,
    ...(item.associated_taric_codes || []),
  ]
    .join(" ")
    .toLowerCase();
}

function sortItems(items) {
  const rank = new Map(EXPORTER_ORDER.map((name, index) => [name, index]));
  return [...items].sort(
    (a, b) => (rank.get(a.exporter) ?? 999) - (rank.get(b.exporter) ?? 999),
  );
}

function renderSummary(items) {
  const totalInitialKg = items.reduce((sum, item) => sum + (item.initial_amount.value || 0), 0);
  const totalUsedKg = items.reduce((sum, item) => sum + (item.used_amount.value || 0), 0);
  const totalInitialMt = items.reduce((sum, item) => sum + (item.initial_amount_mt.value || 0), 0);
  const totalUsedMt = items.reduce((sum, item) => sum + (item.used_amount_mt.value || 0), 0);
  const totalBalanceMt = items.reduce((sum, item) => sum + (item.balance_mt.value || 0), 0);
  const avgUtil = totalInitialKg ? (totalUsedKg / totalInitialKg) * 100 : 0;

  const blocks = [
    { label: "Tracked Quotas", value: `${items.length}` },
    { label: "Total Allocation (MT)", value: numberFmt.format(totalInitialMt) },
    { label: "Total Used (MT)", value: numberFmt.format(totalUsedMt) },
    { label: "Total Balance (MT)", value: numberFmt.format(totalBalanceMt) },
    { label: "Weighted Utilization", value: `${avgUtil.toFixed(2)}%` },
  ];

  document.getElementById("summary").innerHTML = blocks
    .map(
      (block) => `
        <article class="summary-card">
          <span>${block.label}</span>
          <strong>${block.value}</strong>
        </article>
      `,
    )
    .join("");
}

function renderCards(items) {
  const cards = document.getElementById("cards");
  const template = document.getElementById("cardTemplate");
  cards.innerHTML = "";

  items.forEach((item) => {
    const node = template.content.firstElementChild.cloneNode(true);
    node.querySelector(".card__order").textContent = item.order_number;
    node.querySelector(".card__origin").textContent = item.exporter || "-";
    node.querySelector(".card__link").href = item.detail_url;
    node.querySelector(".progress__bar").style.width = `${Math.min(item.utilization_pct || 0, 100)}%`;

    const stats = [
      ["Initial", formatQuantity(item.initial_amount_mt)],
      ["Used", formatQuantity(item.used_amount_mt)],
      ["Balance", formatQuantity(item.balance_mt)],
      ["Utilization", formatPercent(item.utilization_pct)],
      ["Critical", item.critical || "-"],
      ["Awaiting", formatQuantity(item.awaiting_allocation_mt)],
    ];

    node.querySelector(".stats").innerHTML = stats
      .map(([label, value]) => `<div class="stat"><span>${label}</span><strong>${value}</strong></div>`)
      .join("");

    cards.appendChild(node);
  });
}

function renderTable(items) {
  document.getElementById("tableBody").innerHTML = items
    .map(
      (item) => `
        <tr>
          <td>${item.order_number}</td>
          <td>${item.exporter || "-"}</td>
          <td>${formatQuantity(item.initial_amount_mt)}</td>
          <td>${formatQuantity(item.used_amount_mt)}</td>
          <td>${formatQuantity(item.balance_mt)}</td>
          <td>${formatPercent(item.utilization_pct)}</td>
          <td>${item.critical || "-"}</td>
          <td>${item.last_allocation_date || "-"}</td>
          <td>${formatQuantity(item.awaiting_allocation_mt)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderMeta(data) {
  document.getElementById("generatedAt").textContent = new Date(data.generated_at_utc).toLocaleString();
  document.getElementById("quotaMeta").textContent = `${data.quota_section} | ${data.product_group}`;
  document.getElementById("reportPeriod").textContent = data.report_period || "-";

  const sourceUpdate = data.items
    .map((item) => item.source_last_taric_update)
    .filter(Boolean)
    .sort()
    .at(-1);

  document.getElementById("sourceUpdate").textContent = sourceUpdate || "-";
}

function applySearch() {
  const term = document.getElementById("searchInput").value.trim().toLowerCase();
  const items = !term
    ? baseItems
    : baseItems.filter((item) => parseSearchTarget(item).includes(term));
  renderSummary(items);
  renderCards(items);
  renderTable(items);
}

async function loadData() {
  const response = await fetch("./data/orders.json", { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Request failed: ${response.status}`);
  }

  const data = await response.json();
  baseItems = sortItems(data.items);
  renderMeta(data);
  applySearch();
}

async function main() {
  document.getElementById("searchInput").addEventListener("input", applySearch);
  await loadData();
}

main().catch((error) => {
  console.error(error);
});
