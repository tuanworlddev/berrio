const state = {
  shops: [],
  selectedShopId: "",
  page: 1,
  limit: 10,
  totalPages: 1,
};

const els = {
  shopForm: document.querySelector("#shopForm"),
  shopId: document.querySelector("#shopId"),
  shopName: document.querySelector("#shopName"),
  shopMarketplace: document.querySelector("#shopMarketplace"),
  shopApiKey: document.querySelector("#shopApiKey"),
  shopDescription: document.querySelector("#shopDescription"),
  shopIsActive: document.querySelector("#shopIsActive"),
  shopList: document.querySelector("#shopList"),
  shopCount: document.querySelector("#shopCount"),
  shopPageLabel: document.querySelector("#shopPageLabel"),
  prevShopPage: document.querySelector("#prevShopPage"),
  nextShopPage: document.querySelector("#nextShopPage"),
  clearShopButton: document.querySelector("#clearShopButton"),
  refreshShopsButton: document.querySelector("#refreshShopsButton"),
  reportForm: document.querySelector("#reportForm"),
  reportShopSelect: document.querySelector("#reportShopSelect"),
  selectedShopLabel: document.querySelector("#selectedShopLabel"),
  dateFrom: document.querySelector("#dateFrom"),
  dateTo: document.querySelector("#dateTo"),
  tax: document.querySelector("#tax"),
  discount: document.querySelector("#discount"),
  downloadReportButton: document.querySelector("#downloadReportButton"),
  loadOrdersButton: document.querySelector("#loadOrdersButton"),
  totalOrders: document.querySelector("#totalOrders"),
  totalPrevOrders: document.querySelector("#totalPrevOrders"),
  ordersTableBody: document.querySelector("#ordersTableBody"),
  toast: document.querySelector("#toast"),
};

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function daysAgoISO(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return date.toISOString().slice(0, 10);
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("is-visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    els.toast.classList.remove("is-visible");
  }, 2800);
}

async function requestJSON(url, options = {}) {
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });

  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error || `Request failed with ${response.status}`);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

function getSelectedShop() {
  return state.shops.find((shop) => shop.id === state.selectedShopId) || null;
}

function renderShops(data) {
  state.shops = data.items || [];
  state.totalPages = Math.max(data.totalPages || 1, 1);
  els.shopCount.textContent = String(data.total || 0);
  els.shopPageLabel.textContent = `${state.page} / ${state.totalPages}`;
  els.prevShopPage.disabled = state.page <= 1;
  els.nextShopPage.disabled = state.page >= state.totalPages;

  if (!state.selectedShopId && state.shops.length > 0) {
    state.selectedShopId = state.shops[0].id;
  }

  els.shopList.innerHTML = "";
  if (state.shops.length === 0) {
    const empty = document.createElement("div");
    empty.className = "shop-item";
    empty.textContent = "Chưa có shop";
    els.shopList.appendChild(empty);
  }

  for (const shop of state.shops) {
    const item = document.createElement("div");
    item.className = `shop-item${shop.id === state.selectedShopId ? " is-selected" : ""}`;

    const info = document.createElement("button");
    info.className = "shop-select-button";
    info.type = "button";
    info.innerHTML = `<div class="shop-name">${escapeHTML(shop.name)}</div><div class="shop-meta">${escapeHTML(shop.marketplace || "Chưa có sàn")}</div>`;
    info.addEventListener("click", () => selectShop(shop.id));

    const actions = document.createElement("div");
    actions.className = "shop-actions";

    const edit = document.createElement("button");
    edit.className = "btn btn-outline-secondary";
    edit.type = "button";
    edit.textContent = "Sửa";
    edit.addEventListener("click", () => fillShopForm(shop));

    const del = document.createElement("button");
    del.className = "btn btn-outline-danger";
    del.type = "button";
    del.textContent = "Xóa";
    del.addEventListener("click", () => deleteShop(shop));

    actions.append(edit, del);
    item.append(info, actions);
    els.shopList.appendChild(item);
  }

  renderShopSelect();
  renderSelectedShop();
}

function renderShopSelect() {
  els.reportShopSelect.innerHTML = "";
  if (state.shops.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "Chưa có shop";
    els.reportShopSelect.appendChild(option);
    return;
  }

  for (const shop of state.shops) {
    const option = document.createElement("option");
    option.value = shop.id;
    option.textContent = shop.name;
    option.selected = shop.id === state.selectedShopId;
    els.reportShopSelect.appendChild(option);
  }
}

function renderSelectedShop() {
  const shop = getSelectedShop();
  els.selectedShopLabel.textContent = shop ? shop.name : "Chưa chọn shop";
}

async function loadShops() {
  const data = await requestJSON(`/api/v1/shops?page=${state.page}&limit=${state.limit}`);
  renderShops(data);
}

function selectShop(id) {
  state.selectedShopId = id;
  renderShops({ items: state.shops, total: Number(els.shopCount.textContent), totalPages: state.totalPages });
}

function fillShopForm(shop) {
  selectShop(shop.id);
  els.shopId.value = shop.id;
  els.shopName.value = shop.name || "";
  els.shopMarketplace.value = shop.marketplace || "";
  els.shopApiKey.value = shop.apiKey || "";
  els.shopDescription.value = shop.description || "";
  els.shopIsActive.checked = shop.isActive !== false;
}

function clearShopForm() {
  els.shopForm.reset();
  els.shopId.value = "";
  els.shopIsActive.checked = true;
}

function shopPayload() {
  return {
    name: els.shopName.value.trim(),
    marketplace: els.shopMarketplace.value.trim(),
    apiKey: els.shopApiKey.value.trim(),
    description: els.shopDescription.value.trim(),
    isActive: els.shopIsActive.checked,
  };
}

async function saveShop(event) {
  event.preventDefault();
  const id = els.shopId.value;
  const method = id ? "PATCH" : "POST";
  const url = id ? `/api/v1/shops/${id}` : "/api/v1/shops";
  await requestJSON(url, { method, body: JSON.stringify(shopPayload()) });
  clearShopForm();
  await loadShops();
  showToast(id ? "Đã cập nhật shop" : "Đã tạo shop");
}

async function deleteShop(shop) {
  const ok = window.confirm(`Xóa shop "${shop.name}"?`);
  if (!ok) {
    return;
  }

  await requestJSON(`/api/v1/shops/${shop.id}`, { method: "DELETE" });
  if (state.selectedShopId === shop.id) {
    state.selectedShopId = "";
  }
  await loadShops();
  showToast("Đã xóa shop");
}

async function downloadReport(event) {
  event.preventDefault();
  const shop = getSelectedShop();
  if (!shop || !shop.apiKey) {
    showToast("Hãy chọn shop có API key");
    return;
  }

  els.downloadReportButton.disabled = true;
  els.downloadReportButton.textContent = "Đang tạo báo cáo";

  try {
    const response = await fetch("/api/v1/reports", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: shop.apiKey,
        dateFrom: els.dateFrom.value,
        dateTo: els.dateTo.value,
        tax: Number(els.tax.value || 0.06),
        discount: Number(els.discount.value || 3.5),
      }),
    });

    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.error || "Không thể tạo báo cáo");
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `reports-${els.dateFrom.value}-${els.dateTo.value}.zip`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showToast("Đã tải báo cáo");
  } catch (error) {
    showToast(error.message);
  } finally {
    els.downloadReportButton.disabled = false;
    els.downloadReportButton.textContent = "Tải báo cáo ZIP";
  }
}

async function loadOrders() {
  const shop = getSelectedShop();
  if (!shop || !shop.apiKey) {
    showToast("Hãy chọn shop có API key");
    return;
  }

  els.loadOrdersButton.disabled = true;
  try {
    const data = await requestJSON("/api/v1/orders", {
      method: "POST",
      body: JSON.stringify({
        apiKey: shop.apiKey,
        dateFrom: els.dateFrom.value,
        dateTo: els.dateTo.value,
      }),
    });
    renderOrders(data);
  } catch (error) {
    showToast(error.message);
  } finally {
    els.loadOrdersButton.disabled = false;
  }
}

function renderOrders(data) {
  els.totalOrders.textContent = String(data.totalOrders || 0);
  els.totalPrevOrders.textContent = String(data.totalPrevOrders || 0);
  els.ordersTableBody.innerHTML = "";

  const rows = data.chartData || [];
  if (rows.length === 0) {
    els.ordersTableBody.innerHTML = '<tr><td class="empty-row" colspan="5">Không có dữ liệu</td></tr>';
    return;
  }

  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.nmID}</td>
      <td>${escapeHTML(row.vendorCode || "")}</td>
      <td>${row.ordersCount || 0}</td>
      <td>${row.prevOrdersCount || 0}</td>
      <td>${row.ordersSumRub || 0}</td>
    `;
    els.ordersTableBody.appendChild(tr);
  }
}

function escapeHTML(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

els.shopForm.addEventListener("submit", (event) => {
  saveShop(event).catch((error) => showToast(error.message));
});
els.clearShopButton.addEventListener("click", clearShopForm);
els.refreshShopsButton.addEventListener("click", () => loadShops().catch((error) => showToast(error.message)));
els.prevShopPage.addEventListener("click", () => {
  state.page -= 1;
  loadShops().catch((error) => showToast(error.message));
});
els.nextShopPage.addEventListener("click", () => {
  state.page += 1;
  loadShops().catch((error) => showToast(error.message));
});
els.reportShopSelect.addEventListener("change", (event) => selectShop(event.target.value));
els.reportForm.addEventListener("submit", downloadReport);
els.loadOrdersButton.addEventListener("click", loadOrders);

els.dateFrom.value = daysAgoISO(7);
els.dateTo.value = todayISO();
loadShops().catch((error) => showToast(error.message));
