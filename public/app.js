const iconEdit = `
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M12 20h9" />
    <path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z" />
  </svg>
`;

const iconTrash = `
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <path d="M3 6h18" />
    <path d="M8 6V4h8v2" />
    <path d="m19 6-1 14H6L5 6" />
    <path d="M10 11v6" />
    <path d="M14 11v6" />
  </svg>
`;

const state = {
  shops: [],
  tokenStatuses: new Map(),
  selectedShopId: "",
  page: 1,
  limit: 20,
  total: 0,
  totalPages: 1,
  isLoadingShops: false,
};

const els = {
  addShopButton: document.querySelector("#addShopButton"),
  shopForm: document.querySelector("#shopForm"),
  shopModal: document.querySelector("#shopModal"),
  shopModalTitle: document.querySelector("#shopModalTitle"),
  shopId: document.querySelector("#shopId"),
  shopName: document.querySelector("#shopName"),
  shopMarketplace: document.querySelector("#shopMarketplace"),
  shopApiKey: document.querySelector("#shopApiKey"),
  shopDescription: document.querySelector("#shopDescription"),
  shopIsActive: document.querySelector("#shopIsActive"),
  saveShopButton: document.querySelector("#saveShopButton"),
  shopList: document.querySelector("#shopList"),
  shopCount: document.querySelector("#shopCount"),
  shopLoadedLabel: document.querySelector("#shopLoadedLabel"),
  shopLoading: document.querySelector("#shopLoading"),
  reportForm: document.querySelector("#reportForm"),
  reportShopSelect: document.querySelector("#reportShopSelect"),
  apiStatus: document.querySelector("#apiStatus"),
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

const shopModal = bootstrap.Modal.getOrCreateInstance(els.shopModal);

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

function getLocalTokenStatus(shop) {
  if (!shop?.apiKey) {
    return { valid: false, status: "missing", message: "Thiếu API key" };
  }

  const parts = shop.apiKey.split(".");
  if (parts.length < 2) {
    return { valid: false, status: "invalid", message: "Token sai định dạng" };
  }

  try {
    let payloadSegment = parts[1].replaceAll("-", "+").replaceAll("_", "/");
    payloadSegment += "=".repeat((4 - (payloadSegment.length % 4)) % 4);
    const payload = JSON.parse(atob(payloadSegment));
    if (payload.exp && payload.exp * 1000 < Date.now()) {
      return { valid: false, status: "expired", message: "Token đã hết hạn" };
    }
  } catch {
    return { valid: false, status: "invalid", message: "Token sai định dạng" };
  }

  return { valid: true, status: "unchecked", message: "Chưa kiểm tra WB" };
}

function getTokenStatus(shop) {
  return state.tokenStatuses.get(shop.id) || getLocalTokenStatus(shop);
}

function hasMoreShops() {
  return state.page < state.totalPages;
}

async function loadShops({ reset = false } = {}) {
  if (state.isLoadingShops) {
    return;
  }

  state.isLoadingShops = true;
  els.shopLoading.classList.add("is-visible");

  if (reset) {
    state.page = 1;
    state.shops = [];
    state.selectedShopId = "";
    renderShops();
  }

  try {
    const data = await requestJSON(`/api/v1/shops?page=${state.page}&limit=${state.limit}`);
    state.total = data.total || 0;
    state.totalPages = Math.max(data.totalPages || 1, 1);

    const incoming = data.items || [];
    const seen = new Set(state.shops.map((shop) => shop.id));
    state.shops = [...state.shops, ...incoming.filter((shop) => !seen.has(shop.id))];

    if (!state.selectedShopId && state.shops.length > 0) {
      state.selectedShopId = state.shops[0].id;
    }

    renderShops();
  } finally {
    state.isLoadingShops = false;
    els.shopLoading.classList.remove("is-visible");
  }
}

async function loadNextShopPage() {
  if (!hasMoreShops() || state.isLoadingShops) {
    return;
  }

  state.page += 1;
  await loadShops();
}

function renderShops() {
  els.shopCount.textContent = `(${state.total})`;
  els.shopLoadedLabel.textContent = `${state.shops.length} đã tải`;
  els.shopList.innerHTML = "";

  if (state.shops.length === 0) {
    const empty = document.createElement("div");
    empty.className = "shop-item";
    empty.textContent = "Chưa có shop";
    els.shopList.appendChild(empty);
  }

  for (const shop of state.shops) {
    const item = document.createElement("div");
    const tokenStatus = getTokenStatus(shop);
    item.className = `shop-item${shop.id === state.selectedShopId ? " is-selected" : ""}${tokenStatus.valid ? "" : " is-token-invalid"}`;

    const info = document.createElement("button");
    info.className = "shop-select-button";
    info.type = "button";
    info.innerHTML = `
      <div class="shop-name">${escapeHTML(shop.name)}</div>
      <div class="shop-meta">${escapeHTML(tokenStatus.message || (shop.marketplace || "Chưa có sàn"))}</div>
    `;
    info.addEventListener("click", () => selectShop(shop.id));

    const actions = document.createElement("div");
    actions.className = "shop-actions";

    const edit = document.createElement("button");
    edit.className = "btn btn-outline-primary";
    edit.type = "button";
    edit.title = "Sửa shop";
    edit.setAttribute("aria-label", "Sửa shop");
    edit.innerHTML = iconEdit;
    edit.addEventListener("click", () => openEditShopModal(shop));

    const del = document.createElement("button");
    del.className = "btn btn-outline-danger";
    del.type = "button";
    del.title = "Xóa shop";
    del.setAttribute("aria-label", "Xóa shop");
    del.innerHTML = iconTrash;
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
  if (!shop) {
    els.apiStatus.textContent = "Chưa chọn shop";
    els.apiStatus.className = "api-status";
    return;
  }

  const tokenStatus = getTokenStatus(shop);
  const isReady = tokenStatus.valid && tokenStatus.status === "ok";
  els.apiStatus.textContent = isReady ? "Token hoạt động" : tokenStatus.message;
  els.apiStatus.className = `api-status${isReady ? " is-ready" : ""}${tokenStatus.valid ? "" : " is-invalid"}`;
}

function selectShop(id) {
  state.selectedShopId = id;
  renderShops();
  checkSelectedShopToken().catch((error) => showToast(error.message));
}

function clearShopForm() {
  els.shopForm.reset();
  els.shopId.value = "";
  els.shopIsActive.checked = true;
  els.shopModalTitle.textContent = "Thêm shop";
  els.saveShopButton.textContent = "Thêm shop";
}

function openAddShopModal() {
  clearShopForm();
  shopModal.show();
}

function openEditShopModal(shop) {
  selectShop(shop.id);
  els.shopId.value = shop.id;
  els.shopName.value = shop.name || "";
  els.shopMarketplace.value = shop.marketplace || "";
  els.shopApiKey.value = shop.apiKey || "";
  els.shopDescription.value = shop.description || "";
  els.shopIsActive.checked = shop.isActive !== false;
  els.shopModalTitle.textContent = "Cập nhật shop";
  els.saveShopButton.textContent = "Lưu thay đổi";
  shopModal.show();
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
  const saved = await requestJSON(url, { method, body: JSON.stringify(shopPayload()) });

  shopModal.hide();
  if (id) {
    state.shops = state.shops.map((shop) => (shop.id === id ? saved : shop));
    state.selectedShopId = id;
  } else {
    state.shops = [saved, ...state.shops];
    state.total += 1;
    state.selectedShopId = saved.id;
  }

  renderShops();
  showToast(id ? "Đã cập nhật shop" : "Đã thêm shop");
  checkSelectedShopToken().catch((error) => showToast(error.message));
}

async function deleteShop(shop) {
  const ok = window.confirm(`Xóa shop "${shop.name}"?`);
  if (!ok) {
    return;
  }

  await requestJSON(`/api/v1/shops/${shop.id}`, { method: "DELETE" });
  state.shops = state.shops.filter((item) => item.id !== shop.id);
  state.total = Math.max(state.total - 1, 0);
  if (state.selectedShopId === shop.id) {
    state.selectedShopId = state.shops[0]?.id || "";
  }
  renderShops();
  showToast("Đã xóa shop");
}

async function checkSelectedShopToken() {
  const shop = getSelectedShop();
  if (!shop) {
    return;
  }

  const localStatus = getLocalTokenStatus(shop);
  state.tokenStatuses.set(shop.id, localStatus);
  renderShops();

  if (!localStatus.valid) {
    return;
  }

  els.apiStatus.textContent = "Đang kiểm tra token";
  const status = await requestJSON(`/api/v1/shops/${shop.id}/token-status`);
  state.tokenStatuses.set(shop.id, {
    valid: status.valid,
    status: status.status,
    message: status.valid ? "Token hoạt động" : (status.message || "Token không hoạt động"),
  });
  renderShops();
}

async function downloadReport(event) {
  event.preventDefault();
  const shop = getSelectedShop();
  const tokenStatus = shop ? getTokenStatus(shop) : null;
  if (!shop || !shop.apiKey || tokenStatus?.valid === false) {
    showToast(tokenStatus?.message || "Hãy chọn shop có API key hợp lệ");
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
  const tokenStatus = shop ? getTokenStatus(shop) : null;
  if (!shop || !shop.apiKey || tokenStatus?.valid === false) {
    showToast(tokenStatus?.message || "Hãy chọn shop có API key hợp lệ");
    return;
  }

  els.loadOrdersButton.disabled = true;
  els.loadOrdersButton.textContent = "Đang tải";
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
    els.loadOrdersButton.textContent = "Xem đơn hàng";
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

els.addShopButton.addEventListener("click", openAddShopModal);
els.shopForm.addEventListener("submit", (event) => {
  saveShop(event).catch((error) => showToast(error.message));
});
els.shopList.addEventListener("scroll", () => {
  const nearBottom = els.shopList.scrollTop + els.shopList.clientHeight >= els.shopList.scrollHeight - 80;
  if (nearBottom) {
    loadNextShopPage().catch((error) => showToast(error.message));
  }
});
els.reportShopSelect.addEventListener("change", (event) => selectShop(event.target.value));
els.reportForm.addEventListener("submit", downloadReport);
els.loadOrdersButton.addEventListener("click", loadOrders);

els.dateFrom.value = daysAgoISO(7);
els.dateTo.value = todayISO();
loadShops({ reset: true })
  .then(() => checkSelectedShopToken())
  .catch((error) => showToast(error.message));
