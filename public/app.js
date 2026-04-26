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
  reportJobs: [],
  deletedReportJobIds: new Set(),
  ordersRequestKey: "",
  selectedShopId: "",
  tokenModalShopId: "",
  activeReportJobId: "",
  page: 1,
  limit: 20,
  total: 0,
  totalPages: 1,
  isLoadingShops: false,
};

const JOB_POLL_INTERVAL_MS = 30000;

const els = {
  sidebar: document.querySelector("#sidebar"),
  mobileBackdrop: document.querySelector("#mobileBackdrop"),
  menuButton: document.querySelector("#menuButton"),
  jobsButton: document.querySelector("#jobsButton"),
  jobsBadge: document.querySelector("#jobsBadge"),
  jobsDrawer: document.querySelector("#jobsDrawer"),
  jobsList: document.querySelector("#jobsList"),
  closeJobsButton: document.querySelector("#closeJobsButton"),
  currentShopTitle: document.querySelector("#currentShopTitle"),
  editSelectedShopButton: document.querySelector("#editSelectedShopButton"),
  deleteSelectedShopButton: document.querySelector("#deleteSelectedShopButton"),
  addShopButton: document.querySelector("#addShopButton"),
  shopForm: document.querySelector("#shopForm"),
  shopModal: document.querySelector("#shopModal"),
  tokenModal: document.querySelector("#tokenModal"),
  tokenModalMessage: document.querySelector("#tokenModalMessage"),
  updateTokenButton: document.querySelector("#updateTokenButton"),
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
  dateFrom: document.querySelector("#dateFrom"),
  dateTo: document.querySelector("#dateTo"),
  tax: document.querySelector("#tax"),
  discount: document.querySelector("#discount"),
  downloadReportButton: document.querySelector("#downloadReportButton"),
  reportProgress: document.querySelector("#reportProgress"),
  reportProgressText: document.querySelector("#reportProgressText"),
  reportProgressPercent: document.querySelector("#reportProgressPercent"),
  reportProgressBar: document.querySelector("#reportProgressBar"),
  loadOrdersButton: document.querySelector("#loadOrdersButton"),
  totalOrders: document.querySelector("#totalOrders"),
  totalPrevOrders: document.querySelector("#totalPrevOrders"),
  totalOrdersRevenue: document.querySelector("#totalOrdersRevenue"),
  totalPrevOrdersRevenue: document.querySelector("#totalPrevOrdersRevenue"),
  ordersStatus: document.querySelector("#ordersStatus"),
  ordersTableBody: document.querySelector("#ordersTableBody"),
  toast: document.querySelector("#toast"),
};

for (const [name, element] of Object.entries(els)) {
  if (!element) {
    throw new Error(`Missing DOM element: ${name}`);
  }
}

const shopModal = bootstrap.Modal.getOrCreateInstance(els.shopModal);
const tokenModal = bootstrap.Modal.getOrCreateInstance(els.tokenModal);
els.editSelectedShopButton.innerHTML = iconEdit;
els.deleteSelectedShopButton.innerHTML = iconTrash;

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

function updateOverlayState() {
  const isSidebarOpen = els.sidebar.classList.contains("is-open");
  const isJobsOpen = els.jobsDrawer.classList.contains("is-open");
  const hasOverlay = isSidebarOpen || isJobsOpen;

  els.mobileBackdrop.classList.toggle("is-visible", hasOverlay);
  document.body.classList.toggle("is-overlay-open", hasOverlay);
  els.menuButton.setAttribute("aria-expanded", String(isSidebarOpen));
  els.jobsButton.setAttribute("aria-expanded", String(isJobsOpen));
}

function setSidebarOpen(isOpen) {
  if (isOpen) {
    els.jobsDrawer.classList.remove("is-open");
  }
  els.sidebar.classList.toggle("is-open", isOpen);
  updateOverlayState();
}

function setJobsDrawerOpen(isOpen, { refresh = true } = {}) {
  if (isOpen) {
    els.sidebar.classList.remove("is-open");
  }
  els.jobsDrawer.classList.toggle("is-open", isOpen);
  updateOverlayState();
  if (isOpen && refresh) {
    loadReportJobs().catch((error) => showToast(error.message));
  }
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

function currentReportScope() {
  const shop = getSelectedShop();
  return {
    shopId: shop?.id || "",
    dateFrom: els.dateFrom.value,
    dateTo: els.dateTo.value,
    tax: Number(els.tax.value || 0.06),
    discount: Number(els.discount.value || 3.5),
  };
}

function numbersMatch(left, right) {
  return Math.abs(Number(left) - Number(right)) < 0.0001;
}

function jobMatchesCurrentReport(job) {
  const scope = currentReportScope();
  if (!scope.shopId || job.shopId !== scope.shopId || job.dateFrom !== scope.dateFrom || job.dateTo !== scope.dateTo) {
    return false;
  }

  const jobTax = job.tax ?? scope.tax;
  const jobDiscount = job.discount ?? scope.discount;
  return numbersMatch(jobTax, scope.tax) && numbersMatch(jobDiscount, scope.discount);
}

function findCurrentReportJob({ activeOnly = false } = {}) {
  return state.reportJobs.find((job) => jobMatchesCurrentReport(job) && (!activeOnly || isJobActive(job))) || null;
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

  return { valid: true, status: "unchecked", message: "" };
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
  const selectedShop = getSelectedShop();
  els.currentShopTitle.textContent = selectedShop?.name || "Chưa chọn shop";
  els.editSelectedShopButton.disabled = !selectedShop;
  els.deleteSelectedShopButton.disabled = !selectedShop;

  if (state.shops.length === 0) {
    const empty = document.createElement("div");
    empty.className = "shop-item";
    empty.textContent = "Chưa có shop";
    els.shopList.appendChild(empty);
  }

  for (const shop of state.shops) {
    const item = document.createElement("div");
    const tokenStatus = getTokenStatus(shop);
    const hasRunningJob = state.reportJobs.some((job) => job.shopId === shop.id && isJobActive(job));
    item.className = `shop-item${shop.id === state.selectedShopId ? " is-selected" : ""}${tokenStatus.valid ? "" : " is-token-invalid"}`;

    const info = document.createElement("button");
    info.className = "shop-select-button";
    info.type = "button";
    info.innerHTML = `
      <div class="shop-name">${escapeHTML(shop.name)}${hasRunningJob ? '<span class="shop-job-spinner" title="Đang có job báo cáo"></span>' : ""}</div>
    `;
    info.addEventListener("click", () => {
      info.blur();
      selectShop(shop.id);
      setSidebarOpen(false);
    });

    item.append(info);
    els.shopList.appendChild(item);
  }

}

function selectShop(id, { notifyTokenIssue = true } = {}) {
  state.selectedShopId = id;
  renderShops();
  syncReportProgressWithSelection();
  resetOrders();
  checkSelectedShopToken({ notifyTokenIssue }).catch((error) => showToast(error.message));
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
  selectShop(shop.id, { notifyTokenIssue: false });
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

function showTokenInvalidModal(shop, status) {
  if (!shop) {
    return;
  }

  state.tokenModalShopId = shop.id;
  els.tokenModalMessage.textContent = `${shop.name}: ${status?.message || "API key không còn hợp lệ. Vui lòng cập nhật để tiếp tục tải báo cáo."}`;
  tokenModal.show();
}

function canUseSelectedShop() {
  const shop = getSelectedShop();
  if (!shop) {
    showToast("Hãy chọn shop ở danh sách bên trái");
    return false;
  }

  const tokenStatus = getTokenStatus(shop);
  if (!shop.apiKey || tokenStatus?.valid === false) {
    showTokenInvalidModal(shop, tokenStatus);
    return false;
  }

  return true;
}

function isJobActive(job) {
  return job.status === "queued" || job.status === "running";
}

function upsertReportJob(job) {
  if (state.deletedReportJobIds.has(job.id)) {
    return;
  }

  const index = state.reportJobs.findIndex((item) => item.id === job.id);
  const nextJob = index >= 0 ? { ...state.reportJobs[index], ...job } : job;
  if (index >= 0) {
    state.reportJobs[index] = nextJob;
  } else {
    state.reportJobs.unshift(nextJob);
  }
  if (nextJob.id === state.activeReportJobId) {
    updateReportProgress(nextJob);
  }
  renderReportJobs();
  renderShops();
  syncReportProgressWithSelection();
}

function resetReportProgress() {
  state.activeReportJobId = "";
  els.reportProgress.hidden = true;
  els.reportProgressText.textContent = "Đang chuẩn bị báo cáo";
  els.reportProgressPercent.textContent = "0%";
  els.reportProgressBar.style.width = "0%";
  els.reportProgressBar.setAttribute("aria-valuenow", "0");
  els.downloadReportButton.disabled = false;
  els.downloadReportButton.textContent = "Tải báo cáo Excel";
}

function syncReportProgressWithSelection() {
  const matchingJob = findCurrentReportJob();
  if (!matchingJob) {
    resetReportProgress();
    return;
  }

  state.activeReportJobId = matchingJob.id;
  updateReportProgress(matchingJob);
  els.downloadReportButton.disabled = isJobActive(matchingJob);
  if (!isJobActive(matchingJob)) {
    els.downloadReportButton.textContent = matchingJob.status === "done" ? "Tải file đã tạo" : "Tải báo cáo Excel";
  }
}

function renderReportJobs() {
  const activeCount = state.reportJobs.filter(isJobActive).length;
  els.jobsBadge.hidden = activeCount === 0;
  els.jobsBadge.textContent = String(activeCount);
  els.jobsButton.classList.toggle("has-active-jobs", activeCount > 0);
  els.jobsList.innerHTML = "";

  if (state.reportJobs.length === 0) {
    const empty = document.createElement("div");
    empty.className = "job-empty";
    empty.textContent = "Chưa có job báo cáo";
    els.jobsList.appendChild(empty);
    return;
  }

  for (const job of state.reportJobs) {
    const progress = Math.max(0, Math.min(Number(job.progress || 0), 100));
    const item = document.createElement("div");
    item.className = `job-item is-${job.status}`;
    item.innerHTML = `
      <div class="job-title">
        <span>${escapeHTML(job.shopName || "Shop")}</span>
        <div class="job-title-actions">
          <span class="job-status">
            ${isJobActive(job) ? '<span class="job-spinner"></span>' : ""}
            <strong>${progress}%</strong>
          </span>
        </div>
      </div>
      <div class="job-meta">${escapeHTML(job.dateFrom || "")} - ${escapeHTML(job.dateTo || "")}</div>
      <div class="job-step">${escapeHTML(job.error || job.currentStep || "Đang chờ xử lý")}</div>
      <div class="progress job-progress" role="progressbar" aria-valuemin="0" aria-valuemax="100">
        <div class="progress-bar" style="width: ${progress}%"></div>
      </div>
    `;

    const actions = item.querySelector(".job-title-actions");
    const deleteButton = document.createElement("button");
    deleteButton.className = "job-delete-button";
    deleteButton.type = "button";
    deleteButton.title = isJobActive(job) ? "Hủy job" : "Xóa job";
    deleteButton.setAttribute("aria-label", isJobActive(job) ? "Hủy job" : "Xóa job");
    deleteButton.innerHTML = iconTrash;
    deleteButton.addEventListener("click", () => {
      deleteReportJob(job).catch((error) => showToast(error.message));
    });
    actions.appendChild(deleteButton);

    if (job.status === "done" && job.downloadUrl) {
      const button = document.createElement("button");
      button.className = "btn btn-sm btn-outline-primary mt-2";
      button.type = "button";
      button.textContent = "Tải file";
      button.addEventListener("click", () => {
        downloadReportJob(job.downloadUrl, job).catch((error) => showToast(error.message));
      });
      item.appendChild(button);
    }

    els.jobsList.appendChild(item);
  }
}

async function loadReportJobs() {
  const data = await requestJSON("/api/v1/reports/jobs");
  state.reportJobs = data.items || [];
  renderReportJobs();
  renderShops();
  await refreshRunningReportJobs();
  syncReportProgressWithSelection();
}

async function refreshRunningReportJobs() {
  const runningJobs = state.reportJobs.filter(isJobActive);
  for (const job of runningJobs) {
    if (state.deletedReportJobIds.has(job.id)) {
      continue;
    }
    try {
      const latest = await requestJSON(`/api/v1/reports/jobs/${job.id}`);
      upsertReportJob({ ...job, ...latest });
    } catch (error) {
      upsertReportJob({ ...job, status: "failed", error: "Không thể cập nhật trạng thái job" });
    }
  }
}

async function deleteReportJob(job) {
  const message = isJobActive(job) ? "Hủy job báo cáo đang chạy?" : "Xóa job báo cáo?";
  if (!window.confirm(message)) {
    return;
  }

  state.deletedReportJobIds.add(job.id);
  state.reportJobs = state.reportJobs.filter((item) => item.id !== job.id);
  if (state.activeReportJobId === job.id) {
    state.activeReportJobId = "";
    syncReportProgressWithSelection();
    els.reportProgressText.textContent = "Đã hủy job báo cáo";
  }
  renderReportJobs();
  renderShops();

  try {
    await requestJSON(`/api/v1/reports/jobs/${job.id}`, { method: "DELETE" });
    showToast(isJobActive(job) ? "Đã hủy job" : "Đã xóa job");
  } catch (error) {
    state.deletedReportJobIds.delete(job.id);
    upsertReportJob(job);
    throw error;
  }
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
  checkSelectedShopToken({ notifyTokenIssue: false }).catch((error) => showToast(error.message));
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

async function checkSelectedShopToken({ notifyTokenIssue = false } = {}) {
  const shop = getSelectedShop();
  if (!shop) {
    return;
  }

  const localStatus = getLocalTokenStatus(shop);
  state.tokenStatuses.set(shop.id, localStatus);
  renderShops();

  if (!localStatus.valid) {
    if (notifyTokenIssue) {
      showTokenInvalidModal(shop, localStatus);
    }
    return;
  }

  const status = await requestJSON(`/api/v1/shops/${shop.id}/token-status`);
  const remoteStatus = {
    valid: status.valid,
    status: status.status,
    message: status.valid ? "" : (status.message || "API key không còn hợp lệ"),
  };
  state.tokenStatuses.set(shop.id, remoteStatus);
  renderShops();

  if (!remoteStatus.valid && notifyTokenIssue) {
    showTokenInvalidModal(shop, remoteStatus);
  }
}

async function downloadReport(event) {
  event.preventDefault();
  if (!canUseSelectedShop()) {
    return;
  }
  const shop = getSelectedShop();
  const existingJob = findCurrentReportJob();
  if (existingJob) {
    state.activeReportJobId = existingJob.id;
    updateReportProgress(existingJob);

    if (isJobActive(existingJob)) {
      showToast("Job báo cáo này đang chạy");
      els.downloadReportButton.disabled = true;
      try {
        await waitForReportJob(existingJob);
      } catch (error) {
        showToast(error.message);
      } finally {
        syncReportProgressWithSelection();
      }
      return;
    }

    if (existingJob.status === "done" && existingJob.downloadUrl) {
      await downloadReportJob(existingJob.downloadUrl, existingJob);
      showToast("Đã tải báo cáo");
      return;
    }
  }

  els.downloadReportButton.disabled = true;
  els.downloadReportButton.textContent = "Đang tạo job";
  updateReportProgress({ progress: 0, currentStep: "Đang tạo job báo cáo" });

  try {
    const jobMeta = {
      shopId: shop.id,
      shopName: shop.name,
      dateFrom: els.dateFrom.value,
      dateTo: els.dateTo.value,
    };
    const job = await requestJSON("/api/v1/reports/jobs", {
      method: "POST",
      body: JSON.stringify({
        apiKey: shop.apiKey,
        shopId: shop.id,
        shopName: shop.name,
        dateFrom: els.dateFrom.value,
        dateTo: els.dateTo.value,
        tax: Number(els.tax.value || 0.06),
        discount: Number(els.discount.value || 3.5),
      }),
    });
    const reportJob = { ...jobMeta, ...job };
    state.activeReportJobId = job.id;
    upsertReportJob(reportJob);
    await waitForReportJob(reportJob);
    showToast("Đã tải báo cáo");
  } catch (error) {
    showToast(error.message);
  } finally {
    syncReportProgressWithSelection();
  }
}

function updateReportProgress(job) {
  const progress = Math.max(0, Math.min(Number(job.progress || 0), 100));
  els.reportProgress.hidden = false;
  els.reportProgressText.textContent = job.currentStep || "Đang xử lý báo cáo";
  els.reportProgressPercent.textContent = `${progress}%`;
  els.reportProgressBar.style.width = `${progress}%`;
  els.reportProgressBar.setAttribute("aria-valuenow", String(progress));
  els.downloadReportButton.textContent = progress > 0 ? `Đang xử lý ${progress}%` : "Đang xử lý";
}

function isMainReportJob(job) {
  return job.id === state.activeReportJobId && jobMatchesCurrentReport(job);
}

async function waitForReportJob(reportJob) {
  let job = await requestJSON(`/api/v1/reports/jobs/${reportJob.id}`);
  job = { ...reportJob, ...job };
  upsertReportJob(job);
  if (isMainReportJob(job)) {
    updateReportProgress(job);
  }

  while (job.status === "queued" || job.status === "running") {
    await sleep(JOB_POLL_INTERVAL_MS);
    if (state.deletedReportJobIds.has(job.id)) {
      throw new Error("Đã hủy job báo cáo");
    }
    if (!isMainReportJob(job)) {
      throw new Error("Job vẫn đang chạy trong Reports");
    }
    job = { ...job, ...(await requestJSON(`/api/v1/reports/jobs/${job.id}`)) };
    upsertReportJob(job);
    if (isMainReportJob(job)) {
      updateReportProgress(job);
    }
  }

  if (job.status === "failed") {
    throw new Error(job.error || "Không thể tạo báo cáo");
  }
  if (job.status !== "done" || !job.downloadUrl) {
    throw new Error("Báo cáo chưa sẵn sàng");
  }

  if (isMainReportJob(job)) {
    await downloadReportJob(job.downloadUrl, job);
  }
}

async function downloadReportJob(downloadUrl, job = {}) {
  const response = await fetch(downloadUrl);
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error || "Không thể tải báo cáo");
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${sanitizeFileName(job.shopName || getSelectedShop()?.name || "shop")}_report_${job.dateFrom || els.dateFrom.value}_${job.dateTo || els.dateTo.value}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function sanitizeFileName(value) {
  return String(value || "shop")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .replace(/^[. ]+|[. ]+$/g, "") || "shop";
}

function currentOrdersKey() {
  const shop = getSelectedShop();
  return [shop?.id || "", els.dateFrom.value, els.dateTo.value].join("|");
}

function formatNumber(value) {
  return new Intl.NumberFormat("vi-VN").format(Number(value || 0));
}

function setOrdersStatus(message, stateName = "") {
  els.ordersStatus.textContent = message;
  els.ordersStatus.className = `orders-status${stateName ? ` is-${stateName}` : ""}`;
}

function resetOrders() {
  state.ordersRequestKey = "";
  els.totalOrders.textContent = "0";
  els.totalPrevOrders.textContent = "0";
  els.totalOrdersRevenue.textContent = "0";
  els.totalPrevOrdersRevenue.textContent = "0";
  els.ordersTableBody.innerHTML = '<tr><td class="empty-row" colspan="6">Chưa có dữ liệu</td></tr>';
  setOrdersStatus("Chưa tải dữ liệu");
}

async function loadOrders() {
  if (!canUseSelectedShop()) {
    return;
  }
  const shop = getSelectedShop();
  const requestKey = currentOrdersKey();
  state.ordersRequestKey = requestKey;

  els.loadOrdersButton.disabled = true;
  els.loadOrdersButton.textContent = "Đang tải";
  setOrdersStatus("Đang tải dữ liệu", "loading");
  try {
    const data = await requestJSON("/api/v1/orders", {
      method: "POST",
      body: JSON.stringify({
        apiKey: shop.apiKey,
        dateFrom: els.dateFrom.value,
        dateTo: els.dateTo.value,
      }),
    });
    if (state.ordersRequestKey !== requestKey) {
      return;
    }
    renderOrders(data);
  } catch (error) {
    if (state.ordersRequestKey === requestKey) {
      setOrdersStatus("Không thể tải dữ liệu", "error");
      showToast(error.message);
    }
  } finally {
    if (state.ordersRequestKey === requestKey) {
      els.loadOrdersButton.disabled = false;
      els.loadOrdersButton.textContent = "Xem đơn hàng";
    }
  }
}

function renderOrders(data) {
  els.totalOrders.textContent = formatNumber(data.totalOrders);
  els.totalPrevOrders.textContent = formatNumber(data.totalPrevOrders);
  els.totalOrdersRevenue.textContent = `${formatNumber(data.totalOrdersSumRub)} RUB`;
  els.totalPrevOrdersRevenue.textContent = `${formatNumber(data.totalPrevOrdersSumRub)} RUB`;
  els.ordersTableBody.innerHTML = "";

  const rows = data.chartData || [];
  if (rows.length === 0) {
    setOrdersStatus("Không có sản phẩm có đơn trong khoảng ngày này", "warning");
    els.ordersTableBody.innerHTML = '<tr><td class="empty-row" colspan="6">Không có dữ liệu</td></tr>';
    return;
  }

  setOrdersStatus(`Đã tải ${formatNumber(rows.length)} sản phẩm qua ${formatNumber(data.pagesLoaded || 1)} page`, data.hasMorePages ? "warning" : "ready");

  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.nmID}</td>
      <td>${escapeHTML(row.vendorCode || "")}</td>
      <td>${formatNumber(row.ordersCount)}</td>
      <td>${formatNumber(row.prevOrdersCount)}</td>
      <td>${formatNumber(row.ordersSumRub)}</td>
      <td>${formatNumber(row.prevOrdersSumRub)}</td>
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

els.addShopButton.addEventListener("click", () => {
  setSidebarOpen(false);
  openAddShopModal();
});
els.editSelectedShopButton.addEventListener("click", () => {
  const shop = getSelectedShop();
  if (!shop) {
    showToast("Hãy chọn shop trước");
    return;
  }
  openEditShopModal(shop);
});
els.deleteSelectedShopButton.addEventListener("click", () => {
  const shop = getSelectedShop();
  if (!shop) {
    showToast("Hãy chọn shop trước");
    return;
  }
  deleteShop(shop).catch((error) => showToast(error.message));
});
els.menuButton.addEventListener("click", () => setSidebarOpen(true));
els.jobsButton.addEventListener("click", () => setJobsDrawerOpen(true));
els.closeJobsButton.addEventListener("click", () => setJobsDrawerOpen(false));
els.mobileBackdrop.addEventListener("click", () => {
  setSidebarOpen(false);
  setJobsDrawerOpen(false);
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    setSidebarOpen(false);
    setJobsDrawerOpen(false);
  }
});
els.shopForm.addEventListener("submit", (event) => {
  saveShop(event).catch((error) => showToast(error.message));
});
els.shopList.addEventListener("scroll", () => {
  const nearBottom = els.shopList.scrollTop + els.shopList.clientHeight >= els.shopList.scrollHeight - 80;
  if (nearBottom) {
    loadNextShopPage().catch((error) => showToast(error.message));
  }
});
els.updateTokenButton.addEventListener("click", () => {
  const shop = state.shops.find((item) => item.id === state.tokenModalShopId);
  tokenModal.hide();
  if (shop) {
    openEditShopModal(shop);
  }
});
els.reportForm.addEventListener("submit", downloadReport);
els.loadOrdersButton.addEventListener("click", loadOrders);
for (const input of [els.dateFrom, els.dateTo]) {
  input.addEventListener("change", () => {
    syncReportProgressWithSelection();
    resetOrders();
  });
}
for (const input of [els.tax, els.discount]) {
  input.addEventListener("change", syncReportProgressWithSelection);
}

els.dateFrom.value = daysAgoISO(7);
els.dateTo.value = todayISO();
resetOrders();
updateOverlayState();
loadShops({ reset: true })
  .then(() => checkSelectedShopToken({ notifyTokenIssue: false }))
  .catch((error) => showToast(error.message));
loadReportJobs().catch((error) => showToast(error.message));
renderReportJobs();
