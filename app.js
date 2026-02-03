const state = {
  caseInformation: "",
  dateOfSearch: "",
  conclusionDate: "",
  premiseAddress: "",
  partyName: "",
  partyFatherName: "",
  partyAadhaar: "",
  partyAddress: "",
  authorizedOfficer: "",
  examinerName: "",
  employedBy: "EnReach Solution",
  deviceType: "",
  deviceName: "",
  deviceMake: "",
  deviceSize: "",
  deviceModelNo: "",
  deviceSerialNo: "",
  deviceImei1: "",
  deviceImei2: "",
  includeMakeModel: true,
  includeSerialSize: true,
  includeImei: true,
  hashMd5: "",
  hashSha1: "",
  hashSha256: "",
  imagingSoftware: "",
  actionTaken: "",
  masterCopyMake: "",
  masterCopyModel: "",
  masterCopySerial: "",
  masterCopySize: "",
  workingCopyMake: "",
  workingCopyModel: "",
  workingCopySerial: "",
  workingCopySize: "",
  masterCopies: [],
  workingCopies: [],
  stickerWorkingSerial: "",
  stickerMasterSerial: "",
  customEnabled: false,
  customTitle: "",
  customTemplate: "",
  customDocxName: "",
  deviceList: []
};

const escapeHtml = (value) =>
  String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

const formatDate = (value) => {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
};

const renderDeviceList = () => {
  const el = document.getElementById("deviceList");
  el.innerHTML = "";
  state.deviceList.forEach((device, idx) => {
    const chip = document.createElement("div");
    chip.className = "device-chip";
    chip.innerHTML = `${escapeHtml(device.name || "Unnamed device")} <button data-idx="${idx}">x</button>`;
    el.appendChild(chip);
  });
  el.querySelectorAll("button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.idx);
      state.deviceList.splice(idx, 1);
      renderDeviceList();
      renderPreview();
    });
  });
};

const renderCopyLists = () => {
  const masterEl = document.getElementById("masterCopyList");
  const workingEl = document.getElementById("workingCopyList");
  if (masterEl) {
    masterEl.innerHTML = state.masterCopies.map((item, idx) => `
      <div class="device-chip">Master ${idx + 1}: ${escapeHtml(item.make)} ${escapeHtml(item.model)} (${escapeHtml(item.serial)}) ${escapeHtml(item.size)} <button data-type="master" data-idx="${idx}">x</button></div>
    `).join("");
  }
  if (workingEl) {
    workingEl.innerHTML = state.workingCopies.map((item, idx) => `
      <div class="device-chip">Working ${idx + 1}: ${escapeHtml(item.make)} ${escapeHtml(item.model)} (${escapeHtml(item.serial)}) ${escapeHtml(item.size)} <button data-type="working" data-idx="${idx}">x</button></div>
    `).join("");
  }
  document.querySelectorAll("#masterCopyList button, #workingCopyList button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.idx);
      const type = btn.dataset.type;
      if (type === "master") state.masterCopies.splice(idx, 1);
      if (type === "working") state.workingCopies.splice(idx, 1);
      renderCopyLists();
      renderPreview();
    });
  });
};

const getCopyList = (type) => {
  const list = type === "master" ? state.masterCopies : state.workingCopies;
  if (list.length) return list;
  if (type === "master" && (state.masterCopyMake || state.masterCopySerial || state.masterCopySize || state.masterCopyModel)) {
    return [{
      make: state.masterCopyMake,
      model: state.masterCopyModel,
      serial: state.masterCopySerial,
      size: state.masterCopySize
    }];
  }
  if (type === "working" && (state.workingCopyMake || state.workingCopySerial || state.workingCopySize || state.workingCopyModel)) {
    return [{
      make: state.workingCopyMake,
      model: state.workingCopyModel,
      serial: state.workingCopySerial,
      size: state.workingCopySize
    }];
  }
  return [];
};

const renderCopyTable = (title, list) => {
  if (!list.length) return `<div class="section"><strong>${title}:</strong> Not provided.</div>`;
  const rows = list.map((item, idx) => `
    <tr>
      <td>${idx + 1}</td>
      <td>${escapeHtml(item.make)}</td>
      <td>${escapeHtml(item.model)}</td>
      <td>${escapeHtml(item.serial)}</td>
      <td>${escapeHtml(item.size)}</td>
    </tr>
  `).join("");
  return `
    <div class="section">
      <strong>${title}</strong>
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Make</th>
            <th>Model</th>
            <th>Serial</th>
            <th>Size</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
};

const getCommon = () => ({
  caseInformation: escapeHtml(state.caseInformation),
  dateOfSearch: escapeHtml(formatDate(state.dateOfSearch)),
  conclusionDate: escapeHtml(formatDate(state.conclusionDate)),
  premiseAddress: escapeHtml(state.premiseAddress),
  partyName: escapeHtml(state.partyName),
  partyFatherName: escapeHtml(state.partyFatherName),
  partyAadhaar: escapeHtml(state.partyAadhaar),
  partyAddress: escapeHtml(state.partyAddress),
  authorizedOfficer: escapeHtml(state.authorizedOfficer),
  examinerName: escapeHtml(state.examinerName),
  employedBy: escapeHtml(state.employedBy),
  deviceType: escapeHtml(state.deviceType),
  deviceName: escapeHtml(state.deviceName),
  deviceMake: escapeHtml(state.deviceMake),
  deviceSize: escapeHtml(state.deviceSize),
  deviceModelNo: escapeHtml(state.deviceModelNo),
  deviceSerialNo: escapeHtml(state.deviceSerialNo),
  deviceImei1: escapeHtml(state.deviceImei1),
  deviceImei2: escapeHtml(state.deviceImei2),
  hashMd5: escapeHtml(state.hashMd5),
  hashSha1: escapeHtml(state.hashSha1),
  hashSha256: escapeHtml(state.hashSha256),
  imagingSoftware: escapeHtml(state.imagingSoftware),
  actionTaken: escapeHtml(state.actionTaken),
  masterCopyMake: escapeHtml(state.masterCopyMake),
  masterCopyModel: escapeHtml(state.masterCopyModel),
  masterCopySerial: escapeHtml(state.masterCopySerial),
  masterCopySize: escapeHtml(state.masterCopySize),
  workingCopyMake: escapeHtml(state.workingCopyMake),
  workingCopyModel: escapeHtml(state.workingCopyModel),
  workingCopySerial: escapeHtml(state.workingCopySerial),
  workingCopySize: escapeHtml(state.workingCopySize),
  stickerWorkingSerial: escapeHtml(state.stickerWorkingSerial),
  stickerMasterSerial: escapeHtml(state.stickerMasterSerial)
});

const getDevices = () => {
  if (state.deviceList.length) return state.deviceList;
  if (!state.deviceName && !state.deviceType) return [];
  return [{
    type: state.deviceType,
    name: state.deviceName,
    make: state.deviceMake,
    size: state.deviceSize,
    model: state.deviceModelNo,
    serial: state.deviceSerialNo,
    imei1: state.deviceImei1,
    imei2: state.deviceImei2,
    md5: state.hashMd5,
    sha1: state.hashSha1,
    sha256: state.hashSha256,
    imagingSoftware: state.imagingSoftware,
    action: state.actionTaken
  }];
};

const renderDeviceTable = (devices) => {
  if (!devices.length) return "<div>No device added.</div>";
  const rows = devices.map((d, idx) => {
    const includeMakeModel = d.includeMakeModel ?? true;
    const includeSerialSize = d.includeSerialSize ?? true;
    const includeImei = d.includeImei ?? true;
    const rowSpan = 1
      + (includeMakeModel ? 1 : 0)
      + (includeSerialSize ? 1 : 0)
      + (includeImei ? 1 : 0);

    let detailRows = `
      <tr>
        <td rowspan="${rowSpan}">${idx + 1}</td>
        <td>${escapeHtml(d.type)}</td>
        <td colspan="6"><strong>${escapeHtml(d.name)}</strong></td>
      </tr>
    `;

    if (includeMakeModel) {
      detailRows += `
        <tr>
          <td>Make</td>
          <td colspan="2">${escapeHtml(d.make)}</td>
          <td>Model</td>
          <td colspan="2">${escapeHtml(d.model)}</td>
        </tr>
      `;
    }

    if (includeSerialSize) {
      detailRows += `
        <tr>
          <td>Serial</td>
          <td colspan="2">${escapeHtml(d.serial)}</td>
          <td>Size</td>
          <td colspan="2">${escapeHtml(d.size)}</td>
        </tr>
      `;
    }

    if (includeImei) {
      detailRows += `
        <tr>
          <td>IMEI 1</td>
          <td colspan="2">${escapeHtml(d.imei1)}</td>
          <td>IMEI 2</td>
          <td colspan="2">${escapeHtml(d.imei2)}</td>
        </tr>
      `;
    }

    detailRows += `
      <tr>
        <td colspan="8"><strong>Action Taken:</strong> ${escapeHtml(d.action)}</td>
      </tr>
    `;

    return detailRows;
  }).join("");

  return `
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>Type</th>
          <th colspan="6">Device Name</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
};

const renderSignatureTable = ({ includeAuthorizedOfficer = true } = {}) => `
  <table>
    <thead>
      <tr>
        <th>Party Signature</th>
        <th>Witness 1 Signature</th>
        <th>Witness 2 Signature</th>
        <th>Examiner Signature</th>
        ${includeAuthorizedOfficer ? "<th>Authorized Officer Signature</th>" : ""}
      </tr>
    </thead>
    <tbody>
      <tr>
        <td style="height:40px"></td>
        <td style="height:40px"></td>
        <td style="height:40px"></td>
        <td style="height:40px"></td>
        ${includeAuthorizedOfficer ? "<td style='height:40px'></td>" : ""}
      </tr>
    </tbody>
  </table>
`;

const renderHashBlock = (device) => `
  <div class="hash-block">
    <div class="hash-title">Hash Values</div>
    <div class="hash-row"><span>MD5</span> ${escapeHtml(device.md5 || "")}</div>
    <div class="hash-row"><span>SHA1</span> ${escapeHtml(device.sha1 || "")}</div>
    <div class="hash-row"><span>SHA256</span> ${escapeHtml(device.sha256 || "")}</div>
  </div>
`;

const applyTemplate = (template, data) => {
  if (!template) return "";
  return template.replace(/\{\{(.*?)\}\}/g, (_, key) => {
    const normalized = key.trim();
    return data[normalized] ?? "";
  });
};

const renderCustomDoc = () => {
  if (!state.customEnabled) {
    return `<div class="section">Custom document is disabled.</div>`;
  }
  const d = getCommon();
  const devices = getDevices();
  const data = {
    caseInformation: d.caseInformation,
    dateOfSearch: d.dateOfSearch,
    conclusionDate: d.conclusionDate,
    premiseAddress: d.premiseAddress,
    partyName: d.partyName,
    partyFatherName: d.partyFatherName,
    partyAadhaar: d.partyAadhaar,
    partyAddress: d.partyAddress,
    authorizedOfficer: d.authorizedOfficer,
    examinerName: d.examinerName,
    employedBy: d.employedBy,
    deviceType: d.deviceType,
    deviceName: d.deviceName,
    deviceMake: d.deviceMake,
    deviceSize: d.deviceSize,
    deviceModelNo: d.deviceModelNo,
    deviceSerialNo: d.deviceSerialNo,
    deviceImei1: d.deviceImei1,
    deviceImei2: d.deviceImei2,
    hashMd5: d.hashMd5,
    hashSha1: d.hashSha1,
    hashSha256: d.hashSha256,
    imagingSoftware: d.imagingSoftware,
    actionTaken: d.actionTaken,
    devices_table: renderDeviceTable(devices),
    hash_blocks: devices.map((device) => renderHashBlock(device)).join("")
  };
  const title = state.customTitle || "Custom Document";
  const body = applyTemplate(state.customTemplate, data);
  return `
    <div class="doc-title">${escapeHtml(title)}</div>
    <div class="doc-subtitle">Custom Template</div>
    <div class="section">${body || "Paste a template to render."}</div>
    ${state.customDocxName ? `<div class="section muted">Uploaded DOCX: ${escapeHtml(state.customDocxName)}</div>` : ""}
  `;
};

const templates = {
  cert63: () => {
    const d = getCommon();
    const devices = getDevices();
    const masterList = getCopyList("master");
    const workingList = getCopyList("working");
    const partA = `
      <div class="doc-title">Certificate</div>
      <div class="doc-subtitle">Section 63(4)(c) - Part A</div>
      <div class="section">
        <div class="label">Part A (To be filled by the Party)</div>
        <p>
          I, <strong>${d.partyName}</strong> Son of Shri <strong>${d.partyFatherName}</strong>,
          Aadhaar No. <strong>${d.partyAadhaar}</strong>, residing at <strong>${d.partyAddress}</strong>,
          in the case of <strong>${d.caseInformation}</strong>, affirm and sincerely state that we have
          produced electronic record/output from the following device(s)/digital record source.
        </p>
      </div>
      <div class="section">
        <strong>Devices Produced:</strong>
        ${renderDeviceTable(devices)}
      </div>
      <div class="section meta-grid">
        <div><span class="label">Date:</span> ${d.dateOfSearch}</div>
        <div><span class="label">Party Name:</span> ${d.partyName}</div>
      </div>
      <div class="section">
        <div class="signature">Party Signature: ${d.partyName}</div>
      </div>
      <div class="section">
        ${renderSignatureTable({ includeAuthorizedOfficer: false })}
      </div>
    `;

    const partB = `
      <div class="doc-title">Certificate</div>
      <div class="doc-subtitle">Section 63(4)(c) - Part B</div>
      <div class="section">
        <div class="label">Part B (To be filled by the Expert)</div>
        <p>
          I, <strong>${d.examinerName}</strong>, employed by <strong>${d.employedBy}</strong> as a Digital Forensic Examiner,
          do hereby solemnly affirm
          that the produced electronic record/output was obtained from:
        </p>
      </div>
      <div class="section">
        <strong>Device Details:</strong>
        ${renderDeviceTable(devices)}
      </div>
      <div class="section">
        ${devices.map((device) => renderHashBlock(device)).join("")}
      </div>
      ${renderCopyTable("Master Copy Details", masterList)}
      ${renderCopyTable("Working Copy Details", workingList)}
      <div class="section">
        <div class="signature">Examiner Signature: ${d.examinerName}</div>
      </div>
    `;

    return `${partA}<!--PAGEBREAK-->${partB}`;
  },
  coc: () => {
    const d = getCommon();
    const masterList = getCopyList("master");
    const workingList = getCopyList("working");
    return `
      <div class="doc-title">Chain of Custody</div>
      <div class="doc-subtitle">Digital Evidence Handling</div>
      <div class="section">
        <div class="meta-grid">
          <div><span class="label">Case Information:</span> ${d.caseInformation}</div>
          <div><span class="label">Date of Search:</span> ${d.dateOfSearch}</div>
          <div><span class="label">Premise Address:</span> ${d.premiseAddress}</div>
          <div><span class="label">Description:</span> Digital data collected from various devices and stored in following storage device.</div>
        </div>
      </div>
      ${renderCopyTable("Master Copy Details", masterList)}
      ${renderCopyTable("Working Copy Details", workingList)}
      <table>
        <thead>
          <tr>
            <th>Reason/Action</th>
            <th>Received From</th>
            <th>Received By</th>
            <th>Date</th>
            <th>Time</th>
            <th>Sign of Receiver</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>HANDOVER</td>
            <td>${d.authorizedOfficer}</td>
            <td>${d.examinerName}</td>
            <td>${d.dateOfSearch}</td>
            <td></td>
            <td></td>
          </tr>
          <tr>
            <td>HANDOVER</td>
            <td>${d.examinerName}</td>
            <td>${d.authorizedOfficer}</td>
            <td>${d.conclusionDate}</td>
            <td></td>
            <td></td>
          </tr>
        </tbody>
      </table>
      <div class="section">
        ${renderSignatureTable()}
      </div>
    `;
  },
  evidence: () => {
    const d = getCommon();
    const devices = getDevices();
    const masterList = getCopyList("master");
    const workingList = getCopyList("working");
    if (!devices.length) {
      return `
        <div class="doc-title">Digital Evidence Collection Form</div>
        <div class="doc-subtitle">Hash & Imaging Record</div>
        <div class="section">No device added.</div>
      `;
    }

    return devices.map((device, idx) => `
      <section class="page">
        <div class="doc-box">
          <div class="doc-title">Digital Evidence Collection Form</div>
          <div class="doc-subtitle">Hash & Imaging Record</div>
          <div class="section">
            <strong>Hash Value report of Evidence:</strong> ${escapeHtml(device.name)}<br/>
            <strong>Name of Authorized Officer:</strong> ${d.authorizedOfficer}<br/>
            <strong>Date:</strong> ${d.dateOfSearch}<br/>
            <strong>Premise Address:</strong> ${d.premiseAddress}<br/>
            <strong>Examiner’s Name and Details:</strong> ${d.examinerName} (Digital Forensic Examiner)<br/>
            <strong>Case Information:</strong> ${d.caseInformation}<br/>
            <strong>Device #:</strong> ${idx + 1}
          </div>
          <div class="section">
            <strong>Evidence Information</strong>
            ${renderDeviceTable([device])}
          </div>
          <div class="section">
            ${renderHashBlock(device)}
          </div>
        ${renderCopyTable("Master Copy Details", masterList)}
        ${renderCopyTable("Working Copy Details", workingList)}
          <div class="section">
            ${renderSignatureTable()}
          </div>
        </div>
      </section>
    `).join("");
  },
  job: () => {
    const d = getCommon();
    const devices = getDevices();
    return `
      <div class="doc-title">Job Form</div>
      <div class="doc-subtitle">Digital Forensic Examination</div>
      <div class="section">
        <div class="meta-grid">
          <div><span class="label">Date of Submission:</span> ${d.conclusionDate}</div>
          <div><span class="label">Examiner Name:</span> ${d.examinerName}</div>
          <div><span class="label">Authorized Officer:</span> ${d.authorizedOfficer}</div>
          <div><span class="label">Case Information:</span> ${d.caseInformation}</div>
        </div>
      </div>
      <div class="section">
        <strong>Device List</strong><br/>
        ${renderDeviceTable(devices)}
      </div>
      <div class="section">
        <strong>Review & Analysis</strong><br/>
        Start Date: ${d.dateOfSearch}<br/>
        End Date: ${d.conclusionDate}
      </div>
      <div class="section">
        <table>
          <thead>
            <tr>
              <th>Authorized Officer</th>
              <th>Digital Forensic Examiner</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="height:50px"></td>
              <td style="height:50px"></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  },
  sticker: () => {
    const d = getCommon();
    const devices = getDevices();
    const deviceLines = devices.length
      ? devices.map((device, idx) => `<div>${idx + 1}. ${escapeHtml(device.name)}</div>`).join("")
      : "<div>No device added.</div>";
    return `
      <div class="doc-title">Working / Master Sticker</div>
      <div class="doc-subtitle">Device Imaging Identification</div>
      <div class="section sticker-box">
        <strong>WORKING COPY (Serial Number - ${d.stickerWorkingSerial})</strong><br/>
        In the case of ${d.caseInformation}. DOS: ${d.dateOfSearch}<br/>
        This storage medium contains the imaged data of the following devices:
        ${deviceLines}
        <div class="section">
          <div class="signature">W-1 ____________________</div>
          <div class="signature">Party - ${d.partyName}</div>
          <div class="signature">W-2 ____________________</div>
          <div class="signature">AO - ${d.authorizedOfficer}</div>
        </div>
      </div>
      <div class="section sticker-box">
        <strong>MASTER COPY (Serial Number - ${d.stickerMasterSerial})</strong><br/>
        In the case of ${d.caseInformation}. DOS: ${d.dateOfSearch}<br/>
        This storage medium contains the imaged data of the following devices:
        ${deviceLines}
        <div class="section">
          <div class="signature">W-1 ____________________</div>
          <div class="signature">Party - ${d.partyName}</div>
          <div class="signature">W-2 ____________________</div>
          <div class="signature">AO - ${d.authorizedOfficer}</div>
        </div>
      </div>
      <div class="section">
        ${renderSignatureTable()}
      </div>
    `;
  }
  ,
  custom: () => renderCustomDoc()
};

const renderPreview = () => {
  const active = document.querySelector(".tab.active")?.dataset.doc || "cert63";
  let html = templates[active] ? templates[active]() : "";
  html = html.replaceAll("<!--PAGEBREAK-->", "<div class=\"page-sep\">Page Break</div>");
  document.getElementById("preview").innerHTML = html;
  const caseEl = document.getElementById("metaCase");
  const dateEl = document.getElementById("metaDate");
  const devicesEl = document.getElementById("metaDevices");
  const customEl = document.getElementById("metaCustom");
  if (caseEl) caseEl.textContent = `Case: ${state.caseInformation || "—"}`;
  if (dateEl) dateEl.textContent = `DOS: ${state.dateOfSearch ? formatDate(state.dateOfSearch) : "—"}`;
  if (devicesEl) devicesEl.textContent = `Devices: ${getDevices().length}`;
  if (customEl) customEl.textContent = `Custom: ${state.customEnabled ? "On" : "Off"}`;
};

const bindInputs = () => {
  document.querySelectorAll("input, textarea").forEach((input) => {
    input.addEventListener("input", () => {
      state[input.id] = input.type === "checkbox" ? input.checked : input.value;
      renderPreview();
    });
  });
};

const bindTabs = () => {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      renderPreview();
    });
  });
};

const bindDeviceList = () => {
  const btn = document.getElementById("btn-add-device");
  const clearBtn = document.getElementById("btn-clear-device");
  btn.addEventListener("click", () => {
    const device = {
      type: state.deviceType,
      name: state.deviceName,
      make: state.deviceMake,
      size: state.deviceSize,
      model: state.deviceModelNo,
      serial: state.deviceSerialNo,
      imei1: state.deviceImei1,
      imei2: state.deviceImei2,
      includeMakeModel: state.includeMakeModel,
      includeSerialSize: state.includeSerialSize,
      includeImei: state.includeImei,
      md5: state.hashMd5,
      sha1: state.hashSha1,
      sha256: state.hashSha256,
      imagingSoftware: state.imagingSoftware,
      action: state.actionTaken
    };
    if (!device.name && !device.type) return;
    state.deviceList.push(device);
    renderDeviceList();
    renderPreview();
  });
  clearBtn.addEventListener("click", () => {
    [
      "deviceType",
      "deviceName",
      "deviceMake",
      "deviceSize",
      "deviceModelNo",
      "deviceSerialNo",
      "deviceImei1",
      "deviceImei2",
      "includeMakeModel",
      "includeSerialSize",
      "includeImei",
      "hashMd5",
      "hashSha1",
      "hashSha256",
      "imagingSoftware",
      "actionTaken"
    ].forEach((id) => {
      const input = document.getElementById(id);
      if (input) {
        if (input.type === "checkbox") {
          input.checked = true;
          state[id] = true;
        } else {
          input.value = "";
          state[id] = "";
        }
      }
    });
    renderPreview();
  });
};

const bindCopyControls = () => {
  const addMaster = document.getElementById("btn-add-master");
  const addWorking = document.getElementById("btn-add-working");
  const clearCopies = document.getElementById("btn-clear-copies");

  addMaster?.addEventListener("click", () => {
    const item = {
      make: state.masterCopyMake,
      model: state.masterCopyModel,
      serial: state.masterCopySerial,
      size: state.masterCopySize
    };
    if (!item.make && !item.model && !item.serial && !item.size) return;
    state.masterCopies.push(item);
    renderCopyLists();
    renderPreview();
  });

  addWorking?.addEventListener("click", () => {
    const item = {
      make: state.workingCopyMake,
      model: state.workingCopyModel,
      serial: state.workingCopySerial,
      size: state.workingCopySize
    };
    if (!item.make && !item.model && !item.serial && !item.size) return;
    state.workingCopies.push(item);
    renderCopyLists();
    renderPreview();
  });

  clearCopies?.addEventListener("click", () => {
    [
      "masterCopyMake",
      "masterCopyModel",
      "masterCopySerial",
      "masterCopySize",
      "workingCopyMake",
      "workingCopyModel",
      "workingCopySerial",
      "workingCopySize"
    ].forEach((id) => {
      const input = document.getElementById(id);
      if (input) {
        input.value = "";
        state[id] = "";
      }
    });
  });
};

const getPagesFromContent = (content) => {
  if (content.includes("class=\"page\"")) {
    const matches = content.match(/<section class="page">[\s\S]*?<\/section>/g);
    if (matches && matches.length) {
      return matches.map((section) => section.replace("<section class=\"page\">", "").replace("</section>", ""));
    }
  }
  if (content.includes("<!--PAGEBREAK-->")) {
    return content.split("<!--PAGEBREAK-->");
  }
  return [content];
};

const wrapPagesWithHeaderFooter = (pages) => {
  const d = getCommon();
  const total = pages.length;
  return pages.map((content, index) => `
    <section class="page">
      <div class="page-header">
        <div>Case: ${d.caseInformation || "N/A"}</div>
        <div>Date: ${d.dateOfSearch || "N/A"}</div>
      </div>
      ${content}
      <div class="page-footer">
        <div>Prepared by: ${d.examinerName || "N/A"}</div>
        <div>Page ${index + 1} of ${total}</div>
      </div>
    </section>
  `).join("");
};

const downloadCurrent = () => {
  const active = document.querySelector(".tab.active")?.dataset.doc || "cert63";
  const html = templates[active] ? templates[active]() : "";
  const docName = {
    cert63: "63A-certificate",
    coc: "chain-of-custody",
    evidence: "evidence-collection",
    job: "job-form",
    sticker: "working-master-sticker"
  }[active];

  const pages = getPagesFromContent(html);
  const body = wrapPagesWithHeaderFooter(pages);

  const blob = new Blob([
    `<!doctype html><html><head><meta charset='utf-8'>`,
    `<title>${docName}</title>`,
    getPrintStyles(),
    `</head><body>${body}</body></html>`
  ], { type: "text/html" });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${docName}.html`;
  a.click();
  URL.revokeObjectURL(url);
};

const getPrintStyles = () => `
  <style>
    body{font-family:Georgia,serif;line-height:1.5;padding:32px;color:#111;}
    h1,h2,h3{margin:0 0 8px;text-align:center;}
    .doc-title{font-size:20px;font-weight:700;text-align:center;margin:0 0 6px;}
    .doc-subtitle{font-size:12px;text-align:center;margin:0 0 16px;letter-spacing:0.08em;text-transform:uppercase;}
    .section{margin-bottom:16px;}
    table{width:100%;border-collapse:collapse;margin-top:8px;}
    th,td{border:1px solid #333;padding:6px;font-size:12px;text-align:left;}
    .meta-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 18px;font-size:13px;}
    .meta-grid div{padding:6px 8px;border:1px solid #999;}
    .signature{border-top:1px solid #333;padding-top:8px;}
    .hash-block{border:2px solid #333;background:#f8f1e8;padding:10px 12px;border-radius:8px;margin-top:10px;}
    .hash-title{font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:6px;}
    .hash-row{display:grid;grid-template-columns:100px 1fr;gap:8px;font-size:11px;}
    .hash-row span{font-weight:600;}
    .doc-box{border:2px solid #333;padding:14px;border-radius:10px;background:#fffef9;}
    .sticker-box{border:2px dashed #333;padding:12px;border-radius:10px;background:#fff8ee;margin-bottom:12px;}
    .page{page-break-after:always;position:relative;min-height:95vh;padding-top:24px;padding-bottom:28px;}
    .page:last-child{page-break-after:auto;}
    .page-header,.page-footer{font-size:11px;color:#333;display:flex;justify-content:space-between;}
    .page-header{position:absolute;top:0;left:0;right:0;padding-bottom:6px;border-bottom:1px solid #333;}
    .page-footer{position:absolute;bottom:0;left:0;right:0;padding-top:6px;border-top:1px solid #333;}
  </style>
`;

const exportPrintPack = () => {
  const docs = [
    "cert63",
    "coc",
    "evidence",
    "job",
    "sticker",
    state.customEnabled ? "custom" : null
  ].filter(Boolean);
  const pages = docs.flatMap((key) => {
    const content = templates[key] ? templates[key]() : "";
    return getPagesFromContent(content);
  });
  const body = wrapPagesWithHeaderFooter(pages);
  const html = `<!doctype html><html><head><meta charset="utf-8">${getPrintStyles()}</head><body>${body}</body></html>`;
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
  win.print();
};

const exportPrintCurrent = () => {
  const active = document.querySelector(".tab.active")?.dataset.doc || "cert63";
  const html = templates[active] ? templates[active]() : "";
  const pages = getPagesFromContent(html);
  const body = wrapPagesWithHeaderFooter(pages);
  const htmlDoc = `<!doctype html><html><head><meta charset="utf-8">${getPrintStyles()}</head><body>${body}</body></html>`;
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(htmlDoc);
  win.document.close();
  win.focus();
  win.print();
};

const getWordStyles = () => `
  <style>
    body{font-family:Georgia,serif;line-height:1.5;padding:24px;color:#111;}
    h1,h2,h3{margin:0 0 8px;text-align:center;}
    .doc-title{font-size:18px;font-weight:700;text-align:center;margin:0 0 6px;}
    .doc-subtitle{font-size:11px;text-align:center;margin:0 0 14px;letter-spacing:0.08em;text-transform:uppercase;}
    .section{margin-bottom:14px;}
    table{width:100%;border-collapse:collapse;margin-top:8px;}
    th,td{border:1px solid #333;padding:6px;font-size:11px;text-align:left;}
    .meta-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 18px;font-size:12px;}
    .meta-grid div{padding:6px 8px;border:1px solid #999;}
    .signature{border-top:1px solid #333;padding-top:8px;margin-top:10px;}
    .hash-block{border:2px solid #333;background:#f8f1e8;padding:10px 12px;border-radius:8px;margin-top:10px;}
    .hash-title{font-size:10px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:6px;}
    .hash-row{display:grid;grid-template-columns:100px 1fr;gap:8px;font-size:10px;}
    .hash-row span{font-weight:600;}
    .doc-box{border:2px solid #333;padding:14px;border-radius:10px;background:#fffef9;}
    .page{page-break-after:always;}
    .page:last-child{page-break-after:auto;}
  </style>
`;

const exportWordPack = () => {
  const docs = [
    "cert63",
    "coc",
    "evidence",
    "job",
    "sticker",
    state.customEnabled ? "custom" : null
  ].filter(Boolean);
  const pages = docs.flatMap((key) => {
    const content = templates[key] ? templates[key]() : "";
    return getPagesFromContent(content);
  });
  const body = pages.map((content) => `
    <section class="page">
      ${content}
    </section>
  `).join("");
  const html = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office'
          xmlns:w='urn:schemas-microsoft-com:office:word'
          xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>Forensic Documents</title>
        ${getWordStyles()}
      </head>
      <body>${body}</body>
    </html>
  `;
  const blob = new Blob([html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "forensic-documents.doc";
  a.click();
  URL.revokeObjectURL(url);
};

const exportDocxDesktop = async () => {
  if (!window.__TAURI__) {
    alert("DOCX export is available in the desktop app.");
    return;
  }
  try {
    const { invoke, dialog } = window.__TAURI__;
    const savePath = await dialog.save({ filters: [{ name: "Word Document", extensions: ["docx"] }] });
    if (!savePath) return;
    const payload = { ...state };
    const result = await invoke("generate_docx", { payloadJson: JSON.stringify(payload, null, 2), outputPath: savePath });
    const parsed = JSON.parse(result);
    if (parsed.error) throw new Error(parsed.error);
    alert("DOCX saved.");
  } catch (err) {
    alert(`DOCX export failed: ${err.message}`);
  }
};

const exportWordCurrent = () => {
  const active = document.querySelector(".tab.active")?.dataset.doc || "cert63";
  const html = templates[active] ? templates[active]() : "";
  const docName = {
    cert63: "63A-certificate",
    coc: "chain-of-custody",
    evidence: "evidence-collection",
    job: "job-form",
    sticker: "working-master-sticker"
  }[active];
  const pages = getPagesFromContent(html);
  const body = pages.map((content) => `<section class="page">${content}</section>`).join("");
  const htmlDoc = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office'
          xmlns:w='urn:schemas-microsoft-com:office:word'
          xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${docName}</title>
        ${getWordStyles()}
      </head>
      <body>
        ${body}
      </body>
    </html>
  `;
  const blob = new Blob([htmlDoc], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${docName}.doc`;
  a.click();
  URL.revokeObjectURL(url);
};

const downloadJson = () => {
  const payload = {
    ...state,
    dateOfSearch: state.dateOfSearch,
    conclusionDate: state.conclusionDate
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "case-data.json";
  a.click();
  URL.revokeObjectURL(url);
};

const printCurrent = () => {
  window.print();
};

bindInputs();
bindTabs();
bindDeviceList();
bindCopyControls();
renderDeviceList();
renderCopyLists();
renderPreview();

const customDocxInput = document.getElementById("customDocx");
customDocxInput?.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  state.customDocxName = file ? file.name : "";
  const statusEl = document.getElementById("customDocxStatus");
  if (statusEl) statusEl.textContent = state.customDocxName || "No file selected.";
  renderPreview();
});

const isTauri = () => Boolean(window.__TAURI__);

const extractDocxDesktop = async (path) => {
  const { invoke } = window.__TAURI__;
  const response = await invoke("extract_docx", { path });
  const data = JSON.parse(response);
  if (data.error) throw new Error(data.error);
  return data;
};

const bindDesktopDocx = () => {
  const btn = document.getElementById("btn-custom-docx");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    if (!isTauri()) {
      alert("Desktop DOCX selection is available only in the desktop app.");
      return;
    }
    try {
      const { dialog } = window.__TAURI__;
      const path = await dialog.open({ multiple: false, filters: [{ name: "DOCX", extensions: ["docx"] }] });
      if (!path) return;
      const data = await extractDocxDesktop(path);
      state.customDocxName = String(path).split("/").pop() || "selected.docx";
      const statusEl = document.getElementById("customDocxStatus");
      if (statusEl) statusEl.textContent = state.customDocxName;
      // Auto-fill template with extracted text for quick editing
      if (data.text) {
        state.customTemplate = data.text;
        const textarea = document.getElementById("customTemplate");
        if (textarea) textarea.value = data.text;
      }
      renderPreview();
    } catch (err) {
      alert(`DOCX extraction failed: ${err.message}`);
    }
  });
};

bindDesktopDocx();

document.getElementById("btn-export").addEventListener("click", downloadCurrent);
document.getElementById("btn-print").addEventListener("click", printCurrent);
document.getElementById("btn-pack").addEventListener("click", exportPrintPack);
document.getElementById("btn-pdf-one").addEventListener("click", exportPrintCurrent);
document.getElementById("btn-pdf-all").addEventListener("click", exportPrintPack);
document.getElementById("btn-docx-one").addEventListener("click", exportDocxDesktop);
document.getElementById("btn-docx-all").addEventListener("click", exportDocxDesktop);
const jsonBtn = document.createElement("button");
jsonBtn.className = "btn btn-ghost";
jsonBtn.textContent = "Export Case JSON";
jsonBtn.addEventListener("click", downloadJson);
document.querySelector(".top-actions").appendChild(jsonBtn);
