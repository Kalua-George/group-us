let generatedGroups = [];

// Shuffle array helper
function shuffle(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

// Handle CSV/XLSX file
function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (e) {
    let names = [];

    if (file.name.endsWith(".csv")) {
      names = e.target.result
        .split(/[\n,]+/)
        .map((n) => n.trim())
        .filter((n) => n.length > 0);
    } else if (file.name.endsWith(".xlsx")) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      sheetData.forEach((row) => {
        row.forEach((cell) => {
          if (cell) names.push(cell.toString().trim());
        });
      });
    }

    document.getElementById("names").value = names.join(", ");
  };

  if (file.name.endsWith(".csv")) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

// Animated mobile-safe modal
function showModal(message, type = "info", withInput = false) {
  return new Promise((resolve) => {
    let modal = document.getElementById("modal");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal";
      modal.innerHTML = `
        <div id="modal-overlay" style="
          position: fixed; top: 0; left: 0;
          width: 100vw; height: 100vh;
          background: rgba(0,0,0,0.4);
          display: flex; justify-content: center; align-items: center;
          z-index: 9999; opacity: 0; transition: opacity 0.3s ease;">
          <div id="modal-box" style="
            background: #fff; padding: 25px; border-radius: 10px;
            width: 80%; max-width: 350px; text-align: center;
            box-shadow: 0 4px 10px rgba(0,0,0,0.25);
            transform: translateY(20px);
            transition: transform 0.3s ease, box-shadow 0.3s ease;">
            <p id="modal-text" style="margin-bottom: 10px; font-size: 16px;"></p>
            <input type="text" id="modal-input" placeholder="Enter file name"
              style="display:none; width: 90%; padding: 8px; border: 1px solid #ccc; border-radius: 6px;">
            <div style="margin-top:15px; display:flex; justify-content:space-around;">
              <button id="modal-ok" style="background:#007bff;color:white;border:none;padding:8px 15px;border-radius:6px;">OK</button>
              <button id="modal-cancel" style="background:#ccc;border:none;padding:8px 15px;border-radius:6px;display:none;">Cancel</button>
            </div>
          </div>
        </div>`;
      document.body.appendChild(modal);
    }

    const overlay = modal.querySelector("#modal-overlay");
    const box = modal.querySelector("#modal-box");
    const text = modal.querySelector("#modal-text");
    const input = modal.querySelector("#modal-input");
    const ok = modal.querySelector("#modal-ok");
    const cancel = modal.querySelector("#modal-cancel");

    // Set message color by type
    text.textContent = message;
    if (type === "success") text.style.color = "#28a745";
    else if (type === "error") text.style.color = "#dc3545";
    else text.style.color = "#000";

    input.style.display = withInput ? "block" : "none";
    cancel.style.display = withInput ? "inline-block" : "none";
    input.value = "";

    overlay.style.display = "flex";
    requestAnimationFrame(() => {
      overlay.style.opacity = "1";
      box.style.transform = "translateY(0)";
    });

    ok.onclick = () => {
      overlay.style.opacity = "0";
      box.style.transform = "translateY(20px)";
      setTimeout(() => {
        overlay.style.display = "none";
        resolve(withInput ? input.value.trim() || null : true);
      }, 250);
    };

    cancel.onclick = () => {
      overlay.style.opacity = "0";
      box.style.transform = "translateY(20px)";
      setTimeout(() => {
        overlay.style.display = "none";
        resolve(null);
      }, 250);
    };
  });
}

// Generate groups
async function generateGroups() {
  const rawInput = document.getElementById("names").value;
  const names = rawInput
    .split(/[\n,]+/)
    .map((n) => n.trim())
    .filter((n) => n.length > 0);

  const groupSize = parseInt(document.getElementById("groupSize").value);
  const output = document.getElementById("output");
  const summary = document.getElementById("summary");
  output.innerHTML = "";
  summary.innerHTML = "";
  generatedGroups = [];

  if (names.length === 0 || !groupSize || groupSize <= 0) {
    await showModal("Please enter valid names and group size.", "error");
    return;
  }

  shuffle(names);

  let groupNum = 1;
  for (let i = 0; i < names.length; i += groupSize) {
    const group = names.slice(i, i + groupSize);
    generatedGroups.push({ number: groupNum, members: group });

    const div = document.createElement("div");
    div.className = "group";
    div.innerHTML = `<strong>Group ${groupNum}:</strong><ul>${group
      .map((name) => `<li>${name}</li>`)
      .join("")}</ul>`;
    output.appendChild(div);
    groupNum++;
  }

  summary.textContent = `Total Students: ${names.length} | Total Groups: ${generatedGroups.length}`;
  await showModal("Groups generated successfully!", "success");
}

// Save DOCX (formatted bullet list + spacing)
async function saveDocx() {
  if (generatedGroups.length === 0) {
    await showModal("Generate groups first!", "error");
    return;
  }

  const { Document, Packer, Paragraph } = docx;

  const paragraphs = [
    new Paragraph({
      text: "Groups",
      heading: "Heading1",
      spacing: { after: 300 },
    }),
  ];

  generatedGroups.forEach((g) => {
    paragraphs.push(
      new Paragraph({
        text: `Group ${g.number}`,
        heading: "Heading2",
        spacing: { after: 100 },
      })
    );

    g.members.forEach((name) => {
      paragraphs.push(new Paragraph({ text: name, bullet: { level: 0 } }));
    });

    paragraphs.push(new Paragraph({ text: "", spacing: { after: 300 } }));
  });

  const doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });

  const blob = await Packer.toBlob(doc);
  const filename = await showModal("Enter file name:", "info", true);
  if (!filename) return;

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${filename}.docx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);

  await showModal("File saved successfully!", "success");
}
