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

  reader.onload = function(e) {
    let names = [];

    if (file.name.endsWith(".csv")) {
      // CSV parsing
      names = e.target.result
        .split(/[\n,]+/)
        .map(n => n.trim())
        .filter(n => n.length > 0);
    } else if (file.name.endsWith(".xlsx")) {
      // XLSX parsing
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      sheetData.forEach(row => {
        row.forEach(cell => {
          if (cell) names.push(cell.toString().trim());
        });
      });
    }

    // Fill textarea so user sees the loaded names
    document.getElementById("names").value = names.join(", ");
  };

  if (file.name.endsWith(".csv")) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

// Generate groups
function generateGroups() {
  const rawInput = document.getElementById("names").value;
  const names = rawInput
    .split(/[\n,]+/)
    .map(n => n.trim())
    .filter(n => n.length > 0);

  const groupSize = parseInt(document.getElementById("groupSize").value);
  const output = document.getElementById("output");
  const summary = document.getElementById("summary");
  output.innerHTML = "";
  summary.innerHTML = "";
  generatedGroups = [];

  if (names.length === 0 || !groupSize || groupSize <= 0) {
    alert("Please enter valid names and group size.");
    return;
  }

  shuffle(names);

  let groupNum = 1;
  for (let i = 0; i < names.length; i += groupSize) {
    const group = names.slice(i, i + groupSize);
    generatedGroups.push({ number: groupNum, members: group });

    const div = document.createElement("div");
    div.className = "group";
    div.innerHTML = `<strong>Group ${groupNum}:</strong><ul>${group.map(name => `<li>${name}</li>`).join("")}</ul>`;
    output.appendChild(div);
    groupNum++;
  }

  summary.textContent = `Total Students: ${names.length} | Total Groups: ${generatedGroups.length}`;
}

// Save DOCX (formatted with bullet list + spacing)
async function saveDocx() {
  if (generatedGroups.length === 0) {
    alert("Generate groups first!");
    return;
  }

  const { Document, Packer, Paragraph, TextRun } = docx;

  const paragraphs = [
    new Paragraph({
      text: "Groups",
      heading: "Heading1",
      spacing: { after: 300 },
    }),
  ];

  generatedGroups.forEach(g => {
    // Add group title
    paragraphs.push(
      new Paragraph({
        text: `Group ${g.number}`,
        heading: "Heading2",
        spacing: { after: 100 },
      })
    );

    // Add each member as a bulleted list
    g.members.forEach(name => {
      paragraphs.push(
        new Paragraph({
          text: name,
          bullet: { level: 0 },
        })
      );
    });

    // Add space between groups
    paragraphs.push(new Paragraph({ text: "", spacing: { after: 300 } }));
  });

  const doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });

  const blob = await Packer.toBlob(doc);
  const filename = prompt("Enter file name:", "student_groups") || "student_groups";

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${filename}.docx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}
