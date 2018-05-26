let pptxName;
const xlsxPromise = new Promise(resolve => {
  function handleFile(e) {
    const files = e.target.files,
      f = files[0];
    const reader = new FileReader();
    reader.onload = e =>
      resolve(XLSX.read(new Uint8Array(e.target.result), { type: "array" }));
    reader.readAsArrayBuffer(f);
  }

  document.getElementById("xlsx").addEventListener("change", handleFile, false);
});

const pptxPromise = new Promise(resolve => {
  function handleFile(e) {
    const [f] = e.target.files;
    pptxName = f.name;
    const reader = new FileReader();
    reader.onload = e =>
      JSZip.loadAsync(new Uint8Array(e.target.result)).then(zip =>
        resolve(zip)
      );
    reader.readAsArrayBuffer(f);
  }

  document.getElementById("pptx").addEventListener("change", handleFile, false);
});

Promise.all([xlsxPromise, pptxPromise]).then(([workbook, pptxZip]) => {
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  const promises = [];
  pptxZip.forEach(async (path, file) => {
    if (path.substr(-4) !== ".xml") return;
    promises.push(
      (async () => {
        const content = await file.async("string");
        const result = content.replace(/{{\s*[\w\.]+\s*}}/g, cellReference => {
          const cellAddress = cellReference.match(/[\w\.]+/)[0];
          const desiredCell = worksheet[cellAddress];
          return desiredCell ? desiredCell.v : "NOT_FOUND";
        });
        pptxZip.file(path, result);
      })()
    );
  });
  Promise.all(promises)
    .then(() => pptxZip.generateAsync({ type: "blob" }))
    .then(blob => saveAs(blob, pptxName));
});
