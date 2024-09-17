(function () {
  const template = document.createElement('template');
  template.innerHTML = `
        <style>
        </style>
        
        <div id="root" style="width: 100%; height: 100%;">
          <button id="generateWordBtn">Generate Word Document</button>
        </div>
      `

  class Main extends HTMLElement {
    constructor () {
      super();
      this._shadowRoot = this.attachShadow({ mode: 'open' });
      this._shadowRoot.appendChild(template.content.cloneNode(true));

      // Object to store the selected Antrag data
      this.selectedAntrag = {
        createdBy: "",
        createdOn: "",
        totalAmount: 0
      };

      // Adding event listener for the button to generate Word document
      this._shadowRoot.getElementById('generateWordBtn').addEventListener('click', () => {
        this.generateWordDocument();
      });
    }

    // Method to accept data (Antrag information) from the SAC model
    setAntragData(antragData) {
      this.selectedAntrag = antragData;
    }

    // Method to generate Word document using docxtemplater and JSZip
    async generateWordDocument() {
      const { createdBy, createdOn, totalAmount } = this.selectedAntrag;

      // Load the JSZip and docxtemplater libraries dynamically
      const JSZip = await import('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.2.2/jszip.min.js');
      const Docxtemplater = await import('https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.19.6/docxtemplater.min.js');

      // Base template for the Word document
      const template = `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>Created by: ${createdBy}</w:t>
              </w:r>
            </w:p>
            <w:p>
              <w:r>
                <w:t>Created on: ${createdOn}</w:t>
              </w:r>
            </w:p>
            <w:p>
              <w:r>
                <w:t>Total Amount: ${totalAmount}</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;

      // Use JSZip to generate a Word document
      const zip = new JSZip();
      zip.file("word/document.xml", template);

      // Finalize the document and trigger download
      zip.generateAsync({ type: "blob" }).then(function (blob) {
        saveAs(blob, "Antrag_Document.docx");
      });
    }
  }

  customElements.define('com-sap-sac-aa', Main);
})();
