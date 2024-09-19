class Main extends HTMLElement {
  constructor () {
    super();
    console.log('Widget initialized');
    this._shadowRoot = this.attachShadow({ mode: 'open' });
    this._shadowRoot.appendChild(template.content.cloneNode(true));
    this.Response = null;

    // Initialize empty postData for selected data
    this._postData = {};

    // Attach event listener for download link
    this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
      this.generateWordDocument();
    });

    // Load necessary libraries sequentially
    this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
      .then(() => {
        console.log("FileSaver.js loaded successfully!");
        return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pizzip/3.1.1/pizzip.min.js');
      })
      .then(() => {
        console.log("PizZip loaded successfully!");
        return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.22.0/docxtemplater.js');
      })
      .then(() => {
        console.log("docxtemplater loaded successfully!");
      })
      .catch((error) => {
        console.error("Error loading script:", error);
      });
  }

  // Function to dynamically load external script
  loadScript(url) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = url;
      script.async = false;
      script.onload = () => resolve();
      script.onerror = () => reject(`Failed to load script: ${url}`);
      document.head.appendChild(script);
    });
  }

  // Function to generate a real Word document using docxtemplater and PizZip
  generateWordDocument() {
    if (typeof PizZip === 'undefined' || typeof window.docxtemplater === 'undefined') {
      console.error("Required libraries are not loaded yet.");
      alert("The document generation libraries have not been loaded yet. Please try again.");
      return;
    }

    console.log('Generating document with Post Data:', this._postData);

    if (!this._postData || Object.keys(this._postData).length === 0) {
      alert("No data to generate document");
      return;
    }

    // Create an empty Word document using PizZip
    const zip = new PizZip();
    const doc = new window.docxtemplater(zip);

    // Create the content for the Word document
    const content = `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Antrag Document</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r>
              <w:t>------------------------------</w:t>
            </w:r>
          </w:p>
          ${Object.keys(this._postData).map(key => `
            <w:p>
              <w:r>
                <w:t>${key}: ${this._postData[key]}</w:t>
              </w:r>
            </w:p>
          `).join('')}
        </w:body>
      </w:document>
    `;

    doc.loadZip(zip);
    doc.setData(this._postData);

    try {
      doc.render(); // Render the document
      const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      // Trigger the download using FileSaver.js
      saveAs(out, "AntragDocument.docx");

    } catch (error) {
      console.error("Error rendering the document:", error);
      alert("Failed to generate the document.");
    }
  }
}

customElements.define('com-sap-sac-jm', Main);
