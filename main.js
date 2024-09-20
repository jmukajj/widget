(function () {
  const template = document.createElement('template');
  template.innerHTML = `
       <style>
        #root {
            width: 300px;
            justify-content: flex-start;
            align-items: flex-start;
            height: 100vh; /* Ensure the div takes the full height of the viewport */
        }
        .link-container {
            position: relative;
            display: flex;
            flex-direction: column;
            justify-content: left;
            left: 7px;
            align-items: left;
            padding: 20px;
            border: 0.5px solid black;
            background-color: #FCFCFC;
            box-shadow: 0 4px 8px rgba(0, 0, 0, .3);
        }
        .link-container::before {
            content: '';
            position: absolute;
            left: -5.5px;
            top: 50%;
            transform: translateY(-50%) rotate(135deg);
            border: solid black;
            border-width: 0 0.5px 0.5px 0;
            display: inline-block;
            padding: 5px;
            background-color: #FCFCFC;
        }
        .link {
            text-decoration: none;
            color: #5E97C4;
            font-family: Arial, sans-serif;
            margin-bottom: 10px;
            display: block; /* Ensures each link starts on a new line */
        }
       </style>

       <div id="root">
          <div class="link-container" id="links-container">
            <p><a id="link_href" href="#" target="_blank">Download Word Document</a></p>
          </div>
       </div>
  `;

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

      // Load docx library dynamically
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docx/7.1.0/docx.umd.min.js')
        .then(() => {
          if (window.docx) {
            console.log("docx library loaded successfully!", window.docx);
          } else {
            console.error("docx library failed to load.");
          }
        })
        .catch((error) => {
          console.error("Error loading docx library:", error);
        });

      // Load FileSaver.js library dynamically for triggering file downloads
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
        .then(() => {
          console.log("FileSaver.js library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading FileSaver.js library:", error);
        });
    }

    // Function to dynamically load external scripts
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

    // This function will be called when the user selects a row in the SAC table, to pass selected data from the SAC table into the widget.
    sendPostData(selectedRowData) {
      console.log("Received selected row data: ", selectedRowData);

      // Ensure selectedRowData contains valid data
      if (!selectedRowData || Object.keys(selectedRowData).length === 0) {
        console.error("No data provided in selected row", selectedRowData);
        alert("No data to generate document");
        return;
      }

      this._postData = selectedRowData; // Store all selected data for the document generation

      console.log("Post Data after population: ", this._postData);
    }

    // Fallback approach: Function to generate a basic Word document using Blob
    generateWordDocumentFallback() {
      console.log('Generating document with Post Data:', this._postData);

      if (!this._postData || Object.keys(this._postData).length === 0) {
        alert("No data to generate document");
        return;
      }

      // Construct simple Word document XML content
      const docContent = `
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

      // Convert XML content to a Blob with correct MIME type
      const blob = new Blob([docContent], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

      // Trigger download using FileSaver.js
      saveAs(blob, "AntragDocument.docx");
    }

    // Function to generate a Word document using docx (if the library is loaded)
    generateWordDocument() {
      if (!window.docx) {
        console.error("docx library is not loaded, falling back to XML document.");
        this.generateWordDocumentFallback();  // Fallback to manual method if docx is not available
        return;
      }

      console.log('Generating document with Post Data:', this._postData);

      if (!this._postData || Object.keys(this._postData).length === 0) {
        alert("No data to generate document");
        return;
      }

      const { Document, Packer, Paragraph, TextRun } = window.docx;
      
      // Create a new document
      const doc = new Document();

      // Add a title paragraph
      doc.addSection({
        children: [
          new Paragraph({
            children: [new TextRun({ text: 'Antrag Document', bold: true, size: 32 })],
          }),
          new Paragraph({
            children: [new TextRun({ text: '------------------------------', bold: true, size: 24 })],
          }),
          ...Object.keys(this._postData).map(key => 
            new Paragraph({
              children: [new TextRun({ text: `${key}: ${this._postData[key]}`, size: 24 })],
            })
          )
        ],
      });

      // Generate the Word document and trigger the download
      Packer.toBlob(doc).then(blob => {
        saveAs(blob, "AntragDocument.docx");
      }).catch(error => {
        console.error("Error generating the document:", error);
        alert("Failed to generate the document.");
      });
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
