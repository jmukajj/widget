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

      // Load external library (docx) to generate the .docx file
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docx/6.1.0/docx.min.js')
        .then(() => {
          console.log("docx library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading docx library:", error);
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

    // This function will be called when the user selects a row in the SAC table
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

    // Function to generate a Word document using docx
    generateWordDocument() {
      console.log('Generating document with Post Data:', this._postData);

      if (!this._postData || Object.keys(this._postData).length === 0) {
        alert("No data to generate document");
        return;
      }

      // Create a new document
      const { Document, Packer, Paragraph, TextRun } = window.docx;
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
