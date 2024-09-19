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
            <p><a id="link_href" href="#" target="_blank">Download Text Document</a></p>
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

      // Load FileSaver.js to save the blob (a small external script for downloads)
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
        .then(() => {
          console.log("FileSaver.js loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading FileSaver.js:", error);
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

    // This function will be called when the user selects a row in the SAC table
    sendPostData(selectedRowData) {
      console.log("Received selected row data: ", selectedRowData);

      // Check if there is valid data in selectedRowData
      if (!selectedRowData || Object.keys(selectedRowData).length === 0) {
        console.error("No data provided in selected row", selectedRowData);
        alert("No data to generate document");
        return;
      }

      // Create an object to store the column names and their values
      const rowDataWithValues = {};
      for (let key in selectedRowData) {
        if (selectedRowData.hasOwnProperty(key)) {
          // If the key holds an object with a `value` property, use the `value`
          if (selectedRowData[key].value) {
            rowDataWithValues[key] = selectedRowData[key].value;
          } else {
            // Otherwise, assume the raw value is needed
            rowDataWithValues[key] = selectedRowData[key];
          }
        }
      }

      this._postData = rowDataWithValues; // Store the human-readable data for document generation

      console.log("Post Data with column names and values: ", this._postData);
    }

    // Function to generate a real Word document using docxtemplater and PizZip
    generateWordDocument() {
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

      // Populate the document with the data
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
})();

