(function () {
  const template = document.createElement('template');
  template.innerHTML = `
       <style>
        #root {
            width: 300px;
            justify-content: flex-start;
            align-items: flex-start;
            height: 100vh;
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
            display: block;
        }
       </style>

       <div id="root">
          <div class="link-container" id="links-container">
            <p><a id="link_href" href="#" target="_blank">Download Word Document</a></p>
          </div>
       </div>
  `;

  class Main extends HTMLElement {
    constructor() {
      super();
      console.log('Widget initialized');
      this._shadowRoot = this.attachShadow({ mode: 'open' });
      this._shadowRoot.appendChild(template.content.cloneNode(true));
      this._postData = {};

      this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
        this.generateAndDownloadDocument();
      });

      // Load necessary libraries
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.2/mammoth.browser.min.js')
        .then(() => {
          console.log("mammoth.js library loaded successfully!");
          return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pizzip/3.1.1/pizzip.min.js');
        })
        .then(() => {
          console.log("PizZip library loaded successfully!");
          return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.21.2/docxtemplater.min.js');
        })
        .then(() => {
          console.log("docxtemplater library loaded successfully!");
          return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js');
        })
        .then(() => {
          console.log("FileSaver.js library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading a library:", error);
        });
    }

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

    sendPostData(selectedRowData) {
      console.log("Received selected row data: ", selectedRowData);
      if (!selectedRowData || Object.keys(selectedRowData).length === 0) {
        console.error("No data provided in selected row", selectedRowData);
        alert("No data to generate document");
        return;
      }
      this._postData = selectedRowData;
      console.log("Post Data after population: ", this._postData);
    }

    async generateAndDownloadDocument() {
      const data = this._postData;
      if (!data || Object.keys(data).length === 0) {
        alert("No data to generate document");
        return;
      }

      try {
        // Fetch the Word document template using a CORS proxy
        const templateBlob = await this.fetchWordTemplate();

        // Convert DOCX to HTML using Mammoth.js
        mammoth.convertToHtml({ arrayBuffer: templateBlob })
          .then(result => {
            let html = result.value; // HTML representation of the document

            // Replace placeholders in the HTML content
            Object.keys(data).forEach((key) => {
              const placeholder = `{{${key}}}`;
              html = html.replace(new RegExp(placeholder, 'g'), data[key]);
            });

            // Convert the updated HTML back into a Word document
            this.createWordDocumentFromHtml(html);
          })
          .catch(err => console.error('Error converting document:', err));

      } catch (error) {
        console.error("Error generating document:", error);
      }
    }

    async fetchWordTemplate() {
      try {
        const proxyUrl = 'https://corsproxy.io/?'; // Use a CORS proxy URL
        const templateUrl = 'https://github.com/jmukajj/widget/raw/refs/heads/main/template.docx';
        const response = await fetch(proxyUrl + encodeURIComponent(templateUrl), {
          headers: {
            'Origin': 'https://itsvac-test.eu20.hcs.cloud.sap'
          }
        });
        if (!response.ok) {
          throw new Error(`Failed to fetch the Word template: ${response.statusText}`);
        }
        return await response.arrayBuffer();
      } catch (error) {
        console.error('Error fetching template:', error);
        throw error;
      }
    }

    createWordDocumentFromHtml(htmlContent) {
      try {
        // Create a new PizZip instance
        const zip = new PizZip();

        // Create a new docxtemplater document with the updated HTML content
        const doc = new window.docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        // Inject the updated content into the document
        doc.loadZip(zip);
        doc.setData({ content: htmlContent });

        // Render the document
        doc.render();

        // Generate a new Word document and download it
        const out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        saveAs(out, 'populated_document.docx');
      } catch (error) {
        console.error("Error creating Word document from HTML:", error);
      }
    }
  }

  customElements.define('com-sap-sac-jm', Main);

})();
