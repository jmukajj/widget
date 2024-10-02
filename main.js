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

      // Load docx library
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docx/7.0.1/docx.min.js')
        .then(() => {
          console.log("docx library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading the library:", error);
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
        // Load the existing document from a URL or file
        const response = await fetch('https://github.com/jmukajj/widget/raw/refs/heads/main/template.docx');
        if (!response.ok) {
          throw new Error(`Failed to fetch the Word template: ${response.statusText}`);
        }
        const arrayBuffer = await response.arrayBuffer();

        const zip = new JSZip();
        const content = await zip.loadAsync(arrayBuffer);
        const doc = new docx.Document(content);

        // Replace placeholders in the document with the given data
        doc.getParagraphs().forEach((paragraph) => {
          paragraph.getRuns().forEach((run) => {
            let text = run.text;
            Object.keys(data).forEach((key) => {
              const placeholder = `{{${key}}}`;
              if (text.includes(placeholder)) {
                text = text.replace(placeholder, data[key]);
              }
            });
            run.text = text;
          });
        });

        // Generate the new document and download it
        const buffer = await docx.Packer.toBuffer(doc);
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        const link = document.createElement('a');
        link.href = window.URL.createObjectURL(blob);
        link.download = 'populated_document.docx';
        link.click();
      } catch (error) {
        console.error("Error generating document:", error);
      }
    }
  }

  customElements.define('com-sap-sac-jm', Main);

})();
