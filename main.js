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
      this.Response = null;
      this._postData = {};

      this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
        this.generateAndDownloadDocument();
      });

      // Load the libraries in the correct order
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
        .then(() => {
          console.log("FileSaver.js library loaded successfully!");
          return this.loadScript('https://cdn.jsdelivr.net/npm/pizzip@3.1.1/dist/pizzip.min.js');
        })
        .then(() => {
          console.log("PizZip library loaded successfully!");
          return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.21.2/docxtemplater.min.js');
        })
        .then(() => {
          console.log("docxtemplater library loaded successfully!");
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

    generateAndDownloadDocument() {
      const data = this._postData;
      if (!data || Object.keys(data).length === 0) {
        alert("No data to generate document");
        return;
      }

      // Fetch the Word template, populate it, and trigger the download
      fetchWordTemplate()
        .then(templateBlob => populateWordTemplate(templateBlob, data))
        .then(populatedDocument => {
          saveAs(populatedDocument, 'populated_document.docx'); // Use FileSaver.js to save locally
        })
        .catch(error => console.error("Error generating document:", error));
    }
  }

  customElements.define('com-sap-sac-jm', Main);

  // Fetch the Word Template from your GitHub Repo using a CORS Proxy
  async function fetchWordTemplate() {
    try {
      const proxyUrl = 'https://corsproxy.io/?';
      const templateUrl = 'https://github.com/jmukajj/widget/raw/refs/heads/main/template.docx';
      const response = await fetch(proxyUrl + templateUrl, {
        headers: {
          'Origin': 'https://itsvac-test.eu20.hcs.cloud.sap'
        }
      });
      if (!response.ok) {
        throw new Error(`Failed to fetch the Word template: ${response.statusText}`);
      }
  
      // Create a Blob from the response
      const blob = await response.blob();
      
      // Test by creating a link and downloading to ensure correctness
      const link = document.createElement("a");
      link.href = window.URL.createObjectURL(blob);
      link.download = "fetched_template.docx";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
  
      return blob;
    } catch (error) {
      console.error('Error fetching template:', error);
      throw error;
    }
  }
  

  // Populate the Word Template
  function populateWordTemplate(templateBlob, data) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function (event) {
        const arrayBuffer = event.target.result;
        
        let zip;
        try {
          zip = new PizZip(arrayBuffer);
        } catch (error) {
          console.error("Error reading the array buffer with PizZip:", error);
          return reject(error);
        }
    
        let doc;
        try {
          doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
          });
        } catch (error) {
          console.error("Error initializing docxtemplater:", error);
          if (error.properties && error.properties.errors) {
            console.error("Detailed errors:", error.properties.errors);
          }
          return reject(error);
        }
    
        // Ensure data is not undefined or null
        if (!data || typeof data !== 'object') {
          console.error("Data for template is invalid:", data);
          return reject(new Error("Data for template is invalid."));
        }
    
        // Sanitize data to avoid undefined or null values
        const sanitizedData = {
          AccountDescription: data.AccountDescription || '',
          Antrag: data.Antrag || ''
        };
    
        // Set the template variables
        try {
          console.log("Data passed to docxtemplater:", sanitizedData);
          doc.setData(sanitizedData);
        } catch (error) {
          console.error("Error setting data for docxtemplater:", error);
          return reject(error);
        }
    
        // Render the document
        try {
          doc.render();
        } catch (error) {
          if (error.properties && error.properties.errors) {
            console.error("Errors in the template:", error.properties.errors);
          } else {
            console.error("Error rendering document:", error);
          }
          return reject(error);
        }
    
        // Generate the final output
        const out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
    
        resolve(out);
      };
      reader.readAsArrayBuffer(templateBlob);
    });
  }
  

})();
