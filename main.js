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

      // Attach event listener for download link
      this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
        event.preventDefault(); // Prevent the default action
        this.generateWordDocument();
      });

      // Load external scripts in sequence
      this.loadScriptsInOrder([
        'https://cdnjs.cloudflare.com/ajax/libs/pizzip/3.1.1/pizzip.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.21.0/docxtemplater.js',
        'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.7.1/jszip.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js'
      ]).then(() => {
        console.log("All libraries loaded successfully!");
      }).catch((error) => {
        console.error("Error loading scripts:", error);
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

    // Function to load scripts in order
    loadScriptsInOrder(scripts) {
      return scripts.reduce((promise, script) => {
        return promise.then(() => this.loadScript(script));
      }, Promise.resolve());
    }

    // Setter for the link
    setLink (link) {
      this._link = link;
    }

    // Setter for the SAP Server
    setServerSAP (ServerSAP) {
      this._ServerSAP = ServerSAP;
    }

    // Setter for the OData service
    setODataServiceSAP (ODataService) {
      this._ODataService = ODataService;
    }

    // Send post data that includes selected Antrag information
    sendPostData (postData) {
      this._postData = postData; // postData will now contain "Antrag" info
      this.render(); // Trigger the rendering of the widget
    }

    // Send GET request to fetch CSRF token and post request
    sendGet () {
      this.render();
    }

    // Get the response after the request
    getResponse () {
      return this.Response;
    }

    // Get the link
    getLink () {
      return this._link;
    }

    // Core rendering function handling both GET and POST requests
    async render () {
      const url = `https://${this._ServerSAP}/${this._ODataService}`;
      
      // GET request to fetch CSRF token
      var xhrGet = new XMLHttpRequest();
      xhrGet.open('GET', url, true);
      xhrGet.setRequestHeader('X-CSRF-Token', 'Fetch');
      xhrGet.setRequestHeader('Access-Control-Allow-Methods', 'GET');
      xhrGet.setRequestHeader('Access-Control-Allow-Origin', 'https://itsvac-test.eu20.hcs.cloud.sap');
      xhrGet.setRequestHeader('Access-Control-Allow-Credentials', true);
      xhrGet.setRequestHeader('Access-Control-Expose-Headers','X-Csrf-Token,x-csrf-token');
      xhrGet.setRequestHeader('Content-Type', 'application/json');
      xhrGet.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
      xhrGet.withCredentials = true;
      xhrGet.send();
      
      // Wait for the response
      xhrGet.onreadystatechange = () => {
        if (xhrGet.readyState === 4) {
          
          // Parse CSRF Token
          this.Response = JSON.parse(xhrGet.responseText);
          
          if (this._postData) {
            const data = this._postData; // Data containing Antrag info
            const __XCsrfToken = xhrGet.getResponseHeader('x-csrf-token');

            // POST request with selected Antrag data
            var xhr = new XMLHttpRequest();
            xhr.open('POST', url, true);
            xhr.setRequestHeader('Content-type', 'application/json');
            xhr.setRequestHeader('Access-Control-Allow-Credentials', true);
            xhr.setRequestHeader('Cache-Control', 'no-cache');
            xhr.setRequestHeader("X-Referrer-Hash", window.location.hash);
            xhr.setRequestHeader('Access-Control-Allow-Origin', 'https://itsvac-test.eu20.hcs.cloud.sap');
            xhr.setRequestHeader('Access-Control-Allow-Methods', 'POST');
            xhr.setRequestHeader('X-CSRF-Token', __XCsrfToken);
            xhr.withCredentials = true;

            // Convert Antrag data to JSON
            xhr.send(JSON.stringify(data));

            // Capture the response after posting
            xhr.onreadystatechange = () => {
              if (xhr.readyState == 4) {
                if (xhr.status == 201) {
                  this.Response = JSON.parse(xhr.responseText);
                }
              }
            }
          }
        }
      }
      
      // Data binding validation for SAC (if needed)
      const dataBinding = this.dataBinding;
      if (!dataBinding || dataBinding.state !== 'success') {
        return;
      }
    }

    // Function to generate a Word document
    generateWordDocument() {
      if (!this._postData) {
        alert("No data to generate document");
        return;
      }

      const data = this._postData;

      // Template for the Word document
      const content = `
        Antrag Document
        ------------------------------
        Created by: ${data.CreatedBy}
        Created on: ${data.CreatedOn}
        Total Amount: ${data.TotalAmount}
      `;

      // Use JSZip to create a Word file
      const zip = new JSZip();
      const doc = new window.docxtemplater();
      
      // Load the template and replace the content with the provided data
      zip.file("AntragDocument.txt", content);
      
      // Generate the Word document as a blob
      zip.generateAsync({ type: "blob" })
        .then(function (blob) {
          saveAs(blob, "AntragDocument.docx"); // Download the generated document
        });
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
