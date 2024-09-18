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

      // Hardcoded random data for testing without specific fields
      const randomData = {
        AntragID: "12345",  // Random Antrag ID
        Description: "Random Antrag Description",
        TotalAmount: "5000 USD" // Random amount
      };

      // Automatically send hardcoded data as postData to simulate table selection
      console.log("Sending Post Data: ", randomData); // Add log to verify postData
      this.sendPostData(randomData);

      // Attach event listener for download link
      this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
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

    // Send post data that includes random Antrag information
    sendPostData (postData) {
      this._postData = postData; // postData will now contain random Antrag info
      console.log("Post Data set: ", this._postData); // Add log to confirm postData
      this.render(); // Trigger the rendering of the widget
    }

    // Core rendering function handling both GET and POST requests
    async render () {
      // Here you would make your request to the server (if needed)
      console.log("Rendering with postData: ", this._postData);
    }

    // Function to generate a Word document
    generateWordDocument() {
      console.log('Generating document with Post Data:', this._postData);

      // Check if JSZip is available
      if (typeof JSZip === 'undefined') {
        alert('JSZip is not loaded. Please check if the library has loaded correctly.');
        return;
      }

      if (!this._postData) {
        alert("No data to generate document");
        return;
      }

      const data = this._postData;

      // Template for the Word document
      const content = `
        Antrag Document
        ------------------------------
        Antrag ID: ${data.AntragID}
        Description: ${data.Description}
        Total Amount: ${data.TotalAmount}
      `;

      // Use JSZip to create a Word file
      const zip = new JSZip();

      // Create text file inside the Word document
      zip.file("AntragDocument.txt", content);

      // Generate the Word document as a blob
      zip.generateAsync({ type: "blob" })
        .then(function (blob) {
          console.log("Document generated successfully!");
          saveAs(blob, "AntragDocument.docx"); // Download the generated document
        })
        .catch(function (error) {
          console.error("Error generating document: ", error);
        });
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
