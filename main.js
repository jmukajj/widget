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

      // Load the docx library dynamically from a reliable source
      this.loadScriptsInOrder([
        'https://unpkg.com/docx@7.0.0-beta.4/build/index.js'
      ]).then(() => {
        console.log("Docx library loaded successfully!");
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

    // Function to generate a Word document using docx library
    generateWordDocument() {
      console.log('Generating document with Post Data:', this._postData);

      if (!this._postData) {
        alert("No data to generate document");
        return;
      }

      const data = this._postData;

      // Check if docx is available
      if (typeof docx === 'undefined') {
        alert('docx library is not loaded. Please check if the library has loaded correctly.');
        return;
      }

      // Create a new docx document
      const { Document, Packer, Paragraph, TextRun } = docx;

      const doc = new Document({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({
                children: [
                  new TextRun("Antrag Document"),
                  new TextRun({
                    text: "\n------------------------------",
                    break: 1,
                  }),
                  new TextRun(`Antrag ID: ${data.AntragID}`),
                  new TextRun({
                    text: `\nDescription: ${data.Description}`,
                    break: 1,
                  }),
                  new TextRun({
                    text: `\nTotal Amount: ${data.TotalAmount}`,
                    break: 1,
                  }),
                ],
              }),
            ],
          },
        ],
      });

      // Generate the document and trigger the download
      Packer.toBlob(doc).then((blob) => {
        console.log("Document generated successfully!");
        saveAs(blob, "AntragDocument.docx"); // Download the generated document
      }).catch((error) => {
        console.error("Error generating document: ", error);
      });
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
