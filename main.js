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

      // Initialize empty postData for selected Antrag
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
    sendPostData(selectedAntrag) {
      console.log("Received selected Antrag: ", selectedAntrag);

      // Ensure selectedAntrag contains the expected properties
      if (!selectedAntrag || !selectedAntrag.Konto || !selectedAntrag.Antrag || !selectedAntrag.Wert) {
        console.log("Selected row data:", selectedAntrag); 
        console.error("Missing necessary Antrag data", selectedAntrag);

        // Hardcoding values for testing
        this._postData = {
          Konto: "1234",
          Antrag: "Antrag A",
          Wert: "5000"
        };
        console.log("Using hardcoded test data since no proper data was passed: ", this._postData);
        return;
      }

      this._postData = {
        Konto: selectedAntrag.Konto,
        Antrag: selectedAntrag.Antrag,
        Wert: selectedAntrag.Wert
      };

      console.log("Selected Antrag Data after population: ", this._postData);
    }

    // Function to generate a Word document using Blob
    generateWordDocument() {
      console.log('Generating document with Post Data:', this._postData);

      if (!this._postData || !this._postData.Konto || !this._postData.Antrag || !this._postData.Wert) {
        alert("No data or incomplete data to generate document");
        return;
      }

      const data = this._postData;

      // Create the content of the Word document
      const content = `
      Antrag Document
      ------------------------------
      Konto: ${data.Konto}
      Antrag: ${data.Antrag}
      Wert: ${data.Wert}
      `;

      console.log("Document content:", content);

      // Create a Blob from the content
      const blob = new Blob([content], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

      // Use FileSaver.js to trigger download
      saveAs(blob, "AntragDocument.docx");
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
